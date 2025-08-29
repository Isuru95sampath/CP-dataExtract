# streamlit_app.py
# -------------------------------------------------------------
# Streamlit app to extract fields from purchase-order PDFs,
# match product references with an existing Excel file, and add department info.
# Excel file path: C:\Users\Pcadmin\Desktop\New folder\CPEXCEL.xlsx
# -------------------------------------------------------------

import io
import re
import os
import pandas as pd
import streamlit as st
import pdfplumber
from difflib import SequenceMatcher
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

# Default Excel file path
DEFAULT_EXCEL_PATH = r"C:\Users\Pcadmin\Desktop\New folder\CPEXCEL.xlsx"

# ---------------------- Helpers ----------------------

def read_pdf_text(file_bytes: bytes) -> list[str]:
    texts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            text = re.sub(r"\u00A0", " ", text)
            texts.append(text)
    return texts

def full_text(pages: list[str]) -> str:
    return "\n".join(pages)

def load_product_reference_data(file_path: str) -> dict:
    try:
        if not os.path.exists(file_path):
            st.error(f"Excel file not found: {file_path}")
            return {}
        
        df = pd.read_excel(file_path, engine='openpyxl')
        
        product_ref_col = None
        department_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in ['product', 'ref', 'reference', 'item']):
                if product_ref_col is None:
                    product_ref_col = col
            if any(term in col_lower for term in ['department', 'dept', 'division', 'section']):
                if department_col is None:
                    department_col = col
        
        if product_ref_col is None:
            product_ref_col = df.columns[0]
        if department_col is None:
            department_col = df.columns[1]
        
        mapping = {}
        for idx, row in df.iterrows():
            prod_ref = row[product_ref_col]
            dept = row[department_col]
            if pd.notna(prod_ref) and pd.notna(dept):
                mapping[str(prod_ref).strip()] = str(dept).strip()
        return mapping
        
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return {}

def normalize_ref(s: str) -> str:
    s = str(s).upper().strip()
    s = re.sub(r'[^A-Z0-9]', '', s)
    s = s.replace('O', '0')
    return s

def match_product_reference(extracted_refs: str | list[str], reference_mapping: dict) -> tuple[str | None, str | None]:
    if not extracted_refs or not reference_mapping:
        if isinstance(extracted_refs, str):
            return extracted_refs, None
        elif isinstance(extracted_refs, list) and extracted_refs:
            return extracted_refs[0], None
        else:
            return None, None

    if isinstance(extracted_refs, str):
        refs_list = [r.strip() for r in extracted_refs.split(',') if r.strip()]
    else:
        refs_list = [str(r).strip() for r in extracted_refs if str(r).strip()]

    clean_map = {}
    for canonical_ref, dept in reference_mapping.items():
        ck = normalize_ref(canonical_ref)
        if ck:
            clean_map[ck] = (canonical_ref, dept)
    keys_clean = list(clean_map.keys())

    def best_match(one_ref: str):
        ref_clean = normalize_ref(one_ref)
        if not ref_clean:
            return None
        if ref_clean in clean_map:
            return clean_map[ref_clean]
        nums = re.findall(r'\d{4,}', ref_clean)
        if nums:
            for num in nums:
                for ck in keys_clean:
                    if num in ck:
                        return clean_map[ck]
        for ck in keys_clean:
            if (len(ref_clean) >= 5 and ref_clean in ck) or (len(ck) >= 5 and ck in ref_clean):
                return clean_map[ck]
        from difflib import SequenceMatcher
        best_ratio = 0.0
        best_ck = None
        for ck in keys_clean:
            r = SequenceMatcher(None, ref_clean, ck).ratio()
            if r > best_ratio:
                best_ratio = r
                best_ck = ck
        if best_ck and best_ratio >= 0.82:
            return clean_map[best_ck]
        return None

    for ref in refs_list:
        bm = best_match(ref)
        if bm:
            return bm
    return (refs_list[0] if refs_list else None), None

# ------------------ Field Extractors ------------------

PO_NUM_PATTERNS = [
    r"\bPO\s*Number\s*[-:]*\s*(\d+)\b",
    r"\bPO\s*Number\s*\n\s*(\d+)\b",
]

TOTAL_VALUE_PATTERNS = [
    r"\bTotal\s*Value\s*[:\-]?\s*(?:USD\s*)?([0-9]+(?:\.[0-9]+)?)\b",
    r"TOTAL\s+NET\s+VALUE[\s\S]{0,40}?USD\s*([0-9]+(?:\.[0-9]+)?)\b",
]

NOTIFY_BLOCK_RE = re.compile(
    r"Notify\s*Party/Deliver\s*To\s*:([\s\S]{0,300})",
    re.IGNORECASE,
)

PRICE_PER_UNIT_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*/\s*([0-9][0-9,]*)\b"
)


def extract_po_number(text: str) -> str | None:
    for pat in PO_NUM_PATTERNS:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None

def extract_total_value(text: str) -> float | None:
    for pat in TOTAL_VALUE_PATTERNS:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                continue
    return None

def extract_customer(text: str) -> str | None:
    m = NOTIFY_BLOCK_RE.search(text)
    if not m:
        return None
    lines = [line.strip() for line in m.group(1).splitlines() if line.strip()]
    if len(lines) >= 2:
        customer = lines[1].replace("INTERNATIONAL TRIMMINGS &", "").strip()
        return customer
    return None

def extract_price_per_unit_4th(text: str) -> tuple[float, int] | None:
    matches = list(PRICE_PER_UNIT_RE.finditer(text))
    if not matches:
        return None
    idx = 3 if len(matches) >= 4 else len(matches) - 1
    m = matches[idx]
    price = float(m.group(1))
    units = int(m.group(2).replace(",", ""))
    return price, units

def extract_product_references(text: str) -> list[str]:
    refs = []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "ITEM DESCRIPTION" in line.upper() and "SUPPLIER REF" in line.upper():
            for j in range(i+1, min(i+20, len(lines))):
                l = lines[j].strip()
                if not l:
                    continue
                if "   " in l:
                    break
                if re.search(r"LINE#|UNIT|SURCHARGE|DATE|SHIP MODE", l, re.IGNORECASE):
                    continue
                if re.search(r"EX-FACT|TOTAL VALUE|PRICE/PER UNIT", l, re.IGNORECASE):
                    break
                if re.search(r"[A-Za-z0-9]", l):
                    match = re.search(r"^\d+\s+([^/]+)", l)
                    if match:
                        product_ref = match.group(1).strip()
                        if product_ref:
                            refs.append(product_ref)
    return refs

def compute_quantity(total_value: float | None, price_units: tuple[float, int] | None) -> int | None:
    if total_value is None or price_units is None:
        return None
    price, units = price_units
    if price == 0:
        return None
    qty = (total_value / price) * units
    return int(round(qty))

def extract_ex_fact_date_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """
    Extract EX-FACT DATE from PDF bytes using your specific logic:
    Second table, row 5, column 5 (4 row before last column)
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                try:
                    tables = page.extract_tables()
                    if len(tables) >= 2:  # second table exists
                        table = tables[1]   # second table (index 1)
                        if len(table) > 4 and len(table[4]) > 4:  # at least 5 rows & 5 columns
                            date_val = table[4][4]  # 5th row, 5th col (index 4,4)
                            if date_val and date_val.strip():
                                return date_val.strip()
                except Exception as e:
                    continue
                    
            # If table extraction fails, try text extraction
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    return extract_ex_fact_date_from_text(text)
                    
    except Exception as e:
        return f"Error processing PDF: {str(e)}"
    
    return "date not found"

def extract_ex_fact_date_from_text(document_text: str) -> str:
    """
    Extract EX-FACT DATE from text using the specified logic:
    Read from second long square (table), 4 row before last column (EX-FACT DATE column)
    """
    import re
    
    # Split document into lines to find table structure
    lines = document_text.split('\n')
    
    # Find the start of the main item table (after the header table)
    table_start = -1
    for i, line in enumerate(lines):
        if 'SALES ORDER/ LINE ITEM CODE' in line and 'EX-FACT' in line:
            table_start = i
            break
    
    if table_start == -1:
        # Fallback: look for any line with size codes and dates
        for line in lines:
            if any(size in line for size in ['L-R', 'M-R', 'S-R', 'XL-R', 'XS-R']) and 'PC' in line and 'Truck' in line:
                date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                if date_match:
                    return date_match.group(1)
        return "table not found"
    
    # Look for data lines that contain size information and dates
    # These represent the "second long square" (main items table)
    data_lines = []
    for i in range(table_start + 1, len(lines)):
        line = lines[i].strip()
        if line and any(size in line for size in ['L-R', 'M-R', 'S-R', 'XL-R', 'XS-R', 'XXL-R', 'L-S', 'M-S', 'S-S', 'XL-S', 'XS-S', 'M-L', 'S-L', 'XL-L', 'XS-L']):
            # Check if line contains date pattern (DD.MM.YYYY)
            if re.search(r'\d{2}\.\d{2}\.\d{4}', line):
                data_lines.append(line)
    
    if not data_lines:
        return "no data lines found"
    
    # Apply your logic: take the 4th row (index 3) from the data lines
    # This represents "4 row before last column have date"
    if len(data_lines) >= 4:
        target_line = data_lines[3]  # 4th row (0-indexed)
        
        # Extract date from this line (EX-FACT DATE column)
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', target_line)
        if date_match:
            return date_match.group(1)
    
    # Fallback: if less than 4 rows, take the first available date
    if data_lines:
        first_line = data_lines[0]
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', first_line)
        if date_match:
            return date_match.group(1)
    
    return "date not found"

# ---------------------- UI ----------------------

st.set_page_config(page_title="PDF → Excel Extractor", layout="wide")
st.title("📄 PDF → Excel Extractor (PO fields + Department Matching)")

st.write("Upload PDF files to extract purchase order data and match product references with department database.")

reference_mapping = {}
if 'reference_mapping' not in st.session_state:
    with st.spinner("Loading product reference database..."):
        reference_mapping = load_product_reference_data(DEFAULT_EXCEL_PATH)
        st.session_state.reference_mapping = reference_mapping
else:
    reference_mapping = st.session_state.reference_mapping

st.subheader("📁 PDF Upload")
uploaded = st.file_uploader("Upload PDF(s)", type=["pdf"], accept_multiple_files=True)

if uploaded:
    rows = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, up in enumerate(uploaded):
        status_text.text(f"Processing {up.name}...")
        progress_bar.progress(int(((idx) / len(uploaded)) * 100))

        try:
            raw = up.read()
            pages = read_pdf_text(raw)
            text = full_text(pages)

            po_number = extract_po_number(text)
            customer = extract_customer(text)
            # Fixed: Now using PDF bytes for table extraction and text as fallback
            ex_fact_date = extract_ex_fact_date_from_pdf_bytes(raw)
            total_value = extract_total_value(text)
            price_units = extract_price_per_unit_4th(text)

            refs_list = extract_product_references(text)
            original_refs = ", ".join(refs_list) if refs_list else None
            matched_ref, department = match_product_reference(refs_list or None, reference_mapping)
            product_ref_display = matched_ref if department else (refs_list[0] if refs_list else None)
            quantity = compute_quantity(total_value, price_units)

            rows.append(
                {
                    "EX-FACT DATE": ex_fact_date,
                    "Customer": customer,
                    "PO Number": po_number,
                    "Department": department,
                    "Product Reference": product_ref_display,
                    "Total Value": total_value,
                    "Quantity (calculated)": quantity
                }
            )
        except Exception as e:
            st.error(f"Error processing {up.name}: {str(e)}")
            continue

    progress_bar.progress(100)
    status_text.text("Processing complete!")
    
    if rows:
        df = pd.DataFrame(rows)
        display_columns = ["EX-FACT DATE",  "Customer", "PO Number", 
                           "Product Reference", "Department",  "Total Value", "Quantity (calculated)"]
        st.subheader("📋 Extracted Data")
        st.dataframe(df[display_columns], use_container_width=True)

        # Show some statistics
        st.subheader("📊 Processing Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Files Processed", len(rows))
        with col2:
            dept_found = df['Department'].notna().sum()
            st.metric("Departments Found", dept_found)
        with col3:
            dates_found = df[df['EX-FACT DATE'] != 'date not found']['EX-FACT DATE'].count()
            st.metric("Dates Extracted", dates_found)
        with col4:
            total_val = df['Total Value'].sum() if df['Total Value'].notna().any() else 0
            st.metric("Total Value", f"${total_val:,.2f}")

        st.subheader("💾 Download Results")
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Extracted_Data")
        out.seek(0)

        st.download_button(
            label="⬇️ Download Excel with Department Info",
            data=out.getvalue(),
            file_name=f"extracted_with_departments_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("No data could be extracted from the uploaded files.")
else:
    st.info("👆 Upload one or more PDF files to begin extraction.")