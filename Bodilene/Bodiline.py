import io
import re
import os
import pandas as pd
import pdfplumber
from difflib import SequenceMatcher
from datetime import datetime
import streamlit as st

# ---------------------- Common Helpers ----------------------

def convert_date_format(date_str):
    """Convert date from dd-mm-yy to mm-dd-yy format"""
    if not date_str or date_str == "date not found" or date_str.startswith("Error"):
        return date_str
    try:
        # Handle different separators (., /, -)
        normalized = date_str.replace('.', '/').replace('-', '/')
        if len(normalized.split('/')[2]) == 2:  # dd/mm/yy format
            date_obj = datetime.strptime(normalized, "%d/%m/%y")
        else:  # dd/mm/yyyy format
            date_obj = datetime.strptime(normalized, "%d/%m/%Y")
        return date_obj.strftime("%m-%d-%y")
    except ValueError:
        return date_str

def format_date_for_display(date_str):
    """Format date for display as dd/mm/yy"""
    if not date_str or date_str == "date not found" or date_str.startswith("Error"):
        return date_str
    try:
        # Handle different separators (., /, -)
        normalized = date_str.replace('.', '/').replace('-', '/')
        if len(normalized.split('/')[2]) == 2:  # yy format
            date_obj = datetime.strptime(normalized, "%m/%d/%y")
        else:  # yyyy format
            date_obj = datetime.strptime(normalized, "%m/%d/%Y")
        return date_obj.strftime("%d/%m/%y")
    except ValueError:
        try:
            # Try parsing as dd/mm/yy or dd/mm/yyyy
            if len(normalized.split('/')[2]) == 2:  # yy format
                date_obj = datetime.strptime(normalized, "%d/%m/%y")
            else:  # yyyy format
                date_obj = datetime.strptime(normalized, "%d/%m/%Y")
            return date_obj.strftime("%d/%m/%y")
        except ValueError:
            return date_str

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

def match_product_reference(extracted_refs, reference_mapping: dict):
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
        
        # Direct match first
        if ref_clean in clean_map:
            return clean_map[ref_clean]
        
        # Special handling for references like "HGT LB 06731-C", "LBL LB 05735-C", "HSL LB 02691-C"
        special_pattern = re.match(r'^([A-Z]{3})\s*LB\s*0*(\d+)(.*)$', one_ref.upper().strip())
        if special_pattern:
            prefix = special_pattern.group(1)
            number = special_pattern.group(2)
            suffix = special_pattern.group(3)
            
            # First try: Search with leading zero
            padded_number = number.zfill(5)
            search_pattern_with_zero = f"0{number}"
            
            for ck in keys_clean:
                if padded_number in ck or search_pattern_with_zero in ck:
                    return clean_map[ck]
            
            # Second try: Search without leading zero
            search_without_zero = normalize_ref(f"LB {number}{suffix}")
            if search_without_zero in clean_map:
                return clean_map[search_without_zero]
            
            # Third try: Partial matching with the number only
            for ck in keys_clean:
                if number in ck and 'LB' in ck:
                    return clean_map[ck]
        
        # Try removing leading zeros in numeric parts
        tokens = re.split(r'(\d+)', ref_clean)
        new_tokens = []
        for token in tokens:
            if token.isdigit():
                new_token = token.lstrip('0')
                if not new_token:
                    new_token = '0'
                new_tokens.append(new_token)
            else:
                new_tokens.append(token)
        variation = ''.join(new_tokens)
        if variation in clean_map:
            return clean_map[variation]
        
        # Continue with existing strategies
        nums = re.findall(r'\d{4,}', ref_clean)
        if nums:
            for num in nums:
                for ck in keys_clean:
                    if num in ck:
                        return clean_map[ck]
        for ck in keys_clean:
            if (len(ref_clean) >= 5 and ref_clean in ck) or (len(ck) >= 5 and ck in ref_clean):
                return clean_map[ck]
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

# ---------------------- Unichela Extractors ----------------------

PO_NUM_PATTERNS = [
    r"\bPO\s*Number\s*[-:]*\s*(\d+)\b",
    r"\bPO\s*Number\s*\n\s*(\d+)\b",
]

TOTAL_VALUE_PATTERNS = [
    r"\bTotal\s*Value\s*[:\-]?\s*(?:USD\s*)?([0-9][0-9,]*(?:\.[0-9]+)?)\b",
    r"TOTAL\s+NET\s+VALUE[\s\S]{0,40}?(?:USD\s*)?([0-9][0-9,]*(?:\.[0-9]+)?)\b",
]

NOTIFY_BLOCK_RE = re.compile(
    r"Notify\s*Party/Deliver\s*To\s*:([\s\S]{0,300})",
    re.IGNORECASE,
)

def extract_unichela_po_number(text: str) -> str | None:
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
                val = m.group(1)
                val = val.replace(",", "").strip()
                return float(val)
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

def extract_price_per_unit(text: str) -> tuple[float, int] | None:
    """Extract price per unit value from patterns like '15.00  1,000  ' with robust handling"""
    
    # Clean up the text first
    cleaned_text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
    
    # Pattern 1: Price followed by spaces and units with comma (most common)
    # Matches: "15.00 1,000", "15.00  1,000", "15.00   1,000"
    pattern1 = r'([0-9]+(?:\.[0-9]{2})?)\s+([0-9]{1,3}(?:,[0-9]{3})+)\b'
    
    # Pattern 2: Price followed by spaces and 4-digit units (without comma)
    # Matches: "15.00 1000", "15.00  1000"
    pattern2 = r'([0-9]+(?:\.[0-9]{2})?)\s+([0-9]{4})\b'
    
    # Pattern 3: Price with slash and units
    # Matches: "15.00/1000", "15.00 / 1000"
    pattern3 = r'([0-9]+(?:\.[0-9]{2})?)\s*/\s*([0-9]{1,3}(?:,[0-9]{3})+)\b'
    
    # Pattern 4: Specific to PRICE/PER UNIT column context
    # Look for patterns in table-like structures
    lines = text.split('\n')
    for line in lines:
        line_stripped = line.strip()
        
        # Skip if line doesn't contain numbers
        if not re.search(r'[0-9]', line_stripped):
            continue
            
        # Try to find price/unit pattern in this line
        # Pattern for table data: price at specific position, units at another
        # Example: "32D 360.000 6.60 1,000 0.00 15.09.2025 Truck"
        
        # Split by spaces and look for price followed by 1000
        tokens = line_stripped.split()
        
        for i in range(len(tokens) - 1):
            current_token = tokens[i]
            next_token = tokens[i + 1] if i + 1 < len(tokens) else ""
            
            # Check if current token is a price (with decimal)
            if re.match(r'^[0-9]+\.[0-9]{2}$', current_token):
                # Check if next token is 1000 (with or without comma)
                if re.match(r'^1,?000$', next_token):
                    try:
                        price = float(current_token)
                        units = int(next_token.replace(',', ''))
                        return price, units
                    except ValueError:
                        continue
            
            # Check if current token is price and next token starts with 1,000
            if re.match(r'^[0-9]+\.[0-9]{2}$', current_token) and next_token.startswith('1,000'):
                try:
                    price = float(current_token)
                    units = int(next_token.replace(',', ''))
                    return price, units
                except ValueError:
                    continue
    
    # Try regex patterns on cleaned text
    for pattern in [pattern1, pattern2, pattern3]:
        matches = re.finditer(pattern, cleaned_text)
        for m in matches:
            try:
                price_str = m.group(1).replace(",", "").strip()
                units_str = m.group(2).replace(",", "").strip()
                price = float(price_str)
                units = int(units_str)
                
                # Check if units is 1000
                if units == 1000:
                    return price, units
            except ValueError:
                continue
    
    # Special handling: Look for "PRICE/PER UNIT" section
    price_unit_match = re.search(r'PRICE/PER\s+UNIT.*?([0-9]+\.[0-9]{2})\s+([0-9]{1,3}(?:,[0-9]{3})+)', 
                                text, re.IGNORECASE | re.DOTALL)
    if price_unit_match:
        try:
            price_str = price_unit_match.group(1).replace(",", "").strip()
            units_str = price_unit_match.group(2).replace(",", "").strip()
            price = float(price_str)
            units = int(units_str)
            
            if units == 1000:
                return price, units
        except ValueError:
            pass
    
    # Last resort: Look for any occurrence of 1000 with a nearby price
    # Find all 1000 occurrences and check nearby tokens
    thousand_matches = list(re.finditer(r'1,?000', text))
    for match in thousand_matches:
        # Get context around the 1000 (50 characters before and after)
        start = max(0, match.start() - 50)
        end = min(len(text), match.end() + 50)
        context = text[start:end]
        
        # Look for price pattern in this context
        price_match = re.search(r'([0-9]+\.[0-9]{2})', context)
        if price_match:
            try:
                price = float(price_match.group(1))
                units = int(match.group(0).replace(',', ''))
                return price, units
            except ValueError:
                continue
    
    return None

def extract_product_references(text: str) -> list[str]:
    """Extract product reference from lines starting with 'LB'"""
    refs = []
    lines = text.splitlines()
    
    # Search for lines that start with "LB"
    for line in lines:
        line_stripped = line.strip()
        # Check if line starts with "LB" (case-insensitive)
        if line_stripped.upper().startswith("LB"):
            # Extract the entire line as the product reference
            product_ref = line_stripped.strip()
            if product_ref:
                refs.append(product_ref)
                return refs  # Return the first match found
    
    # Fallback: Look for any line containing "LB" at the beginning of a word
    for line in lines:
        if re.match(r'^\s*LB\s+', line.upper()):
            product_ref = line.strip()
            if product_ref:
                refs.append(product_ref)
                return refs
    
    # Additional fallback: Look for lines with "LB" pattern anywhere
    for line in lines:
        if "LB" in line.upper() and len(line.strip()) > 5:  # Ensure it's not just "LB"
            product_ref = line.strip()
            if product_ref:
                refs.append(product_ref)
                return refs
    
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
    """Extract EX-FACT DATE from PDF bytes in MM-DD-YY format"""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            # First try table extraction
            for page in pdf.pages:
                try:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        
                        # Find header row containing "EX-FACT"
                        header_row_idx = None
                        ex_fact_col_idx = None
                        
                        for i, row in enumerate(table):
                            if row and any(cell and "EX-FACT" in str(cell).upper() for cell in row):
                                header_row_idx = i
                                # Find the exact column index for EX-FACT DATE
                                for j, cell in enumerate(row):
                                    if cell and "EX-FACT" in str(cell).upper():
                                        ex_fact_col_idx = j
                                        break
                                break
                        
                        if header_row_idx is not None and ex_fact_col_idx is not None:
                            # Look through rows below header for DD/MM/YYYY format dates
                            for row_idx in range(header_row_idx + 1, len(table)):
                                if row_idx < len(table) and ex_fact_col_idx < len(table[row_idx]):
                                    cell_value = table[row_idx][ex_fact_col_idx]
                                    if cell_value and isinstance(cell_value, str):
                                        # Check for DD/MM/YYYY format (like 04.08.2025 or 04/08/2025)
                                        date_match = re.search(r'\b(\d{2}[./]\d{2}[./]\d{4})\b', cell_value.strip())
                                        if date_match:
                                            # Convert to MM-DD-YY format
                                            return convert_date_format(date_match.group(1))
                except Exception:
                    continue
            
            # Fallback to text extraction if table extraction fails
            for page in pdf.pages:
                try:
                    text = page.extract_text()
                    if text:
                        # Look for EX-FACT DATE pattern followed by DD/MM/YYYY
                        lines = text.split('\n')
                        po_date_found = None
                        
                        # First identify the PO date to avoid it
                        for line in lines[:10]:  # Check first 10 lines for PO date
                            if any(keyword in line.upper() for keyword in ['PO NUMBER', 'PURCHASE ORDER', 'DATE']):
                                po_date_match = re.search(r'\b(\d{2}[./]\d{2}[./]\d{4})\b', line)
                                if po_date_match:
                                    po_date_found = po_date_match.group(1)
                                    break
                        
                        # Now look for EX-FACT DATE, avoiding the PO date
                        for i, line in enumerate(lines):
                            if 'EX-FACT' in line.upper() and 'DATE' in line.upper():
                                # Check this line and next few lines for date pattern
                                for j in range(i, min(i + 15, len(lines))):
                                    date_matches = re.findall(r'\b(\d{2}[./]\d{2}[./]\d{4})\b', lines[j])
                                    for date_match in date_matches:
                                        # Skip if this is the PO date
                                        if po_date_found and date_match == po_date_found:
                                            continue
                                        # Convert to MM-DD-YY format
                                        return convert_date_format(date_match)
                        
                        # As last resort, find dates that are NOT in the header section
                        # Skip first 15 lines to avoid header dates
                        lower_text = '\n'.join(lines[15:])
                        date_matches = re.findall(r'\b(\d{2}[./]\d{2}[./]\d{4})\b', lower_text)
                        if date_matches:
                            # Return the first date found in the body (not header)
                            # Convert to MM-DD-YY format
                            return convert_date_format(date_matches[0])
                        
                except Exception:
                    continue
                    
    except Exception as e:
        return f"Error processing PDF: {str(e)}"
    
    return "date not found"

# ---------------------- Main Processing Functions ----------------------

def process_unichela_pdf(uploaded_file, reference_mapping):
    """Process Unichela PDF and return extracted data"""
    try:
        raw = uploaded_file.read()
        pages = read_pdf_text(raw)
        text = full_text(pages)

        po_number = extract_unichela_po_number(text)
        customer = extract_customer(text)
        ex_fact_date = extract_ex_fact_date_from_pdf_bytes(raw)
        total_value = extract_total_value(text)
        price_units = extract_price_per_unit(text)

        # Capture price and units for table display
        ppu_price = price_units[0] if price_units else None
        ppu_units = price_units[1] if price_units else None

        refs_list = extract_product_references(text)
        matched_ref, department = match_product_reference(refs_list or None, reference_mapping)
        product_ref_display = matched_ref if department else (refs_list[0] if refs_list else None)
        quantity = compute_quantity(total_value, price_units)

        return {
            "EX-FACT DATE": ex_fact_date,
            "Customer": customer,
            "PO Number": po_number,
            "Department": department,
            "Product Reference": product_ref_display,
            "Total Value": total_value,
            "PRICE/PER UNIT Price": ppu_price,
            "PRICE/PER UNIT Units": ppu_units,
            "Quantity (calculated)": quantity
        }
    except Exception as e:
        st.error(f"Error processing Bodilyne PDF {uploaded_file.name}: {str(e)}")
        return None

# ---------------------- Streamlit UI ----------------------

st.set_page_config(
    page_title="Bodilyne Data Extractor",
    page_icon="🚀",
    layout="wide"
)

st.title("🚀 Bodilyne PDF Data Extractor")
st.markdown("""
**Advanced Bodilyne PDF Analysis Technology**

Upload Bodilyne PDF files to automatically extract data.
""")

# Load department mapping
DEFAULT_EXCEL_PATH = r"C:\Users\Pcadmin\Documents\ITL\CP-SO-Tracker\CPEXCEL.xlsx"
reference_mapping = {}
if 'reference_mapping' not in st.session_state:
    with st.spinner("Loading product reference database..."):
        reference_mapping = load_product_reference_data(DEFAULT_EXCEL_PATH)
        st.session_state.reference_mapping = reference_mapping
else:
    reference_mapping = st.session_state.reference_mapping

st.subheader("📁 Bodilyne PDF Upload")
uploaded_files = st.file_uploader("Upload Unichela PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    results = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Processing {uploaded_file.name}...")
        progress_bar.progress(int(((idx) / len(uploaded_files)) * 100))
        
        try:
            result = process_unichela_pdf(uploaded_file, reference_mapping)
            if result:
                results.append(result)
                    
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")
            continue
    
    progress_bar.progress(100)
    status_text.text("Processing complete!")
    
    # Display results
    if results:
        st.subheader("📋 Unichela Data")
        df = pd.DataFrame(results)
        
        # Format dates for display
        df_display = df.copy()
        if 'EX-FACT DATE' in df_display.columns:
            df_display['EX-FACT DATE'] = df_display['EX-FACT DATE'].apply(format_date_for_display)
        
        display_columns = [
            "EX-FACT DATE", "Customer", "PO Number",
            "Product Reference", "Department",
            "Total Value", "PRICE/PER UNIT Price", "PRICE/PER UNIT Units",
            "Quantity (calculated)"
        ]
        st.dataframe(df_display[display_columns], use_container_width=True)
        
        # Download button
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f'unichela_data_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.csv',
            mime='text/csv',
            use_container_width=True
        )
        
        # Show processing summary
        st.subheader("📊 Processing Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Files Processed", len(uploaded_files))
        with col2:
            st.metric("Successful Extractions", len(results))
    else:
        st.warning("No data could be extracted from the uploaded files.")
else:
    st.info("👆 Upload one or more Unichela PDF files to begin extraction.")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; padding: 20px; color: #666;">
    <p>
        🚀 <strong>Unichela PDF Data Extractor</strong> | 
        Powered by <strong>Razz....</strong> | 
        Advanced PDF Analysis Technology
    </p>
</div>
""", unsafe_allow_html=True)