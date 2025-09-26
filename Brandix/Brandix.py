import streamlit as st
import pdfplumber
import re
import pandas as pd
from datetime import datetime
import os
from difflib import SequenceMatcher

# -------------------- Department Filtering Functions --------------------

def normalize_ref(s: str) -> str:
    s = str(s).upper().strip()
    s = re.sub(r'[^A-Z0-9]', '', s)
    s = s.replace('O', '0')
    return s

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

def match_product_reference(extracted_refs, reference_mapping: dict):
    if not extracted_refs or not reference_mapping:
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
            return None, None
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
        best_ratio = 0.0
        best_ck = None
        for ck in keys_clean:
            r = SequenceMatcher(None, ref_clean, ck).ratio()
            if r > best_ratio:
                best_ratio = r
                best_ck = ck
        if best_ck and best_ratio >= 0.82:
            return clean_map[best_ck]
        return None, None

    for ref in refs_list:
        bm = best_match(ref)
        if bm and bm[0] is not None:
            return bm
    return None, None

# -------------------- Original Helpers --------------------

def convert_date_format(date_str):
    """Convert date from dd-mm-yy to mm-dd-yy format"""
    if not date_str:
        return ""
    try:
        date_obj = datetime.strptime(date_str, "%d-%m-%y")
        return date_obj.strftime("%m-%d-%y")
    except ValueError:
        return date_str

def extract_brandix_from_beginning(text):
    """Extract Brandix information from the beginning of the PO text"""
    try:
        lines = text.split('\n')
        brandix_info = ""
        for line in lines[:10]:
            if "Brandix" in line:
                brandix_info = line.strip()
                break
        if not brandix_info:
            m = re.search(r'Brandix[^\n]+', text, re.IGNORECASE)
            if m:
                brandix_info = m.group(0).strip()
        return brandix_info
    except Exception as e:
        st.error(f"Error extracting Brandix: {e}")
        return ""

def extract_po_number(text):
    """Extract PO number from different formats"""
    try:
        # First try: Look for "PO No." followed by 7 digits
        po_match = re.search(r'PO\s*No\.?\s*[:\-]?\s*(\d{7})', text, re.IGNORECASE)
        if po_match:
            return po_match.group(1)
            
        # Second try: Look for "BFF" followed by 7 digits
        bff_match = re.search(r'BFF\s+(\d{7})', text)
        if bff_match:
            return bff_match.group(1)
            
        # Third try: Look for "PO Number" followed by 7 digits
        po_match2 = re.search(r'PO\s*Number[^\d]*(\d{7})', text, re.IGNORECASE)
        if po_match2:
            return po_match2.group(1)
            
        # Fourth try: Look for any 7-digit number after "PO"
        lines = text.split('\n')
        for line in lines:
            if re.search(r'PO', line, re.IGNORECASE):
                digits = re.findall(r'\d{7}', line)
                if digits:
                    return digits[0]
        
        return ""
    except Exception as e:
        st.error(f"Error extracting PO Number: {e}")
        return ""

def extract_po_total_line_from_last_page(pdf):
    """Try to capture PO Total line from last page"""
    try:
        if not pdf.pages:
            return ""
        last_page = pdf.pages[-1]
        text = last_page.extract_text() or ""
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

        for line in lines:
            if re.search(r'PO\s*TOTAL\s*AMOUNT', line, re.IGNORECASE):
                return line
        # Fallback: "PO Total Amount" variant
        for line in lines:
            if re.search(r'PO\s*Total\s*Amount', line, re.IGNORECASE):
                return line
        return ""
    except Exception as e:
        st.error(f"Error extracting PO TOTAL AMOUNT line: {e}")
        return ""

def clean_po_total_line(line):
    if not line:
        return "", ""
    cleaned = re.sub(r'PO\s*TOTAL\s*AMOUNT\s*:?', '', line, flags=re.IGNORECASE)
    cleaned = cleaned.replace("USD", "").strip()
    numbers = re.findall(r'[\d,]+(?:\.\d+)?', cleaned)
    if not numbers:
        return cleaned.strip(), ""
    total = numbers[-1].replace(",", "")
    line_amounts = [n.replace(",", "") for n in numbers[:-1]]
    line_amount = ", ".join(line_amounts) if line_amounts else ""
    return total, line_amount

# -------------------- Updated Product Reference Extraction --------------------

def extract_product_reference_from_structured_line(text, reference_mapping):
    """Extract product reference only from the structured format line"""
    try:
        lines = text.split('\n')
        
        # Look for the header line: "No Item Quantity UM Price Price Unit Line Amount"
        for i, line in enumerate(lines):
            if re.search(r'No\s+Item\s+Quantity\s+UM\s+Price', line, re.IGNORECASE):
                # Look at the next few lines for the actual data
                for j in range(i + 1, min(i + 5, len(lines))):
                    data_line = lines[j].strip()
                    if not data_line:
                        continue
                    
                    # Split the line by spaces to process tokens
                    tokens = data_line.split()
                    
                    # Skip the first token (line number)
                    if len(tokens) < 2:
                        continue
                        
                    # Collect tokens until we hit a 5-6 digit number (quantity)
                    product_ref_tokens = []
                    for token in tokens[1:]:
                        # Check if token is a 5-6 digit number (quantity)
                        if re.match(r'^\d{5,6}$', token):
                            break
                        product_ref_tokens.append(token)
                    
                    # Join the tokens to form the product reference
                    if product_ref_tokens:
                        potential_ref = ' '.join(product_ref_tokens)
                        
                        # Try to match the entire string (with the number)
                        result = match_product_reference(potential_ref, reference_mapping)
                        if result and result[0] is not None:
                            matched_ref, department = result
                            return matched_ref, department
                        
                        # If that fails, try just the first token (product code)
                        if product_ref_tokens:
                            code_part = product_ref_tokens[0]
                            result = match_product_reference(code_part, reference_mapping)
                            if result and result[0] is not None:
                                matched_ref, department = result
                                return matched_ref, department
                        
                        # If neither matches, return the captured reference and None for department
                        return potential_ref, None
                break
        
        # If we didn't find the structured format, try to find any line with product reference pattern
        for line in lines:
            # Look for pattern like "LBL.MAIN_LB" or similar
            match = re.search(r'([A-Z]{3}\.[A-Z]{4,}(?:\s+\d+)?)', line)
            if match:
                potential_ref = match.group(1).strip()
                result = match_product_reference(potential_ref, reference_mapping)
                if result and result[0] is not None:
                    matched_ref, department = result
                    return matched_ref, department
                else:
                    return potential_ref, None
        
        # If still no match found, return None
        return None, None
    except Exception as e:
        st.error(f"Error extracting product reference: {e}")
        return None, None

# -------------------- Main extraction --------------------

def extract_product_code_and_xmill_date(pdf_file, reference_mapping):
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            if len(pdf.pages) == 0:
                return None, None, None, None, "", "", None

            first_page = pdf.pages[0]
            first_text = first_page.extract_text() or ""

            po_number = extract_po_number(first_text)
            brandix = extract_brandix_from_beginning(first_text)

            product_code = ""
            department = None
            
            # Extract product reference from the structured format only
            result = extract_product_reference_from_structured_line(first_text, reference_mapping)
            if result and result[0] is not None:
                product_code, department = result
            elif result and result[0] is None:
                # If product_code is None but result is not None, it means we have a product reference but no department
                product_code = result[1] if result[1] else ""
            else:
                # If result is None, try to extract product reference without matching
                product_code = ""
                department = None

            # X-Mill Date extraction
            x_mill_date = ""    
            date_match = re.search(r'X-Mill Date\(dd-mm-yy\)\s*[:\-]?\s*(\d{2}-\d{2}-\d{2})', first_text, re.IGNORECASE)
            if not date_match:
                date_match = re.search(r'XMill Date\s*[:\-]?\s*(\d{2}-\d{2}-\d{2})', first_text, re.IGNORECASE)
            if date_match:
                x_mill_date = convert_date_format(date_match.group(1))

            po_total_line = extract_po_total_line_from_last_page(pdf)
            po_total, line_amount = clean_po_total_line(po_total_line)

            return product_code, x_mill_date, brandix, po_number, po_total, line_amount, department
    except Exception as e:
        st.error(f"Error extracting data: {e}")
        return None, None, None, None, "", "", None

# -------------------- Streamlit UI --------------------

st.set_page_config(
    page_title="PO Details Extractor",
    page_icon="📋",
    layout="wide"
)

st.title("📋 PO Details Extractor")
st.markdown("""
Upload multiple PO PDF files to extract the Product Code, X-Mill Date, Brandix, PO Number, 
the **PO TOTAL AMOUNT**, and any other numbers in the same line (as Line Amount).
""")

DEFAULT_EXCEL_PATH = r"C:\Users\Pcadmin\Documents\ITL\CP-SO-Tracker\CPEXCEL.xlsx"

reference_mapping = {}
if 'reference_mapping' not in st.session_state:
    with st.spinner("Loading product reference database..."):
        reference_mapping = load_product_reference_data(DEFAULT_EXCEL_PATH)
        st.session_state.reference_mapping = reference_mapping
else:
    reference_mapping = st.session_state.reference_mapping

uploaded_files = st.file_uploader("Choose PO PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_results = []
    with st.spinner(f"Extracting PO details from {len(uploaded_files)} files..."):
        try:
            for uploaded_file in uploaded_files:
                product_code, x_mill_date, brandix, po_number, po_total, line_amount, department = extract_product_code_and_xmill_date(uploaded_file, reference_mapping)

                all_results.append({
                    "X-Mill Date": x_mill_date or "",
                    "Customer": brandix or "",
                    "Department": department or "",
                    "Product Reference": product_code or "",
                    "PO Number": po_number or "",
                    "Quantity": line_amount or "",
                    "Total value": po_total or ""
                })

            if all_results:
                df = pd.DataFrame(all_results)
                st.success(f"✅ Successfully extracted details from {len(all_results)} files")
                st.dataframe(df, use_container_width=True, hide_index=True)

                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv,
                    file_name='po_details.csv',
                    mime='text/csv',
                    use_container_width=True
                )
            else:
                st.warning("⚠ Could not find PO details in any of the uploaded files.")
                st.info("Please ensure the PO files contain the required information.")
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
            st.info("Please ensure you're uploading valid PO PDF files.")