import streamlit as st
import pandas as pd
import pdfplumber
import re
from difflib import get_close_matches
from io import BytesIO
from openpyxl import load_workbook
import os

# ---------------- Extraction functions ----------------
def extract_sb_data(pdf_path):
    sb_data = []

    sb_regex = r'\b\d{5,8}\b'
    date_regex = r'\b\d{2}-[A-Z]{3}-\d{2}\b'
    iec_regex = r'IEC/Br\s*[:\-]?\s*([A-Z0-9]+)'
    gstin_regex = r'GSTIN/TYPE\s*[:\-]?\s*([A-Z0-9]+)'
    cbcode_regex = r'CB CODE\s*[:\-]?\s*([A-Z0-9]+)'

    country_list = [
        "INDIA", "SWEDEN", "GERMANY", "USA", "UNITED STATES", "FRANCE", "ITALY", 
        "CHINA", "JAPAN", "SOUTH KOREA", "THAILAND", "SINGAPORE", "UAE", "BRAZIL",
        "UK", "UNITED KINGDOM", "NORWAY", "FINLAND", "DENMARK", "NETHERLANDS",
        "POLAND", "SPAIN", "CANADA", "AUSTRALIA", "SWITZERLAND", "BELGIUM"
    ]

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

        iec_value = re.search(iec_regex, text)
        iec_value = iec_value.group(1) if iec_value else ""
        gstin_value = re.search(gstin_regex, text)
        gstin_value = gstin_value.group(1) if gstin_value else ""
        cbcode_value = re.search(cbcode_regex, text)
        cbcode_value = cbcode_value.group(1) if cbcode_value else ""

        final_dest_value = ""
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if re.search(r'13\.*\s*COUNTRY\s*OF\s*FINALDESTINATIO', line, re.IGNORECASE):
                candidates = []
                after = line.split("13.COUNTRY OF FINALDESTINATIO")[-1].strip()
                if after:
                    candidates.append(after)
                for next_line in lines[i+1:i+5]:
                    clean_next = next_line.strip()
                    if clean_next:
                        candidates.append(clean_next)
                for cand in candidates:
                    cand_clean = re.sub(r'[^A-Z\s]', '', cand.upper()).strip()
                    if cand_clean:
                        match = get_close_matches(cand_clean, country_list, n=1, cutoff=0.5)
                        if match:
                            final_dest_value = match[0]
                            break
                if not final_dest_value and candidates:
                    final_dest_value = candidates[0].strip()
                break

        for i, line in enumerate(lines):
            if "Port Code SB No SB Date" in line:
                if i + 1 < len(lines):
                    next_line = lines[i + 1]
                    sb_numbers = re.findall(sb_regex, next_line)
                    dates = re.findall(date_regex, next_line)
                    port_code = ""
                    if sb_numbers:
                        sb_index = next_line.find(sb_numbers[0])
                        port_code_candidate = next_line[:sb_index].strip()
                        port_code = port_code_candidate.split()[-1] if port_code_candidate else ""
                    for j in range(max(len(sb_numbers), len(dates))):
                        sb_data.append({
                            "PORT CODE(FROM)": port_code,
                            "SHIPPINGBILL NO": sb_numbers[j] if j < len(sb_numbers) else "",
                            "SHIPPING BILL DATE": dates[j] if j < len(dates) else "",
                            "IE CODE": iec_value,
                            "GSTIN/TYPE": gstin_value,
                            "CB CODE": cbcode_value,
                            "FINAL DESTINATION": final_dest_value,
                            "INVOICE NO": ""
                        })

    return pd.DataFrame(sb_data) if sb_data else None

def extract_invoice_tables(pdf_path):
    page_tables_dict = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            tables = page.extract_tables()
            if (text and "PART - II - INVOICE DETAILS" in text) or i == 1:
                if tables:
                    page_df = pd.concat([pd.DataFrame(tbl) for tbl in tables], ignore_index=True)
                    page_tables_dict[f"Page_{i}"] = page_df
                else:
                    page_tables_dict[f"Page_{i}"] = pd.DataFrame([["No table found on this page"]])
    return page_tables_dict

def extract_invoice_details_from_all_pages(tables_dict, sb_df=None):
    if sb_df is None or sb_df.empty:
        sb_df = pd.DataFrame()
    invoice_rows = []

    for page_name, page_df in tables_dict.items():
        if page_name == "Page_1":
            continue
        page_df = page_df.fillna("").astype(str)
        try:
            if page_df.shape[0] >= 13 and page_df.shape[1] >= 10:
                check_cell = page_df.iat[12, 9].strip().upper()
                if "2.BUYER'S NAME & ADDRESS".upper() in check_cell:
                    invoice_no, invoice_date = "", ""
                    if page_df.shape[0] >= 12 and page_df.shape[1] >= 3:
                        cell_val = page_df.iat[11, 2].strip()
                        match = re.match(r"([A-Za-z0-9/\\-]+)\s+(\d{2}/\d{2}/\d{4})", cell_val)
                        if match:
                            invoice_no = match.group(1)
                            invoice_date = match.group(2)
                        else:
                            invoice_no = cell_val
                    drawee_name = page_df.iat[13, 9].strip() if page_df.shape[0] >= 14 else ""
                    drawee_address = " ".join([page_df.iat[r, 9].strip() for r in range(14, min(19, page_df.shape[0])) if len(page_df.iat[r, 9].strip()) > 2])
                    goods_desc = page_df.iat[28, 4].strip() if page_df.shape[0] >= 29 else ""
                    port_of_dest = ""
                    if "Page_1" in tables_dict:
                        page1_df = tables_dict["Page_1"].fillna("").astype(str)
                        if page1_df.shape[0] >= 14 and page1_df.shape[1] >= 30:
                            port_of_dest = page1_df.iat[13, 29].strip()
                    invoice_rows.append({
                        "INVOICE NO": invoice_no,
                        "INVOICE DATE": invoice_date,
                        "DRAWEE NAME": drawee_name,
                        "DRAWEE ADDRESS": drawee_address,
                        "GOODS DESCRIPTION": goods_desc,
                        "PORT OF DESTINATION": port_of_dest
                    })
        except Exception as e:
            print(f"Error processing {page_name}: {e}")

    combined_rows = []
    if not sb_df.empty:
        for _, sb_row in sb_df.iterrows():
            for inv in invoice_rows:
                combined_rows.append({**sb_row.to_dict(), **inv})
    else:
        combined_rows = invoice_rows

    return pd.DataFrame(combined_rows)

# ---------------- Streamlit App ----------------
st.set_page_config(page_title="Multi-PDF SB Data Extractor", layout="wide")
st.title("📄 Multi-PDF SB Data Extractor")

# --- Upload PDFs ---
uploaded_files = st.file_uploader(
    "Upload PDF files (multiple allowed)", type=["pdf"], accept_multiple_files=True
)

combined_sb_df = pd.DataFrame()
if uploaded_files:
    st.info("Processing PDFs...")
    for uploaded_file in uploaded_files:
        pdf_path = f"temp_uploaded.pdf"
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.read())
        sb_df = extract_sb_data(pdf_path)
        invoice_tables_dict = extract_invoice_tables(pdf_path)
        sb_df = extract_invoice_details_from_all_pages(invoice_tables_dict, sb_df=sb_df)
        if sb_df is not None and not sb_df.empty:
            combined_sb_df = pd.concat([combined_sb_df, sb_df], ignore_index=True)

    if not combined_sb_df.empty:
        combined_sb_df = combined_sb_df.ffill()
        st.subheader("📊 Combined SB Data")
        st.dataframe(combined_sb_df)

        # Download combined SB data
        towrite = BytesIO()
        combined_sb_df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            label="⬇️ Download Combined SB Data as Excel",
            data=towrite,
            file_name="Combined_SB_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------------- Upload Template to Fill ----------------
st.subheader("📥 Upload Excel Template to Fill Data")
uploaded_excel = st.file_uploader(
    "Upload Excel file containing 'SHIPPINGBILL NO' column", type=["xlsx"]
)

if uploaded_excel and not combined_sb_df.empty:
    template_df = pd.read_excel(uploaded_excel)
    if "SHIPPINGBILL NO" not in template_df.columns:
        st.error("❌ The Excel must contain 'SHIPPINGBILL NO' column.")
    else:
        # Normalize SB numbers
        template_df["SHIPPINGBILL NO"] = template_df["SHIPPINGBILL NO"].astype(str).str.strip()
        combined_sb_df["SHIPPINGBILL NO"] = combined_sb_df["SHIPPINGBILL NO"].astype(str).str.strip()

        # Merge extracted data into template
        filled_df = template_df.merge(
            combined_sb_df, on="SHIPPINGBILL NO", how="left", suffixes=("", "_extracted")
        )

        st.subheader("✅ Filled Template Preview")
        st.dataframe(filled_df)

        # Download filled Excel
        towrite = BytesIO()
        filled_df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            label="⬇️ Download Filled Excel",
            data=towrite,
            file_name="Filled_SB_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
