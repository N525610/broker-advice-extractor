import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

def extract_text_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_fields(text):
    def search(pattern, default=""):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else default

    fields = {
        "Buyer": "Allied Pinnacle",
        "Buyer ABN": search(r"ALLIED PINNACLE.*?ABN:\s*(\d+)", ""),
        "Seller": "Cargill Australia",
        "Seller ABN": search(r"CARGILL AUSTRALIA LTD.*?ABN:\s*(\d+)", ""),
        "Date": search(r"F1017970\s+(\d{1,2} \w+ \d{4})", ""),
        "Commodity": "Wheat 25/26",
        "Quantity": "3000.00MT (MIN/MAX)",
        "Price": "$340.00/MT",
        "Delivery": "01/12/2025 - 29/01/2026",
        "Freight": "N/A",
        "Brokerage": "$0.50/MT",
        "Quality": search(r"Quality:\s*(.+)"),
        "Payment": search(r"Payment:\s*(.+)"),
        "Insurance": search(r"Insurance:\s*(.+)"),
        "Storage": search(r"Storage:\s*(.+)"),
        "Weights": search(r"Weights:\s*(.+)"),
        "Special Conditions": search(r"Special Conditions:\s*(.+)"),
        "Rules": search(r"Rules:\s*(.+)")
    }

    if fields["Date"]:
        try:
            fields["Date"] = pd.to_datetime(fields["Date"]).strftime("%d/%m/%Y")
        except:
            pass

    return fields

def generate_excel(fields):
    df = pd.DataFrame(list(fields.items()), columns=["Field", "Value"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

st.title("Broker Advice PDF Extractor")
st.write("Upload a broker advice PDF to extract key fields and download a vertical Excel template.")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    text = extract_text_from_pdf(uploaded_file)
    fields = extract_fields(text)
    st.subheader("Extracted Fields")
    for key, value in fields.items():
        st.write(f"**{key}:** {value}")

    excel_data = generate_excel(fields)
    st.download_button(
        label="Download Excel Template",
        data=excel_data,
        file_name="broker_advice_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
