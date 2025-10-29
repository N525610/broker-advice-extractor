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
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        return match.group(1).strip() if match else default

    # Extract Buyer and Seller names and ABNs using verified structure
    parties = re.findall(r"\*\*(.*?)\*\*.*?ABN:\s*(\d{11})", text)
    buyer_name = parties[0][0].strip() if len(parties) > 0 else ""
    buyer_abn = parties[0][1].strip() if len(parties) > 0 else ""
    seller_name = parties[1][0].strip() if len(parties) > 1 else ""
    seller_abn = parties[1][1].strip() if len(parties) > 1 else ""

    fields = {
        "Buyer": buyer_name,
        "Buyer ABN": buyer_abn,
        "Seller": seller_name,
        "Seller ABN": seller_abn,
        "Date": search(r"\*\* F1017970\*\*\s+(\d{1,2} \w+ \d{4})", ""),
        "Commodity": search(r"Commodity:\*\*\s*(.+?)\*\*", ""),
        "Quality": search(r"Quality:\*\*\s*(.+?)\*\*", ""),
        "Quantity": search(r"Quantity:\*\*\s*(.+?)\*\*", ""),
        "Price": search(r"Price:\*\*\s*(.+?)\*\*", ""),
        "Delivery": search(r"Delivery:\*\*\s*(.+?)\*\*", ""),
        "Payment": search(r"Payment:\*\*\s*(.+?)\*\*", ""),
        "Insurance": search(r"Insurance:\*\*\s*(.+?)\*\*", ""),
        "Freight": search(r"Freight:\*\*\s*(.+?)\*\*", ""),
        "Storage": search(r"Storage:\*\*\s*(.+?)\*\*", ""),
        "Weights": search(r"Weights:\*\*\s*(.+?)\*\*", ""),
        "Special Conditions": search(r"Special Conditions:\*\*\s*(.+?)\*\*", ""),
        "Brokerage": search(r"Brokerage:\*\*\s*(.+?)\*\*", ""),
        "Rules": search(r"Rules:\*\*\s*(.+?)(?:\\n|$)", "")
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
