import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
from datetime import datetime

# -----------------------------------------
# Constants for Party Detection
# -----------------------------------------
ADDRESS_WORDS = set("""
PO BOX GPO LOCKED BAG CONTACT PHONE FAX EMAIL NSW VIC QLD SA WA TAS ACT NT
MELBOURNE SYDNEY BRISBANE ADELAIDE PERTH HOBART CANBERRA RHODES TOOWOOMBA
""".split())
NEXT_SEP_RE = re.compile(r"(LOCKED|GPO|PO\s+BOX|ABN)\b", re.IGNORECASE)
CONTACT_RE  = re.compile(r"CONTACT\b", re.IGNORECASE)
SUFFIXES = {"PTY","LTD","LIMITED","TRUST","COMPANY","CO","INC","LLC","LLP","AUSTRALIA"}

# -----------------------------------------
# PDF Text Extraction & Normalization
# -----------------------------------------
def extract_text_from_pdf(file) -> str:
    data = file.read()
    doc = fitz.open(stream=data, filetype="pdf")
    text = "\n".join(page.get_text("text") for page in doc)
    # Normalize for robust regex matching
    text = text.replace("\xa0", " ")
    text = re.sub(r"[：﹕]", ":", text)   # normalize colon variants
    text = re.sub(r"[–—−]", "-", text)   # normalize dash variants
    text = re.sub(r"[ \t]+", " ", text)  # collapse spaces
    text = re.sub(r"\s*\n\s*", "\n", text)  # tidy newlines
    return text

# -----------------------------------------
# Party (Buyer/Seller) Extraction
# -----------------------------------------
def extract_parties(text: str):
    """
    Detect Buyer/Seller by pairing each ABN with the nearest preceding uppercase
    company phrase, ignoring address/location tokens and contact names.
    """
    abn_matches = list(re.finditer(r"\bABN\s*:\s*((?:\d\s*){11})\b", text, re.IGNORECASE))
    parties = []

    for m in abn_matches:
        idx = m.start()
        abn = re.sub(r"\D", "", m.group(1))

        # Skip obvious header ABN (broker letterhead) near the top
        if idx < 300:
            continue

        # Look back a bit from the ABN
        win = text[max(0, idx - 280): idx]

        # If 'CONTACT' appears in the window, start after the last one (avoid person names)
        contact_hits = list(CONTACT_RE.finditer(win))
        if contact_hits:
            win = win[contact_hits[-1].end():]

        # Keep text before the next address marker (LOCKED/GPO/PO BOX/ABN)
        segment = NEXT_SEP_RE.split(win, maxsplit=1)[0]

        # Tokenize and keep ALL-CAPS tokens that are not address words
        tokens = [t for t in re.split(r"[^A-Za-z]+", segment) if t]
        uc_tokens = [t for t in tokens if t.isupper() and t not in ADDRESS_WORDS]

        name = ""
        if uc_tokens:
            # Prefer a company suffix (LTD/PTY/etc.): take suffix + up to 2 tokens before it
            suffix_index = -1
            for i, t in enumerate(uc_tokens):
                if t in SUFFIXES:
                    suffix_index = i
            if suffix_index != -1:
                start = max(0, suffix_index - 2)
                name = " ".join(uc_tokens[start: suffix_index + 1])
            else:
                # Otherwise, last up to 3 tokens (covers names like "ALLIED PINNACLE")
                name = " ".join(uc_tokens[-min(3, len(uc_tokens)):])

        # Basic sanity: require at least 2 words to reduce false positives
        if name and len(name.split()) >= 2:
            parties.append((name, abn, idx))

    # De-duplicate by ABN and sort by position
    seen, ordered = set(), []
    for n, a, i in parties:
        if a not in seen:
            ordered.append((n, a, i))
            seen.add(a)
    ordered.sort(key=lambda x: x[2])

    buyer_name = ordered[0][0] if len(ordered) > 0 else ""
    buyer_abn  = ordered[0][1] if len(ordered) > 0 else ""
    seller_name = ordered[-1][0] if len(ordered) > 1 else ""
    seller_abn  = ordered[-1][1] if len(ordered) > 1 else ""
    return buyer_name, buyer_abn, seller_name, seller_abn

# -----------------------------------------
# Field Extraction
# -----------------------------------------
def extract_fields(text: str) -> dict:
    # Parties
    buyer_name, buyer_abn, seller_name, seller_abn = extract_parties(text)

    # Date (generic patterns, normalize to dd/mm/YYYY)
    date_str = ""
    mdate = re.search(r"\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\b", text)
    if mdate:
        raw = mdate.group(1)
        for fmt in ("%d %b %Y", "%d %B %Y"):
            try:
                date_str = datetime.strptime(raw, fmt).strftime("%d/%m/%Y")
                break
            except Exception:
                pass
        if not date_str:
            date_str = raw

    # Colon-based labels → capture until next label-like line or EOF
    labels = [
        "Commodity","Quality","Quantity","Price","Delivery","Payment",
        "Insurance","Freight","Storage","Weights","Special Conditions","Brokerage","Rules"
    ]
    fields = {
        "Buyer": buyer_name,
        "Buyer ABN": buyer_abn,
        "Seller": seller_name,
        "Seller ABN": seller_abn,
        "Date": date_str,
    }
    for label in labels:
        pat = re.compile(rf"{label}:\s*(.+?)(?=\n[A-Z][A-Za-z ]{{1,40}}:\s|$)", re.DOTALL)
        m = pat.search(text)
        fields[label] = m.group(1).strip() if m else ""

    # Light tidy: first line for Price, compact whitespace
    if fields["Price"]:
        fields["Price"] = fields["Price"].splitlines()[0].strip()
    for k, v in list(fields.items()):
        if isinstance(v, str):
            fields[k] = re.sub(r"[ \t]+", " ", v.strip())

    return fields

# -----------------------------------------
# Excel Output
# -----------------------------------------
def generate_excel(fields: dict) -> io.BytesIO:
    df = pd.DataFrame(list(fields.items()), columns=["Field", "Value"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

# -----------------------------------------
# Streamlit UI
# -----------------------------------------
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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
