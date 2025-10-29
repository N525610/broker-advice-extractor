import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
from datetime import datetime

# ---------------------------
# PDF TEXT EXTRACTION
# ---------------------------

def extract_text_from_pdf(file):
    """
    Extract text with PyMuPDF and normalize.
    - Replace non-breaking spaces
    - Normalize punctuation
    - Convert '**' (bold markers that often appear in PDF text dumps) into newlines to prevent block concatenation
    """
    data = file.read()
    doc = fitz.open(stream=data, filetype="pdf")
    pages = []
    for page in doc:
        pages.append(page.get_text("text"))  # try "layout" if label order is broken
    raw = "\n".join(pages)

    text = raw.replace("\xa0", " ")
    text = re.sub(r"[：﹕]", ":", text)  # normalize colon
    text = re.sub(r"[–—−]", "-", text)  # normalize dash
    text = text.replace("**", "\n")     # <-- critical: keep block boundaries
    # clean whitespace/newlines
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    text = "\n".join(ln.rstrip() for ln in text.splitlines())
    return text


# ---------------------------
# PARTY & FIELD EXTRACTION
# ---------------------------

# Optional labeled-section support if some templates do have "Buyer:" / "Seller:" headings
BUYER_LABELS  = [r"Buyer", r"Purchaser", r"Buyer\s*\(Principal\)", r"Principal\s*\(Buyer\)"]
SELLER_LABELS = [r"Seller", r"Vendor", r"Seller\s*\(Counterparty\)", r"Counterparty\s*\(Seller\)"]

def compile_all_label_pattern(field_label_variants):
    """Utility to build 'next label' pattern for lookaheads (colon required)."""
    tokens = []
    for v in BUYER_LABELS + SELLER_LABELS:
        tokens.append(fr"(?:{v})")
    for variants in field_label_variants.values():
        for v in variants:
            tokens.append(fr"(?:{v})")
    # next label = any of the tokens followed by a colon
    return fr"(?:(?:{'|'.join(tokens)})\s*:)"

def find_block_after_label(text: str, label_variants) -> str:
    """Return text block after a line-start label until the next label or EOF."""
    label_alt = "|".join(label_variants)
    m = re.search(fr"^(?:{label_alt})\s*:?\s*", text, flags=re.IGNORECASE | re.MULTILINE)
    if not m:
        return ""
    # Next label occurrence (start of next block)
    next_label_pattern = re.compile(fr"^(?:{label_alt})\s*:?\s*", flags=re.IGNORECASE | re.MULTILINE)
    # This next search is generic across *all* labels (Buyer/Seller + field labels),
    # so we’ll compile that later and slice with it in extract_fields where we have the full token set.
    return text[m.end():]  # raw tail; truncation is handled in extract_fields

def extract_parties(text: str):
    """
    Party fallback for templates without 'Buyer:' / 'Seller:':
    Detect uppercase company lines followed soon after by ABN on following lines.
    Returns list of (name, abn) in document order.
    """
    party_pat = re.compile(
        r"^([A-Z][A-Z &'./()\-\&]{2,})\s*$\n.*?\bABN\b\s*:\s*((?:\d\s*){11})",
        flags=re.IGNORECASE | re.MULTILINE | re.DOTALL
    )
    parties = []
    for m in party_pat.finditer(text):
        name = m.group(1).strip()
        # Keep only ALL-CAPS headers as company names
        if name and name.upper() == name:
            abn = re.sub(r"\D", "", m.group(2))
            parties.append((name, abn))
    # Dedup by ABN in order
    seen = set()
    parties = [(n, a) for n, a in parties if not (a in seen or seen.add(a))]
    return parties  # expected: [(BuyerName, BuyerABN), (SellerName, SellerABN)]


FIELD_LABELS = {
    "Commodity": [r"Commodity"],
    "Quality": [r"Quality", r"Spec(?:ification)?"],
    "Quantity": [r"Quantity", r"Qty"],
    "Price": [r"Price", r"Price Basis", r"Contract Price"],
    "Delivery": [r"Delivery", r"Delivery Terms", r"Delivery Period"],
    "Payment": [r"Payment", r"Payment Terms"],
    "Insurance": [r"Insurance"],
    "Freight": [r"Freight"],
    "Storage": [r"Storage"],
    "Weights": [r"Weights", r"Weight Basis", r"Weight(?:s)?\s*&\s*Measures"],
    "Special Conditions": [r"Special Conditions", r"Specials", r"Notes"],
    "Brokerage": [r"Brokerage", r"Commission"],
    "Rules": [r"Rules", r"Contract Rules", r"Terms & Conditions"],
}

def try_parse_date(text_fragment: str) -> str:
    pats = [
        r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",
        r"\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\b",
        r"\b([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b",
        r"\b(\d{4}-\d{2}-\d{2})\b",
    ]
    raw = ""
    for p in pats:
        m = re.search(p, text_fragment, flags=re.IGNORECASE)
        if m:
            raw = m.group(1).strip()
            break
    if not raw:
        return ""
    fmts = ["%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
            "%d %b %Y", "%d %B %Y",
            "%b %d, %Y", "%B %d, %Y",
            "%Y-%m-%d"]
    for f in fmts:
        try:
            return datetime.strptime(raw, f).strftime("%d/%m/%Y")
        except Exception:
            pass
    return ""

def normalize_common_fields(fields: dict) -> dict:
    out = dict(fields)
    # Quantity: first line → number + unit if possible
    q = out.get("Quantity", "")
    if q:
        q_line = q.splitlines()[0]
        m = re.search(r"([\d,\.]+)\s*([A-Za-z/%]+)?", q_line)
        out["Quantity"] = (m.group(0).strip() if m else q_line).strip()
    # Price: keep first line (often contains full basis)
    p = out.get("Price", "")
    if p:
        out["Price"] = p.splitlines()[0].strip()
    # Trim whitespace
    for k, v in out.items():
        if isinstance(v, str):
            v = re.sub(r"[ \t]+", " ", v.strip())
            v = re.sub(r"\n{3,}", "\n\n", v)
            out[k] = v.strip()
    return out

def extract_fields(text: str) -> dict:
    # 0) Parties
    parties = extract_parties(text)

    buyer_name = parties[0][0] if len(parties) > 0 else ""
    buyer_abn  = parties[0][1] if len(parties) > 0 else ""
    seller_name = parties[1][0] if len(parties) > 1 else ""
    seller_abn  = parties[1][1] if len(parties) > 1 else ""

    # 1) Field capture: colon REQUIRED, until next colon-label (any field/party label)
    next_label = compile_all_label_pattern(FIELD_LABELS)

    def capture(canonical: str) -> str:
        variants = FIELD_LABELS[canonical]
        label_alt = "|".join(variants)
        pat = re.compile(
            fr"\b(?:{label_alt})\s*:\s*(.+?)(?={next_label}|$)",
            flags=re.IGNORECASE | re.DOTALL
        )
        m = pat.search(text)
        return m.group(1).strip() if m else ""

    # 2) Date (unlabeled fallback near top)
    head = "\n".join(text.splitlines()[:100])
    date_norm = try_parse_date(head)

    fields = {
        "Buyer": buyer_name,
        "Buyer ABN": buyer_abn,
        "Seller": seller_name,
        "Seller ABN": seller_abn,
        "Date": date_norm,
        "Commodity": capture("Commodity"),
        "Quality": capture("Quality"),
        "Quantity": capture("Quantity"),
        "Price": capture("Price"),
        "Delivery": capture("Delivery"),
        "Payment": capture("Payment"),
        "Insurance": capture("Insurance"),
        "Freight": capture("Freight"),
        "Storage": capture("Storage"),
        "Weights": capture("Weights"),
        "Special Conditions": capture("Special Conditions"),
        "Brokerage": capture("Brokerage"),
        "Rules": capture("Rules"),
    }

    return normalize_common_fields(fields)


# ---------------------------
# EXCEL GENERATION
# ---------------------------

def generate_excel(fields):
    df = pd.DataFrame(list(fields.items()), columns=["Field", "Value"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


# ---------------------------
# STREAMLIT APP
# ---------------------------

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
