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
    Extract text from PDF using PyMuPDF, with whitespace normalization.
    """
    data = file.read()
    doc = fitz.open(stream=data, filetype="pdf")
    chunks = []
    for page in doc:
        # "text" gives reading order; try "layout" if your PDFs lose label order
        chunks.append(page.get_text("text"))
    raw = "\n".join(chunks)

    # Normalize whitespace & punctuation
    text = raw.replace("\xa0", " ")
    text = re.sub(r"[：﹕]", ":", text)      # normalize colons
    text = re.sub(r"[–—−]", "-", text)      # normalize dashes
    text = "\n".join([ln.rstrip() for ln in text.splitlines()])  # strip trailing spaces per line
    return text


# ---------------------------
# PARTY & FIELD EXTRACTION
# ---------------------------

# Label vocab
BUYER_LABELS  = [r"Buyer", r"Purchaser", r"Buyer\s*\(Principal\)", r"Principal\s*\(Buyer\)"]
SELLER_LABELS = [r"Seller", r"Vendor", r"Seller\s*\(Counterparty\)", r"Counterparty\s*\(Seller\)"]

FIELD_LABELS = {
    "Date": [r"Date", r"Contract Date", r"Trade Date", r"Issue Date"],
    "Commodity": [r"Commodity"],
    "Quality": [r"Quality", r"Spec(?:ification)?"],
    "Quantity": [r"Quantity", r"Qty"],
    "Price": [r"Price", r"Price Basis", r"Contract Price"],
    "Delivery": [r"Delivery", r"Delivery Terms", r"Delivery Period"],
    "Payment": [r"Payment", r"Payment Terms"],
    "Insurance": [r"Insurance"],
    "Freight": [r"Freight"],
    "Storage": [r"Storage"],
    "Weights": [r"Weights", r"Weight(?:s)?\s*&\s*Measures", r"Weight Basis"],
    "Special Conditions": [r"Special Conditions", r"Specials", r"Notes"],
    "Brokerage": [r"Brokerage", r"Commission"],
    "Rules": [r"Rules", r"Contract Rules", r"Terms & Conditions"],
}

# Regex for ABN/ACN (supports spaced digits like "12 345 678 901")
ABN_RE = re.compile(r"\bA\s*B\s*N\s*[:\-]?\s*((?:\d\s*){11})\b", re.IGNORECASE)
ACN_RE = re.compile(r"\bA\s*C\s*N\s*[:\-]?\s*((?:\d\s*){9})\b", re.IGNORECASE)

META_PREFIX = re.compile(r"^(ABN|A\.B\.N\.|ACN|Address|Phone|Email|Tel|Fax|Contact)\b", re.IGNORECASE)

def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def choose_party_name(block: str) -> str:
    """
    Heuristics to pick a company/party name from a labeled party block.
    Prefers lines with company-ish cues; otherwise first non-meta line.
    """
    lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
    if not lines:
        return ""

    # Prefer lines that look like company names
    for ln in lines:
        if META_PREFIX.match(ln):
            continue
        if re.search(r"\b(Pty\.?\s*Ltd|Pty|Ltd|Limited|Trust|Company|Co\.?|Australia|Trading)\b", ln, re.IGNORECASE):
            return ln

    # Next: first non-meta line not starting with bullets/leading digits
    for ln in lines:
        if META_PREFIX.match(ln):
            continue
        if not re.match(r"^[\d\-\u2022•]", ln):  # ignore leading bullets/digits
            return ln

    # Fallback: the first non-empty line
    return lines[0]

def compile_all_label_pattern():
    """
    Build a single pattern that matches any label (party or field) at line start
    with optional colon after it. Used as lookahead stopper for block captures.
    """
    tokens = []
    for v in BUYER_LABELS + SELLER_LABELS:
        tokens.append(fr"(?:{v})")
    for variants in FIELD_LABELS.values():
        for v in variants:
            tokens.append(fr"(?:{v})")
    # ^ label [optional colon] optional spaces
    pattern = re.compile(fr"(?im)^(?:{'|'.join(tokens)})\s*:?\s*")
    return pattern
ALL_LABEL_AT_LINE_START = compile_all_label_pattern()

def find_block_after_label(text: str, label_variants) -> str:
    """
    Find the text block after a given label (e.g., Buyer) until the next label or EOF.
    Labels are matched at the **start of a line**, case-insensitive, colon optional.
    """
    label_alt = "|".join(label_variants)
    # Find the label at line start
    m = re.search(fr"(?im)^(?:{label_alt})\s*:?\s*", text)
    if not m:
        return ""
    start = m.end()
    # Find the next label occurrence from 'start'
    m2 = ALL_LABEL_AT_LINE_START.search(text, pos=start)
    end = m2.start() if m2 else len(text)
    return text[start:end].strip()

def extract_abn_acn(block: str):
    abn = ""
    acn = ""
    m_abn = ABN_RE.search(block)
    if m_abn:
        abn = normalize_digits(m_abn.group(1))
    m_acn = ACN_RE.search(block)
    if m_acn:
        acn = normalize_digits(m_acn.group(1))
    return abn, acn

def fallback_parties_from_abns(text: str):
    """
    If labeled 'Buyer'/'Seller' blocks aren't present, try to infer parties by
    scanning for ABNs and using nearby lines as names. Returns list of tuples.
    """
    results = []
    for m in ABN_RE.finditer(text):
        abn = normalize_digits(m.group(1))
        # grab ~5 lines before the ABN occurrence
        start = max(0, text.rfind("\n", 0, m.start()))
        context_start = text.rfind("\n", 0, start, )  # step back one more line start
        context_start = text.rfind("\n", 0, context_start) if context_start != -1 else start
        context_start = 0 if context_start == -1 else context_start
        window = text[context_start:m.start()]
        name = choose_party_name(window)
        results.append((name, abn))
    # Deduplicate by ABN while preserving order
    seen = set()
    deduped = []
    for name, abn in results:
        if abn and abn not in seen:
            deduped.append((name, abn))
            seen.add(abn)
    return deduped

def try_parse_date(text_fragment: str) -> str:
    """
    Attempt to parse a date from arbitrary text; returns dd/mm/YYYY or "".
    """
    candidates = [
        r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",         # 12/10/2025 or 12-10-25
        r"\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\b",       # 12 Oct 2025
        r"\b([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b",      # Oct 12, 2025
        r"\b(\d{4}-\d{2}-\d{2})\b",                     # 2025-10-12
    ]
    raw = ""
    for pat in candidates:
        m = re.search(pat, text_fragment, re.IGNORECASE)
        if m:
            raw = m.group(1).strip()
            break
    if not raw:
        return ""

    fmts = [
        "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
        "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y",
        "%Y-%m-%d",
    ]
    for f in fmts:
        try:
            dt = datetime.strptime(raw, f)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            continue
    return ""

def normalize_common_fields(fields: dict) -> dict:
    """
    Light cleanup on extracted values.
    """
    cleaned = dict(fields)

    # Quantity: prefer first line, number + unit
    q = cleaned.get("Quantity", "")
    if q:
        q_line = q.splitlines()[0]
        m = re.search(r"([\d,\.]+)\s*([A-Za-z/%]+)?", q_line)
        cleaned["Quantity"] = (m.group(0).strip() if m else q_line).strip()

    # Price: keep first line; often contains full basis (e.g., A$ 320/MT FIS)
    p = cleaned.get("Price", "")
    if p:
        cleaned["Price"] = p.splitlines()[0].strip()

    # Collapse excessive whitespace
    for k, v in cleaned.items():
        if isinstance(v, str):
            v = re.sub(r"[ \t]+", " ", v.strip())
            v = re.sub(r"\n{3,}", "\n\n", v)
            cleaned[k] = v.strip()

    return cleaned

def extract_fields(text: str) -> dict:
    """
    Main extraction: generic Buyer/Seller blocks, ABN/ACN, and multi-line fields.
    """
    # 1) Parties via labeled sections
    buyer_block = find_block_after_label(text, BUYER_LABELS)
    seller_block = find_block_after_label(text, SELLER_LABELS)

    buyer_name = choose_party_name(buyer_block) if buyer_block else ""
    seller_name = choose_party_name(seller_block) if seller_block else ""

    buyer_abn, buyer_acn = extract_abn_acn(buyer_block) if buyer_block else ("", "")
    seller_abn, seller_acn = extract_abn_acn(seller_block) if seller_block else ("", "")

    # 1b) Fallback: if any party missing, try to infer from ABN occurrences
    if not (buyer_name and buyer_abn) or not (seller_name and seller_abn):
        abn_candidates = fallback_parties_from_abns(text)
        # Assign first to Buyer if missing, second to Seller if missing
        if not (buyer_name and buyer_abn) and len(abn_candidates) >= 1:
            buyer_name = buyer_name or abn_candidates[0][0]
            buyer_abn = buyer_abn or abn_candidates[0][1]
        if not (seller_name and seller_abn) and len(abn_candidates) >= 2:
            seller_name = seller_name or abn_candidates[1][0]
            seller_abn = seller_abn or abn_candidates[1][1]

    # 2) Build combined next-label lookahead for field capture (multi-line until next label)
    all_tokens = []
    for v in BUYER_LABELS + SELLER_LABELS:
        all_tokens.append(fr"(?:{v})")
    for variants in FIELD_LABELS.values():
        for v in variants:
            all_tokens.append(fr"(?:{v})")
    next_label_or_eod = fr"(?:(?im)^(?:{'|'.join(all_tokens)})\s*:?\s*|$)"

    def capture_field(canonical: str) -> str:
        variants = FIELD_LABELS[canonical]
        label_alt = "|".join(variants)
        pat = re.compile(fr"(?im)^(?:{label_alt})\s*:?\s*(.+?)(?={next_label_or_eod})", re.DOTALL)
        m = pat.search(text)
        if not m:
            return ""
        val = m.group(1).strip()
        val = re.sub(r"[ \t]+\n", "\n", val)
        return val.strip()

    # 3) Date: try labeled; else generic date near top
    date_raw = capture_field("Date")
    if not date_raw:
        # search in the top ~50 lines
        head = "\n".join(text.splitlines()[:50])
        date_raw = try_parse_date(head)

    date_norm = try_parse_date(date_raw) if date_raw else ""

    # 4) Capture remaining fields
    fields = {
        "Buyer": buyer_name,
        "Buyer ABN": buyer_abn,
        "Seller": seller_name,
        "Seller ABN": seller_abn,
        # Optionally expose ACNs as separate fields (comment out if not needed)
        # "Buyer ACN": buyer_acn,
        # "Seller ACN": seller_acn,
        "Date": date_norm or date_raw or "",
        "Commodity": capture_field("Commodity"),
        "Quality": capture_field("Quality"),
        "Quantity": capture_field("Quantity"),
        "Price": capture_field("Price"),
        "Delivery": capture_field("Delivery"),
        "Payment": capture_field("Payment"),
        "Insurance": capture_field("Insurance"),
        "Freight": capture_field("Freight"),
        "Storage": capture_field("Storage"),
        "Weights": capture_field("Weights"),
        "Special Conditions": capture_field("Special Conditions"),
        "Brokerage": capture_field("Brokerage"),
        "Rules": capture_field("Rules"),
    }

    fields = normalize_common_fields(fields)
    return fields


# ---------------------------
# EXCEL GENERATION
# ---------------------------

def generate_excel(fields):
    df = pd.DataFrame(list(fields.items()), columns=["Field", "Value"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
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

    # For debugging: uncomment to see raw text (helps tune label variants)
    # with st.expander("Show raw text"):
    #     st.text(text)

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
