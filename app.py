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
ADDRESS_HINTS = re.compile(
    r"\b(PO BOX|GPO|LOCKED BAG|STREET|ROAD|RD|AVE|AVENUE|DRIVE|DR|LANE|LN|"
    r"QLD|NSW|VIC|SA|WA|TAS|ACT|NT|\d{3,4})\b",
    re.IGNORECASE
)

# -----------------------------------------
# PDF Text Extraction & Normalization
# -----------------------------------------
def extract_text_from_pdf(file) -> str:
    data = file.read()
    doc = fitz.open(stream=data, filetype="pdf")
    text = "\n".join(page.get_text("text") for page in doc)
    # Normalize for robust regex matching
    text = text.replace("\xa0", " ")
    text = text.replace("**", "\n")        # <<< fix: keep block boundaries from bold markers
    text = re.sub(r"[：﹕]", ":", text)     # normalize colon variants
    text = re.sub(r"[–—−]", "-", text)     # normalize dash variants
    text = re.sub(r"[ \t]+", " ", text)    # collapse spaces
    text = re.sub(r"\s*\n\s*", "\n", text) # tidy newlines
    return text


# -----------------------------------------
# Party (Buyer/Seller) Extraction
# -----------------------------------------
def extract_parties(text: str):
    """
    Detect Buyer/Seller by pairing each ABN with the nearest preceding uppercase
    company line. Prefer the full uppercase line; fall back to suffix/token logic.
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

        # Start after last CONTACT in the window to avoid person names
        contact_hits = list(CONTACT_RE.finditer(win))
        if contact_hits:
            win = win[contact_hits[-1].end():]

        # Keep text before the next address marker (LOCKED/GPO/PO BOX/ABN)
        segment = NEXT_SEP_RE.split(win, maxsplit=1)[0].strip()

        # ---------- PRIMARY: line-based capture ----------
        # Split the segment by lines and pick the last uppercase, non-address line
        company_line = ""
        lines = [ln.strip() for ln in segment.split("\n") if ln.strip()]
        for ln in reversed(lines):
            up = ln.upper()
            if "CONTACT" in up:
                continue
            if not up.isupper():
                continue
            if ADDRESS_HINTS.search(ln):
                continue
            # require at least two words to reduce false positives
            if len(ln.split()) >= 2:
                company_line = ln
                break

        if company_line:
            name = company_line
        else:
            # ---------- FALLBACK: suffix/token capture ----------
            tokens = [t for t in re.split(r"[^A-Za-z]+", segment) if t]
            uc_tokens = [t for t in tokens if t.isupper() and t not in ADDRESS_WORDS]
            name = ""
            if uc_tokens:
                suffix_positions = [i for i, t in enumerate(uc_tokens) if t in SUFFIXES]
                if suffix_positions:
                    suf_idx = suffix_positions[-1]
                    # include suffix + up to 3 tokens before (ensures at least one non-suffix)
                    start = max(0, suf_idx - 3)
                    slice_tokens = uc_tokens[start:suf_idx + 1]
                    # ensure at least one non-suffix present
                    if all(tok in SUFFIXES for tok in slice_tokens) and start > 0:
                        slice_tokens = uc_tokens[start - 1:suf_idx + 1]
                    name = " ".join(slice_tokens)
                else:
                    # no explicit suffix; take last up to 4 tokens
                    name = " ".join(uc_tokens[-min(4, len(uc_tokens)):])

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
# Formatting Helpers
# -----------------------------------------
_MONTHS = {
    'JANUARY': '01','FEBRUARY': '02','MARCH': '03','APRIL': '04','MAY': '05','JUNE': '06',
    'JULY': '07','AUGUST': '08','SEPTEMBER': '09','OCTOBER': '10','NOVEMBER': '11','DECEMBER': '12',
    'JAN': '01','FEB': '02','MAR': '03','APR': '04','MAY': '05','JUN': '06',
    'JUL': '07','AUG': '08','SEP': '09','SEPT': '09','OCT': '10','NOV': '11','DEC': '12',
}

def _strip_day_suffix(s: str) -> str:
    return re.sub(r"(\d{1,2})(st|nd|rd|th)", r"\1", s, flags=re.IGNORECASE)

def _parse_text_date(s: str) -> str:
    """
    Parse dates like 'DECEMBER 1ST 2025' or '1 Dec 2025' to 'dd/mm/YYYY'.
    Returns '' if no parse.
    """
    s = _strip_day_suffix(s).strip()
    m1 = re.search(r"\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})\b", s, re.IGNORECASE)
    m2 = re.search(r"\b([A-Za-z]{3,9})\s+(\d{1,2})\s+(\d{4})\b", s, re.IGNORECASE)
    dd = mm = yyyy = None
    if m1:
        dd, mon, yyyy = m1.group(1), m1.group(2), m1.group(3)
    elif m2:
        mon, dd, yyyy = m2.group(1), m2.group(2), m2.group(3)
    if not (dd and yyyy and mon):
        return ""
    mm = _MONTHS.get(mon.upper(), "")
    if not mm:
        return ""
    dd = f"{int(dd):02d}"
    return f"{dd}/{mm}/{yyyy}"

def _format_price(val: str) -> str:
    # pick first currency amount like $340.00 or A$ 340
    m = re.search(r"(A\$|\$)\s*([0-9][0-9,]*\.?\d*)", val, re.IGNORECASE)
    if not m:
        return val
    amount = m.group(2).replace(",", "")
    try:
        amount = f"{float(amount):.2f}"
    except:
        pass
    # per your spec: lowercase /mt
    return f"${amount}/mt"

def _format_quantity(val: str) -> str:
    # find first number; keep (MIN/MAX) if present
    m = re.search(r"([0-9][0-9,]*\.?\d*)", val)
    out = val
    if m:
        num = m.group(1).replace(",", "")
        try:
            num = f"{float(num):.2f}"
        except:
            pass
        out = f"{num}mt"
        if re.search(r"\bMIN/MAX\b", val, re.IGNORECASE):
            out += " (MIN/MAX)"
    return out

def _format_delivery(val: str) -> str:
    # extract two dates from a phrase like 'DECEMBER 1ST 2025 TO JANUARY 29TH 2026'
    # or '1 Dec 2025 to 29 Jan 2026'; if month-only, output MM/YYYY - MM/YYYY
    seg = val.replace("—", "-").replace(" to ", " TO ").replace("–", "-")
    # First try to find two day-level dates
    date_tokens = []
    for m in re.finditer(r"([A-Za-z]{3,9}\s+\d{1,2}(?:st|nd|rd|th)?\s+\d{4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
                         seg, re.IGNORECASE):
        parsed = _parse_text_date(m.group(0))
        if parsed:
            date_tokens.append(parsed)
        if len(date_tokens) == 2:
            break
    if len(date_tokens) == 2:
        return f"{date_tokens[0]} - {date_tokens[1]}"
    # Else, try Month Year -> Month Year
    m = re.search(r"([A-Za-z]{3,9})\s+(\d{4})\s+TO\s+([A-Za-z]{3,9})\s+(\d{4})", seg, re.IGNORECASE)
    if m:
        m1, y1, m2, y2 = m.group(1), m.group(2), m.group(3), m.group(4)
        mm1 = _MONTHS.get(m1.upper(), "")
        mm2 = _MONTHS.get(m2.upper(), "")
        if mm1 and mm2:
            return f"{mm1}/{y1} - {mm2}/{y2}"
    # fallback: leave original if nothing matched
    return val

def _format_brokerage(val: str) -> str:
    # find A$ or $ amount; output A$X.XX/MT (EXCL GST)
    m = re.search(r"(A\$|\$)\s*([0-9][0-9,]*\.?\d*)", val, re.IGNORECASE)
    if not m:
        return val
    amount = m.group(2).replace(",", "")
    try:
        amount = f"{float(amount):.2f}"
    except:
        pass
    return f"A${amount}/MT (EXCL GST)"

def format_output(fields: dict) -> dict:
    out = dict(fields)
    # Price
    if out.get("Price"):
        out["Price"] = _format_price(out["Price"])
    # Quantity
    if out.get("Quantity"):
        out["Quantity"] = _format_quantity(out["Quantity"])
    # Delivery
    if out.get("Delivery"):
        out["Delivery"] = _format_delivery(out["Delivery"])
    # Brokerage
    if out.get("Brokerage"):
        out["Brokerage"] = _format_brokerage(out["Brokerage"])
    return out

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
    fields = format_output(fields)   # <-- Apply your display formatting here

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
