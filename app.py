import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io
from datetime import datetime
from collections import deque

# ============================================================
#                      CONFIG / CONSTANTS
# ============================================================

# Words that are typically part of addresses / headers, not company names
_ADDRESS_WORDS = set("""
PO BOX GPO LOCKED BAG CONTACT PHONE FAX EMAIL NSW VIC QLD SA WA TAS ACT NT
AUSTRALIA MELBOURNE SYDNEY BRISBANE ADELAIDE PERTH HOBART CANBERRA RHODES
TOOWOOMBA
""".split())

# Words that often appear between a company line and its ABN (used as separators)
_SEPARATORS = r"(?:LOCKED|GPO|PO\s+BOX|ABN|CONTACT|PHONE|FAX|EMAIL)\b"

# Company suffix tokens — helpful when deciding if a line is a company name
_COMPANY_SUFFIX = r"(?:PTY|LTD|LIMITED|TRUST|COMPANY|CO|INC|LLC|LLP|AUSTRALIA)\b"


# ============================================================
#                     PDF TEXT EXTRACTION
# ============================================================

def extract_text_from_pdf(file) -> str:
    """
    Extract text using PyMuPDF and normalize the result:
      - Replace non-breaking spaces
      - Normalize punctuation
      - Convert ** (bold markers common in PDF text dumps) to newlines to reintroduce block boundaries
      - Tidy whitespace and line endings
    """
    data = file.read()
    doc = fitz.open(stream=data, filetype="pdf")
    pages = []
    for page in doc:
        # If you see label order issues, consider: page.get_text("layout")
        pages.append(page.get_text("text"))
    raw = "\n".join(pages)

    text = raw.replace("\xa0", " ")
    text = re.sub(r"[：﹕]", ":", text)   # normalize colon
    text = re.sub(r"[–—−]", "-", text)   # normalize dash
    text = text.replace("**", "\n")      # critical: prevents blocks from sticking together

    # normalize whitespace/newlines
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    text = "\n".join(ln.rstrip() for ln in text.splitlines())
    return text


# ============================================================
#            PARTY (Buyer/Seller) EXTRACTION (ROBUST)
# ============================================================

def _extract_company_before_window(window: str) -> str:
    """
    From a lookback window ending just before 'ABN: ...',
    pick the rightmost uppercase company-looking phrase.

    Strategy:
      - Tokenize into alphabetic words and walk backwards.
      - Collect ALL-CAPS tokens that are not address words.
      - Stop when we hit a lowercase or a boundary.
      - Join tokens back in normal order.
    """
    tokens = [t for t in re.split(r"[^A-Za-z]+", window) if t]
    collected = deque()
    started = False
    for tok in reversed(tokens):
        if not tok.isupper():  # boundary (e.g., lowercase token)
            if started:
                break
            else:
                continue

        # Filter out address/location tokens; allow known company suffix tokens
        if tok in _ADDRESS_WORDS:
            if started:
                break
            else:
                continue

        collected.appendleft(tok)
        started = True

    name = " ".join(collected).strip()
    return name


def extract_parties(text: str):
    """
    Robust party extraction that works without 'Buyer:'/'Seller:' labels.

    - Finds all ABN occurrences (ABN: 11 digits).
    - Skips very-early header ABNs near the top (likely broker letterhead).
    - For each ABN, looks ~220 chars back, trims to after the last separator
      (LOCKED, GPO, PO BOX, ABN, CONTACT, etc.), and extracts the company phrase.
    - Dedupes by ABN and assumes doc order = Buyer then Seller.
    - Fallback: if two ABNs exist but names are missing, still map first->Buyer, last->Seller.
    """
    abn_iter = list(re.finditer(r"\bABN\s*:\s*((?:\d\s*){11})\b", text, flags=re.IGNORECASE))
    candidates = []

    for m in abn_iter:
        abn = re.sub(r"\D", "", m.group(1))
        start_idx = m.start()

        # Skip obvious header ABNs near very top (tweak threshold if needed)
        if start_idx < 300:
            continue

        # Window just before ABN
        lookback = 220
        window_start = max(0, start_idx - lookback)
        window = text[window_start:start_idx]

        # Trim to after last separator if present
        sep_match = list(re.finditer(_SEPARATORS, window, flags=re.IGNORECASE))
        if sep_match:
            window = window[sep_match[-1].end():]

        name = _extract_company_before_window(window)

        # If still empty, try a looser uppercase-phrase before separator
        if not name:
            simple_pat = re.compile(
                rf"([A-Z][A-Z&'./()\-\s]{{3,}}?(?:\s+{_COMPANY_SUFFIX})?)\s+(?={_SEPARATORS})",
                flags=re.IGNORECASE
            )
            matches = list(simple_pat.finditer(window))
            if matches:
                sm = matches[-1]
                name = re.sub(r"\s{2,}", " ", sm.group(1)).strip()

        name = re.sub(r"\s{2,}", " ", (name or "")).strip(" ,.-")

        # Require at least two tokens to reduce false positives
        if name and len(name.split()) >= 2:
            candidates.append((name, abn, start_idx))

    # De-dup by ABN (keep first)
    seen = set()
    parties = []
    for name, abn, idx in candidates:
        if abn not in seen:
            parties.append((name, abn, idx))
            seen.add(abn)

    # If we have Buyer and Seller, return them ordered by document position
    if len(parties) >= 2:
        parties_sorted = sorted(parties, key=lambda t: t[2])
        buyer_name, buyer_abn = parties_sorted[0][0], parties_sorted[0][1]
        seller_name, seller_abn = parties_sorted[-1][0], parties_sorted[-1][1]
        return [(buyer_name, buyer_abn), (seller_name, seller_abn)]

    # Fallback: if two ABNs exist but names unavailable, still map first->Buyer, last->Seller
    if len(abn_iter) >= 2 and not parties:
        abns = [re.sub(r"\D", "", m.group(1)) for m in abn_iter]
        return [("", abns[0]), ("", abns[-1])]

    # Else: return whatever we found (maybe only Buyer, or only Seller)
    return [(n, a) for (n, a, _) in parties]


def extract_parties_debug(text: str):
    """
    Helper to inspect the windows used for party detection.
    Use inside a Streamlit expander to troubleshoot tricky templates.
    """
    abn_iter = list(re.finditer(r"\bABN\s*:\s*((?:\d\s*){11})\b", text, flags=re.IGNORECASE))
    debug_rows = []
    for m in abn_iter:
        abn = re.sub(r"\D", "", m.group(1))
        idx = m.start()
        lookback = 220
        window = text[max(0, idx - lookback): idx]
        debug_rows.append({
            "abn": abn,
            "start_idx": idx,
            "window_tail": window[-260:],  # safe slice for display
            "name_guess": _extract_company_before_window(window)
        })
    return debug_rows


# ============================================================
#             FIELD (Commodity, Delivery, etc.) CAPTURE
# ============================================================

# Optional label vocab for Buyer/Seller sections (used only if you later add labeled blocks)
BUYER_LABELS  = [r"Buyer", r"Purchaser", r"Buyer\s*\(Principal\)", r"Principal\s*\(Buyer\)"]
SELLER_LABELS = [r"Seller", r"Vendor", r"Seller\s*\(Counterparty\)", r"Counterparty\s*\(Seller\)"]

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

def compile_next_label_pattern():
    """
    Build a pattern that represents "the next label" anywhere (colon required).
    We include party labels and field labels to serve as multi-line stop markers.
    """
    tokens = []
    for v in BUYER_LABELS + SELLER_LABELS:
        tokens.append(fr"(?:{v})")
    for variants in FIELD_LABELS.values():
        for v in variants:
            tokens.append(fr"(?:{v})")
    # 'next label' means any of these tokens followed by a colon
    return fr"(?:(?:{'|'.join(tokens)})\s*:)"


def capture_field(text: str, canonical: str, next_label_pat: str) -> str:
    """
    Capture field value for a given canonical label, requiring a colon,
    and grabbing text lazily until the next label (or end).
    """
    variants = FIELD_LABELS[canonical]
    label_alt = "|".join(variants)
    pat = re.compile(
        fr"\b(?:{label_alt})\s*:\s*(.+?)(?={next_label_pat}|$)",
        flags=re.IGNORECASE | re.DOTALL
    )
    m = pat.search(text)
    return m.group(1).strip() if m else ""


def try_parse_date(text_fragment: str) -> str:
    """
    Attempt to find and normalize a date to dd/mm/YYYY.
    """
    candidates = [
        r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",
        r"\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\b",
        r"\b([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b",
        r"\b(\d{4}-\d{2}-\d{2})\b",
    ]
    raw = ""
    for pat in candidates:
        m = re.search(pat, text_fragment, flags=re.IGNORECASE)
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
            return datetime.strptime(raw, f).strftime("%d/%m/%Y")
        except Exception:
            continue
    return ""


def normalize_common_fields(fields: dict) -> dict:
    """
    Light clean-up: quantity, price, whitespace.
    """
    out = dict(fields)

    # Quantity -> prefer first line, number + unit
    q = out.get("Quantity", "")
    if q:
        q_line = q.splitlines()[0]
        m = re.search(r"([\d,\.]+)\s*([A-Za-z/%]+)?", q_line)
        out["Quantity"] = (m.group(0).strip() if m else q_line).strip()

    # Price -> first line
    p = out.get("Price", "")
    if p:
        out["Price"] = p.splitlines()[0].strip()

    # Collapse whitespace
    for k, v in out.items():
        if isinstance(v, str):
            v = re.sub(r"[ \t]+", " ", v.strip())
            v = re.sub(r"\n{3,}", "\n\n", v)
            out[k] = v.strip()

    return out


def extract_fields(text: str) -> dict:
    """
    Main extraction:
      - Parties (robust, unlabeled)
      - Date (unlabeled heuristic near top)
      - Multi-line fields captured until the next colon-label
    """
    # Parties
    parties = extract_parties(text)
    buyer_name = parties[0][0] if len(parties) > 0 else ""
    buyer_abn  = parties[0][1] if len(parties) > 0 else ""
    seller_name = parties[1][0] if len(parties) > 1 else ""
    seller_abn  = parties[1][1] if len(parties) > 1 else ""

    # Next-label pattern for multi-line fields
    next_label_pat = compile_next_label_pattern()

    # Date (search near top)
    head = "\n".join(text.splitlines()[:120])
    date_norm = try_parse_date(head)

    # Capture labeled fields
    fields = {
        "Buyer": buyer_name,
        "Buyer ABN": buyer_abn,
        "Seller": seller_name,
        "Seller ABN": seller_abn,
        "Date": date_norm,
        "Commodity": capture_field(text, "Commodity", next_label_pat),
        "Quality": capture_field(text, "Quality", next_label_pat),
        "Quantity": capture_field(text, "Quantity", next_label_pat),
        "Price": capture_field(text, "Price", next_label_pat),
        "Delivery": capture_field(text, "Delivery", next_label_pat),
        "Payment": capture_field(text, "Payment", next_label_pat),
        "Insurance": capture_field(text, "Insurance", next_label_pat),
        "Freight": capture_field(text, "Freight", next_label_pat),
        "Storage": capture_field(text, "Storage", next_label_pat),
        "Weights": capture_field(text, "Weights", next_label_pat),
        "Special Conditions": capture_field(text, "Special Conditions", next_label_pat),
        "Brokerage": capture_field(text, "Brokerage", next_label_pat),
        "Rules": capture_field(text, "Rules", next_label_pat),
    }

    return normalize_common_fields(fields)


# ============================================================
#                     EXCEL GENERATION
# ============================================================

def generate_excel(fields: dict) -> io.BytesIO:
    df = pd.DataFrame(list(fields.items()), columns=["Field", "Value"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


# ============================================================
#                        STREAMLIT UI
# ============================================================

st.title("Broker Advice PDF Extractor")
st.write("Upload a broker advice PDF to extract key fields and download a vertical Excel template.")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    text = extract_text_from_pdf(uploaded_file)

    # Optional: debugging helpers
    with st.expander("Show raw text (debug)"):
        st.text(text[:5000])  # clip for display

    with st.expander("Party detection debug (ABN windows)"):
        for row in extract_parties_debug(text):
            st.code(
                f"ABN: {row['abn']}\n"
                f"Idx: {row['start_idx']}\n"
                f"Name guess: {row['name_guess']}\n"
                f"Window tail:\n{row['window_tail']}"
            )

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
