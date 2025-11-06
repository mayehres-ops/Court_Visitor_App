# -*- coding: utf-8 -*-
"""
DocStrange cloud extractor -> spreadsheet-shaped CSV for ARP + Order.

- Uses JSON Schema extraction (more reliable than 'specified_fields')
- Skips "approval" PDFs
- For ARP: pulls info commonly found on pages 1–2
- For Order: pulls case # and signed date
- Adds PDF-text fallbacks for case number and FILED date
- Writes one CSV with the exact columns in 'ward_guardian_info2test.xlsx'
"""

import os, re, json, csv, sys, datetime
from pathlib import Path

# ---------- YOUR FOLDERS ----------
INPUT_DIR = r"C:\GoogleSync\Guardianship Files\aa New"
OUT_DIR   = r"C:\GoogleSync\Guardianship Files\Extracted"
CSV_PATH  = Path(OUT_DIR) / "ward_guardian_info_out.csv"

# ---------- ENV KEY LOADER ----------
def load_api_key() -> str:
    # Prefer file path env (what you already use)
    key_file = os.environ.get("DOCSTRANGE_API_KEY_FILE")
    if key_file and Path(key_file).exists():
        return Path(key_file).read_text(encoding="utf-8").strip()
    # Or raw env var
    return os.environ.get("DOCSTRANGE_API_KEY", "").strip()

# ---------- SAFE IMPORTS ----------
try:
    from docstrange import DocumentExtractor
except Exception as e:
    print("DocStrange import failed:", e)
    sys.exit(1)

# PyMuPDF (fitz) is already a dependency of docstrange; used for fallbacks.
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

# ---------- JSON SCHEMA ----------
# Keep this flat to match your sheet; DocStrange will fill strings/booleans/numbers.
# (We collect a little more than we write so we can derive liveswith, etc.)
JSON_SCHEMA = {
    # Header / identifiers
    "case_number": "string",                              # e.g., "C-1-PB-19-000694"
    "probate_court_no": "string",
    "county": "string",
    "state": "string",

    # File stamps
    "filed_for_record_date": "string",                    # stamped date (preferred)
    "filed_on_date": "string",                            # sometimes printed as "Filed:"
    "updated_date_note": "string",                        # ignore for DateARPfiled

    # Ward block
    "ward_name": "string",
    "ward_phone": "string",
    "ward_address_no_po_box": "string",
    "ward_city_state_zip": "string",
    "ward_dob": "string",

    # Guardian 1 block (first listed)
    "guardian1_name": "string",
    "guardian1_age": "string",
    "guardian1_dob": "string",
    "guardian1_email": "string",
    "guardian1_address_no_po_box": "string",
    "guardian1_city_state_zip": "string",
    "guardian1_phone": "string",
    "guardian1_relationship_to_ward": "string",

    # Guardian 2 block (co-guardian if present)
    "guardian2_name": "string",
    "guardian2_age": "string",
    "guardian2_dob": "string",
    "guardian2_email": "string",
    "guardian2_address_no_po_box": "string",
    "guardian2_city_state_zip": "string",
    "guardian2_phone": "string",
    "guardian2_relationship_to_ward": "string",

    # ARP Q4 (reside w/ ward)
    "guardian_resides_with_ward_yes": "boolean",
    "guardian_resides_with_ward_no": "boolean",

    # Page 2 #5 residence choices
    "residence_type_checked": "string",                   # e.g., "Guardian's home", "Ward's home", etc.
    "residence_facility_name": "string",

    # ORDER page items
    "order_signed_on_date": "string",                     # "Signed on:" on the Order
}

# ---------- SHEET COLUMNS (exact from your workbook) ----------
SHEET_COLUMNS = [
    "wlast","wfirst","causeno","visitdate","visittime","wtele","liveswith",
    "waddress","wcitystatezip","wob",
    "guardian1","gaddress","gcitystatezip","gemail","gtele","Relationship","gdob",
    "Guardian2","g2address","g2citystatezip","g2email","g2tele","g2Relationship",
    "datesubmited","Dateappointed","expensesubmited","expenspd","DateARPfiled",
    "last_updated","Comments",
]

# ---------- HELPERS ----------
NAME_SPLIT_RE = re.compile(r"^\s*(.+?)\s+([A-Za-z'’-]+)$")

def split_name(full: str):
    if not full:
        return ("","")
    s = " ".join(full.split())
    m = NAME_SPLIT_RE.match(s)
    if m:
        return (m.group(1), m.group(2))
    # Fallback: first token = first name(s) except last token
    parts = s.split()
    if len(parts) >= 2:
        return (" ".join(parts[:-1]), parts[-1])
    return (s, "")

def normalize_phone(s: str) -> str:
    if not s: return ""
    digits = re.sub(r"\D+", "", s)
    if len(digits) == 10:
        return f"({digits[0:3]}) {digits[3:6]}-{digits[6:]}"
    if len(digits) == 7:
        return f"{digits[0:3]}-{digits[3:]}"
    return s.strip()

def text_from_pdf_first_pages(pdf_path: Path, max_pages=2) -> str:
    if not fitz:
        return ""
    try:
        with fitz.open(str(pdf_path)) as doc:
            pages = min(max_pages, len(doc))
            return "\n".join(doc[i].get_text() for i in range(pages))
    except Exception:
        return ""

CASE_RE = re.compile(r"(C-1-PB-\s*\d{2}\s*-\s*\d{6})")
STAMP_RE = re.compile(
    r"(FILED\s+FOR\s+RECORD|Filed\s*:|Filed\s+on\s*:)\s*([A-Z][a-z]{2,}\s+\d{1,2},?\s+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4}|[A-Z]{3}\s+\d{1,2}\s+\d{2,4})",
    re.IGNORECASE
)

def derive_liveswith(d: dict) -> str:
    # Priority: explicit checkbox Y/N on ARP Q4
    if d.get("guardian_resides_with_ward_yes") is True:
        return "guardian"
    if d.get("guardian_resides_with_ward_no") is True:
        # fall through to residence_type if available
        pass
    t = (d.get("residence_type_checked") or "").lower()
    if "guardian" in t:
        return "guardian"
    if "ward" in t:
        return "ward"
    if "foster" in t:
        return "foster"
    if "group home" in t or "boarding" in t or "hospital" in t or "facility" in t or "state supported" in t:
        return "facility"
    return ""

def pick_first(*vals):
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""

def to_row(d: dict, pdf_name: str, is_order: bool) -> dict:
    # Ward
    wfirst, wlast = split_name(d.get("ward_name",""))
    # Guardian 1/2 strings
    g1 = d.get("guardian1_name","").strip()
    g2 = d.get("guardian2_name","").strip()

    row = {k:"" for k in SHEET_COLUMNS}
    row.update({
        "wfirst": wfirst, "wlast": wlast,
        "wtele": normalize_phone(d.get("ward_phone","")),
        "waddress": d.get("ward_address_no_po_box",""),
        "wcitystatezip": d.get("ward_city_state_zip",""),
        "wob": d.get("ward_dob",""),
        "guardian1": g1,
        "gaddress": d.get("guardian1_address_no_po_box",""),
        "gcitystatezip": d.get("guardian1_city_state_zip",""),
        "gemail": d.get("guardian1_email",""),
        "gtele": normalize_phone(d.get("guardian1_phone","")),
        "Relationship": d.get("guardian1_relationship_to_ward",""),
        "gdob": d.get("guardian1_dob",""),
        "Guardian2": g2,
        "g2address": d.get("guardian2_address_no_po_box",""),
        "g2citystatezip": d.get("guardian2_city_state_zip",""),
        "g2email": d.get("guardian2_email",""),
        "g2tele": normalize_phone(d.get("guardian2_phone","")),
        "g2Relationship": d.get("guardian2_relationship_to_ward",""),
        "last_updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Comments": f"src={pdf_name}",
    })

    # Case number + filed date (ARP) / signed date (Order)
    row["causeno"] = d.get("case_number","")

    # Filed for record date (ARP page stamp), prefer explicit fields; fall back to PDF text grep
    row["DateARPfiled"] = pick_first(d.get("filed_for_record_date"), d.get("filed_on_date"))

    # Optional dates the sheet includes but we don't extract here
    row["visitdate"] = ""
    row["visittime"] = ""
    row["datesubmited"] = ""        # your sheet label spelling
    row["Dateappointed"] = d.get("order_signed_on_date","") if is_order else ""
    row["expensesubmited"] = ""
    row["expenspd"] = ""

    # liveswith inference (guardian/ward/facility/…)
    row["liveswith"] = derive_liveswith(d)

    return row

def ensure_csv_header():
    CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not CSV_PATH.exists():
        with open(CSV_PATH, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=SHEET_COLUMNS)
            w.writeheader()

def append_row(row: dict):
    with open(CSV_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=SHEET_COLUMNS)
        w.writerow(row)

def fallback_enrich_from_pdftext(pdf_path: Path, row: dict, is_order: bool):
    txt = text_from_pdf_first_pages(pdf_path, max_pages=1 if is_order else 2)
    if not txt:
        return
    if not row.get("causeno"):
        m = CASE_RE.search(txt)
        if m:
            row["causeno"] = m.group(1).replace(" ", "")
    if not row.get("DateARPfiled"):
        m = STAMP_RE.search(txt)
        if m:
            row["DateARPfiled"] = m.group(2)

def main():
    api_key = load_api_key()
    if not api_key:
        print("No API key. Set DOCSTRANGE_API_KEY or DOCSTRANGE_API_KEY_FILE.")
        sys.exit(2)

    extractor = DocumentExtractor(api_key=api_key)  # CLOUD
    inp = Path(INPUT_DIR)
    out = Path(OUT_DIR)
    out.mkdir(parents=True, exist_ok=True)

    ensure_csv_header()

    pdfs = sorted([p for p in inp.glob("*.pdf") if p.is_file()])
    if not pdfs:
        print(f"No PDFs in {inp}")
        return

    for pdf in pdfs:
        stem_lower = pdf.stem.lower()
        if "approval" in stem_lower:
            print(f"⏭️  Skipping {pdf.name} (approval)")
            continue

        is_order = "order" in stem_lower
        doc_type = "Order" if is_order else "ARP"
        print(f"☁️  Cloud extracting: {pdf.name} [{doc_type}]")

        try:
            result = extractor.extract(str(pdf))
            data = result.extract_data(json_schema=JSON_SCHEMA)  # <- JSON schema mode
            # Save raw structured fields (for troubleshooting)
            fields_path = out / f"{pdf.stem}.cloud.fields.json"
            fields_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

            # Flatten the actual values dict
            extracted = data.get("structured_data") or data.get("extracted_fields") or data

            row = to_row(extracted, pdf.name, is_order=is_order)
            # PDF-text fallbacks for case # and filed date
            fallback_enrich_from_pdftext(pdf, row, is_order=is_order)

            append_row(row)
            print(f"   ✅ wrote {fields_path.name} and appended to {CSV_PATH.name}")

        except Exception as e:
            print(f"   ✗ Error on {pdf.name}: {e}")

if __name__ == "__main__":
    main()
