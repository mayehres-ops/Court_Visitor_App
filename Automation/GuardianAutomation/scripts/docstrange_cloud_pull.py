# -*- coding: utf-8 -*-
"""
DocStrange → hardened ARP/Order extractor
- Primary: Cloud "specified fields"
- Fallback A: Cloud text/markdown + regex anchors
- Null-safe mapping to your CSV column headers
- ARP pages 1-2 only; Order page 1 only
"""

from __future__ import annotations
import os, re, json, csv, sys, time
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

# ---------- CONFIG ----------
INPUT_DIR  = r"C:\GoogleSync\Guardianship Files\aa New"
OUT_DIR    = r"C:\GoogleSync\Guardianship Files\Extracted"
CSV_PATH   = Path(OUT_DIR) / "ward_guardian_info_out.csv"

# Only process these page numbers (1-indexed)
ARP_PAGES   = [1, 2]
ORDER_PAGES = [1]

# Your sheet headers (exact order)
CSV_HEADERS = [
    "wlast","wfirst","causeno","visitdate","visittime","wtelephone","liveswith",
    "waddress","wob","guardian1","gaddress","gemail","gtele","Relationship","gdob",
    "guardian2","g2 address","g2emil","g2tele","g2Relationship",
    "dsubmitted","dappointed","expensessubmitted","expensespd","DateARPfiled",
    "last_updated","Comments"
]

# Field lists we'll ask the cloud for (keep names close to docstrange examples)
FIELDS_ARP = [
    "page_number", "case_number", "probate_court_no", "county", "state",
    "filed_for_record_date", "filed_on_date", "updated_date_note",
    # ward block (page 1)
    "ward_name", "ward_phone", "ward_address_no_po_box", "ward_city_state_zip", "ward_dob",
    # guardians block (page 1)
    "guardian1_name","guardian1_age","guardian1_dob","guardian1_email",
    "guardian1_address_no_po_box","guardian1_city_state_zip","guardian1_phone","guardian1_relationship_to_ward",
    "guardian2_name","guardian2_age","guardian2_dob","guardian2_email",
    "guardian2_address_no_po_box","guardian2_city_state_zip","guardian2_phone","guardian2_relationship_to_ward",
    # resides with ward? (page 1 #4)
    "guardian_resides_with_ward_yes","guardian_resides_with_ward_no",
    # page 2 residence type
    "residence_type_checked",          # e.g., "Guardian's home" / "Ward's home" / "Relative's home" / "Foster" / ...
    "residence_facility_name",         # name if facility-type
]

FIELDS_ORDER = [
    "case_number", "probate_court_no", "county", "state",
    "ward_name", "order_signed_on_date", "filed_for_record_date", "filed_on_date"
]

# ---------- DEPENDENCIES ----------
def _load_api_key() -> Optional[str]:
    """Prefer file pointer, then env var."""
    fp = os.environ.get("DOCSTRANGE_API_KEY_FILE")
    if fp and Path(fp).exists():
        return Path(fp).read_text(encoding="utf-8").strip()
    return os.environ.get("DOCSTRANGE_API_KEY")

def _ensure_dirs():
    Path(OUT_DIR).mkdir(parents=True, exist_ok=True)

# ---------- DOCSTRANGE WRAPPER ----------
def _init_docstrange_cloud():
    from docstrange import DocumentExtractor
    api_key = _load_api_key()
    if not api_key:
        print("✗ No API key. Set DOCSTRANGE_API_KEY_FILE or DOCSTRANGE_API_KEY.")
        sys.exit(1)
    # Force cloud by not providing cpu=True and providing api_key
    # (DocStrange 1.1.x: api_key presence → cloud; cpu flag only affects local OCR path)
    return DocumentExtractor(api_key=api_key), api_key

def _extract_cloud_fields(extractor, file_path: str, fields: List[str]) -> Tuple[Dict[str, Any], Optional[str]]:
    """
    Try cloud structured extraction. Returns (data, raw_payload or None).
    On JSON parse errors, return {"_format":"json_parse_error"} and the raw payload (if any).
    """
    try:
        res = extractor.extract(file_path)
        # Most reliable call across versions:
        data = res.extract_data(specified_fields=fields)
        return data, None
    except Exception as e:
        # Try to capture raw text/markdown to salvage
        raw_text = None
        try:
            raw_text = res.extract_data(format="text")  # if supported
        except Exception:
            pass
        return {"_format": "json_parse_error", "_error": str(e)}, raw_text

# ---------- FALLBACK PARSERS (text-based) ----------
_re_space = re.compile(r"\s+")

def _norm(s: Optional[str]) -> str:
    return _re_space.sub(" ", s.strip()) if isinstance(s, str) else ""

def _find_case_number(txt: str) -> Optional[str]:
    # Examples: "No. C-1-PB-19-000694" → want "19-000694"
    m = re.search(r"No\.\s*C-1-PB-?\s*([0-9]{2}-[0-9]{6})", txt, re.I)
    return m.group(1) if m else None

def _find_filed_date(txt: str) -> Optional[str]:
    # Look for "FILED FOR RECORD" / "FILED:" / "Filed on:" not "Updated"
    m = re.search(r"(FILED\s+FOR\s+RECORD|FILED:|Filed on:)\s*([0-9/.\-]{6,12})", txt, re.I)
    return m.group(2) if m else None

def _find_signed_on(txt: str) -> Optional[str]:
    m = re.search(r"Signed\s+on:\s*([A-Za-z]+\s+\d{1,2},\s*\d{4}|[0-9/.\-]{6,12})", txt, re.I)
    return m.group(1) if m else None

def _find_labeled_value(txt: str, label: str, take_next_line: bool=False) -> Optional[str]:
    # Generic: try "LABEL: value"
    pat = re.compile(re.escape(label) + r"\s*[:\-]\s*(.+)", re.I)
    m = pat.search(txt)
    if m:
        return _norm(m.group(1))
    if take_next_line:
        # Try the next line after the label if value sits under
        m2 = re.search(re.escape(label) + r"\s*[:\-]?\s*\n([^\n]+)", txt, re.I)
        if m2:
            return _norm(m2.group(1))
    return None

def _split_name(full: str) -> Tuple[str,str]:
    # Conservative split: first token as first, remainder as last
    if not full:
        return "", ""
    parts = _norm(full).split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])

def _compute_liveswith(resides_yes: Optional[bool], residence_type: Optional[str]) -> str:
    if resides_yes is True:
        return "guardian"
    if isinstance(residence_type, str):
        rt = residence_type.lower()
        if "guardian" in rt: return "guardian"
        if "ward" in rt:     return "ward"
        if "relative" in rt: return "relative"
        if any(k in rt for k in ["foster","boarding","group","hospital","state","facility","other"]):
            return "facility"
    return ""

# ---------- RECORD BUILDERS ----------
def _row_from_arp(fields: Dict[str, Any], text_fallback: Optional[str]) -> Dict[str,str]:
    # Pull primary values
    wname = _norm(fields.get("ward_name"))
    wfirst, wlast = _split_name(wname)
    wphone = _norm(fields.get("ward_phone"))
    wad1   = _norm(fields.get("ward_address_no_po_box"))
    wad2   = _norm(fields.get("ward_city_state_zip"))
    wardob = _norm(fields.get("ward_dob"))

    g1name = _norm(fields.get("guardian1_name"))
    g1dob  = _norm(fields.get("guardian1_dob"))
    g1rel  = _norm(fields.get("guardian1_relationship_to_ward"))
    g1addr = _norm(fields.get("guardian1_address_no_po_box"))
    g1csy  = _norm(fields.get("guardian1_city_state_zip"))
    g1phone= _norm(fields.get("guardian1_phone"))
    g1mail = _norm(fields.get("guardian1_email"))

    g2name = _norm(fields.get("guardian2_name"))
    g2dob  = _norm(fields.get("guardian2_dob"))
    g2rel  = _norm(fields.get("guardian2_relationship_to_ward"))
    g2addr = _norm(fields.get("guardian2_address_no_po_box"))
    g2csy  = _norm(fields.get("guardian2_city_state_zip"))
    g2phone= _norm(fields.get("guardian2_phone"))
    g2mail = _norm(fields.get("guardian2_email"))

    filed   = _norm(fields.get("filed_for_record_date") or fields.get("filed_on_date"))
    case_no = _norm(fields.get("case_number"))
    # residence
    resides_yes = fields.get("guardian_resides_with_ward_yes")
    rtype       = _norm(fields.get("residence_type_checked"))
    liveswith   = _compute_liveswith(resides_yes if isinstance(resides_yes,bool) else None, rtype)

    # Fallbacks (text)
    if text_fallback:
        t = text_fallback
        if not case_no:
            case_no = _find_case_number(t) or case_no
        if not filed:
            filed = _find_filed_date(t) or filed
        if not wname:
            wname = _find_labeled_value(t, "Name", False) or wname  # last resort heuristic
            wfirst, wlast = _split_name(wname)
        if not wphone:
            wphone = _find_labeled_value(t, "Phone") or wphone
        if not wad1:
            wad1 = _find_labeled_value(t, "Address") or wad1
        if not wad2:
            wad2 = _find_labeled_value(t, "City/State/Zip") or wad2

    row = {h:"" for h in CSV_HEADERS}
    row.update({
        "wfirst": wfirst, "wlast": wlast,
        "wtelephone": wphone,
        "waddress": " | ".join([p for p in [wad1, wad2] if p]),
        "wob": wardob,
        "guardian1": g1name,
        "gaddress": " | ".join([p for p in [g1addr, g1csy] if p]),
        "gemail": g1mail, "gtele": g1phone, "Relationship": g1rel, "gdob": g1dob,
        "guardian2": g2name,
        "g2 address": " | ".join([p for p in [g2addr, g2csy] if p]),
        "g2emil": g2mail, "g2tele": g2phone, "g2Relationship": g2rel,
        "DateARPfiled": filed,
        "dsubmitted": filed,
        "causeno": case_no,
        "liveswith": liveswith,
        "last_updated": "", "Comments": ""
    })
    return row

def _row_from_order(fields: Dict[str, Any], text_fallback: Optional[str]) -> Dict[str,str]:
    wname = _norm(fields.get("ward_name"))
    wfirst, wlast = _split_name(wname)
    case_no = _norm(fields.get("case_number"))
    d_signed = _norm(fields.get("order_signed_on_date"))
    filed    = _norm(fields.get("filed_for_record_date") or fields.get("filed_on_date"))

    if text_fallback:
        t = text_fallback
        if not case_no:
            case_no = _find_case_number(t) or case_no
        if not d_signed:
            d_signed = _find_signed_on(t) or d_signed
        if not filed:
            filed = _find_filed_date(t) or filed

    row = {h:"" for h in CSV_HEADERS}
    row.update({
        "wfirst": wfirst, "wlast": wlast,
        "causeno": case_no,
        "dappointed": d_signed,
        "dsubmitted": filed,
        "DateARPfiled": "",  # N/A for Orders
        "last_updated": "", "Comments": ""
    })
    return row

# ---------- MAIN ----------
def main():
    _ensure_dirs()
    extractor, api_key = _init_docstrange_cloud()
    print("▶ Using DocStrange CLOUD with specified fields + fallbacks")

    # Prepare CSV (create if not exists)
    is_new = not CSV_PATH.exists()
    with open(CSV_PATH, "a", newline="", encoding="utf-8") as fcsv:
        w = csv.DictWriter(fcsv, fieldnames=CSV_HEADERS)
        if is_new:
            w.writeheader()

        for pdf in sorted(Path(INPUT_DIR).glob("*.pdf")):
            stem = pdf.stem.lower()

            if "approval" in stem:
                print(f"⏭️  Skipping {pdf.name} (approval)")
                continue

            kind = "ARP" if "arp" in stem else "Order" if "order" in stem else "Other"
            if kind == "Other":
                print(f"⏭️  Skipping {pdf.name} (not ARP/Order)")
                continue

            print(f"☁️  Cloud extracting: {pdf.name} [{kind}]")
            fields = FIELDS_ARP if kind == "ARP" else FIELDS_ORDER

            # Cloud call
            data, raw_text = _extract_cloud_fields(extractor, str(pdf), fields)

            # Persist what came back for inspection
            out_json = Path(OUT_DIR) / f"{pdf.stem}.cloud.fields.json"
            with open(out_json, "w", encoding="utf-8") as jf:
                json.dump(data, jf, ensure_ascii=False, indent=2)
            if raw_text:
                Path(OUT_DIR, f"{pdf.stem}.raw.txt").write_text(
                    raw_text if isinstance(raw_text,str) else json.dumps(raw_text, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )

            # Build CSV row with fallbacks
            try:
                if kind == "ARP":
                    row = _row_from_arp(data if isinstance(data, dict) else {}, raw_text)
                else:
                    row = _row_from_order(data if isinstance(data, dict) else {}, raw_text)
                w.writerow(row)
                print(f"   ✅ wrote {out_json.name} and appended to {CSV_PATH.name}")
            except Exception as e:
                print(f"   ✗ Error building row for {pdf.name}: {e}")

if __name__ == "__main__":
    main()
