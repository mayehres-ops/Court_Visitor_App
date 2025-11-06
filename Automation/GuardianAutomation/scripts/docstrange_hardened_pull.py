# -*- coding: utf-8 -*-
from __future__ import annotations
import os, re, json, csv, sys
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

INPUT_DIR  = r"C:\GoogleSync\Guardianship Files\aa New"
OUT_DIR    = r"C:\GoogleSync\Guardianship Files\Extracted"
CSV_PATH   = Path(OUT_DIR) / "ward_guardian_info_out.csv"

CSV_HEADERS = [
    "wlast","wfirst","causeno","visitdate","visittime","wtelephone","liveswith",
    "waddress","wob","guardian1","gaddress","gemail","gtele","Relationship","gdob",
    "guardian2","g2 address","g2emil","g2tele","g2Relationship",
    "dsubmitted","dappointed","expensessubmitted","expensespd","DateARPfiled",
    "last_updated","Comments"
]

FIELDS_ARP = [
    "page_number", "case_number", "probate_court_no", "county", "state",
    "filed_for_record_date", "filed_on_date", "updated_date_note",
    "ward_name", "ward_phone", "ward_address_no_po_box", "ward_city_state_zip", "ward_dob",
    "guardian1_name","guardian1_age","guardian1_dob","guardian1_email",
    "guardian1_address_no_po_box","guardian1_city_state_zip","guardian1_phone","guardian1_relationship_to_ward",
    "guardian2_name","guardian2_age","guardian2_dob","guardian2_email",
    "guardian2_address_no_po_box","guardian2_city_state_zip","guardian2_phone","guardian2_relationship_to_ward",
    "guardian_resides_with_ward_yes","guardian_resides_with_ward_no",
    "residence_type_checked","residence_facility_name",
]

FIELDS_ORDER = [
    "case_number", "probate_court_no", "county", "state",
    "ward_name", "order_signed_on_date", "filed_for_record_date", "filed_on_date"
]

def _load_api_key() -> Optional[str]:
    fp = os.environ.get("DOCSTRANGE_API_KEY_FILE")
    if fp and Path(fp).exists():
        return Path(fp).read_text(encoding="utf-8").strip()
    return os.environ.get("DOCSTRANGE_API_KEY")

def _ensure_dirs():
    Path(OUT_DIR).mkdir(parents=True, exist_ok=True)

def _init_docstrange_cloud():
    from docstrange import DocumentExtractor
    api_key = _load_api_key()
    if not api_key:
        print("✗ No API key. Set DOCSTRANGE_API_KEY_FILE or DOCSTRANGE_API_KEY.")
        sys.exit(1)
    return DocumentExtractor(api_key=api_key)

def _extract_cloud_fields(extractor, file_path: str, fields: List[str]) -> Tuple[Dict[str, Any], Optional[str]]:
    try:
        res = extractor.extract(file_path)
        data = res.extract_data(specified_fields=fields)
        return data if isinstance(data, dict) else {}, None
    except Exception as e:
        # try to salvage plain text for regex fallbacks
        try:
            text = res.extract_data(format="text")
        except Exception:
            text = None
        return {"_format":"json_parse_error","_error":str(e)}, (text if isinstance(text,str) else None)

# --------- text helpers ----------
_re_space = re.compile(r"\s+")
def _norm(s: Optional[str]) -> str:
    return _re_space.sub(" ", s.strip()) if isinstance(s, str) else ""

def _split_name(full: str):
    if not full: return "",""
    parts = _norm(full).split()
    if len(parts)==1: return parts[0],""
    return parts[0], " ".join(parts[1:])

def _find_case_number(txt: str) -> Optional[str]:
    m = re.search(r"No\.\s*C-1-PB-?\s*([0-9]{2}-[0-9]{6})", txt, re.I)
    return m.group(1) if m else None

def _find_filed_date(txt: str) -> Optional[str]:
    m = re.search(r"(FILED\s+FOR\s+RECORD|FILED:|Filed on:)\s*([0-9/.\-]{6,12})", txt, re.I)
    return m.group(2) if m else None

def _find_signed_on(txt: str) -> Optional[str]:
    m = re.search(r"Signed\s+on:\s*([A-Za-z]+\s+\d{1,2},\s*\d{4}|[0-9/.\-]{6,12})", txt, re.I)
    return m.group(1) if m else None

def _find_labeled_value(txt: str, label: str) -> Optional[str]:
    m = re.search(re.escape(label)+r"\s*[:\-]\s*(.+)", txt, re.I)
    return _norm(m.group(1)) if m else None

def _compute_liveswith(resides_yes: Optional[bool], residence_type: Optional[str]) -> str:
    if resides_yes is True: return "guardian"
    rt = (residence_type or "").lower()
    if "guardian" in rt: return "guardian"
    if "ward" in rt:     return "ward"
    if "relative" in rt: return "relative"
    if any(k in rt for k in ["foster","boarding","group","hospital","state","facility","other"]):
        return "facility"
    return ""

# --------- row builders ----------
def _row_from_arp(fields: Dict[str, Any], text_fallback: Optional[str]) -> Dict[str,str]:
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

    resides_yes = fields.get("guardian_resides_with_ward_yes")
    rtype       = _norm(fields.get("residence_type_checked"))
    liveswith   = _compute_liveswith(resides_yes if isinstance(resides_yes,bool) else None, rtype)

    if text_fallback:
        t = text_fallback
        case_no = case_no or _find_case_number(t)
        filed   = filed   or _find_filed_date(t)
        wphone  = wphone  or _find_labeled_value(t,"Phone")
        wad1    = wad1    or _find_labeled_value(t,"Address")
        wad2    = wad2    or _find_labeled_value(t,"City/State/Zip")

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
        case_no = case_no or _find_case_number(t)
        d_signed= d_signed or _find_signed_on(t)
        filed   = filed or _find_filed_date(t)

    row = {h:"" for h in CSV_HEADERS}
    row.update({
        "wfirst": wfirst, "wlast": wlast,
        "causeno": case_no,
        "dappointed": d_signed,
        "dsubmitted": filed,
        "DateARPfiled": "",
        "last_updated": "", "Comments": ""
    })
    return row

def main():
    _ensure_dirs()
    ex = _init_docstrange_cloud()
    print("▶ Using DocStrange CLOUD with fallbacks")

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
            fields = FIELDS_ARP if kind=="ARP" else FIELDS_ORDER
            data, raw_text = _extract_cloud_fields(ex, str(pdf), fields)

            # persist responses for debugging
            (Path(OUT_DIR)/f"{pdf.stem}.cloud.fields.json").write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            if raw_text:
                (Path(OUT_DIR)/f"{pdf.stem}.raw.txt").write_text(raw_text, encoding="utf-8")

            try:
                row = _row_from_arp(data, raw_text) if kind=="ARP" else _row_from_order(data, raw_text)
                w.writerow(row)
                print(f"   ✅ wrote {pdf.stem}.cloud.fields.json and appended to {CSV_PATH.name}")
            except Exception as e:
                print(f"   ✗ Error building row for {pdf.name}: {e}")

if __name__ == "__main__":
    main()
