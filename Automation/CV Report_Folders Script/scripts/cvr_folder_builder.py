#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
CVR Folder Builder - Creates client folders from New Files and moves to New Clients

WORKFLOW:
1. Reads PDFs from "New Files" folder (ARPs, Orders after OCR extraction)
2. Groups files by cause number
3. Creates folder structure: "LastName, FirstName - CauseNo"
4. Moves organized folders to "New Clients" (ready for CVR creation)
5. Never touches "Completed" folder

- Buckets by trailing number ONLY: "", "1", "2", ...
  => ARP, Order, Approval with same number are grouped.
- If any file in a bucket yields a cause, all in that bucket inherit it.
- If a bucket still lacks a cause, try ward last-name match from Excel on ARP text.
- Matches C-?-PB-YY-XXXXXX to YY-XXXXXX by stripping Travis prefix.
"""

import argparse, re, sys, shutil, time, logging
from pathlib import Path

# Dynamic path management - works from any installation location
_script_dir = Path(__file__).parent.parent.parent.parent  # Go up to app root

try:
    # Try to import app_paths for dynamic path detection
    sys.path.insert(0, str(_script_dir / "Scripts"))
    from app_paths import get_app_paths

    # Pass the detected app root explicitly
    _app_paths = get_app_paths(str(_script_dir))

    WORKBOOK_PATH = _app_paths.EXCEL_PATH
    INBOX_DIR     = _app_paths.NEW_FILES_DIR
    GUARDIAN_BASE = _app_paths.NEW_CLIENTS_DIR
    COMPLETED_DIR = _app_paths.COMPLETED_DIR
except Exception:
    # Fallback to hardcoded paths if app_paths not available
    WORKBOOK_PATH = Path(r"C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx")
    INBOX_DIR     = Path(r"C:\GoogleSync\GuardianShip_App\New Files")        # Read PDFs from here
    GUARDIAN_BASE = Path(r"C:\GoogleSync\GuardianShip_App\New Clients")      # Create case folders here
    COMPLETED_DIR = Path(r"C:\GoogleSync\GuardianShip_App\Completed")        # Reference only - don't modify

TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # or None
POPPLER_BIN   = r"C:\poppler\Library\bin"                        # or None

EXCEL_COLUMNS = ["wardlast", "wardfirst", "wardmiddle", "causeno"]
UNMATCHED_DIR_NAME = "_Unmatched"


def _init_ocr_paths():
    try:
        import pytesseract  # noqa
        if TESSERACT_EXE:
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
    except Exception:
        pass

def sanitize_for_fs(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name).strip()

CAUSE_PATTERNS = [
    r"\b[A-Z]-\d-[A-Z]{2}-\d{2}-\d{5,6}\b",  # C-1-PB-25-000123
    r"\b\d{2}-\d{5,6}\b",                    # 21-001411
    r"\b\d{5,8}\b",                          # fallback numeric run
]
NEAR_CAUSE_HINTS = [r"cause\s*no\.?", r"cause\s*number", r"case\s*no\.?", r"case\s*number", r"cause", r"case"]

def strip_travis_prefix(s: str) -> str:
    return re.sub(r"^[A-Z]-\d-[A-Z]{2}-", "", s or "", flags=re.IGNORECASE)

def normalize_alnum(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", s or "").upper()

def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def causes_equal(a: str, b: str) -> bool:
    if not a or not b: return False
    a2, b2 = strip_travis_prefix(a), strip_travis_prefix(b)
    return (normalize_alnum(a2) == normalize_alnum(b2)) or (normalize_digits(a2) == normalize_digits(b2))

def find_cause_number_in_text(text: str) -> str | None:
    if not text: return None
    lowered = text.lower()
    windows = []
    for hint in NEAR_CAUSE_HINTS:
        for m in re.finditer(hint, lowered, flags=re.IGNORECASE):
            start = max(0, m.start()-50); end = min(len(text), m.end()+50)
            windows.append(text[start:end])
    for chunk in windows + [text]:
        for pat in CAUSE_PATTERNS:
            m = re.search(pat, chunk, flags=re.IGNORECASE)
            if m: return m.group(0)
    return None

def extract_text_pdfplumber(pdf_path: Path, max_pages: int = 3) -> str:
    try:
        import pdfplumber
    except ImportError:
        return ""
    try:
        parts = []
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages[:max_pages]:
                parts.append(page.extract_text() or "")
        return "\n".join(parts)
    except Exception:
        return ""

def ocr_first_pages(pdf_path: Path, max_pages: int = 2) -> str:
    try:
        from pdf2image import convert_from_path
        import pytesseract
    except ImportError:
        return ""
    try:
        images = convert_from_path(str(pdf_path), first_page=1, last_page=max_pages, dpi=300, fmt="png",
                                   poppler_path=POPPLER_BIN if POPPLER_BIN else None)
    except Exception:
        return ""
    texts = []
    for im in images:
        try:
            t = pytesseract.image_to_string(im)
            if t: texts.append(t)
        except Exception:
            continue
    return "\n".join(texts)

def extract_text_any(pdf_path: Path) -> str:
    """Combine fast text and (if needed) OCR for a single best-effort text string."""
    txt = extract_text_pdfplumber(pdf_path, max_pages=3)
    if txt: return txt
    _init_ocr_paths()
    return ocr_first_pages(pdf_path, max_pages=2)

def extract_cause_number(pdf_path: Path) -> str | None:
    txt = extract_text_any(pdf_path)
    return find_cause_number_in_text(txt)

def build_mapping_from_excel(xlsx_path: Path) -> tuple[dict[str, str], dict[str, list[str]]]:
    """Returns:
       - cause_to_foldername: causeno -> base folder name
       - lastname_to_causes: lowercase last name -> list of causes (to check ambiguity)
    """
    import pandas as pd
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    df.columns = [str(c).strip().lower() for c in df.columns]
    for col in EXCEL_COLUMNS:
        if col not in df.columns:
            raise ValueError(f"Excel missing required column '{col}'. Found: {list(df.columns)}")

    cause_to_foldername: dict[str, str] = {}
    lastname_to_causes: dict[str, list[str]] = {}
    for _, row in df.iterrows():
        last  = "" if pd.isna(row.get("wardlast"))   else str(row.get("wardlast")).strip()
        first = "" if pd.isna(row.get("wardfirst"))  else str(row.get("wardfirst")).strip()
        mid   = "" if pd.isna(row.get("wardmiddle")) else str(row.get("wardmiddle")).strip()
        cause = "" if pd.isna(row.get("causeno"))    else str(row.get("causeno")).strip()
        if not cause: continue
        midsp = f" {mid}" if mid else ""
        folder_name = sanitize_for_fs(f"{last}, {first}{midsp} - {cause}")
        cause_to_foldername[cause] = folder_name
        if last:
            lastname_to_causes.setdefault(last.lower(), []).append(cause)
    return cause_to_foldername, lastname_to_causes

def scan_cause_folders(parent: Path) -> dict[str, Path]:
    out: dict[str, Path] = {}
    if not parent.exists(): return out
    for child in parent.iterdir():
        if not child.is_dir(): continue
        name = child.name
        found = None
        for pat in CAUSE_PATTERNS:
            m = re.search(pat, name, flags=re.IGNORECASE)
            if m: found = m.group(0); break
        if not found:
            m2 = re.search(r"-\s*([A-Z]-\d-[A-Z]{2}-\d{2}-\d{5,6}|\d{2}-\d{5,6}|\d{5,8})$", name, flags=re.IGNORECASE)
            if m2: found = m2.group(1)
        if found:
            norm = normalize_alnum(strip_travis_prefix(found))
            if norm not in {normalize_alnum(strip_travis_prefix(k)) for k in out.keys()}:
                out[found] = child
    return out

def ensure_folder(path: Path):
    path.mkdir(parents=True, exist_ok=True)

# ---------- Bucketing (by number only) ----------
DOC_RE = re.compile(r"(?i)^(arp|order|approval)[ _-]?(\d+)?(?:\.pdf)?$")

def get_bucket_num_key(p: Path) -> str:
    """
    Returns bucket number key: '' for unnumbered, '1' for *_1, etc.
    Files that don't match known patterns get a unique key so they don't mingle.
    """
    m = DOC_RE.match(p.stem)
    if not m:
        return f"misc::{p.stem.lower()}"
    return m.group(2) or ""  # '' = unnumbered

def deduce_doc_tag(filename: str) -> str:
    m = DOC_RE.match(Path(filename).stem)
    if m:
        typ = m.group(1).upper()
        num = m.group(2) or ""
        return f"{typ}{num}"
    return re.sub(r"[^A-Za-z0-9_-]+", "", Path(filename).stem)[:24] or "DOC"

# ---------- Matching helpers ----------
def resolve_match_in_map(cause: str, mp: dict[str, Path]) -> Path | None:
    for k, p in mp.items():
        if causes_equal(cause, k):
            return p
    return None

def resolve_excel_folder_name(cause: str, excel_map: dict[str, str]) -> str | None:
    for k, name in excel_map.items():
        if causes_equal(cause, k):
            return name
    return None

def unique_destination(dst_folder: Path, filename: str) -> Path:
    base = Path(filename).stem; ext = Path(filename).suffix
    candidate = dst_folder / f"{base}{ext}"
    n = 1
    while candidate.exists():
        candidate = dst_folder / f"{base} ({n}){ext}"
        n += 1
    return candidate

# ---------- Main ----------
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Preview only; no filesystem changes")
    args = parser.parse_args()

    ts = time.strftime("%Y%m%d_%H%M%S")
    log_dir = Path(__file__).parent.parent / "logs"
    ensure_folder(log_dir)
    log_path = log_dir / f"cvr_folder_builder_{ts}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s: %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )

    logging.info("Starting CVR Folder Builder (bucket-by-number; last-name fallback)")
    logging.info(f"Workbook: {WORKBOOK_PATH}")
    logging.info(f"Guardian Base: {GUARDIAN_BASE}")
    logging.info(f"Inbox: {INBOX_DIR}")
    logging.info(f"Completed: {COMPLETED_DIR}")
    logging.info(f"Dry Run: {args.dry_run}")

    if not WORKBOOK_PATH.exists():
        logging.error(f"Workbook not found: {WORKBOOK_PATH}"); sys.exit(1)
    if not INBOX_DIR.exists():
        logging.error(f"Inbox folder not found: {INBOX_DIR}"); sys.exit(1)

    # Excel + existing folders
    try:
        cause_to_foldername, lastname_to_causes = build_mapping_from_excel(WORKBOOK_PATH)
    except Exception as e:
        logging.exception(f"Failed reading Excel: {e}"); sys.exit(1)
    completed_map = scan_cause_folders(COMPLETED_DIR)
    base_map      = scan_cause_folders(GUARDIAN_BASE)

    # Gather inbox PDFs
    pdfs = sorted([p for p in INBOX_DIR.glob("*.pdf") if p.is_file()])
    logging.info(f"Found {len(pdfs)} PDF(s) in inbox.")

    # Pass 1: extract cause per file + build number-buckets
    pdf_cause: dict[Path, str | None] = {}
    bucket_to_causes: dict[str, set[str]] = {}
    for pdf in pdfs:
        cause = extract_cause_number(pdf)
        pdf_cause[pdf] = cause
        bnum = get_bucket_num_key(pdf)
        bucket_to_causes.setdefault(bnum, set())
        if cause:
            bucket_to_causes[bnum].add(cause)

    # Pass 2: inherit cause inside each number-bucket
    for pdf in pdfs:
        if pdf_cause[pdf]:
            continue
        bnum = get_bucket_num_key(pdf)
        if bucket_to_causes.get(bnum):
            rep_cause = next(iter(bucket_to_causes[bnum]))
            pdf_cause[pdf] = rep_cause
            logging.info(f"[INHERIT] {pdf.name} got cause from bucket #{bnum or 'blank'}: {rep_cause}")

    # Pass 2b: last-name fallback for buckets still unknown (scan ARP text only)
    for pdf in pdfs:
        if pdf_cause[pdf]:
            continue
        # Only try on ARP-like files
        if not re.match(r"(?i)^arp([ _-]?\d+)?$", pdf.stem):
            continue
        txt = extract_text_any(pdf).lower()
        if not txt:
            continue
        candidates = []
        for lname, causes in lastname_to_causes.items():
            if lname and re.search(rf"\b{re.escape(lname)}\b", txt):
                # use only if this last name maps to exactly one cause (avoid ambiguity)
                if len(causes) == 1:
                    candidates.append(causes[0])
        # pick if exactly one unique cause found
        uniq = list({c for c in candidates})
        if len(uniq) == 1:
            found_cause = uniq[0]
            pdf_cause[pdf] = found_cause
            # also seed the bucket so siblings inherit
            bnum = get_bucket_num_key(pdf)
            bucket_to_causes.setdefault(bnum, set()).add(found_cause)
            logging.info(f"[LASTNAME-FALLBACK] {pdf.name} matched last name -> cause {found_cause}")

    # Determine which causes to ensure (seen in inbox after inheritance/fallback)
    causes_seen = {c for c in pdf_cause.values() if c}
    created = 0
    for cause in sorted(causes_seen):
        if resolve_match_in_map(cause, completed_map):
            continue  # lives in Completed -> do not create
        if resolve_match_in_map(cause, base_map):
            continue  # already exists in base
        folder_name = resolve_excel_folder_name(cause, cause_to_foldername)
        if not folder_name:
            continue
        dst = GUARDIAN_BASE / folder_name
        logging.info(f"[ENSURE-NEW] {dst}")
        if not args.dry_run:
            ensure_folder(dst)
        created += 1
        base_map[cause] = dst
    logging.info(f"Ensured {created} new ward folder(s).")

    # Unmatched folder
    unmatched_dir = GUARDIAN_BASE / UNMATCHED_DIR_NAME
    if not unmatched_dir.exists() and not args.dry_run:
        ensure_folder(unmatched_dir)

    # Pass 3: move PDFs
    moved, unmatched, skipped = 0, 0, 0
    for pdf in pdfs:
        try:
            cause = pdf_cause.get(pdf)
            if not cause:
                logging.warning(f"[NO-CAUSE] {pdf.name} -> {UNMATCHED_DIR_NAME}")
                if not args.dry_run:
                    dst = unique_destination(unmatched_dir, pdf.name)
                    shutil.move(str(pdf), str(dst))
                unmatched += 1
                continue

            # prefer base folder
            base_dst = resolve_match_in_map(cause, base_map)
            if base_dst:
                dst_folder = base_dst
            else:
                if resolve_match_in_map(cause, completed_map):
                    logging.info(f"[SKIP-COMPLETED-CASE] {pdf.name} cause={cause} -> left in inbox")
                    skipped += 1
                    continue
                folder_name = resolve_excel_folder_name(cause, cause_to_foldername)
                if not folder_name:
                    logging.warning(f"[NO-MATCH-IN-EXCEL] {pdf.name} cause={cause} -> {UNMATCHED_DIR_NAME}")
                    if not args.dry_run:
                        dst = unique_destination(unmatched_dir, pdf.name)
                        shutil.move(str(pdf), str(dst))
                    unmatched += 1
                    continue
                dst_folder = GUARDIAN_BASE / folder_name
                if not dst_folder.exists() and not args.dry_run:
                    ensure_folder(dst_folder)
                base_map[cause] = dst_folder

            tag = deduce_doc_tag(pdf.name)  # ARP/ORDER/APPROVAL + number if any
            new_name = f"{tag} - {pdf.name}"
            dst = unique_destination(dst_folder, new_name)
            logging.info(f"[MOVE] {pdf.name} (cause={cause}) -> {dst_folder.parent.name}\\{dst_folder.name}\\{dst.name}")
            if not args.dry_run:
                shutil.move(str(pdf), str(dst))
            moved += 1

        except Exception as e:
            logging.exception(f"Error handling {pdf.name}: {e}")
            if not args.dry_run:
                dst = unique_destination(unmatched_dir, pdf.name)
                shutil.move(str(pdf), str(dst))
            unmatched += 1

    logging.info(f"Done. Moved={moved}, SkippedCompleted={skipped}, Unmatched={unmatched}. Log: {log_path}")

if __name__ == "__main__":
    main()
