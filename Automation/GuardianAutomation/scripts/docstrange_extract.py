# CPU-only high-quality extraction for ARP (pages 1–2) and ORDER (page 1).
# Renders target pages to PNG at 300 DPI, extracts structured fields with DocStrange (local/Ollama),
# writes per-page fields (with metadata) and a merged per-document fields JSON (plain key→value dict).
#
# Usage:
#   py -3.12 "C:\GoogleSync\Automation\GuardianAutomation\scripts\docstrange_quick_pages.py"

import json
import warnings
from pathlib import Path
import fitz  # PyMuPDF
from docstrange import DocumentExtractor

# --- Paths & settings ---
INPUT_DIR  = Path(r"C:\GoogleSync\Guardianship Files\aa New")
OUTPUT_DIR = Path(r"C:\GoogleSync\Guardianship Files\Extracted")
DPI        = 300  # 300 DPI = better OCR; slower on CPU but higher accuracy

# Silence the harmless "pin_memory ... no accelerator" warning
warnings.filterwarnings("ignore", message=".*pin_memory.*", category=UserWarning)

# === Fields we want ======================================================
# Page 1 (identity/contact)
FIELDS_P1 = [
    "page_number", "document_id", "updated_date", "probate_court_county",
    "case_number",
    "ward_name", "ward_age", "ward_dob", "ward_phone",
    "ward_address_no_po_box", "ward_city_state_zip",
    "guardian_names", "guardian_ages", "guardian_dobs", "guardian_phone",
    "guardian_email",
    "guardian_address_no_po_box", "guardian_city_state_zip",
    "guardian_relationship_to_ward",
]

# Page 2 (residence section only — per user request)
FIELDS_P2 = [
    "5_wards_residence_type",
    "5_facility_type_nursing_home",
    "5_facility_type_group_home",
    "5_facility_type_hospital_medical_facility",
    "5_facility_type_state_supported_living_center",
    "5_facility_type_other",
    "5_facility_name",
]

# For ORDER (page 1 only) we re-use page-1 identity fields
FIELDS_ORDER = FIELDS_P1[:]

# ========================================================================

def render_pages_to_png(pdf_path: Path, out_dir: Path, dpi: int, page_indices: list[int]) -> list[Path]:
    """Render selected 0-based pages to PNG; return list of PNG file paths."""
    out_dir.mkdir(parents=True, exist_ok=True)
    pngs: list[Path] = []
    with fitz.open(pdf_path) as doc:
        for i in page_indices:
            if i < 0 or i >= len(doc):
                continue
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
            out_png = out_dir / f"{pdf_path.stem}_p{i+1}.png"
            pix.save(out_png)
            pngs.append(out_png)
    return pngs

def extract_fields_from_image(extractor: DocumentExtractor, png_path: Path, requested_fields: list[str]) -> dict:
    """
    Return the full DocStrange result dict for this image (contains 'extracted_fields', 'requested_fields', etc.).
    """
    res = extractor.extract(str(png_path))
    return res.extract_data(specified_fields=requested_fields)

def non_empty_merge(pages: list[dict]) -> dict:
    """
    Merge the INNER 'extracted_fields' dicts from page results.
    Rule: Only write non-empty values; never let empty overwrite something good.
    Return a PLAIN dict of merged fields (no wrapper), ideal for feeding your Excel step.
    """
    EMPTY = (None, "", [], {})
    merged: dict = {}
    for page in pages:
        ef = page.get("extracted_fields", {}) if isinstance(page, dict) else {}
        for k, v in (ef or {}).items():
            if v not in EMPTY:
                merged[k] = v  # last non-empty wins
    return merged

def process_pdf(extractor: DocumentExtractor, pdf: Path):
    name = pdf.stem.lower()

    # Skip approvals entirely
    if "approval" in name:
        print(f"⏭️  Skipping {pdf.name} (approval)")
        return

    # Target pages and fields per doc type
    if "arp" in name:
        page_plan = [(0, FIELDS_P1), (1, FIELDS_P2)]  # pages 1–2
    elif "order" in name:
        page_plan = [(0, FIELDS_ORDER)]               # page 1
    else:
        print(f"⏭️  Skipping {pdf.name} (not ARP/ORDER)")
        return

    all_page_results: list[dict] = []
    print(f"🖼️  Rendering {pdf.name} at {DPI} DPI …")
    pngs = render_pages_to_png(pdf, OUTPUT_DIR, DPI, [p for p, _ in page_plan])
    if not pngs:
        print(f"❌ No images rendered for {pdf.name}")
        return

    # Extract per page with its own requested fields
    for (page_index, fields), png in zip(page_plan, pngs):
        print(f"🔎 Extracting {png.name} (page {page_index + 1}) …")
        try:
            result = extract_fields_from_image(extractor, png, fields)
            all_page_results.append(result)

            # Write per-page result WITH wrapper (useful for debugging)
            page_json = OUTPUT_DIR / f"{png.stem}.fields.json"
            page_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"✅ Wrote {page_json.name}")
        except Exception as e:
            print(f"  ✗ Error on {png.name}: {e}")

    # Write merged PLAIN dict (best for downstream mapping)
    if all_page_results:
        merged_plain = non_empty_merge(all_page_results)
        merged_json = OUTPUT_DIR / f"{pdf.stem}.fields.json"
        merged_json.write_text(json.dumps(merged_plain, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"📦 Merged fields → {merged_json.name} (plain key→value)")

def main():
    if not INPUT_DIR.exists():
        print(f"❌ Input folder not found: {INPUT_DIR}")
        return
    print("▶ Local mode (cpu=True) with Ollama; high DPI on CPU. This can take a bit…")
    extractor = DocumentExtractor(cpu=True)

    pdfs = sorted(p for p in INPUT_DIR.glob("*.pdf") if p.is_file())
    for pdf in pdfs:
        process_pdf(extractor, pdf)

    print("🎉 Done.")

if __name__ == "__main__":
    main()
