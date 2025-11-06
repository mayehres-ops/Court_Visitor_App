# CPU-only high quality run: ARP pages 1–2, ORDER page 1, skip approvals.
# Renders selected pages at 300 DPI to PNG, then extracts fields locally (Ollama).

import json, os
from pathlib import Path
import fitz  # PyMuPDF
from docstrange import DocumentExtractor

INPUT_DIR = r"C:\GoogleSync\Guardianship Files\aa New"
OUTPUT_DIR = r"C:\GoogleSync\Guardianship Files\Extracted"
DPI = 300  # 300 = higher quality OCR on CPU (slower but better). Drop to 200 if needed.

# Fields we care about (tweak as needed)
FIELDS = [
    "page_number", "document_id", "updated_date", "probate_court_county",
    "case_number",
    "ward_name", "ward_age", "ward_dob", "ward_phone",
    "ward_address_no_po_box", "ward_city_state_zip",
    "guardian_names", "guardian_ages", "guardian_dobs", "guardian_phone",
    "guardian_email",
    "guardian_address_no_po_box", "guardian_city_state_zip",
    "guardian_relationship_to_ward",
    # add more if you want (from your online sample)
]

def render_pages_to_png(pdf_path: Path, out_dir: Path, dpi: int, page_indices: list[int]) -> list[Path]:
    """Render specified 0-based pages to PNG files; return list of PNG paths."""
    out_dir.mkdir(parents=True, exist_ok=True)
    png_paths = []
    with fitz.open(pdf_path) as doc:
        for i in page_indices:
            if i < 0 or i >= len(doc):
                continue
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
            png_path = out_dir / f"{pdf_path.stem}_p{i+1}.png"
            pix.save(png_path)
            png_paths.append(png_path)
    return png_paths

def merge_fields(dicts: list[dict]) -> dict:
    """Union merge of page field dicts (last non-empty wins)."""
    out = {}
    for d in dicts:
        for k, v in (d or {}).items():
            if v not in (None, "", [], {}):
                out[k] = v
    return out

def process_doc(extractor: DocumentExtractor, pdf: Path):
    name = pdf.stem.lower()
    if "approval" in name:
        print(f"⏭️  Skipping {pdf.name} (approval)")
        return

    # Target pages by document type
    if "arp" in name:
        pages = [0, 1]        # pages 1–2
    elif "order" in name:
        pages = [0]           # page 1
    else:
        print(f"⏭️  Skipping {pdf.name} (not ARP/ORDER)")
        return

    print(f"🖼️  Rendering {pdf.name} at {DPI} DPI for pages { [p+1 for p in pages] } …")
    pngs = render_pages_to_png(pdf, Path(OUTPUT_DIR), DPI, pages)
    if not pngs:
        print(f"❌ No images rendered for {pdf.name}")
        return

    page_fields = []
    for png in pngs:
        print(f"🔎 Extracting from {png.name} …")
        res = extractor.extract(str(png))
        f = res.extract_data(specified_fields=FIELDS)
        page_fields.append(f)
        out_page = Path(OUTPUT_DIR) / f"{png.stem}.fields.json"
        out_page.write_text(json.dumps(f, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ Wrote {out_page.name}")

    # Also write a merged fields file per document
    merged = merge_fields(page_fields)
    out_merged = Path(OUTPUT_DIR) / f"{pdf.stem}.fields.json"
    out_merged.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"📦 Merged fields → {out_merged.name}")

def main():
    inp = Path(INPUT_DIR)
    if not inp.exists():
        print(f"❌ Input folder not found: {inp}")
        return

    print("▶ Local mode (cpu=True) with Ollama; high DPI on CPU. This can take a bit…")
    ex = DocumentExtractor(cpu=True)  # local/CPU; no quotas

    pdfs = sorted(p for p in inp.glob("*.pdf") if p.is_file())
    for pdf in pdfs:
        process_doc(ex, pdf)

    print("🎉 Done.")

if __name__ == "__main__":
    main()
