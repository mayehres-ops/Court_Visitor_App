# One-page smoke test: render page 1 of ARP.pdf to an image and extract fields locally (Ollama).
# Fast on CPU, no cloud/quota involved.

import json, os
from pathlib import Path
import fitz  # PyMuPDF
from docstrange import DocumentExtractor

PDF = r"C:\GoogleSync\Guardianship Files\aa New\ARP.pdf"
PNG = r"C:\GoogleSync\Guardianship Files\Extracted\ARP_p1.png"
OUT = r"C:\GoogleSync\Guardianship Files\Extracted\ARP_p1.fields.json"
DPI = 150  # bump to 200 later if you want sharper OCR

FIELDS = [
    "case_number",
    "ward_name", "ward_dob", "ward_phone",
    "ward_address_no_po_box", "ward_city_state_zip",
    "guardian_names", "guardian_phone",
    "guardian_address_no_po_box",
]

def render_first_page_to_png(pdf_path: str, png_path: str, dpi: int = 150):
    pdf = Path(pdf_path)
    if not pdf.exists():
        raise FileNotFoundError(f"PDF not found: {pdf}")
    Path(png_path).parent.mkdir(parents=True, exist_ok=True)
    print(f"🖼️  Rendering page 1 at {dpi} DPI …")
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)  # first page (0-index)
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    pix.save(png_path)
    doc.close()
    print(f"✅ Wrote image: {png_path}")

def main():
    # Force local mode explicitly (Ollama); no cloud/quota involved.
    print("▶ Initializing local extractor (cpu=True) …")
    ex = DocumentExtractor(cpu=True)  # needs Ollama with a small model (e.g., llama3.2) installed
    print("✅ Local extractor ready")

    render_first_page_to_png(PDF, PNG, DPI)

    print("📄 Extracting fields from image …")
    res = ex.extract(PNG)
    fields = res.extract_data(specified_fields=FIELDS)

    Path(OUT).write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
    print("✅ Wrote:", OUT)
    print(json.dumps(fields, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
