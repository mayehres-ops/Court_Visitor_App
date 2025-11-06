"""
Convert README_FIRST.md to README_FIRST.pdf with nice formatting
"""

import markdown
from weasyprint import HTML, CSS
from pathlib import Path

def convert_readme_to_pdf():
    """Convert README_FIRST.md to a nicely formatted PDF"""

    readme_path = Path(__file__).parent / "README_FIRST.md"
    pdf_path = Path(__file__).parent / "README_FIRST.pdf"

    # Read the markdown
    with open(readme_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # Convert markdown to HTML
    html_content = markdown.markdown(
        md_content,
        extensions=['tables', 'fenced_code', 'nl2br', 'toc']
    )

    # Add nice CSS styling
    styled_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: letter;
                margin: 0.75in;
                @bottom-center {{
                    content: counter(page) " of " counter(pages);
                    font-size: 9pt;
                    color: #666;
                }}
            }}

            body {{
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 11pt;
                line-height: 1.6;
                color: #333;
            }}

            h1 {{
                color: #1e3a8a;
                font-size: 24pt;
                border-bottom: 3px solid #1e3a8a;
                padding-bottom: 10px;
                margin-top: 30px;
                page-break-before: always;
            }}

            h1:first-of-type {{
                page-break-before: avoid;
            }}

            h2 {{
                color: #2563eb;
                font-size: 18pt;
                margin-top: 24px;
                border-bottom: 2px solid #dbeafe;
                padding-bottom: 6px;
            }}

            h3 {{
                color: #3b82f6;
                font-size: 14pt;
                margin-top: 18px;
            }}

            h4 {{
                color: #60a5fa;
                font-size: 12pt;
                margin-top: 14px;
            }}

            p {{
                margin: 10px 0;
            }}

            ul, ol {{
                margin: 10px 0;
                padding-left: 25px;
            }}

            li {{
                margin: 5px 0;
            }}

            code {{
                background-color: #f3f4f6;
                padding: 2px 6px;
                border-radius: 3px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 10pt;
            }}

            pre {{
                background-color: #f3f4f6;
                padding: 15px;
                border-radius: 5px;
                border-left: 4px solid #3b82f6;
                overflow-x: auto;
                font-size: 9pt;
            }}

            pre code {{
                background: none;
                padding: 0;
            }}

            table {{
                border-collapse: collapse;
                width: 100%;
                margin: 15px 0;
                font-size: 10pt;
            }}

            th {{
                background-color: #1e3a8a;
                color: white;
                padding: 10px;
                text-align: left;
                font-weight: bold;
            }}

            td {{
                border: 1px solid #ddd;
                padding: 8px;
            }}

            tr:nth-child(even) {{
                background-color: #f9fafb;
            }}

            blockquote {{
                border-left: 4px solid #fbbf24;
                background-color: #fef3c7;
                padding: 15px;
                margin: 15px 0;
                font-style: italic;
            }}

            strong {{
                color: #1e3a8a;
                font-weight: 600;
            }}

            hr {{
                border: none;
                border-top: 2px solid #e5e7eb;
                margin: 30px 0;
            }}

            a {{
                color: #2563eb;
                text-decoration: none;
            }}

            .warning {{
                background-color: #fef2f2;
                border-left: 4px solid #dc2626;
                padding: 15px;
                margin: 15px 0;
            }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """

    # Convert HTML to PDF
    HTML(string=styled_html).write_pdf(pdf_path)

    print(f"✅ Successfully created: {pdf_path}")
    return pdf_path

if __name__ == "__main__":
    try:
        convert_readme_to_pdf()
        print("\n[OK] README to PDF conversion SUCCESS")
    except Exception as e:
        print(f"\n[FAIL] README to PDF conversion FAILED: {e}")
        import traceback
        traceback.print_exc()
