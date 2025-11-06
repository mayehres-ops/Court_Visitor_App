"""
List all unnamed controls in the CVR template with surrounding context.
Creates a detailed document showing each unnamed control's location and nearby text.
"""

import win32com.client

TEMPLATE_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx"
OUTPUT_PATH = r"C:\GoogleSync\GuardianShip_App\Unnamed_Controls_List.docx"

def get_surrounding_text(doc, cc, context_chars=100):
    """Get text before and after a content control"""
    try:
        before_text = ""
        after_text = ""

        # Get text before control
        if cc.Range.Start > context_chars:
            temp_range = doc.Range(cc.Range.Start - context_chars, cc.Range.Start)
            before_text = temp_range.Text
        elif cc.Range.Start > 0:
            temp_range = doc.Range(0, cc.Range.Start)
            before_text = temp_range.Text

        # Get text after control
        if cc.Range.End + context_chars < doc.Content.End:
            temp_range = doc.Range(cc.Range.End, cc.Range.End + context_chars)
            after_text = temp_range.Text
        else:
            temp_range = doc.Range(cc.Range.End, doc.Content.End)
            after_text = temp_range.Text

        # Get control content
        control_text = cc.Range.Text if cc.Range.Text else "[empty]"

        # Clean up text
        before_text = before_text.replace('\r', ' ').replace('\x07', '').strip()
        control_text = control_text.replace('\r', ' ').replace('\x07', '').strip()
        after_text = after_text.replace('\r', ' ').replace('\x07', '').strip()

        return before_text, control_text, after_text
    except:
        return "[error getting context]", "[error]", "[error getting context]"

def create_unnamed_controls_document():
    """Create detailed document listing all unnamed controls"""

    print("Analyzing CVR template for unnamed controls...")

    # Open template
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(TEMPLATE_PATH)

    # Find all unnamed controls
    unnamed_controls = []

    type_names = {
        0: 'Unknown',
        1: 'RichText',
        2: 'Text',
        3: 'Checkbox',
        4: 'Picture',
        5: 'Date',
        6: 'Dropdown',
        7: 'ComboBox',
        8: 'BuildingBlock'
    }

    for i, cc in enumerate(doc.ContentControls, 1):
        name = (cc.Title or cc.Tag or '').strip()

        if not name:
            # This is an unnamed control
            control_type = type_names.get(cc.Type, f'Type{cc.Type}')
            before, content, after = get_surrounding_text(doc, cc)

            unnamed_controls.append({
                'number': i,
                'type': control_type,
                'before': before[-80:] if len(before) > 80 else before,  # Last 80 chars
                'content': content[:50] if len(content) > 50 else content,  # First 50 chars
                'after': after[:80] if len(after) > 80 else after  # First 80 chars
            })

    doc.Close(False)

    print(f"Found {len(unnamed_controls)} unnamed controls")

    # Create output document
    print("Creating detailed list document...")

    output_doc = word.Documents.Add()

    # Title
    title = output_doc.Range(0, 0)
    title.Text = f"Unnamed Controls in CVR Template ({len(unnamed_controls)} total)\n\n"
    title.Font.Size = 16
    title.Font.Bold = True
    title.InsertParagraphAfter()

    # Introduction
    intro = output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1)
    intro.Text = (
        "This document lists all content controls in the CVR template that do not have "
        "a Title or Tag set. These controls cannot be auto-filled by the scripts.\n\n"
        "For each control, you'll see:\n"
        "  - Control number (position in document)\n"
        "  - Control type\n"
        "  - Text BEFORE the control (to help locate it)\n"
        "  - Text INSIDE the control (if any)\n"
        "  - Text AFTER the control\n\n"
        "You can use this information to decide which controls need names.\n\n"
    )
    intro.InsertParagraphAfter()
    intro.InsertParagraphAfter()

    # List each unnamed control
    for ctrl in unnamed_controls:
        # Control header
        header = output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1)
        header.Text = f"Control #{ctrl['number']} ({ctrl['type']})"
        header.Font.Size = 12
        header.Font.Bold = True
        header.Font.Color = 0x0000FF  # Blue
        header.InsertParagraphAfter()

        # Before text
        before_para = output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1)
        before_para.Text = f"  Before: ...{ctrl['before']}"
        before_para.Font.Color = 0x808080  # Gray
        before_para.InsertParagraphAfter()

        # Content
        content_para = output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1)
        content_para.Text = f"  Content: [{ctrl['content']}]"
        content_para.Font.Bold = True
        content_para.InsertParagraphAfter()

        # After text
        after_para = output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1)
        after_para.Text = f"  After: {ctrl['after']}..."
        after_para.Font.Color = 0x808080  # Gray
        after_para.InsertParagraphAfter()

        # Blank line between controls
        output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1).InsertParagraphAfter()

    # Add summary at end
    output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1).InsertParagraphAfter()
    summary = output_doc.Range(output_doc.Content.End - 1, output_doc.Content.End - 1)
    summary.Text = (
        f"Total unnamed controls: {len(unnamed_controls)}\n\n"
        "To name a control:\n"
        "  1. Open the CVR template in Word\n"
        "  2. Use the context above to locate the control\n"
        "  3. Click on the control\n"
        "  4. Click Developer tab -> Properties\n"
        "  5. Enter a name in the 'Title' field\n"
        "  6. Click OK"
    )
    summary.Font.Size = 10

    # Save document
    output_doc.SaveAs(OUTPUT_PATH)
    output_doc.Close()
    word.Quit()

    print(f"Unnamed controls list saved to: {OUTPUT_PATH}")

def main():
    print("=" * 80)
    print("UNNAMED CONTROLS LIST GENERATOR")
    print("=" * 80)

    create_unnamed_controls_document()

    print("\nDone!")
    print(f"Open: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
