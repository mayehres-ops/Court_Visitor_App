"""
Analyze CVR template and create a detailed field breakdown document.

Creates a Word document with:
1. Table with 3 columns:
   - Column 1: Basic fields filled when report initially created (Step 8)
   - Column 2: Fields filled in by Google Form (Step 10)
   - Column 3: Fields not filled by Google Form but have control names
2. List of unnamed controls at the end
"""

import win32com.client
import json
from pathlib import Path

# Paths
TEMPLATE_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx"
MAPPING_PATH = r"C:\GoogleSync\GuardianShip_App\Config\cvr_google_form_mapping.json"
OUTPUT_PATH = r"C:\GoogleSync\GuardianShip_App\CVR_Field_Analysis.docx"

# Step 8 fields (filled from Excel)
STEP_8_FIELDS = {
    'causeno': 'Cause Number',
    'wardfirst': 'Ward First Name',
    'wardmiddle': 'Ward Middle Name',
    'wardlast': 'Ward Last Name',
    'wfirst': 'Ward First Name (alias)',
    'wlast': 'Ward Last Name (alias)',
    'visitdate': 'Visit Date',
    'visittime': 'Visit Time',
    'waddress': 'Ward Address',
    'datefiled': 'Date Filed'
}

def load_google_form_mapping():
    """Load Google Form field mapping"""
    try:
        with open(MAPPING_PATH, 'r', encoding='utf-8') as f:
            mapping = json.load(f)

        # Extract control names and their descriptions
        google_fields = {}
        for question, config in mapping.items():
            if question.startswith('_'):
                continue

            cvr_control = config.get('cvr_control', '')
            field_type = config.get('type', '')
            note = config.get('note', '')

            if cvr_control:
                # Handle multi-select and yes/no variations
                if 'Multi-select' in note or 'Single-select' in note:
                    # Extract control names from note
                    if 'access_' in note:
                        google_fields['access_phone'] = f"{question} - Telephone"
                        google_fields['access_tv'] = f"{question} - Television"
                        google_fields['access_radio'] = f"{question} - Radio"
                        google_fields['access_computer'] = f"{question} - Computer"
                    elif 'activities_' in note:
                        google_fields['activities_school'] = f"{question} - School"
                        google_fields['activities_work'] = f"{question} - Work"
                        google_fields['activities_daycare'] = f"{question} - Day Care"
                    elif 'help_' in note:
                        google_fields['help_bathing'] = f"{question} - Bathing"
                        google_fields['help_dressing'] = f"{question} - Dressing"
                        google_fields['help_eating'] = f"{question} - Eating"
                        google_fields['help_walking'] = f"{question} - Walking"
                        google_fields['help_bathroom'] = f"{question} - Bathroom"
                    elif 'cond_' in note:
                        google_fields['cond_hearing'] = f"{question} - Hearing impairment"
                        google_fields['cond_speech'] = f"{question} - Speech impairment"
                        google_fields['cond_unwilling_speak'] = f"{question} - Unwilling to speak"
                        google_fields['cond_unable_speak'] = f"{question} - Unable to speak"
                        google_fields['cond_unresponsive'] = f"{question} - Unresponsive"
                        google_fields['cond_voice'] = f"{question} - Able to communicate with voice"
                        google_fields['cond_gestures'] = f"{question} - Nonverbal gestures"
                        google_fields['cond_bed'] = f"{question} - In bed most/all time"
                        google_fields['cond_walk_assist'] = f"{question} - Walk with assistance"
                    elif 'residence_' in note:
                        google_fields['residence_own'] = f"{question} - Own Home/Apartment"
                        google_fields['residence_guardian'] = f"{question} - Guardian's Home"
                        google_fields['residence_relative'] = f"{question} - Other Relative's Home"
                        google_fields['residence_nursing'] = f"{question} - Nursing home/Assisted Living"
                        google_fields['residence_group'] = f"{question} - Group Home"
                        google_fields['residence_hospital'] = f"{question} - Hospital/Medical facility"
                        google_fields['residence_state'] = f"{question} - State Supported Living Center"
                        google_fields['residence_other'] = f"{question} - Other"
                elif field_type == 'yesno':
                    # Yes/No checkboxes have _yes and _no variants
                    google_fields[f"{cvr_control}_yes"] = f"{question} - Yes"
                    google_fields[f"{cvr_control}_no"] = f"{question} - No"
                else:
                    google_fields[cvr_control] = question

        return google_fields
    except Exception as e:
        print(f"Error loading mapping: {e}")
        return {}

def analyze_template():
    """Analyze CVR template and categorize all controls"""

    print("Analyzing CVR template...")

    # Load Google Form mapping
    google_fields = load_google_form_mapping()

    # Open template
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(TEMPLATE_PATH)

    # Categorize controls
    step8_controls = []
    google_form_controls = []
    other_named_controls = []
    unnamed_controls = []

    for i, cc in enumerate(doc.ContentControls, 1):
        name = (cc.Title or cc.Tag or '').strip().lower()

        if not name:
            unnamed_controls.append(i)
            continue

        # Check which category
        if name in [k.lower() for k in STEP_8_FIELDS.keys()]:
            # Find original case
            for orig_name, desc in STEP_8_FIELDS.items():
                if orig_name.lower() == name:
                    step8_controls.append((orig_name, desc))
                    break
        elif name in [k.lower() for k in google_fields.keys()]:
            # Find original case
            for orig_name, desc in google_fields.items():
                if orig_name.lower() == name:
                    google_form_controls.append((orig_name, desc))
                    break
        else:
            other_named_controls.append((name, f"Named control: {name}"))

    doc.Close(False)
    word.Quit()

    # Remove duplicates
    step8_controls = list(set(step8_controls))
    google_form_controls = list(set(google_form_controls))
    other_named_controls = list(set(other_named_controls))

    return step8_controls, google_form_controls, other_named_controls, unnamed_controls

def create_analysis_document(step8_controls, google_form_controls, other_named_controls, unnamed_controls):
    """Create Word document with analysis"""

    print("Creating analysis document...")

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Add()

    # Add title
    title = doc.Range(0, 0)
    title.Text = "CVR Template Field Analysis\n\n"
    title.Font.Size = 16
    title.Font.Bold = True
    title.InsertParagraphAfter()

    # Add introduction
    intro = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    intro.Text = (
        "This document shows all content controls in the Court Visitor Report template "
        "and how they are filled.\n\n"
        "Summary:\n"
        f"  - Step 8 (Excel) fills: {len(step8_controls)} fields\n"
        f"  - Step 10 (Google Form) fills: {len(google_form_controls)} fields\n"
        f"  - Other named controls: {len(other_named_controls)} fields\n"
        f"  - Unnamed controls: {len(unnamed_controls)} controls\n\n"
    )
    intro.InsertParagraphAfter()

    # Create table
    num_rows = max(len(step8_controls), len(google_form_controls), len(other_named_controls)) + 1
    table = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), num_rows, 3)
    table.Borders.Enable = True

    # Header row
    table.Cell(1, 1).Range.Text = "Step 8: Filled from Excel"
    table.Cell(1, 2).Range.Text = "Step 10: Filled from Google Form"
    table.Cell(1, 3).Range.Text = "Other Named Controls"

    # Make header bold
    for col in range(1, 4):
        table.Cell(1, col).Range.Font.Bold = True
        table.Cell(1, col).Shading.BackgroundPatternColor = 0xDDDDDD

    # Fill table
    max_rows = max(len(step8_controls), len(google_form_controls), len(other_named_controls))

    for i in range(max_rows):
        row = i + 2  # +1 for 1-based indexing, +1 for header

        # Column 1: Step 8
        if i < len(step8_controls):
            name, desc = step8_controls[i]
            table.Cell(row, 1).Range.Text = f"{name}\n{desc}"

        # Column 2: Google Form
        if i < len(google_form_controls):
            name, desc = google_form_controls[i]
            table.Cell(row, 2).Range.Text = f"{name}\n{desc}"

        # Column 3: Other named
        if i < len(other_named_controls):
            name, desc = other_named_controls[i]
            table.Cell(row, 3).Range.Text = f"{name}"

    # Add unnamed controls section
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertParagraphAfter()
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertParagraphAfter()

    unnamed_header = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    unnamed_header.Text = f"Unnamed Controls ({len(unnamed_controls)} total)\n"
    unnamed_header.Font.Size = 14
    unnamed_header.Font.Bold = True
    unnamed_header.InsertParagraphAfter()

    unnamed_text = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    unnamed_text.Text = (
        "These content controls do not have a Title or Tag set, so they cannot be auto-filled. "
        "They must be filled manually or you can name them to enable auto-fill.\n\n"
        f"Control numbers: {', '.join(map(str, unnamed_controls[:20]))}"
    )
    if len(unnamed_controls) > 20:
        unnamed_text.Text += f", ... and {len(unnamed_controls) - 20} more"
    unnamed_text.InsertParagraphAfter()

    # Save document
    doc.SaveAs(OUTPUT_PATH)
    doc.Close()
    word.Quit()

    print(f"Analysis saved to: {OUTPUT_PATH}")

def main():
    """Main function"""
    print("=" * 80)
    print("CVR TEMPLATE FIELD ANALYSIS")
    print("=" * 80)

    # Analyze template
    step8, google_form, other_named, unnamed = analyze_template()

    # Create document
    create_analysis_document(step8, google_form, other_named, unnamed)

    print("\nDone!")
    print(f"Open: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
