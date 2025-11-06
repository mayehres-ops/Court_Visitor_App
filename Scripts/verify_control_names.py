"""
Verify control names in CVR template match Google Form mapping.
Shows which controls are correctly named, which are missing, and which don't match.
"""

import win32com.client
import json
from pathlib import Path

TEMPLATE_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx"
MAPPING_PATH = r"C:\GoogleSync\GuardianShip_App\Config\cvr_google_form_mapping.json"
OUTPUT_PATH = r"C:\GoogleSync\GuardianShip_App\Control_Names_Verification.docx"

def load_expected_controls():
    """Load expected control names from Google Form mapping"""
    expected = set()

    with open(MAPPING_PATH, 'r', encoding='utf-8') as f:
        mapping = json.load(f)

    for question, config in mapping.items():
        if question.startswith('_'):
            continue

        note = config.get('note', '')
        cvr_control = config.get('cvr_control', '')
        field_type = config.get('type', '')

        # Handle different field types
        if 'access_phone' in note or 'access_tv' in note:
            expected.update(['access_phone', 'access_tv', 'access_radio', 'access_computer'])
        elif 'activities_' in note:
            expected.update(['activities_school', 'activities_work', 'activities_daycare'])
        elif 'help_' in note:
            expected.update(['help_bathing', 'help_dressing', 'help_eating', 'help_walking', 'help_bathroom'])
        elif 'cond_' in note:
            expected.update(['cond_hearing', 'cond_speech', 'cond_unwilling_speak', 'cond_unable_speak',
                           'cond_unresponsive', 'cond_voice', 'cond_gestures', 'cond_bed', 'cond_walk_assist'])
        elif 'residence_' in note:
            expected.update(['residence_own', 'residence_guardian', 'residence_relative', 'residence_nursing',
                           'residence_group', 'residence_hospital', 'residence_state', 'residence_other'])
        elif field_type == 'yesno':
            expected.add(f"{cvr_control}_yes")
            expected.add(f"{cvr_control}_no")
        elif cvr_control:
            expected.add(cvr_control)

    return expected

def get_actual_controls():
    """Get all named controls from CVR template"""
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(TEMPLATE_PATH)

    actual = {}

    for cc in doc.ContentControls:
        name = (cc.Title or cc.Tag or '').strip()
        if name:
            if name not in actual:
                actual[name] = 0
            actual[name] += 1

    doc.Close(False)
    word.Quit()

    return actual

def create_verification_report(expected, actual):
    """Create verification report document"""

    # Categorize controls
    expected_lower = {e.lower(): e for e in expected}
    actual_lower = {a.lower(): a for a in actual.keys()}

    correctly_named = []
    missing = []
    extra = []

    # Check what's correctly named
    for exp_lower, exp_name in expected_lower.items():
        if exp_lower in actual_lower:
            actual_name = actual_lower[exp_lower]
            count = actual[actual_name]
            correctly_named.append((exp_name, actual_name, count))

    # Check what's missing
    for exp_lower, exp_name in expected_lower.items():
        if exp_lower not in actual_lower:
            missing.append(exp_name)

    # Check for extra controls (not in mapping)
    step8_fields = ['causeno', 'wardfirst', 'wardmiddle', 'wardlast', 'wfirst', 'wlast',
                    'visitdate', 'visittime', 'waddress', 'datefiled']

    for act_lower, act_name in actual_lower.items():
        if act_lower not in expected_lower and act_lower not in [s.lower() for s in step8_fields]:
            extra.append((act_name, actual[act_name]))

    # Create Word document
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Add()

    # Title
    title = doc.Range(0, 0)
    title.Text = "CVR Template Control Names Verification\n\n"
    title.Font.Size = 16
    title.Font.Bold = True
    title.InsertParagraphAfter()

    # Summary
    summary = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    summary.Text = (
        f"Summary:\n"
        f"  ✓ Correctly named: {len(correctly_named)} controls\n"
        f"  ✗ Missing (need to add): {len(missing)} controls\n"
        f"  ? Extra (not in mapping): {len(extra)} controls\n\n"
    )
    summary.InsertParagraphAfter()

    # Section 1: Correctly Named
    if correctly_named:
        header1 = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        header1.Text = f"✓ Correctly Named Controls ({len(correctly_named)})\n"
        header1.Font.Size = 14
        header1.Font.Bold = True
        header1.Font.Color = 0x008000  # Green
        header1.InsertParagraphAfter()

        for expected_name, actual_name, count in sorted(correctly_named):
            item = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            match_str = "✓" if expected_name == actual_name else f"✓ (you used: {actual_name})"
            if count > 1:
                item.Text = f"  • {expected_name} {match_str} - appears {count} times\n"
            else:
                item.Text = f"  • {expected_name} {match_str}\n"
            item.InsertParagraphAfter()

        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertParagraphAfter()

    # Section 2: Missing
    if missing:
        header2 = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        header2.Text = f"✗ Missing Controls - Need to Add ({len(missing)})\n"
        header2.Font.Size = 14
        header2.Font.Bold = True
        header2.Font.Color = 0x0000FF  # Red
        header2.InsertParagraphAfter()

        intro = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        intro.Text = "These controls are in the Google Form mapping but not found in the CVR template:\n\n"
        intro.InsertParagraphAfter()

        for name in sorted(missing):
            item = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            item.Text = f"  • {name}\n"
            item.Font.Color = 0x0000FF
            item.InsertParagraphAfter()

        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertParagraphAfter()

    # Section 3: Extra
    if extra:
        header3 = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        header3.Text = f"? Extra Controls - Not in Google Form Mapping ({len(extra)})\n"
        header3.Font.Size = 14
        header3.Font.Bold = True
        header3.Font.Color = 0x808080  # Gray
        header3.InsertParagraphAfter()

        intro = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        intro.Text = (
            "These controls are named in the CVR template but not in the Google Form mapping.\n"
            "They won't be auto-filled unless you add them to the mapping or they might be for manual entry:\n\n"
        )
        intro.InsertParagraphAfter()

        for name, count in sorted(extra):
            item = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            if count > 1:
                item.Text = f"  • {name} - appears {count} times\n"
            else:
                item.Text = f"  • {name}\n"
            item.Font.Color = 0x808080
            item.InsertParagraphAfter()

    # Save
    doc.SaveAs(OUTPUT_PATH)
    doc.Close()
    word.Quit()

    print(f"\nVerification report saved to: {OUTPUT_PATH}")

    return len(correctly_named), len(missing), len(extra)

def main():
    print("=" * 80)
    print("CONTROL NAMES VERIFICATION")
    print("=" * 80)

    print("\nLoading expected control names from Google Form mapping...")
    expected = load_expected_controls()
    print(f"  Expected: {len(expected)} control names")

    print("\nReading actual control names from CVR template...")
    actual = get_actual_controls()
    print(f"  Found: {len(actual)} named controls in template")

    print("\nCreating verification report...")
    correct, missing, extra = create_verification_report(expected, actual)

    print("\n" + "=" * 80)
    print("RESULTS")
    print("=" * 80)
    print(f"✓ Correctly named: {correct}")
    print(f"✗ Missing: {missing}")
    print(f"? Extra (not in mapping): {extra}")
    print("=" * 80)

    if missing == 0:
        print("\n✓✓✓ All required controls are named! Ready to test Step 10! ✓✓✓")
    else:
        print(f"\nYou still need to name {missing} controls. See the report for details.")

    print(f"\nOpen report: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
