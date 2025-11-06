"""
Automatically name content controls in a Word document based on surrounding text.

This script can automatically name unnamed content controls by analyzing the text
around them and matching to a provided mapping.

WARNING: This modifies the Word document! Make a backup first!
"""

import win32com.client
import json
import re
from pathlib import Path

def normalize_text(text):
    """Normalize text for matching"""
    return re.sub(r'[^\w\s]', '', text.lower()).strip()

def find_best_match(before_text, after_text, mapping_phrases):
    """Find best matching control name based on surrounding text"""
    combined = f"{before_text} {after_text}".lower()

    for control_name, phrases in mapping_phrases.items():
        for phrase in phrases:
            if phrase.lower() in combined:
                return control_name

    return None

def auto_name_controls(template_path, mapping_file=None, dry_run=True):
    """
    Automatically name unnamed controls in a Word document.

    Args:
        template_path: Path to Word document
        mapping_file: Optional JSON file with control name mappings
        dry_run: If True, don't actually modify the document (just show what would change)
    """

    # Load mapping if provided
    if mapping_file and Path(mapping_file).exists():
        with open(mapping_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
    else:
        # Default mapping based on Google Form
        config = {
            "supplemented_by": ["information was supplemented by", "supplemented by"],
            "ward_description": ["tell me a little about", "about the person"],
            "visit_info": ["information needed to visit", "visit information"],
            "guardian_present_yes": ["will you be present", "guardian present.*yes"],
            "guardian_present_no": ["will you be present", "guardian present.*no"],
            "livetogether_yes": ["do you live together.*yes", "live together.*yes"],
            "livetogether_no": ["do you live together.*no", "live together.*no"],
            "ownbed_yes": ["own bed.*yes", "his/her own bed.*yes"],
            "ownbed_no": ["own bed.*no", "his/her own bed.*no"],
            "hotwater_yes": ["hot water.*yes"],
            "hotwater_no": ["hot water.*no"],
            "hvac_yes": ["air conditioning.*yes", "heating.*yes"],
            "hvac_no": ["air conditioning.*no", "heating.*no"],
            "accessible_yes": ["accessible.*yes", "areas accessible.*yes"],
            "accessible_no": ["accessible.*no", "areas accessible.*no"],
            "access_phone": ["telephone"],
            "access_tv": ["television"],
            "access_radio": ["radio"],
            "access_computer": ["computer"],
            "residence_own": ["own home", "apartment"],
            "residence_guardian": ["guardian's home", "guardian home"],
            "residence_relative": ["relative's home", "other relative"],
            "residence_nursing": ["nursing", "assisted living"],
            "residence_group": ["group home"],
            "residence_hospital": ["hospital", "medical facility"],
            "residence_state": ["state school", "state supported"],
            "residence_other": ["other:"],
            "relative_relationship": ["relationship.*click"],
            "facility_name": ["facility.*name", "place the ward lives"],
            "activities_school": ["school"],
            "activities_work": ["work"],
            "activities_daycare": ["day care"],
            "has_visitors_yes": ["visitors.*yes"],
            "has_visitors_no": ["visitors.*no"],
            "exercises_yes": ["exercise.*yes"],
            "exercises_no": ["exercise.*no"],
            "transportation_yes": ["transportation.*yes", "trips.*yes"],
            "transportation_no": ["transportation.*no", "trips.*no"],
            "fire_extinguisher_yes": ["fire extinguisher.*yes"],
            "fire_extinguisher_no": ["fire extinguisher.*no"],
            "safe_alone_yes": ["safe.*alone.*yes"],
            "safe_alone_no": ["safe.*alone.*no"],
            "can_read_yes": ["can.*read.*yes"],
            "can_read_no": ["can.*read.*no"],
            "can_write_yes": ["can.*write.*yes"],
            "can_write_no": ["can.*write.*no"],
            "oriented_yes": ["oriented.*yes"],
            "oriented_no": ["oriented.*no"],
        }

    print(f"Opening template: {template_path}")

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(str(Path(template_path).absolute()))

    changes = []
    unnamed_count = 0

    for i, cc in enumerate(doc.ContentControls, 1):
        name = (cc.Title or cc.Tag or '').strip()

        if not name:
            unnamed_count += 1

            # Get surrounding text
            try:
                before_text = ""
                after_text = ""

                if cc.Range.Start > 100:
                    temp_range = doc.Range(cc.Range.Start - 100, cc.Range.Start)
                    before_text = temp_range.Text

                if cc.Range.End + 100 < doc.Content.End:
                    temp_range = doc.Range(cc.Range.End, cc.Range.End + 100)
                    after_text = temp_range.Text

                # Clean text
                before_text = before_text.replace('\r', ' ').replace('\x07', '').strip()
                after_text = after_text.replace('\r', ' ').replace('\x07', '').strip()

                # Find best match
                suggested_name = find_best_match(before_text, after_text, config)

                if suggested_name:
                    changes.append({
                        'control_num': i,
                        'suggested_name': suggested_name,
                        'before': before_text[-50:] if len(before_text) > 50 else before_text,
                        'after': after_text[:50] if len(after_text) > 50 else after_text,
                        'control': cc
                    })

            except Exception as e:
                print(f"  Error analyzing control #{i}: {e}")

    # Show proposed changes
    print(f"\nFound {unnamed_count} unnamed controls")
    print(f"Suggested names for {len(changes)} controls:\n")

    for change in changes:
        print(f"Control #{change['control_num']}:")
        print(f"  Suggested name: {change['suggested_name']}")
        print(f"  Context: ...{change['before']} [{change['suggested_name']}] {change['after']}...")
        print()

    # Apply changes if not dry run
    if not dry_run:
        print(f"\nApplying {len(changes)} changes...")
        for change in changes:
            try:
                change['control'].Title = change['suggested_name']
                print(f"  ✓ Named control #{change['control_num']} as '{change['suggested_name']}'")
            except Exception as e:
                print(f"  ✗ Error naming control #{change['control_num']}: {e}")

        doc.Save()
        print(f"\n✓ Document saved with {len(changes)} controls renamed")
    else:
        print("\n[DRY RUN] No changes made. Run with dry_run=False to apply changes.")

    doc.Close()
    word.Quit()

    return len(changes)

def main():
    import argparse

    parser = argparse.ArgumentParser(description='Auto-name content controls in Word document')
    parser.add_argument('template', help='Path to Word document template')
    parser.add_argument('--mapping', help='Optional JSON mapping file')
    parser.add_argument('--apply', action='store_true', help='Actually apply changes (default is dry-run)')

    args = parser.parse_args()

    auto_name_controls(
        template_path=args.template,
        mapping_file=args.mapping,
        dry_run=not args.apply
    )

if __name__ == "__main__":
    # Example usage without command line
    TEMPLATE = r"C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx"

    print("=" * 80)
    print("AUTO-NAME CONTENT CONTROLS")
    print("=" * 80)
    print("\nWARNING: This will modify your Word document!")
    print("Make sure you have a backup!\n")

    response = input("Run in DRY-RUN mode to see proposed changes? (y/n): ")

    if response.lower() == 'y':
        auto_name_controls(TEMPLATE, dry_run=True)
    else:
        print("Cancelled.")
