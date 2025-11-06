"""
Analyze CVR template and create a guide for naming content controls.
This script helps map Google Form questions to Word content controls.
"""

import win32com.client
import json

# Google Form questions (from WebFetch analysis)
GOOGLE_FORM_QUESTIONS = [
    {"q": "What is the person name under your care?", "type": "text", "section": "Info"},
    {"q": "Would you please tell me a little about the person under your care?", "type": "longtext", "section": "Info"},
    {"q": "What is your name?", "type": "text", "section": "Info"},
    {"q": "Will you be present during the meeting?", "type": "yesno", "section": "Info"},
    {"q": "If no, please provide the name and relationship of any other individuals...", "type": "text", "section": "Info"},
    {"q": "Please provide any information needed to visit...", "type": "longtext", "section": "Info"},
    {"q": "Do you live together?", "type": "yesno", "section": "Home"},
    {"q": "Does he/she have their own bed?", "type": "yesno", "section": "Home"},
    {"q": "Is hot water available?", "type": "yesno", "section": "Home"},
    {"q": "Does the residence have air conditioning/heating", "type": "yesno", "section": "Home"},
    {"q": "Are most areas accessible to the person under your care?", "type": "yesno", "section": "Home"},
    {"q": "If no, explain which areas are not accessible and why.", "type": "text", "section": "Home"},
    {"q": "Does the individual have access to the following:", "type": "checkbox", "options": ["Telephone", "Television", "Radio", "Computer", "Other"], "section": "Home"},
    {"q": "Does the person under your care live:", "type": "choice", "options": ["Own Home/Apartment", "Lives in your home with you", "Lives in another relative's home", "Nursing home/Assisted Living", "Group Home", "Hospital/Medical Facility", "State School", "Other"], "section": "Home"},
    {"q": "If Ward resides in Other Relatives Home state the name and relationship", "type": "text", "section": "Home"},
    {"q": "If [a] facility, the name of the place the ward lives.", "type": "text", "section": "Home"},
    {"q": "Does the person you care for participate in any of the following", "type": "checkbox", "options": ["School", "Work", "Day Care"], "section": "Activities"},
    {"q": "Does the individual have any visitors?", "type": "yesno", "section": "Activities"},
    {"q": "Does the individual exercise regularly?", "type": "yesno", "section": "Activities"},
    {"q": "If the individual wants to go on trips, is transportation provided?", "type": "yesno", "section": "Activities"},
    {"q": "Are fire extinguishers available?", "type": "yesno", "section": "Safety"},
    {"q": "Is the person under your care able to understand how to use a fire...", "type": "yesno", "section": "Safety"},
    {"q": "Does the individual know where the fire extinguisher is located?", "type": "yesno", "section": "Safety"},
    {"q": "Is the person under your care safe in the home alone?", "type": "yesno", "section": "Safety"},
    {"q": "Can the individual read:", "type": "yesno", "section": "Health"},
    {"q": "Can the individual write:", "type": "yesno", "section": "Health"},
    {"q": "Is the individual oriented in time and space?", "type": "yesno", "section": "Health"},
    {"q": "Does the person in your care need help with any of the following:", "type": "checkbox", "options": ["Bathing", "Dressing", "Eating", "Walking", "Bathroom"], "section": "Health"},
    {"q": "Does the person in your care have any of the following?", "type": "checkbox", "options": ["Hearing impairment", "Speech impairment", "Unwilling to speak", "Unable to speak", "Unresponsive", "Able to communicate with voice", "Communicates with nonverbal gestures", "In bed most/all of the time", "Walk with assistance"], "section": "Health"},
]

# Suggested control names (short, alphanumeric)
SUGGESTED_CONTROL_NAMES = {
    "Do you live together?": "livetogether",
    "Does he/she have their own bed?": "ownbed",
    "Is hot water available?": "hotwater",
    "Does the residence have air conditioning/heating": "hvac",
    "Are most areas accessible to the person under your care?": "accessible",
    "If no, explain which areas are not accessible and why.": "accessible_explain",
    "Does the individual have access to the following:": "access_items",  # Multi checkbox
    "Does the person under your care live:": "living_situation",
    "If Ward resides in Other Relatives Home state the name and relationship": "relative_info",
    "If [a] facility, the name of the place the ward lives.": "facility_name",
    "Does the person you care for participate in any of the following": "activities",  # Multi checkbox
    "Does the individual have any visitors?": "has_visitors",
    "Does the individual exercise regularly?": "exercises",
    "If the individual wants to go on trips, is transportation provided?": "transportation",
    "Are fire extinguishers available?": "fire_extinguisher",
    "Is the person under your care able to understand how to use a fire...": "understands_fire",
    "Does the individual know where the fire extinguisher is located?": "knows_fire_location",
    "Is the person under your care safe in the home alone?": "safe_alone",
    "Can the individual read:": "can_read",
    "Can the individual write:": "can_write",
    "Is the individual oriented in time and space?": "oriented",
    "Does the person in your care need help with any of the following:": "needs_help",  # Multi checkbox
    "Does the person in your care have any of the following?": "conditions",  # Multi checkbox
    "Would you please tell me a little about the person under your care?": "ward_description",
    "Please provide any information needed to visit...": "visit_info",
    "Will you be present during the meeting?": "guardian_present",
    "If no, please provide the name and relationship of any other individuals...": "other_attendees",
}

def analyze_cvr_template(template_path):
    """Analyze CVR template and show what needs to be named"""

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(template_path)

        # Already named controls (from Step 8)
        named_controls = set()
        unnamed_controls = []

        for i, cc in enumerate(doc.ContentControls, 1):
            title = (cc.Title or '').strip()
            tag = (cc.Tag or '').strip()
            name = title or tag

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
            type_name = type_names.get(cc.Type, f'Type{cc.Type}')

            if name:
                named_controls.add(name)
            else:
                # Try to get surrounding text for context
                try:
                    range_text = cc.Range.Text[:50] if cc.Range.Text else ""
                    before_text = ""
                    after_text = ""

                    # Get text before control
                    if cc.Range.Start > 10:
                        temp_range = doc.Range(max(0, cc.Range.Start - 100), cc.Range.Start)
                        before_text = temp_range.Text[-50:] if temp_range.Text else ""

                    # Get text after control
                    if cc.Range.End < doc.Content.End - 10:
                        temp_range = doc.Range(cc.Range.End, min(doc.Content.End, cc.Range.End + 100))
                        after_text = temp_range.Text[:50] if temp_range.Text else ""

                    context = {
                        'index': i,
                        'type': type_name,
                        'before': before_text.strip(),
                        'content': range_text.strip(),
                        'after': after_text.strip(),
                    }
                    unnamed_controls.append(context)
                except:
                    unnamed_controls.append({
                        'index': i,
                        'type': type_name,
                        'before': '',
                        'content': '',
                        'after': '',
                    })

        print("=" * 80)
        print("CVR TEMPLATE ANALYSIS")
        print("=" * 80)

        print(f"\nNamed Controls (already filled by Step 8):")
        print("-" * 80)
        for name in sorted(named_controls):
            print(f"  [OK] {name}")

        print(f"\n\nUnnamed Controls ({len(unnamed_controls)} total):")
        print("-" * 80)
        print("These need to be named for Google Form auto-fill to work.\n")

        for ctrl in unnamed_controls[:20]:  # Show first 20
            # Sanitize text for console output
            before = ctrl['before'].encode('ascii', 'replace').decode('ascii') if ctrl['before'] else ''
            content = ctrl['content'].encode('ascii', 'replace').decode('ascii') if ctrl['content'] else ''
            after = ctrl['after'].encode('ascii', 'replace').decode('ascii') if ctrl['after'] else ''

            print(f"Control #{ctrl['index']} ({ctrl['type']}):")
            if before:
                print(f"  Before: ...{before}")
            if content:
                print(f"  Content: {content}")
            if after:
                print(f"  After: {after}...")
            print()

        if len(unnamed_controls) > 20:
            print(f"... and {len(unnamed_controls) - 20} more unnamed controls")

        # Generate recommended mapping
        print("\n" + "=" * 80)
        print("RECOMMENDED CONTROL NAMES FOR GOOGLE FORM QUESTIONS")
        print("=" * 80)
        print("\nYou need to name content controls in the Word template with these names:")
        print()

        mapping = {}
        for q_data in GOOGLE_FORM_QUESTIONS:
            q = q_data['q']
            if q in SUGGESTED_CONTROL_NAMES:
                control_name = SUGGESTED_CONTROL_NAMES[q]
                mapping[q] = {
                    "cvr_control": control_name,
                    "type": q_data['type'],
                    "section": q_data['section']
                }

                if q_data['type'] == 'checkbox':
                    options = q_data.get('options', [])
                    print(f"[OK] {control_name} ({q_data['type']})")
                    print(f"    Question: {q}")
                    print(f"    Options: {', '.join(options)}")
                    print(f"    Note: Create separate controls for each option OR one text control for comma-separated list")
                elif q_data['type'] == 'yesno':
                    print(f"[OK] {control_name} (checkbox or text)")
                    print(f"    Question: {q}")
                else:
                    print(f"[OK] {control_name} ({q_data['type']})")
                    print(f"    Question: {q}")
                print()

        # Save mapping to JSON
        mapping_path = r"C:\GoogleSync\GuardianShip_App\Config\cvr_google_form_mapping.json"
        with open(mapping_path, 'w') as f:
            json.dump(mapping, f, indent=2)
        print(f"\n[OK] Mapping saved to: {mapping_path}")

        doc.Close(False)

    finally:
        word.Quit()

if __name__ == "__main__":
    template_path = r"C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx"
    analyze_cvr_template(template_path)
