"""
Show which controls are missing from the CVR
"""
import win32com.client
import json

cvr_path = r"C:\GoogleSync\GuardianShip_App\New Files\Ximenez, Jose - 99-071988\Ximenez, Jose, 99-071988 Court Visitor Report.docx"
config_path = r"C:\GoogleSync\GuardianShip_App\Config\cvr_google_form_mapping.json"

# Get actual controls in CVR
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(cvr_path)

actual_controls = set()
for cc in doc.ContentControls:
    name = (cc.Title or cc.Tag or '').strip()
    if name:
        actual_controls.add(name.lower())

doc.Close(False)
word.Quit()

# Load config to see what we expect
with open(config_path) as f:
    config = json.load(f)

missing = [
    "ward_description",
    "other_attendees",
    "visit_info",
    "livetogether_no",
    "accessible_explain",
    "exercises_yes",
    "understands_fire_no",
    "knows_fire_location_no",
    "safe_alone_yes"
]

print("MISSING CONTROLS - What to do:")
print("=" * 70)
print()

for ctrl in missing:
    # Check if a similar name exists
    similar = [name for name in actual_controls if ctrl.replace('_yes', '').replace('_no', '') in name]

    if similar:
        print(f"✓ {ctrl}")
        print(f"   Found similar: {', '.join(similar)}")
        print(f"   ACTION: Rename in CVR to '{ctrl}' or update config")
    else:
        print(f"✗ {ctrl}")
        print(f"   Not found at all in CVR")
        print(f"   ACTION: Add this content control to your CVR template")
    print()

print()
print("ALL CONTROLS IN CVR (for reference):")
print("-" * 70)
for name in sorted(actual_controls):
    print(f"  {name}")
