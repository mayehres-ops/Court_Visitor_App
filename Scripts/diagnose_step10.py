"""
Diagnostic script to check why Step 10 isn't filling Google Form fields.
This will compare what we EXPECT vs what's ACTUALLY in the CVR.
"""

import win32com.client
import json

# The CVR that Step 10 just tried to fill
cvr_path = r"C:\GoogleSync\GuardianShip_App\New Files\Ximenez, Jose - 99-071988\Ximenez, Jose, 99-071988 Court Visitor Report.docx"

# The mapping config
config_path = r"C:\GoogleSync\GuardianShip_App\Config\cvr_google_form_mapping.json"

print("=" * 80)
print("STEP 10 DIAGNOSTIC TOOL")
print("=" * 80)
print()

# Load the mapping config
with open(config_path, 'r') as f:
    config = json.load(f)

print(f"1. CONFIG FILE ({config_path})")
print("-" * 80)
print("Expected control names from config (first 10):")
count = 0
for question, mapping_info in config.items():
    if question.startswith('_'):
        continue
    control_name = mapping_info.get('cvr_control', '')
    field_type = mapping_info.get('type', 'text')
    print(f"  - {control_name} (type: {field_type})")
    count += 1
    if count >= 10:
        break
print()

# Open the CVR and list ALL content controls
print(f"2. CVR DOCUMENT ({cvr_path})")
print("-" * 80)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

doc = word.Documents.Open(cvr_path)

# Get all content control names
actual_controls = {}
for cc in doc.ContentControls:
    name = (cc.Title or cc.Tag or '').strip()
    if name:
        actual_controls[name.lower()] = {
            'name': name,
            'type': cc.Type,  # 5 = checkbox, others = text
            'value': cc.Range.Text[:50] if cc.Type != 5 else ('Checked' if cc.Checked else 'Unchecked')
        }

print(f"Found {len(actual_controls)} named content controls in CVR:")
print()
for name in sorted(list(actual_controls.keys())[:20]):  # First 20
    ctrl = actual_controls[name]
    type_str = 'checkbox' if ctrl['type'] == 5 else 'text'
    print(f"  - {ctrl['name']} ({type_str}) = {ctrl['value']}")
print()

# Now check which expected controls are MISSING
print("3. MAPPING CHECK")
print("-" * 80)
print("Checking if expected controls exist in CVR...")
print()

missing = []
found = []

for question, mapping_info in config.items():
    if question.startswith('_'):
        continue

    control_name = mapping_info.get('cvr_control', '')
    field_type = mapping_info.get('type', 'text')

    # For yesno, check for _yes and _no variants
    if field_type == 'yesno':
        if f"{control_name}_yes".lower() in actual_controls:
            found.append(f"{control_name}_yes ✓")
        else:
            missing.append(f"{control_name}_yes")

        if f"{control_name}_no".lower() in actual_controls:
            found.append(f"{control_name}_no ✓")
        else:
            missing.append(f"{control_name}_no")
    else:
        if control_name.lower() in actual_controls:
            found.append(f"{control_name} ✓")
        else:
            missing.append(control_name)

print(f"FOUND {len(found)} expected controls:")
for name in found[:10]:
    print(f"  ✓ {name}")
if len(found) > 10:
    print(f"  ... and {len(found) - 10} more")
print()

print(f"MISSING {len(missing)} expected controls:")
for name in missing[:15]:
    print(f"  ✗ {name}")
if len(missing) > 15:
    print(f"  ... and {len(missing) - 15} more")
print()

# Clean up
doc.Close(False)
word.Quit()

print("=" * 80)
print("DIAGNOSIS COMPLETE")
print("=" * 80)
print()

if len(missing) > len(found):
    print("⚠️  PROBLEM: Most expected controls are MISSING from the CVR!")
    print("   This means the control names in the CVR don't match the mapping config.")
    print()
    print("   SOLUTION: You need to rename the content controls in your CVR template")
    print("   to match what's in the config file.")
elif len(found) > 0:
    print("✓ GOOD: Found some expected controls in the CVR.")
    print()
    print("  Next step: Check why the script isn't filling them.")
    print("  The filling logic might have an error.")
