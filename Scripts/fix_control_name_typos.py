"""
Fix the 5 typos in CVR template control names
"""

import win32com.client

TEMPLATE_PATH = r"C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx"

# Typo fixes
FIXES = {
    'acessible_yes': 'accessible_yes',
    'fire_extinguishers_no': 'fire_extinguisher_no',
    'has_visitor_no': 'has_visitors_no',
    'attend_yes': 'guardian_present_yes',
    'attend_no': 'guardian_present_no',
}

def fix_typos():
    print("Opening CVR template...")

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    doc = word.Documents.Open(TEMPLATE_PATH)

    fixed_count = 0

    for cc in doc.ContentControls:
        name = (cc.Title or '').strip()

        if name in FIXES:
            new_name = FIXES[name]
            cc.Title = new_name
            print(f"  Fixed: '{name}' -> '{new_name}'")
            fixed_count += 1

    if fixed_count > 0:
        doc.Save()
        print(f"\n✓ Fixed {fixed_count} control names and saved template")
    else:
        print("\nNo typos found to fix")

    doc.Close()
    word.Quit()

if __name__ == "__main__":
    print("=" * 80)
    print("FIX CONTROL NAME TYPOS")
    print("=" * 80)

    fix_typos()

    print("\nDone!")
