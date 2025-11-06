"""
Fix all automation script paths to use GuardianShip_App instead of GuardianShip_Easy_App
"""
import os
import re
from pathlib import Path

# Base directory
APP_DIR = Path(r"C:\GoogleSync\GuardianShip_App")
AUTOMATION_DIR = APP_DIR / "Automation"

# Path replacements
REPLACEMENTS = [
    # Old base path -> New base path
    (r"C:\\GoogleSync\\GuardianShip_Easy_App", r"C:\\GoogleSync\\GuardianShip_App"),
    (r"C:/GoogleSync/GuardianShip_Easy_App", r"C:/GoogleSync/GuardianShip_App"),

    # Relative Excel paths -> Absolute paths
    (r'WORKBOOK_PATH\s*=\s*r?"ward_guardian_info\.xlsx"', r'WORKBOOK_PATH = r"C:\\GoogleSync\\GuardianShip_App\\App Data\\ward_guardian_info.xlsx"'),
    (r'EXCEL_PATH\s*=\s*r?"ward_guardian_info\.xlsx"', r'EXCEL_PATH = r"C:\\GoogleSync\\GuardianShip_App\\App Data\\ward_guardian_info.xlsx"'),

    # Guardian base paths
    (r'GUARDIAN_BASE\s*=\s*r?"C:\\GoogleSync\\GuardianShip_Easy_App"', r'GUARDIAN_BASE = r"C:\\GoogleSync\\GuardianShip_App"'),
]

def fix_script(script_path):
    """Fix paths in a single script"""
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()

        original = content

        # Apply all replacements
        for old_pattern, new_value in REPLACEMENTS:
            content = re.sub(old_pattern, new_value, content)

        # Only write if changed
        if content != original:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"[FIXED] {script_path.relative_to(APP_DIR)}")
            return True
        else:
            print(f"[SKIP] No changes needed: {script_path.relative_to(APP_DIR)}")
            return False

    except Exception as e:
        print(f"[ERROR] Error fixing {script_path}: {e}")
        return False

def main():
    print("=" * 70)
    print("Fixing All Automation Script Paths")
    print("=" * 70)
    print()

    # Find all .py files in Automation folder
    scripts = list(AUTOMATION_DIR.rglob("*.py"))
    print(f"Found {len(scripts)} Python scripts in Automation folder")
    print()

    fixed_count = 0
    for script in scripts:
        if fix_script(script):
            fixed_count += 1

    print()
    print("=" * 70)
    print(f"Complete! Fixed {fixed_count} out of {len(scripts)} scripts")
    print("=" * 70)

if __name__ == "__main__":
    main()
