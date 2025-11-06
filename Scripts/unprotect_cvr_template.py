"""
Unprotect the CVR template so that CVRs created from it are editable.
Run this once, then Step 8 will create unprotected CVRs.
"""

import win32com.client
import os

template_path = r"C:\GoogleSync\GuardianShip_App\Templates\Court Visitor Report fillable new.docx"

print(f"Opening template: {template_path}")
print()

try:
    word = win32com.client.GetActiveObject("Word.Application")
    print("Using existing Word instance")
except:
    word = win32com.client.Dispatch("Word.Application")
    print("Created new Word instance")

try:
    word.Visible = True
except:
    pass  # Ignore visibility errors

doc = word.Documents.Open(os.path.abspath(template_path))

# Check if protected
if doc.ProtectionType != -1:  # -1 = wdNoProtection
    print("Template IS protected. Removing protection...")
    doc.Unprotect()
    print("Protection removed!")

    # Save the template
    doc.Save()
    print("Template saved without protection.")
else:
    print("Template is NOT protected (already unprotected).")

doc.Close()
word.Quit()

print()
print("Done! Now when Step 8 creates CVRs, they will be unprotected.")
print("Step 10 can fill them, and then Step 10 will protect them at the end.")
