"""
Debug why Namico) line isn't being detected in Toney ARP
"""
import sys
import os
import re
from pathlib import Path

# Read Document AI output
doc_ai_file = Path(r"C:\GoogleSync\GuardianShip_App\document_ai_output.txt")
with open(doc_ai_file, 'r', encoding='utf-8') as f:
    text = f.read()

print("=" * 80)
print("SEARCHING FOR GUARDIAN SECTION IN DOCUMENT AI OUTPUT")
print("=" * 80)

# Look for GUARDIAN keyword
lines = text.split('\n')
for i, line in enumerate(lines):
    if 'GUARDIAN' in line.upper() or 'WARD' in line.upper() and i < 50:
        print(f"\nLine {i}: {line[:100]}")

# Look for the Namico line specifically
print("\n" + "=" * 80)
print("SEARCHING FOR 'Namico)' LINE")
print("=" * 80)

for i, line in enumerate(lines):
    if 'namico' in line.lower() or 'DERRICK' in line.upper() or 'SARAJANE' in line.upper():
        print(f"\nLine {i}: {line}")
        print(f"  Contains 'nam': {bool(re.search(r'nam', line, re.I))}")
        print(f"  Pattern match nam(?:ico?|e)?\\)\\s*: {bool(re.search(r'nam(?:ico?|e)?\\)\\s*', line, re.I))}")
        print(f"  Pattern match name(?:\\(s\\))?s?\\s*: {bool(re.search(r'name(?:\\(s\\))?s?\\s*(?:[:\\-]|\\s)', line, re.I))}")

# Now let's simulate the guardian section slicer
print("\n" + "=" * 80)
print("SIMULATING GUARDIAN SECTION SLICER")
print("=" * 80)

# Simple version - find GUARDIAN to end or next major section
start_idx = None
end_idx = None

for i, line in enumerate(lines):
    if start_idx is None and 'GUARDIAN' in line.upper():
        start_idx = i
        print(f"Found GUARDIAN section start at line {i}: {line[:50]}")

    if start_idx is not None and i > start_idx + 3:
        # Look for section endings
        if any(keyword in line.upper() for keyword in ['VISIT', 'SOCIAL', 'HEALTH', 'MEDICAL', 'FINAL REPORT']):
            end_idx = i
            print(f"Found section end at line {i}: {line[:50]}")
            break

if start_idx and not end_idx:
    end_idx = min(start_idx + 50, len(lines))

if start_idx:
    guardian_section = lines[start_idx:end_idx]
    print(f"\nGuardian section spans lines {start_idx} to {end_idx}")
    print(f"Total lines in section: {len(guardian_section)}")

    print("\n" + "=" * 80)
    print("GUARDIAN SECTION CONTENTS:")
    print("=" * 80)
    for i, line in enumerate(guardian_section[:30], start=start_idx):
        print(f"{i:4d}: {line}")

    # Check if Namico line is in this section
    print("\n" + "=" * 80)
    print("IS 'Namico)' IN GUARDIAN SECTION?")
    print("=" * 80)
    namico_found = any('namico' in line.lower() for line in guardian_section)
    derrick_found = any('DERRICK' in line.upper() for line in guardian_section)
    print(f"Namico found: {namico_found}")
    print(f"DERRICK found: {derrick_found}")
