# Testing Path Fixes - Complete Strategy

**Date:** November 5, 2024
**Problem:** How do I test if dynamic paths work when my machine has hardcoded paths?
**Solution:** Create test installations that simulate end-user environments

---

## The Testing Problem

### Your Current Situation:
```
C:\GoogleSync\GuardianShip_App\     ← Your development location
└── All files with HARDCODED paths pointing here
```

### The Question:
"How do I test if `app_paths.py` works when my machine has the hardcoded `C:\GoogleSync\GuardianShip_App\` path?"

### The Answer:
**Create SEPARATE test installations in different locations and run the app from there!**

---

## Testing Strategy - 3 Test Locations

### Test Location 1: Simulated End User Install
```
C:\CourtVisitorApp_TEST\            ← Simulates default end-user install
```

### Test Location 2: Different Drive
```
D:\TestApps\CourtVisitorApp\        ← Simulates user choosing different drive
```

### Test Location 3: User Documents Folder
```
C:\Users\[YourName]\Documents\CourtVisitorApp_TEST\
                                    ← Simulates corporate/restricted environment
```

**If app works from all 3 locations, it will work for ANY end user!**

---

## Step-by-Step Testing Process

### Phase 1: Create Backup (FIRST - DO THIS NOW!)

```bash
cd C:\GoogleSync\GuardianShip_App
python Scripts\create_verified_backup.py --description "Before path testing - original working state"
```

**Why first:** If testing breaks something, you can restore immediately.

---

### Phase 2: Create Test Installation

#### Option A: Manual Test Installation (Recommended for First Test)

```bash
# 1. Create test directory
mkdir C:\CourtVisitorApp_TEST

# 2. Copy entire app to test location
xcopy C:\GoogleSync\GuardianShip_App C:\CourtVisitorApp_TEST /E /I /EXCLUDE:exclude.txt

# 3. Navigate to test location
cd C:\CourtVisitorApp_TEST

# 4. Run app from test location
python guardianship_app.py
```

**What to exclude (create exclude.txt):**
```
__pycache__
.pyc
.git
.vscode
GuardianShip_App_Backups
```

#### Option B: Use Distribution Package (More Realistic)

```bash
# 1. Create distribution package from your dev folder
cd C:\GoogleSync\GuardianShip_App
python create_distribution_package.py

# 2. Extract to test location
# Extract: Distribution\CourtVisitorApp_v1.0.0.zip
# To: C:\CourtVisitorApp_TEST\

# 3. Test from there
cd C:\CourtVisitorApp_TEST
python guardianship_app.py
```

---

### Phase 3: Test app_paths.py Detection

**Before fixing ANY files, test that path detection works:**

```bash
# Navigate to test installation
cd C:\CourtVisitorApp_TEST

# Test app_paths.py
python Scripts\app_paths.py
```

**Expected Output (GOOD):**
```
======================================================================
Court Visitor App - Path Configuration
======================================================================

App Root: C:\CourtVisitorApp_TEST
Valid: True

--- Critical Paths ---
✓ Excel Database
✓ Config Directory
✓ App Data Directory
✓ New Files Directory
✓ Scripts Directory
✓ Automation Directory

--- Key Files ---
Excel: C:\CourtVisitorApp_TEST\App Data\ward_guardian_info.xlsx
CVR Template: C:\CourtVisitorApp_TEST\Templates\Court Visitor Report fillable new.docx
Config: C:\CourtVisitorApp_TEST\Config\app_settings.json
```

**Bad Output (PROBLEM):**
```
App Root: C:\CourtVisitorApp_TEST
Valid: False

✗ Excel Database
✗ Config Directory
```

**If you see ✗ marks:** Path detection isn't finding files - need to fix app_paths.py before continuing.

---

### Phase 4: Test Individual Scripts After Fixing

**After you fix a script's paths, test it from BOTH locations:**

#### Test 1: From Development Location (Should Still Work)
```bash
# Your dev location (should work as before)
cd C:\GoogleSync\GuardianShip_App
python guardianship_app.py
# Test Step 1, Step 2, etc.
```

#### Test 2: From Test Location (NEW - Must Also Work)
```bash
# Test location (must work now!)
cd C:\CourtVisitorApp_TEST
python guardianship_app.py
# Test same steps
```

**Success = Works from BOTH locations!**

---

## Complete Test Workflow (For Each File You Fix)

### Step 1: Backup Current State
```bash
cd C:\GoogleSync\GuardianShip_App
python Scripts\create_verified_backup.py --description "Before fixing [filename]"
```

### Step 2: Fix ONE File
```bash
# Edit the file to use app_paths.py
# Example: guardianship_app.py
```

### Step 3: Test From Dev Location
```bash
cd C:\GoogleSync\GuardianShip_App
python guardianship_app.py
# Test relevant functionality
```

**If it breaks:** Restore immediately!
```bash
copy guardianship_app_BEFORE_PATH_FIX_20241105.py guardianship_app.py
```

### Step 4: Copy Fixed File to Test Location
```bash
# Copy just the fixed file
copy C:\GoogleSync\GuardianShip_App\guardianship_app.py C:\CourtVisitorApp_TEST\guardianship_app.py
```

### Step 5: Test From Test Location
```bash
cd C:\CourtVisitorApp_TEST
python guardianship_app.py
# Test same functionality
```

### Step 6: Verify Both Work
- ✅ Works from `C:\GoogleSync\GuardianShip_App\` (your dev)
- ✅ Works from `C:\CourtVisitorApp_TEST\` (simulated end user)

**Only if BOTH pass:** Mark file as complete and move to next file.

### Step 7: Backup Successful Change
```bash
cd C:\GoogleSync\GuardianShip_App
python Scripts\create_verified_backup.py --description "After fixing [filename] - tested both locations OK"
```

---

## Automated Test Script

Let me create a script that tests everything automatically:

**Create: `Scripts\test_from_multiple_locations.py`**

```python
"""
Test Court Visitor App from multiple installation locations.
Simulates end-user installations.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

class MultiLocationTester:
    def __init__(self):
        self.dev_location = Path(r"C:\GoogleSync\GuardianShip_App")
        self.test_locations = [
            Path(r"C:\CourtVisitorApp_TEST"),
            Path(r"D:\CourtVisitorApp_TEST"),  # If D: drive exists
            Path.home() / "Documents" / "CourtVisitorApp_TEST",
        ]

    def setup_test_location(self, test_location):
        """Copy app to test location."""
        print(f"\n📋 Setting up test location: {test_location}")

        # Clean existing
        if test_location.exists():
            print(f"   Removing existing test location...")
            shutil.rmtree(test_location)

        # Copy app
        print(f"   Copying app files...")
        shutil.copytree(
            self.dev_location,
            test_location,
            ignore=shutil.ignore_patterns(
                '__pycache__', '*.pyc', '.git', '.vscode',
                'GuardianShip_App_Backups', '*_TEST'
            )
        )

        print(f"   ✅ Test location ready")

    def test_path_detection(self, location):
        """Test if app_paths.py correctly detects location."""
        print(f"\n🔍 Testing path detection from: {location}")

        os.chdir(location)

        # Run app_paths.py
        result = subprocess.run(
            [sys.executable, "Scripts/app_paths.py"],
            capture_output=True,
            text=True
        )

        # Check output
        if "Valid: True" in result.stdout:
            print(f"   ✅ Path detection PASSED")
            return True
        else:
            print(f"   ❌ Path detection FAILED")
            print(result.stdout)
            return False

    def test_app_launch(self, location):
        """Test if main app launches from location."""
        print(f"\n🚀 Testing app launch from: {location}")

        os.chdir(location)

        # Try to import and initialize
        try:
            # Add to path
            sys.path.insert(0, str(location / "Scripts"))

            # Try importing path system
            from app_paths import get_app_paths

            paths = get_app_paths()

            # Verify key paths exist
            checks = {
                "App Root": paths.APP_ROOT.exists(),
                "Config Dir": paths.CONFIG_DIR.exists(),
                "App Data Dir": paths.APP_DATA_DIR.exists(),
            }

            all_passed = all(checks.values())

            for name, passed in checks.items():
                status = "✅" if passed else "❌"
                print(f"   {status} {name}")

            return all_passed

        except Exception as e:
            print(f"   ❌ Launch failed: {e}")
            return False

    def run_all_tests(self):
        """Run complete test suite."""
        print("="*70)
        print("Court Visitor App - Multi-Location Test Suite")
        print("="*70)

        results = {}

        # Test each location
        for test_location in self.test_locations:
            # Skip if drive doesn't exist
            if not test_location.drive or not Path(test_location.drive).exists():
                print(f"\n⏭️  Skipping {test_location} (drive doesn't exist)")
                continue

            try:
                # Setup
                self.setup_test_location(test_location)

                # Test path detection
                path_ok = self.test_path_detection(test_location)

                # Test app launch
                launch_ok = self.test_app_launch(test_location)

                # Record results
                results[str(test_location)] = {
                    'path_detection': path_ok,
                    'app_launch': launch_ok,
                    'overall': path_ok and launch_ok
                }

            except Exception as e:
                print(f"\n❌ Error testing {test_location}: {e}")
                results[str(test_location)] = {
                    'path_detection': False,
                    'app_launch': False,
                    'overall': False,
                    'error': str(e)
                }

        # Summary
        print("\n" + "="*70)
        print("TEST RESULTS SUMMARY")
        print("="*70)

        for location, result in results.items():
            status = "✅ PASS" if result['overall'] else "❌ FAIL"
            print(f"\n{status} - {location}")
            if not result['overall']:
                print(f"   Path Detection: {'✅' if result.get('path_detection') else '❌'}")
                print(f"   App Launch: {'✅' if result.get('app_launch') else '❌'}")
                if 'error' in result:
                    print(f"   Error: {result['error']}")

        # Overall result
        all_passed = all(r['overall'] for r in results.values())

        print("\n" + "="*70)
        if all_passed:
            print("✅ ALL TESTS PASSED - App works from all locations!")
        else:
            print("❌ SOME TESTS FAILED - Fix issues before distributing")
        print("="*70)

        return all_passed

if __name__ == "__main__":
    tester = MultiLocationTester()
    success = tester.run_all_tests()
    sys.exit(0 if success else 1)
```

---

## Using the Automated Test Script

### Run Tests After Each File Fix:

```bash
cd C:\GoogleSync\GuardianShip_App
python Scripts\test_from_multiple_locations.py
```

**What it does:**
1. Creates test installations in 3 different locations
2. Tests path detection from each
3. Tests app launch from each
4. Shows summary of results

**Expected Output (GOOD):**
```
✅ PASS - C:\CourtVisitorApp_TEST
✅ PASS - D:\CourtVisitorApp_TEST
✅ PASS - C:\Users\You\Documents\CourtVisitorApp_TEST

✅ ALL TESTS PASSED - App works from all locations!
```

**If you see failures:**
```
❌ FAIL - C:\CourtVisitorApp_TEST
   Path Detection: ❌
   App Launch: ❌

❌ SOME TESTS FAILED - Fix issues before distributing
```

This means the path fix isn't working yet - need to debug.

---

## Quick Manual Test (Fastest)

If you want to test quickly without automation:

```bash
# 1. Create test folder
mkdir C:\CourtVisitorApp_TEST

# 2. Copy everything
xcopy C:\GoogleSync\GuardianShip_App C:\CourtVisitorApp_TEST /E /I

# 3. Go there
cd C:\CourtVisitorApp_TEST

# 4. Test path detection
python Scripts\app_paths.py

# 5. Try launching app
python guardianship_app.py

# 6. Test Step 1 (OCR)
# Place test PDF in New Files
# Click Step 1 button
# Check if it works
```

**If Step 1 works from test location → Path fix is working!**

---

## Testing Checklist for Each Fixed File

After fixing a file's paths:

- [ ] Backup created before changes
- [ ] File updated to use `app_paths.py`
- [ ] Test from dev location (`C:\GoogleSync\GuardianShip_App\`)
- [ ] Works from dev location ✅
- [ ] Copy to test location (`C:\CourtVisitorApp_TEST\`)
- [ ] Test from test location
- [ ] Works from test location ✅
- [ ] Run automated test script
- [ ] All locations pass ✅
- [ ] Backup successful change
- [ ] Document in change log

**Only proceed to next file if ALL checkboxes are ✅**

---

## What "Working" Means for Each Test

### For app_paths.py Test:
- ✅ Shows correct App Root (test location, not dev location)
- ✅ All critical paths marked with ✓
- ✅ Excel path points to test location
- ✅ No errors or exceptions

### For Main App Test:
- ✅ App launches without errors
- ✅ All 14 step buttons visible
- ✅ No "File not found" errors
- ✅ Can navigate interface

### For Step Test (e.g., Step 1):
- ✅ Button clicks without error
- ✅ Finds PDF files in New Files folder
- ✅ Processes PDF successfully
- ✅ Updates Excel in test location
- ✅ Shows success message

---

## Summary

**Your Question:** "How do I test if paths work when my machine has hardcoded paths?"

**Answer:** Create separate test installations and run the app from there!

**Simple Process:**
1. Copy app to `C:\CourtVisitorApp_TEST\`
2. Run app from test location
3. If it works there → paths are dynamic ✅
4. If it breaks → still hardcoded ❌

**Why this works:**
- Test location doesn't have `C:\GoogleSync\GuardianShip_App\`
- So hardcoded paths would fail
- If it works, paths must be dynamic!

---

## Next Steps

1. **First:** Create backup
   ```bash
   cd C:\GoogleSync\GuardianShip_App
   python Scripts\create_verified_backup.py --description "Before path testing"
   ```

2. **Second:** Test current app_paths.py
   ```bash
   mkdir C:\CourtVisitorApp_TEST
   xcopy C:\GoogleSync\GuardianShip_App C:\CourtVisitorApp_TEST /E /I
   cd C:\CourtVisitorApp_TEST
   python Scripts\app_paths.py
   ```

3. **Third:** If path detection works, start fixing files one by one

**Ready to create that first backup and test location?**
