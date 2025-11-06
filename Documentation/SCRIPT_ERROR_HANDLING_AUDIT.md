# Script Error Handling Audit

## Issue

Several automation scripts do NOT properly exit with error code `1` when they fail. This causes the GUI to show "[OK] completed successfully!" even when the script actually failed.

## Root Cause

The GUI checks `process.returncode == 0` to determine success/failure. If a script encounters an error but exits normally (return code 0), the GUI thinks it succeeded.

## Solution

All scripts must call `sys.exit(1)` when they encounter errors or fail to complete their task.

---

## Scripts Audited

### ✅ FIXED - Properly exits with error code on failure

1. **Step 10: google_sheets_cvr_integration_fixed.py**
   - Status: ✅ FIXED
   - Now calls `sys.exit(1)` when authentication or processing fails
   - Shows clear error messages

### ⚠️ NEEDS REVIEW - May not exit with error code

2. **Step 2: cvr_folder_builder.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

3. **Step 3: build_map_sheet.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

4. **Step 4: send_guardian_emails.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

5. **Step 5: add_guardians_to_contacts.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

6. **Step 6: send_confirmation_email.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

7. **Step 7: create_calendar_event.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

8. **Step 8: build_cvr_from_excel_cc_working.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

9. **Step 9: build_court_visitor_summary.py**
   - Status: ⚠️ NEEDS REVIEW
   - Returns early on errors (line 349) but doesn't call sys.exit(1)
   - Should exit with code 1 when workbook not found

11. **Step 11: send_followups_picker.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

12. **Step 12: email_cvr_to_supervisor.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

13. **Step 13: build_payment_forms_sdt.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

14. **Step 14: build_mileage_forms.py**
   - Status: ⚠️ NEEDS REVIEW
   - Check if it exits with code 1 on failures

---

## Recommended Pattern

All scripts should follow this pattern:

```python
import sys

def main():
    try:
        # Do work
        if some_error_condition:
            print("ERROR: Something went wrong")
            return False

        # Success
        return True

    except Exception as e:
        print(f"ERROR: {e}")
        return False

if __name__ == "__main__":
    success = main()

    if not success:
        print("\n❌ FAILED - see errors above")
        sys.exit(1)
    else:
        print("\n✅ SUCCESS")
        sys.exit(0)
```

---

## Testing Checklist

For each script, test:

1. ✅ **Success case** - Should show "[OK] completed successfully!"
2. ✅ **Failure case** - Should show "[FAIL] Process failed with exit code 1"
3. ✅ **Error messages** - Should display clear error messages in output

---

## Priority Fixes

**HIGH PRIORITY** (user reported as confusing):
- ✅ Step 10: google_sheets_cvr_integration_fixed.py - FIXED
- ⚠️ Step 9: build_court_visitor_summary.py - NEEDS FIX

**MEDIUM PRIORITY** (test during normal use):
- All other steps that involve API calls or external resources

**LOW PRIORITY** (less likely to fail silently):
- Steps that only manipulate local files

---

## Notes

- The GUI already has proper error handling (lines 738-747 in guardianship_app.py)
- The issue is scripts not returning proper exit codes
- This is a common oversight when converting interactive scripts to background automation
