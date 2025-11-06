# GuardianShip Easy App - Release Preparation Checklist

## ✅ COMPLETED ITEMS

### Working Features (Verified)
- [x] Step 1: OCR Extract Guardian Info - Working
- [x] Step 2: Build Folders - Working
- [x] Step 3: Build Map Sheet - Working
- [x] Step 8: Create CVR - Working
- [x] Step 9: Fill CVR Content Controls - Working
- [x] Step 10: Complete CVR with Google Form Data - **FIXED & WORKING** (16/17 fields)
- [x] Step 11: Send Follow-up Emails - Working
- [x] Step 12: Email CVR to Supervisor - Working
- [x] Step 13: Build Payment Forms - Working (perfect)
- [x] Step 14: Build Mileage Forms - Working (charm)

---

## 📋 TO-DO BEFORE SHARING

### 1. **Remove Personal/Test Data** ⚠️ CRITICAL

#### A. Clean Excel File (`ward_guardian_info.xlsx`)
- [ ] Remove all real guardian/ward data
- [ ] Add 2-3 sample rows with fake data (e.g., "John Smith", "123 Test St")
- [ ] Keep column headers intact
- [ ] Save as template version

#### B. Clean Templates Folder
- [ ] `Court Visitor Report fillable new.docx` - Remove any filled test data
- [ ] `Payment Form Template.docx` - Remove test data
- [ ] `Mileage Reimbursement Template.docx` - Remove test data
- [ ] Ensure all templates are blank/placeholder data only

#### C. Clean Test Files
- [ ] Delete all files in `C:\GoogleSync\GuardianShip_App\New Files\*` (test CVRs)
- [ ] Delete all files in `C:\GoogleSync\GuardianShip_App\Completed\*`
- [ ] Delete all files in `C:\GoogleSync\GuardianShip_App\App Data\Inbox\*` (test PDFs)
- [ ] Keep folder structure, just empty the contents

#### D. Remove Personal API Keys
- [ ] Delete `C:\configlocal\API\client_secret_*.json` (if bundling)
- [ ] Delete `C:\configlocal\API\*_token.pickle` (OAuth tokens)
- [ ] Delete `C:\configlocal\Keys\google_maps_api_key.txt` (if exists)
- [ ] Add `.gitignore` or equivalent to prevent API keys from being shared

### 2. **Fix Error Handling & Success Messages**

#### Scripts to Update:
- [ ] **Step 1** (`guardian_extractor_claudecode20251023_bestever_11pm.py`)
  - Verify exit code 1 on failure, 0 on success
  - Add clear "SUCCESS" or "FAILED" message at end

- [ ] **Step 2** (`cvr_folder_builder.py`)
  - Verify proper exit codes
  - Add success/failure summary

- [ ] **Step 3** (`build_map_sheet.py`)
  - Verify exit codes
  - Clear success message

- [ ] **Step 4** (`send_guardian_emails.py`)
  - Already has good error handling? Verify

- [ ] **Step 5** (`add_guardians_to_contacts.py`)
  - Verify exit codes

- [ ] **Step 6** (`send_confirmation_email.py`)
  - Verify exit codes

- [ ] **Step 7** (`create_calendar_event.py`)
  - Verify exit codes

- [ ] **Step 8** (`build_cvr_from_excel_cc_working.py`)
  - Already good? Verify

- [ ] **Step 9** (`build_court_visitor_summary.py`)
  - Verify exit codes

- [ ] **Step 10** (`google_sheets_cvr_integration_fixed.py`)
  - ✅ Already returns proper exit codes
  - ✅ Has clear SUCCESS/FAILED messages

- [ ] **Step 11** (`send_followups_picker.py`)
  - Verify exit codes

- [ ] **Step 12** (`email_cvr_to_supervisor.py`)
  - Verify exit codes

- [ ] **Step 13** (`build_payment_forms_sdt.py`)
  - Verify exit codes

- [ ] **Step 14** (`build_mileage_forms.py`)
  - Verify exit codes

### 3. **Test All Sidebar Buttons** ✨

#### Quick Access Buttons (Top Section):
- [ ] **📖 Getting Started** - Opens getting started guide?
- [ ] **📊 Excel File** - Opens `ward_guardian_info.xlsx`?
- [ ] **📁 Guardian Folders** - Opens `New Files` folder?

#### API Setup Buttons:
- [ ] **🔧 Setup Vision API** - Opens Vision API setup wizard?
- [ ] **🗺️ Setup Maps API** - Opens Maps API setup wizard?
- [ ] **📧 Setup Gmail API** - Opens Gmail API setup wizard?
- [ ] **👥 Setup People/Calendar** - Opens People/Calendar setup wizard?

#### Help & Support Buttons:
- [ ] **🆘 Live Tech Support** - Opens AI help window?
- [ ] **👥 Contacts** - Opens Windows Contacts?
- [ ] **📧 Email** - Opens default email client?
- [ ] **❓ Help** - Shows help documentation?
- [ ] **📖 Manual** - Opens user manual?
- [ ] **🐛 Report Bug** - Opens bug report form/email?
- [ ] **💡 Request Feature** - Opens feature request form/email?

### 4. **Documentation**

- [ ] Create `README.md` with:
  - System requirements (Windows, Python 3.x, Office)
  - Installation steps
  - Google API setup instructions
  - Basic usage guide
  - Troubleshooting section

- [ ] Create `USER_MANUAL.md` with:
  - Detailed step-by-step workflow
  - Screenshots of each step
  - Common errors and solutions
  - FAQ section

- [ ] Create `SETUP_GUIDE.md` for:
  - Google Cloud Console setup
  - API enablement steps
  - OAuth consent screen configuration
  - Service account/OAuth client creation

- [ ] Create `CHANGELOG.md` tracking:
  - Version history
  - Bug fixes (especially Step 10 checkbox fix!)
  - New features

### 5. **Code Cleanup**

- [ ] Remove debug `print()` statements from all scripts
- [ ] Add proper logging instead of print statements
- [ ] Remove commented-out old code
- [ ] Add docstrings to all functions
- [ ] Ensure consistent code formatting

### 6. **Configuration Files**

- [ ] Create `config.example.json` showing structure:
  ```json
  {
    "google_sheets_id": "YOUR_SPREADSHEET_ID_HERE",
    "supervisor_email": "supervisor@example.com",
    "smtp_settings": {
      "server": "smtp.gmail.com",
      "port": 587
    }
  }
  ```

- [ ] Update paths to be relative or configurable:
  - Change hardcoded `C:\GoogleSync\GuardianShip_App` paths
  - Use `os.path.join()` for cross-platform compatibility
  - Add config file for user-specific paths

### 7. **Dependency Management**

- [ ] Create `requirements.txt`:
  ```
  pywin32
  openpyxl
  google-api-python-client
  google-auth-httplib2
  google-auth-oauthlib
  Pillow
  ```

- [ ] Test installation on clean machine
- [ ] Document Python version requirement (3.8+?)

### 8. **Packaging**

- [ ] Create installer script or instructions
- [ ] Bundle templates in proper folder structure
- [ ] Include sample data files
- [ ] Create shortcuts for easy launch

### 9. **Testing on Clean System**

- [ ] Test on machine without GuardianShip App installed
- [ ] Verify all folders are created on first run
- [ ] Test each step with sample data
- [ ] Verify all sidebar buttons work
- [ ] Test API setup wizards
- [ ] Check error messages are helpful

### 10. **Security Review**

- [ ] Ensure no API keys in source code
- [ ] Add warning about API costs to documentation
- [ ] Add privacy notice about data handling
- [ ] Verify OAuth tokens are stored securely
- [ ] Check file permissions on sensitive folders

### 11. **Legal & Licensing**

- [ ] Update copyright notices
- [ ] Choose license (MIT, GPL, Proprietary, etc.)
- [ ] Add LICENSE file
- [ ] Add EULA if commercial
- [ ] Add privacy policy
- [ ] Add terms of service

---

## 🔍 CURRENT STATUS OF SIDEBAR BUTTONS

Based on code review, here are the sidebar buttons and their likely status:

### ✅ Likely Working:
- **📊 Excel File** - Opens `ward_guardian_info.xlsx` (simple file open)
- **📁 Guardian Folders** - Opens `New Files` folder (simple folder open)
- **👥 Contacts** - Opens Windows Contacts
- **📧 Email** - Opens default email client

### ⚠️ Need Testing:
- **📖 Getting Started** - Custom dialog/window
- **🔧 Setup Vision API** - Multi-step wizard
- **🗺️ Setup Maps API** - Multi-step wizard
- **📧 Setup Gmail API** - Multi-step wizard
- **👥 Setup People/Calendar** - Multi-step wizard
- **🆘 Live Tech Support** - AI help integration
- **❓ Help** - Help documentation window
- **📖 Manual** - User manual window
- **🐛 Report Bug** - Bug report form
- **💡 Request Feature** - Feature request form

---

## 📊 PRIORITY ORDER

### 🔥 HIGH PRIORITY (Must Do):
1. Remove all personal data from Excel, templates, and test files
2. Remove/secure API keys and tokens
3. Test all 14 automation steps work end-to-end
4. Create basic README with setup instructions

### 🟡 MEDIUM PRIORITY (Should Do):
5. Fix exit codes and success/failure messages in all scripts
6. Test all sidebar buttons
7. Create comprehensive documentation
8. Add requirements.txt and setup instructions

### 🟢 LOW PRIORITY (Nice to Have):
9. Code cleanup and logging
10. Create installer
11. Test on clean system
12. Professional packaging

---

## 📝 NOTES

- **Step 10 Fix**: The critical checkbox Type fix (8 vs 5) and unlock logic was THE breakthrough
- **Success Rate**: Step 10 now fills 16/17 fields (94% automated, 1 manual checkbox)
- **Main Script**: `guardianship_app.py` is the GUI launcher
- **All Steps Working**: 1-14 are functional (13/14 = 93% automated, Step 10 = 94% automated)

---

## ✨ READY FOR RELEASE WHEN:

- [ ] All HIGH PRIORITY items complete
- [ ] At least 80% of MEDIUM PRIORITY items complete
- [ ] Successfully tested on 1-2 external machines
- [ ] Documentation is clear enough for non-technical user
- [ ] All API setup wizards work properly
