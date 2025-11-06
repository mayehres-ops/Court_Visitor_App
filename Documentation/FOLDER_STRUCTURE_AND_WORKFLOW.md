# GuardianShip App - Folder Structure and Workflow

## **NEW FOLDER STRUCTURE**

```
GuardianShip_App/
├── New Files/           ← PDFs arrive here, OCR extraction happens here
├── New Clients/         ← Organized folders ready for CVR creation (NEW!)
├── Completed/           ← Folders moved here after CVR emailed
└── App Data/
    ├── ward_guardian_info.xlsx    ← Central data store
    ├── Templates/                  ← CVR template
    ├── Backup/
    └── Staging/
```

## **COMPLETE WORKFLOW (14 Steps)**

### **STEP 1: OCR Extraction**
**Script**: `guardian_extractor_claudecode20251023_bestever_11pm.py`
- **Reads from**: `New Files/` (root PDFs only)
- **Writes to**: `App Data/ward_guardian_info.xlsx`
- **Action**: Extracts data from ARPs and ORDERs using OCR
- **Status**: ✅ CORRECT - Already uses "New Files"

### **STEP 2: Folder Organization**
**Script**: `Scripts/cvr_folder_builder.py`
- **Reads from**: `New Files/` (root PDFs)
- **Creates folders in**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/`
- **Action**: Groups PDFs by cause number, creates "LastName, FirstName - CauseNo" folders
- **Status**: ❌ NEEDS UPDATE - Currently creates folders in "New Files", should create in "New Clients"

### **STEP 3: Build Map Sheet**
**Script**: `Automation/Build Map Sheet/Scripts/build_map_sheet.py`
- **Reads from**: Excel data
- **Searches folders**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/`
- **Action**: Creates route map for visits
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 4: Email Meeting Requests**
**Script**: `Automation/Email Meeting Request/scripts/send_guardian_emails.py`
- **Reads from**: Excel data
- **Searches folders**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/`
- **Action**: Sends meeting request emails to guardians
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 5: Add Guardians to Contacts**
**Script**: `Automation/Contacts - Guardians/scripts/add_guardians_to_contacts.py`
- **Reads from**: Excel data
- **Action**: Adds guardian contact info to Google Contacts
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 6: Send Appointment Confirmation Email**
**Script**: `Automation/Appt Email Confirm/scripts/send_confirmation_email.py`
- **Reads from**: Excel data
- **Action**: Sends confirmation emails after scheduling
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 7: Create Calendar Event**
**Script**: `Automation/Calendar appt send email conf/scripts/create_calendar_event.py`
- **Reads from**: Excel data
- **Action**: Creates Google Calendar events
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 8: Generate CVR** ⭐ CRITICAL
**Script**: `Automation/Create CV report_move to folder/Scripts/build_cvr_from_excel_cc_working.py`
- **Reads from**: Excel data
- **Saves CVR to**: ~~`New Files/[CauseFolder]/`~~ → **NEEDS CHANGE** → `New Clients/[CauseFolder]/`
- **Action**: Creates Court Visitor Report from Excel data
- **Status**: ❌ NEEDS UPDATE - Currently saves to "New Files" subfolders

### **STEP 9: Court Visitor Summary**
**Script**: `Automation/Court Visitor Summary/build_court_visitor_summary.py`
- **Reads from**: Excel data
- **Searches folders**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/`
- **Action**: Creates summary of all CVRs
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 10: CVR Integration with Form Data**
**Script**: `google_sheets_cvr_integration_fixed.py`
- **Reads from**: Excel + Google Sheets
- **Searches folders**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/`
- **Action**: Updates CVR with form submission data
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 11: Send Follow-up TX Email**
**Script**: `Automation/TX email to guardian/send_followups_picker.py`
- **Reads from**: Excel data
- **Action**: Sends follow-up emails to guardians
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 12: Email CVR to Supervisor** ⭐ CRITICAL
**Script**: `email_cvr_to_supervisor.py`
- **Reads CVR from**: ~~`New Files/[CauseFolder]/`~~ → **NEEDS CHANGE** → `New Clients/[CauseFolder]/`
- **Action**: Emails completed CVR to supervisor
- **After email**: ~~Leaves in New Files~~ → **NEEDS CHANGE** → Move folder to `Completed/`
- **Status**: ❌ NEEDS UPDATE

### **STEP 13: Build Payment Forms**
**Script**: `Automation/CV Payment Form Script/scripts/build_payment_forms_sdt.py`
- **Reads from**: Excel data
- **Searches folders**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/` or `Completed/`
- **Action**: Creates payment forms
- **Status**: ⚠️ NEEDS REVIEW

### **STEP 14: Build Mileage Forms**
**Script**: `Automation/Mileage Reimbursement Script/scripts/build_mileage_forms.py`
- **Reads from**: Excel data
- **Searches folders**: ~~`New Files/`~~ → **NEEDS CHANGE** → `New Clients/` or `Completed/`
- **Action**: Creates mileage reimbursement forms
- **Status**: ⚠️ NEEDS REVIEW

---

## **SCRIPTS REQUIRING CHANGES**

### **Priority 1 (Critical Path)**
1. ❌ **Step 2**: `Scripts/cvr_folder_builder.py` - Create folders in "New Clients" instead of "New Files"
2. ❌ **Step 8**: `build_cvr_from_excel_cc_working.py` - Save CVR to "New Clients" folders
3. ❌ **Step 12**: `email_cvr_to_supervisor.py` - Read from "New Clients", move to "Completed" after email

### **Priority 2 (Supporting Scripts)**
4. ⚠️ **Step 3**: `build_map_sheet.py` - Search "New Clients" for client folders
5. ⚠️ **Step 4**: `send_guardian_emails.py` - Search "New Clients"
6. ⚠️ **Step 9**: `build_court_visitor_summary.py` - Search "New Clients"
7. ⚠️ **Step 10**: `google_sheets_cvr_integration_fixed.py` - Search "New Clients"

### **Priority 3 (Post-CVR Scripts)**
8. ⚠️ **Step 13**: `build_payment_forms_sdt.py` - May need to search both "New Clients" and "Completed"
9. ⚠️ **Step 14**: `build_mileage_forms.py` - May need to search both "New Clients" and "Completed"

---

## **IMPLEMENTATION PLAN**

### Phase 1: Update Critical Path (Steps 2, 8, 12)
1. Backup all 3 scripts
2. Update Step 2 to create folders in "New Clients"
3. Update Step 8 to save CVR to "New Clients" folders
4. Update Step 12 to read from "New Clients" and move to "Completed"
5. Test complete flow: New Files → New Clients → CVR → Email → Completed

### Phase 2: Update Supporting Scripts (Steps 3, 4, 9, 10)
1. Review each script to find folder search logic
2. Update paths from "New Files" to "New Clients"
3. Test each script individually

### Phase 3: Update Post-CVR Scripts (Steps 13, 14)
1. Determine if they should search "New Clients", "Completed", or both
2. Update paths accordingly
3. Test

### Phase 4: Update Main App
1. Update `guardianship_app.py` line 40 if needed
2. Update folder creation logic (line 78-83)
3. Add "New Clients" folder creation

---

## **TESTING CHECKLIST**

- [ ] Step 1: OCR extracts data to Excel from "New Files"
- [ ] Step 2: Folders created in "New Clients" (not "New Files")
- [ ] Step 2: PDFs moved from "New Files" to "New Clients/[CauseFolder]/"
- [ ] Step 8: CVR saved to "New Clients/[CauseFolder]/"
- [ ] Step 12: CVR read from "New Clients/[CauseFolder]/"
- [ ] Step 12: Folder moved to "Completed/[CauseFolder]/" after email
- [ ] All other scripts find folders in "New Clients" or "Completed"

---

**Last Updated**: 2025-10-28
**Status**: Documentation complete, ready to begin Phase 1 implementation
