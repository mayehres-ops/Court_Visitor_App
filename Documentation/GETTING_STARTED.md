# GuardianShip App - Getting Started

## Welcome!

This is your **self-contained** Court Visitor App. All files, scripts, and folders are organized in one place for easy management and distribution.

## Folder Structure

```
GuardianShip_App\
├── guardianship_app.py          # Main application (double-click launch_app.bat to run)
├── launch_app.bat               # Quick launcher
├── Court_Visitor_App_Manual.pdf # User manual
│
├── App Data\
│   ├── Inbox\                   # DROP NEW PDFs HERE (ARP, Order, Approval)
│   ├── ward_guardian_info.xlsx  # Case database
│   ├── Backup\                  # Excel backups (auto-created)
│   ├── Staging\                 # Temp files (auto-created)
│   └── Templates\               # Word/Excel templates
│
├── New Files\                   # Ward case folders (auto-created by Step 2)
│   ├── Smith, John - 25-001234\
│   │   ├── ARP - ARP_25-001234.pdf
│   │   ├── ORDER - Order_25-001234.pdf
│   │   └── APPROVAL - Approval_25-001234.pdf
│   └── ...
│
├── Completed\                   # Move folders here when CVR done
│   └── (manually moved by you)
│
├── Automation\                  # All 14 automation scripts
└── Scripts\                     # Step 2 folder organization script
```

## Quick Start

### 1. First Time Setup

Before using the app, you need to:

1. **Install Python Dependencies** (REQUIRED - Do this first!)
   - Double-click **`INSTALL_DEPENDENCIES.bat`**
   - This installs all required Python libraries
   - Takes 5-10 minutes
   - Only needs to be done once

2. **Google Vision API Setup** (for Step 1 - OCR extraction)
   - Click "Google API Setup" in the sidebar
   - Follow the 5-step wizard
   - Download your credentials JSON key
   - Install it using the wizard

3. **Google Maps API Setup** (CRITICAL for Step 3 - Route Maps)
   - Double-click **`SET_GOOGLE_MAPS_KEY.bat`**
   - Paste your Google Maps API key
   - Restart your computer (required!)
   - See `GOOGLE_MAPS_SETUP.md` for detailed instructions
   - **Without this:** Maps show dots on blank background (useless!)
   - **With this:** Maps show dots on actual street maps (essential for route planning!)

4. **Templates Check**
   - Verify `App Data\Templates\Court Visitor Report fillable new.docx` exists
   - This is used by Step 8 to generate CVR documents

### 2. Daily Workflow

#### Step 1: Extract Guardian Data (OCR)
1. Put ARP/Order PDFs in `App Data\Inbox\`
2. Click **Step 1: OCR Guardian Data** button
3. App extracts ward names, guardian info, cause numbers to Excel

#### Step 2: Organize Case Files
1. Make sure PDFs are in `App Data\Inbox\`
2. Click **Step 2: Organize Case Files** button
3. App creates ward folders in `New Files\` and moves PDFs there
4. Each folder is named: `LastName, FirstName - CauseNumber`
5. Files are renamed: `ARP - filename.pdf`, `ORDER - filename.pdf`, etc.

#### Steps 3-7: Meeting Scheduling
- Step 3: Generate route map for visits
- Step 4: Send meeting requests to guardians
- Step 5: Add guardians to contacts
- Step 6: Send appointment confirmations
- Step 7: Add events to Google Calendar

#### Step 8: Generate Court Visitor Reports
- Creates Word documents from template
- Fills in data from Excel
- Saves reports in each ward's folder

#### Steps 9-14: Reporting & Payment
- Step 9: Create Court Visitor Summary
- Step 10: Upload CVRs to Google Sheets
- Step 11: Send follow-up emails
- Step 12: Email CVR to supervisor
- Step 13: Generate payment invoices
- Step 14: Generate mileage reimbursement forms

### 3. When Cases Are Completed

When you finish a case and submit the CVR:
1. Go to `New Files\`
2. Find the ward's folder
3. **Manually move** the entire folder to `Completed\`

This keeps your active cases separate from completed ones.

## Important Notes

### Step 2 Behavior
- Only creates folders for cases **NOT in Completed folder**
- If a case exists in Completed, files are left in Inbox (skipped)
- Creates `New Files\_Unmatched\` folder for PDFs that can't be matched

### Excel Database
- Located at: `App Data\ward_guardian_info.xlsx`
- Automatically backed up before modifications
- Backups go to: `App Data\Backup\`

### Logs
- Step 2 creates detailed logs at: `New Files\_logs\`
- Check logs if files aren't being organized correctly

## Troubleshooting

### Step 2 isn't moving files
1. Check if PDFs are in `App Data\Inbox\`
2. Check if the case already exists in `Completed\` folder
3. Check if cause number is in `ward_guardian_info.xlsx`
4. Look at the log file in `New Files\_logs\`

### Files go to _Unmatched folder
- Cause number not found in PDF (OCR failed)
- Cause number not in Excel database
- Solution: Run Step 1 first, or manually add case to Excel

### Need Help?
- Click **Live Tech Support** button in sidebar
- Choose AI assistant (Claude recommended)
- Copy the context and paste to AI chat

## Distribution

To give this app to someone else:
1. Zip the entire `GuardianShip_App\` folder
2. Send the zip file
3. They unzip and run `launch_app.bat`

All paths are relative - the app works anywhere!

---

**Copyright © 2024 GuardianShip Easy, LLC. All rights reserved.**
