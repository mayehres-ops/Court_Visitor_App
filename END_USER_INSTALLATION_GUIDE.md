# Court Visitor App - Installation Guide

## Welcome!

This guide will help you install and set up the Court Visitor application on your computer.

**Estimated setup time:** 15-20 minutes

---

## System Requirements

- **Operating System:** Windows 10 or Windows 11
- **RAM:** 4GB minimum (8GB recommended)
- **Disk Space:** 500MB free space
- **Internet:** Required for Google API features
- **Microsoft Word:** Required for document generation (Steps 8, 13, 14)
- **Google Account:** Required for email and calendar features

---

## Installation Steps

### Step 1: Download the Application

1. Visit the download page: [Your download link here]
2. Click "Download CourtVisitorApp.exe"
3. Save the file to your **Downloads** folder

**Note:** Windows Defender might show a warning for unsigned applications. This is normal. Click "More info" → "Run anyway"

---

### Step 2: Create Application Folder

1. Open File Explorer
2. Navigate to `C:\`
3. Create a new folder called `CourtVisitorApp`
4. Move `CourtVisitorApp.exe` into this folder

**Final location:** `C:\CourtVisitorApp\CourtVisitorApp.exe`

---

### Step 3: Create Required Folders

The app will create most folders automatically, but create these manually:

```
C:\CourtVisitorApp\
├── CourtVisitorApp.exe
├── Config\
│   └── API\          (Create this folder for Google credentials)
├── App Data\
│   └── Templates\    (Word document templates)
└── New Files\        (Incoming PDFs)
```

---

### Step 4: Setup Google Cloud API

The app uses Google services for email, calendar, and contacts. You need to set up API access:

#### 4a. Create Google Cloud Project

1. Go to https://console.cloud.google.com
2. Click "Select a project" → "New Project"
3. Name: "Court Visitor App"
4. Click "Create"

#### 4b. Enable Required APIs

Enable these APIs (click "Enable APIs and Services"):
- Gmail API
- Google Calendar API
- Google People API (Contacts)
- Google Sheets API
- Google Drive API
- Google Maps Geocoding API (optional)

#### 4c. Create OAuth Credentials

1. Go to "APIs & Services" → "Credentials"
2. Click "Create Credentials" → "OAuth client ID"
3. If prompted, configure OAuth consent screen:
   - User Type: External
   - App name: Court Visitor App
   - User support email: [Your email]
   - Developer contact: [Your email]
   - Scopes: Add Gmail, Calendar, Contacts, Drive, Sheets
4. Application type: "Desktop app"
5. Name: "Court Visitor Desktop Client"
6. Click "Create"
7. Download the JSON file
8. Rename it to: `client_secret.json`
9. Place it in: `C:\CourtVisitorApp\Config\API\client_secret.json`

#### 4d. Setup Google Vision API (for OCR)

1. In Google Cloud Console, enable "Cloud Vision API"
2. Create Service Account:
   - Go to "IAM & Admin" → "Service Accounts"
   - Click "Create Service Account"
   - Name: "court-visitor-ocr"
   - Role: "Cloud Vision API User"
3. Click on the service account → "Keys" → "Add Key" → "Create new key"
4. Choose JSON format
5. Download and save as: `C:\CourtVisitorApp\Config\API\google_vision_credentials.json`

#### 4e. Setup Billing (Required for Vision API)

Google Vision API requires billing enabled:
1. Go to "Billing" in Cloud Console
2. Link a payment method (credit card)
3. **Cost:** First 1,000 OCR requests/month are FREE
4. After that: ~$1.50 per 1,000 requests

**For typical usage (50 files/month):** FREE

---

### Step 5: Setup Excel Database

1. Download the template: `ward_guardian_info.xlsx`
2. Place it in: `C:\CourtVisitorApp\App Data\ward_guardian_info.xlsx`
3. Open in Excel and fill in your ward information

**Required columns:**
- causeno (Case number)
- wardfirst, wardlast (Ward name)
- waddress (Ward address)
- guardian1, gemail (Primary guardian)
- Guardian2, g2email (Secondary guardian)
- visitdate, visittime (Appointment scheduling)

---

### Step 6: Setup Word Templates

The app needs Word templates for document generation:

1. Place these templates in `C:\CourtVisitorApp\App Data\Templates\`:
   - `CVR_Template.docx` (Court Visitor Report)
   - `Payment_Form_Template.docx` (Payment forms)
   - `Mileage_Form_Template.docx` (Mileage reimbursement)

2. Templates should use **Content Controls** for auto-fill fields
3. Contact support if you need template examples

---

### Step 7: First Run

1. Double-click `CourtVisitorApp.exe`
2. The app will open and show the 14-step workflow
3. On first run, you'll be prompted to authenticate with Google (for each service):
   - A browser window will open
   - Sign in with your Google account
   - Click "Allow" to grant permissions
   - Browser will show "Success!" - you can close it
   - Return to the app

**Note:** You only need to authenticate once. The app saves tokens for future use.

---

## Usage Overview

### The 14 Steps

1. **Extract Data from PDFs** - OCR to extract case information
2. **Organize Case Files** - Create folders and organize documents
3. **Generate Route Map** - Create map for home visits
4. **Send Meeting Requests** - Email guardians to schedule visits
5. **Add Contacts** - Sync guardians to Google Contacts
6. **Send Confirmations** - Email appointment confirmations
7. **Schedule Events** - Add appointments to Google Calendar
8. **Generate CVR** - Create Court Visitor Reports
9. **Generate Summaries** - Create visit summaries
10. **Complete CVR** - Fill CVR with Google Form responses
11. **Send Follow-ups** - Send follow-up emails after visits
12. **Email CVR** - Send completed reports to supervisor
13. **Generate Payment Forms** - Create payment request forms
14. **Generate Mileage Log** - Create mileage reimbursement forms

### Basic Workflow

1. Place new PDF files (ARPs, Orders, Approvals) in `New Files` folder
2. Click Step 1 to extract data
3. Follow steps 2-14 in order
4. Each step processes data and updates the Excel file

---

## Troubleshooting

### "No Python installation found"

The app includes Python, so this shouldn't happen. If it does:
- Re-download the `.exe` file
- Make sure you're not running it from a USB drive
- Try running as Administrator (right-click → Run as administrator)

### "Token has expired or revoked"

The app now handles this automatically:
- It will delete the expired token
- Open your browser for re-authentication
- Just sign in again

### "OCR failed" or "Vision API error"

Check:
- Google Cloud billing is enabled
- Vision API is enabled in your project
- Service account JSON file is in the correct location
- Internet connection is working

### Word document generation fails

Check:
- Microsoft Word is installed
- Templates are in the correct folder
- Templates use Content Controls (not plain text fields)
- Word is not already open with documents

### Emails not sending

Check:
- Gmail API is enabled
- OAuth credentials file exists
- You've authenticated (browser popup completed)
- Email addresses in Excel are valid

---

## Updates

The app automatically checks for updates on startup.

**To update:**
1. App shows "Update Available" dialog
2. Click "Yes" to download
3. Download new `CourtVisitorApp.exe`
4. Close the current app
5. Replace old `.exe` with new one
6. Restart the app

**Your data is preserved:**
- Excel database
- Config files
- API tokens
- All documents

---

## Data Backup

**Important:** Regularly backup these folders:
- `App Data\ward_guardian_info.xlsx` (Your database)
- `Config\API\` (Your credentials)
- `New Clients\` (Case folders)
- `Completed\` (Completed cases)

**Recommended:** Use OneDrive, Google Drive, or Dropbox to automatically backup your `C:\CourtVisitorApp` folder.

---

## Security & Privacy

- **Credentials:** Your Google credentials are stored locally (never uploaded)
- **Data:** All ward information stays on your computer
- **Emails:** Sent directly from your Gmail account (not through our servers)
- **Updates:** Downloaded from GitHub (secure HTTPS)

**Best practices:**
- Use a strong password for your Google account
- Enable 2-factor authentication on Google
- Don't share your `Config\API\` folder
- Encrypt your computer with BitLocker (Windows Pro)

---

## Support

### Getting Help

- **Email:** [Your support email]
- **Phone:** [Your support phone]
- **Hours:** Monday-Friday, 9 AM - 5 PM CST

### What to include in support requests:

1. Screenshot of the error message
2. Which step you were running
3. Your Windows version (Settings → System → About)
4. App version (shown in title bar)

---

## Uninstallation

To remove the app:

1. Close the application
2. Delete the folder: `C:\CourtVisitorApp\`
3. (Optional) Revoke Google API access:
   - Go to https://myaccount.google.com/permissions
   - Find "Court Visitor App"
   - Click "Remove Access"

**Note:** This does NOT delete your case files or ward database. Back them up first!

---

## License

Copyright © 2024 GuardianShip Easy, LLC. All rights reserved.

This software is proprietary. Unauthorized copying, distribution, or modification is strictly prohibited.

---

## Version History

### Version 1.0.0 (Current)
- Initial release
- 14-step workflow automation
- Google API integration
- Auto-update system
- OAuth token auto-refresh

---

**Thank you for using Court Visitor App!**

We're here to help make your court visitor work more efficient and organized.
