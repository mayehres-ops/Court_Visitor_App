# Court Visitor App - Installation Guide

## System Requirements

- Windows 10 or Windows 11
- 4GB RAM minimum (8GB recommended)
- 500MB free disk space
- Internet connection (for Google API features)
- Microsoft Office (Word, Excel) or compatible alternatives
- Google account (Gmail) for email and calendar features

---

## Installation Steps

### 1. Download the Application

You should have received a ZIP file containing:
- `CourtVisitorApp.exe` - The main application
- `Templates/` folder - Form templates
- `Documentation/` folder - User guides
- `EULA.txt` - License agreement

**Extract the ZIP file** to a location on your computer, such as:
- `C:\Court Visitor App\`
- `C:\Users\[YourName]\Documents\Court Visitor App\`

---

### 2. Windows Security Warning (IMPORTANT - READ CAREFULLY)

**When you first run the application, Windows will show a security warning.** This is NORMAL and expected for new applications that are not yet digitally signed.

#### What You'll See:

```
⚠ Windows protected your PC

Microsoft Defender SmartScreen prevented an unrecognized app from starting.
Running this app might put your PC at risk.

App: CourtVisitorApp.exe
Publisher: Unknown publisher

[Don't run]  [More info]
```

#### Why This Happens:

- The application is **brand new** and hasn't built up a "reputation" with Microsoft yet
- The app is **not digitally signed** (signing certificates cost $200-400/year)
- This warning appears for **all unsigned applications**, even safe ones

#### This App IS Safe Because:

✅ You received it from GuardianShip Easy, LLC (authorized source)
✅ The app was scanned by antivirus before distribution
✅ The app only accesses your local files and Google account (with your permission)
✅ No data is sent to external servers
✅ Source code is reviewed and tested

#### How to Proceed Safely:

**Step 1:** When you see the warning, click **"More info"**

**Step 2:** A new button will appear: **"Run anyway"**

**Step 3:** Click **"Run anyway"**

**Step 4:** If asked "Do you want to allow this app to make changes to your device?", click **"Yes"**

---

### 3. Antivirus Software

Some antivirus programs may also flag the application. If this happens:

1. **Add to Exclusions/Whitelist:**
   - Open your antivirus settings
   - Find "Exclusions" or "Whitelist"
   - Add `CourtVisitorApp.exe` to the list

2. **Common Antivirus Instructions:**
   - **Windows Defender:** Settings → Virus & threat protection → Exclusions → Add
   - **Norton:** Settings → Antivirus → Exclusions → Add
   - **McAfee:** Settings → Real-Time Scanning → Excluded Files → Add
   - **Avast:** Settings → General → Exclusions → Add

---

### 4. First Launch - License Agreement

On your first launch, you will see the **End User License Agreement (EULA)**.

**You MUST accept the EULA to use the software.**

**Key Points:**
- The app is licensed for your use as a Court Visitor
- **NOT FOR RESALE** - You cannot sell or redistribute this software
- You must maintain confidentiality of all data
- Comply with HIPAA and privacy laws

**Steps:**
1. Read the license agreement
2. Scroll to the bottom
3. Check the box "I have read and agree..."
4. Click "Accept and Continue"

If you decline, the application will exit.

---

### 5. First Launch - Personal Information Setup

After accepting the EULA, you'll be asked to enter your **Court Visitor Information**:

- Full Name
- Vendor Number
- GL Number
- Cost Center Number
- Address Line 1
- Address Line 2

**This information will be automatically filled into your forms** (mileage logs, payment forms, CVR).

You can update this information later by clicking the **⚙️ Settings** button in the app.

---

### 6. Google Account Authorization

When you use features that require Google APIs (Gmail, Calendar, Sheets), you'll be prompted to authorize the app:

**Steps:**
1. A web browser will open
2. Select your Google account
3. Click "Allow" to grant permissions
4. Close the browser window
5. Return to the app

**Important:**
- Make sure you've been added to the authorized users list by your administrator
- You only need to authorize once (unless you change accounts)

---

### 7. Set Up Google API (One-Time Setup)

Some features require Google Cloud API credentials. Follow the on-screen setup wizard for:

- Gmail API (for sending emails)
- Calendar API (for creating appointments)
- Sheets API (for data access)
- Maps API (for mileage calculations)

**Your administrator should provide the necessary credential files.**

---

## Folder Structure

After installation, your app folder should look like this:

```
Court Visitor App/
├── CourtVisitorApp.exe          (Main application)
├── EULA.txt                     (License agreement)
├── Templates/                   (Form templates)
│   ├── Mileage_Reimbursement_Form.xlsx
│   ├── Court_Visitor_Payment_Invoice.docx
│   └── Court Visitor Report fillable new.docx
├── App Data/                    (Your data - created on first run)
│   ├── ward_guardian_info.xlsx
│   └── Output/
├── Config/                      (Settings - created on first run)
│   ├── court_visitor_info.json
│   ├── API/
│   └── Keys/
└── Documentation/               (User guides)
```

---

## Common Installation Issues

### Issue: "Windows protected your PC" won't go away

**Solution:** Make sure you clicked "More info" first, then "Run anyway". If still having issues, temporarily disable Windows SmartScreen (not recommended long-term).

### Issue: Antivirus quarantines the app

**Solution:** Restore from quarantine and add to antivirus exclusions list (see Section 3 above).

### Issue: App won't start / crashes immediately

**Solutions:**
1. Right-click the app → Properties → check "Unblock" if present
2. Run as Administrator (right-click → Run as administrator)
3. Check Windows Event Viewer for error details
4. Contact support

### Issue: "EULA file not found"

**Solution:** Make sure you extracted ALL files from the ZIP, not just the .exe file. Re-extract if necessary.

### Issue: Google authorization fails

**Solution:**
1. Check that you've been added to authorized users by administrator
2. Clear browser cookies and try again
3. Try a different browser
4. Check that credential files are in `Config/API/` folder

---

## Getting Help

If you encounter issues not covered here:

1. **Check the Help section in the app** (Help menu)
2. **Review the troubleshooting guide** (Documentation folder)
3. **Contact support:**
   - Email: support@guardianshipeasy.com
   - Include: Error message, screenshot, what you were doing when error occurred

---

## Uninstallation

To uninstall:

1. Close the application if running
2. Delete the application folder
3. Optionally: Delete your data in `App Data/` (if you want to remove all data)

**Note:** Uninstalling does NOT revoke Google API permissions. To revoke:
1. Go to: https://myaccount.google.com/permissions
2. Find "Court Visitor App"
3. Click "Remove Access"

---

## Security & Privacy

**What data does the app collect?**
- **NONE.** GuardianShip Easy, LLC does not collect any data from your use of the app.
- All data stays on YOUR computer or YOUR Google account.
- No analytics, no telemetry, no tracking.

**Is my data secure?**
- Your data is stored locally on your computer
- Google API data is secured by Google's encryption
- You are responsible for keeping your computer and Google account secure

**Who can see my data?**
- Only YOU can access your data
- GuardianShip Easy cannot access your data
- Follow proper procedures for handling confidential ward/guardian information

---

## Legal

This software is:
- **Copyright © 2024-2025 GuardianShip Easy, LLC**
- **All rights reserved**
- **NOT FOR RESALE**
- **Licensed for authorized Court Visitors only**

See EULA.txt for full license terms.

---

**Installation Guide Version:** 1.0
**Last Updated:** November 6, 2024
**For App Version:** 1.0.0
