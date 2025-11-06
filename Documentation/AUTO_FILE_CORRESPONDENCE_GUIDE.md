# Auto-File Correspondence Guide

## Overview

The auto-file correspondence system automatically moves sent emails from `_Correspondence_Pending` to the correct client's `Correspondence` folder with proper naming.

## How It Works

### 1. Email Creation (by automation scripts)

When Step 4 or Step 6 creates a draft email, it should include metadata in the subject line:

**Subject Format:**
```
Meeting Request [Smith_John_21-00123]
Appointment Confirmation [Jones_Mary_22-00456]
Follow-up [Brown_25-00789]
```

**Pattern:** `[LastName_FirstName_CauseNo]` or `[LastName_CauseNo]`

### 2. You Review and Send

1. Open your email drafts folder
2. Review the email
3. Send it
4. The email stays in `_Correspondence_Pending` waiting to be filed

### 3. Auto-Filing

Run the auto-file script:

```bash
python C:\GoogleSync\GuardianShip_App\Scripts\auto_file_correspondence.py
```

**What it does:**
1. Scans all `.msg` and `.eml` files in `_Correspondence_Pending`
2. Extracts metadata from subject line: `[LastName_FirstName_CauseNo]`
3. Finds matching client folder in `New Clients` or `Completed Cases`
4. Extracts guardian name from email recipient
5. Renames to: `YYYYMMDD_HHmm_email_guardianname.msg`
6. Moves to: `New Clients/[ClientFolder]/Correspondence/`
7. Creates `Correspondence` folder if it doesn't exist

### 4. Result

**Before:**
```
_Correspondence_Pending/
├── Draft - Meeting Request [Smith_John_21-00123].msg
└── Draft - Confirmation [Jones_Mary_22-00456].msg
```

**After:**
```
New Clients/Smith_John_21-00123/Correspondence/
└── 20251030_0915_email_smith.msg

New Clients/Jones_Mary_22-00456/Correspondence/
└── 20251030_1430_email_jones.msg

_Correspondence_Pending/
└── (empty - goal achieved!)
```

## Usage

### Dry-Run (Preview Mode)

Test what would happen without actually moving files:

```bash
python Scripts\auto_file_correspondence.py --dry-run
```

This shows you:
- Which emails would be filed
- Where they would go
- What they would be renamed to
- Any errors or missing client folders

### Actual Filing

Once you're ready to file the emails:

```bash
python Scripts\auto_file_correspondence.py
```

## Troubleshooting

### Email Not Filed - No Metadata Found

**Problem:** Email subject doesn't contain `[LastName_CauseNo]` tag

**Solution:**
- Add the tag to the subject line manually: `[Smith_21-00123]`
- Or move the email manually to the client folder

### Email Not Filed - Client Folder Not Found

**Problem:** Client folder doesn't exist or name doesn't match

**Solution:**
1. Check if client folder exists in `New Clients` or `Completed Cases`
2. Verify folder naming matches: `LastName_FirstName_CauseNo`
3. Run Step 2 (CVR Folder Builder) to create missing folders
4. Or move the email manually

### Email Not Filed - Permission Error

**Problem:** File is locked or in use

**Solution:**
- Close Outlook or email client
- Run the script again

## Integration with Email Scripts

### Step 4: Email Meeting Request

**Script creates Gmail draft with:**
- **Subject:** `Court Visitor meeting for FirstName LastName - CauseNo`
- **Example:** `Court Visitor meeting for Jon Halford - 22-001301`
- **Opens in:** Gmail drafts (web browser)

### Step 6: Appointment Confirmation Email

**Script should create email with:**
- **Subject:** `Appointment Confirmation [LastName_FirstName_CauseNo]`
- **Save to:** `_Correspondence_Pending/`
- **Opens in:** Outlook drafts

## Running Automatically

### Option 1: Manual Run (Recommended)

Run the script when you're done sending emails for the day:

```bash
python Scripts\auto_file_correspondence.py
```

### Option 2: Scheduled Task (Advanced)

Set up Windows Task Scheduler to run the script:
- **Frequency:** Daily at 5:00 PM (or after you finish work)
- **Command:** `python C:\GoogleSync\GuardianShip_App\Scripts\auto_file_correspondence.py`

### Option 3: GUI Button (Future Enhancement)

Add a button to the GUI:
- **Label:** "File Pending Correspondence"
- **Action:** Runs the auto-file script
- **Shows:** Summary of what was filed

## Best Practices

1. **Review before sending** - Always check draft emails before hitting send
2. **Run auto-file daily** - Keep `_Correspondence_Pending` empty
3. **Check the log** - Review what was filed to catch any errors
4. **Verify metadata** - Ensure subject line has correct `[Name_CauseNo]` tag
5. **Manual fallback** - If auto-filing fails, move files manually

## File Naming Convention

**Format:** `YYYYMMDD_HHmm_email_guardianname.msg`

**Examples:**
- `20251030_0915_email_smith.msg` - Email sent Oct 30, 2025 at 9:15 AM to Smith
- `20251030_1430_email_jones.msg` - Email sent Oct 30, 2025 at 2:30 PM to Jones

**Timestamp:** Uses the file's modification time (when email was created/sent)

**Guardian name:** Extracted from email recipient (To: field)

## What Gets Filed

- ✅ `.msg` files (Outlook emails)
- ✅ `.eml` files (Standard email format)
- ✅ `.emlx` files (Apple Mail format)
- ❌ Other file types are ignored

## Summary

**Goal:** Keep `_Correspondence_Pending` empty = all emails filed!

**Workflow:**
1. Script creates draft → `_Correspondence_Pending`
2. You review and send
3. Run auto-file script
4. Email moved to client's `Correspondence` folder with proper naming
5. `_Correspondence_Pending` is empty ✓
