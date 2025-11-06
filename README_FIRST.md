# Court Visitor App - Getting Started Guide

**READ THIS FIRST** before using the Court Visitor Application.

**📚 Additional Resources:**
- **Court Visitor App Manual** - For comprehensive documentation and advanced topics, click the "📖 Manual" button in the app's Quick Access sidebar
- This guide covers essential setup and workflow - the Manual provides deeper technical details

---

## ⚡ Quick Start - Essential Information (For the Time-Pressed)

*We understand you're busy! Here are the absolute must-know items. However, we strongly encourage reading the full guide when you have time - it will save you hours of troubleshooting later.*

### The 5 Critical Things You Must Know:

1. **🔐 Google APIs Required** - Gmail account needed. Click "Setup Google Vision" in the app sidebar BEFORE Step 1.

2. **📊 Close Excel Before OCR** - Step 1 CANNOT run if Excel is open. Save & close it first.

3. **✅ Verify OCR Results** - After Step 1, ALWAYS check the Excel file. OCR isn't perfect - verify names, dates, and numbers match the PDFs.

4. **📁 Check "Unmatched" Folder** - After Step 2, check [New Files\Unmatched](file:///C:/GoogleSync/GuardianShip_App/New%20Files/Unmatched) folder. Should be empty. If not, manually move files to correct ward folders.

5. **🔒 Security Warnings = Normal** - When you see "Do you trust this document?" → Answer **YES**. When Google asks to re-authorize → Click **ALLOW**. These are expected for automation to work.

### Quick Troubleshooting:
- **Step won't run?** → Check Excel is closed, PDFs are closed, folders aren't open
- **Need to re-send email/recreate CVR?** → Clear the status column in Excel (see full guide)
- **OCR wrong?** → Manually fix Excel AND report the error so we can improve it

**📖 Now please read the full guide below - it contains critical details that will prevent problems!**

---

## ⚠️ Beta Software Notice

**This application is currently in BETA.**

What this means for you:
- The software is fully functional but still being refined
- Your feedback is invaluable and directly shapes improvements
- **Please report ALL errors** - even small ones help us improve
- **Feature requests welcome** - tell us what would make your workflow easier
- Updates and improvements are released regularly

**How to help:**
- Use the "🐛 Report Bug" button in the app for any issues
- Use the "💡 Request Feature" button for suggestions
- Be specific: screenshots, error messages, and steps to reproduce help tremendously

Thank you for being an early adopter and helping make this tool better for everyone!

---

## Copyright & License

**Court Visitor App** © 2024-2025. All rights reserved.

This software is proprietary and confidential. Unauthorized copying, distribution, or modification is prohibited.

**Third-Party Components:**
- Python © Python Software Foundation
- Google Cloud APIs © Google LLC
- Microsoft Office Interop © Microsoft Corporation
- Other open-source libraries used under their respective licenses

**For licensing inquiries or permissions, contact:** [Your contact information]

---

## Table of Contents

1. [Prerequisites & Initial Setup](#prerequisites--initial-setup)
2. [Understanding the Excel Database](#understanding-the-excel-database)
3. [Step-by-Step Workflow Instructions](#step-by-step-workflow-instructions)
4. [Important Validation & Security Notices](#important-validation--security-notices)
5. [Troubleshooting Common Issues](#troubleshooting-common-issues)
6. [Best Practices](#best-practices)

---

## Prerequisites & Initial Setup

### Required: Gmail Account

This application requires a Gmail account to function properly with email automation, calendar integration, and Google Forms features.

**If you don't have a Gmail account:**
1. Click the Gmail setup wizard in the app's sidebar
2. Create a new Gmail account specifically for your Court Visitor work
3. Keep your work and personal email separate as a security best practice

**If you use Outlook:**
- Add your Gmail account to Outlook so all emails appear in one place
- Instructions: Outlook → File → Add Account → Enter Gmail address

### Required: Google Cloud APIs

The following Google Cloud APIs must be set up before using the app:

#### 1. Google Vision API (Required for Step 1: OCR)
- **Purpose:** Extracts text from ORDER.pdf and ARP.pdf files
- **Cost:** FREE for first 1,000 requests/month (sufficient for typical usage)
- **Setup:** Click "🔧 Setup Google Vision" in the app sidebar and follow the 5-step wizard

#### 2. Gmail API (Required for Steps 4, 6, 11, 12)
- **Purpose:** Send emails, create Gmail drafts, send confirmations
- **Cost:** FREE
- **Setup:** Automatic on first use - follow OAuth authorization prompts

#### 3. Google Calendar API (Required for Step 7)
- **Purpose:** Create calendar events with attachments and invitations
- **Cost:** FREE
- **Setup:** Automatic on first use - follow OAuth authorization prompts

**Steps that work WITHOUT APIs:**
- Step 2: Organize Case Files
- Step 3: Generate Route Map
- Step 5: Add Contacts
- Step 8: Generate CVR
- Step 9: Visit Summary
- Step 13: Payment Forms
- Step 14: Mileage Logs

---

## Understanding the Excel Database

### Excel File Location
[C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx](file:///C:/GoogleSync/GuardianShip_App/App%20Data/ward_guardian_info.xlsx)

**Quick access:** Click "📊 Excel File" button in the app sidebar

### Status Tracking Columns (End of Excel Sheet)

The columns at the end of the Excel file prevent duplicate emails, folders, and CVR documents from being created. Understanding these columns is critical for proper workflow management.

| Column | Purpose | Used By | What It Does |
|--------|---------|---------|--------------|
| **emailsent** | Tracks meeting request emails | Step 4 | Prevents sending duplicate meeting requests. Contains the date the initial email was sent. |
| **confirmemail** | Tracks confirmation emails | Step 6 | Prevents sending duplicate confirmation emails. Stores confirmation date. |
| **foldercreated** | Tracks case folder creation | Step 2 | Prevents creating duplicate folders. Marks "Y" when folder exists. |
| **CVR created?** | Tracks CVR document creation | Step 8 | Prevents creating duplicate CVR documents. Marks "Y" when CVR is generated. |
| **datesubmitted** | Tracks CVR submission to supervisor | Step 12 | Prevents re-sending the same CVR. Contains the date CVR was emailed to supervisor. |
| **followupsent** | Tracks follow-up emails | Step 11 | Prevents duplicate follow-ups. Stores follow-up email date. |

### When to Clear Status Columns

You may need to **clear a cell** to re-run a step. Common scenarios:

| Scenario | Column to Clear | Reason |
|----------|----------------|---------|
| Need to resend meeting request email | `emailsent` | Guardian didn't receive original email, or email needs to be sent to new guardian |
| Need to regenerate CVR document | `CVR created?` | CVR was accidentally deleted or corrupted and needs to be recreated |
| Need to re-email CVR to supervisor | `datesubmitted` | Supervisor didn't receive it, or corrections were made to the CVR |
| Need to resend confirmation email | `confirmemail` | Guardian requested a new confirmation or meeting details changed |
| Need to recreate case folder | `foldercreated` | Folder was accidentally deleted or is in wrong location |
| Need to send another follow-up | `followupsent` | Additional follow-up is needed after initial follow-up |

**How to clear a cell:**
1. Open Excel file (close it first if the app is running)
2. Locate the appropriate row (by ward name or cause number)
3. Find the status column
4. Delete the cell contents (date or "Y")
5. Save the Excel file
6. Close Excel before running the app step

---

## Step-by-Step Workflow Instructions

### Step 1: Process New Cases (OCR Extraction)

**What this step does:**
- Extracts data from ORDER.pdf and ARP.pdf files using Google Vision OCR
- Populates the Excel database with case information
- Creates initial records for new guardianship cases

**CRITICAL REQUIREMENTS:**

1. **Close the Excel file before running this step**
   - OCR cannot write to Excel while it's open
   - Save any changes and close Excel completely
   - Error message will appear if Excel is open

2. **Verify ORDER and ARP match:**
   - **Order Number** in ORDER.pdf must match **ARP Number** in ARP.pdf
   - **Approval Number** must be present in both documents
   - **Both files must be present** (ORDER.pdf AND ARP.pdf)
   - If only one file is present, the Excel row will be incomplete

3. **PDF file requirements:**
   - Files must be in: [C:\GoogleSync\GuardianShip_App\New Files](file:///C:/GoogleSync/GuardianShip_App/New%20Files)
   - File names must contain "ORDER" or "ARP" (case-insensitive)
   - PDFs should be clear and legible for accurate OCR

**IMPORTANT - Verify OCR Accuracy:**

After Step 1 completes, you **must manually verify** the OCR results:

1. Open the Excel file
2. Check the newly added rows
3. Compare against the original ORDER and ARP PDFs
4. Verify all fields are filled in completely:
   - Cause number
   - Ward first name, last name, middle name
   - Guardian names and contact information
   - Order number and approval number
   - Dates (appointment date, court visit date)

**If you find OCR errors or omissions:**
- Manually correct the Excel file
- Report the errors so improvements can be made to the OCR system
- Common issues to watch for:
  - Misspelled names
  - Incorrect dates
  - Missing phone numbers or addresses
  - Transposed numbers in cause numbers

### Step 2: Organize Case Files

**What this step does:**
- Creates case folders in the "New Clients" directory
- Organizes documents by ward
- Moves ORDER and ARP files to appropriate folders

**CRITICAL - Close any open PDFs:**
- Windows file locks prevent moving files that are open
- Close all PDF viewers before running this step
- Close the "New Files" folder window if open

**After Step 2 - Check Unmatched Folder:**

Location: [C:\GoogleSync\GuardianShip_App\New Files\Unmatched](file:///C:/GoogleSync/GuardianShip_App/New%20Files/Unmatched)

**This folder should be EMPTY after Step 2.**

If files appear in Unmatched:
1. The OCR couldn't match them to an Excel record
2. You need to manually move them to the correct ward folder in: [C:\GoogleSync\GuardianShip_App\New Clients](file:///C:/GoogleSync/GuardianShip_App/New%20Clients)

**Creating folders manually:**
- **CRITICAL:** Folder names must follow this exact format:
  ```
  [Last Name], [First Name] [Middle Name] - [Cause Number]
  ```
- **Example:** `Warren, David Patrick - 23-000357`
- **Why this matters:** Other automation steps search for folders by cause number. Incorrect naming breaks Steps 4, 6, 7, 8, 11, and 12.

### Step 3: Generate Route Map

**What this step does:**
- Creates a visual map showing ward locations
- Helps plan efficient visit routes
- Opens automatically in your default image viewer

**No special requirements** - This step works offline.

### Step 4: Send Meeting Requests

**What this step does:**
- Sends initial meeting request emails to guardians
- Attaches the ORDER.pdf to the email
- Creates Gmail drafts for review before sending
- Saves a text copy of the email to the case folder

**What happens (detailed):**
1. Reads Excel for cases where `emailsent` column is empty
2. Finds the case folder in New Clients directory
3. Locates the ORDER.pdf in that folder
4. Creates a Gmail draft with:
   - Professional meeting request message
   - ORDER.pdf attached
   - Guardian's email address as recipient
5. Opens Gmail in your browser showing the drafts
6. You review and send the drafts manually
7. Updates `emailsent` column with today's date when sent

**Important notes:**
- Emails are created as **drafts** for your review
- You must manually send each draft from Gmail
- This prevents accidental emails and allows customization
- The ORDER.pdf attachment provides guardians with case details

### Step 5: Add Contacts

**What this step does:**
- Adds guardian email addresses and names to your Windows Contacts
- Creates contact records for easy identification

**Why add contacts:**
- When you receive guardian replies, you'll see their name (not just email address)
- Helps identify legitimate emails vs. spam
- Makes it easier to find guardian contact information
- Contacts appear in Outlook, Gmail, and other email clients

**No special requirements** - Simply click and run.

### Step 6: Confirm Appointment

**What this step does:**
- Sends appointment confirmation emails to guardians
- Includes meeting date, time, and location details
- Saves confirmation copy to case folder

**Prerequisites:**
- Step 4 must be completed (initial meeting request sent)
- Guardian must have responded with availability

**What to watch for:**
- Only sends to cases where `confirmemail` column is empty
- Updates `confirmemail` column after sending

### Step 7: Schedule Calendar Event

**What this step does:**
- Creates Google Calendar event for the court visit
- Attaches both ORDER.pdf and ARP.pdf to the calendar event
- Sends calendar invitation to guardian's email
- Includes clickable directions link for easy navigation

**What happens (detailed):**
1. Reads visit date and time from Excel
2. Creates calendar event with:
   - Ward's name as event title
   - Visit date and time
   - Ward's address as location
   - **Clickable Google Maps directions link**
   - ORDER.pdf attachment (for reference during visit)
   - ARP.pdf attachment (for reference during visit)
3. Sends invitation to guardian's email
4. Event appears in your Google Calendar

**Benefits:**
- All visit information in one place
- Click directions link for turn-by-turn navigation to ward's location
- Access documents during visit from your phone/tablet
- Guardian also receives calendar invitation

### Step 8: Generate CVR (Court Visitor Report)

**What this step does:**
- Creates blank Court Visitor Report (CVR) Word documents
- Pre-fills basic information from Excel (name, cause number, dates)
- Saves CVR to case folder in New Clients directory

**IMPORTANT - Before first use:**

You must customize the CVR template with your name:

1. Open the template: [C:\GoogleSync\GuardianShip_App\Templates\Court_Visitor_Report_Template.docx](file:///C:/GoogleSync/GuardianShip_App/Templates/Court_Visitor_Report_Template.docx)
2. Find the two locations marked "Court Visitor:"
3. Replace with your name
4. **Save the document**
5. **Be extremely careful:**
   - Do not change any content control names or formatting
   - Only change your name in the designated spots
   - Breaking the template will prevent automation from working

**If you see template errors:**
- Do NOT attempt to fix them yourself
- Report the issue immediately
- We will correct the template to prevent workflow disruption

**What to watch for:**
- CVRs are created only for cases where `CVR created?` column is empty
- Marks `CVR created? = Y` after generation
- Clear this column if you need to regenerate a CVR

### Step 9: Visit Summary

**What this step does:**
- Creates a one-page visit summary document
- Provides quick reference for visit details
- Opens automatically for printing

**No special requirements** - Simply run when needed.

### Step 10: Complete CVR with Google Form Data

**What this step does:**
- Reads responses from Google Forms submitted by guardians
- Auto-fills the CVR document with form data
- Matches form responses to CVR files by cause number or ward name
- Preserves data already filled in by Step 8 (doesn't overwrite)

**How it works:**
1. Guardian receives email with Google Form link (from Step 4 or 6)
2. Guardian fills out the form with ward information:
   - Physical condition of ward (walking, hearing, speech, etc.)
   - Living situation and accessibility
   - Social interaction and activities
   - Safety and fire safety
3. Form response is saved to Google Sheets
4. Step 10 reads the Google Sheets response
5. Matches the response to the correct CVR file
6. Fills in the blank fields in the CVR (doesn't overwrite Step 8 data)

**What gets filled from Google Forms:**
- Physical condition checkboxes (11 fields)
- Living situation details
- Yes/No questions about ward's environment
- Social interaction information
- Supplemented by (your name from the form)

**Important:**
- Form responses are matched by cause number or ward name
- If guardian doesn't fill out form, CVR will have blanks for those sections
- You can manually fill any remaining blanks before submitting

### Step 11: Send Follow-up Emails

**What this step does:**
- Sends follow-up emails to guardians who haven't responded
- Tracks follow-up dates in Excel
- Prevents duplicate follow-ups

**What to watch for:**
- Updates `followupsent` column with date
- Clear column to send additional follow-ups if needed

### Step 12: Email CVR to Supervisor

**What this step does:**
- Emails completed CVR documents to your supervisor
- Shows preview dialog with editable email address
- Updates Excel with submission date
- Moves case folder to "Completed" directory

**Important - Email Address:**

The supervisor email address will stay in the system until you change it.

**How to change the supervisor email:**
1. Run Step 12
2. Preview dialog appears showing current email address
3. Email field is **editable** - click and type new address
4. Click "Send Email"
5. New email is saved for future use

**First-time behavior:**
- Default email: al.benedict@traviscountytx.gov
- You'll be prompted to confirm or change this

**What happens after sending:**
1. Email sent to supervisor with CVR attached
2. `datesubmitted` column updated with today's date
3. Case folder moved from "New Clients" to "Completed"
4. Future runs skip cases that already have `datesubmitted` filled

**To re-send a CVR:**
- Clear the `datesubmitted` cell in Excel
- Move folder back to "New Clients" if needed

### Step 13: Payment Forms

**What this step does:**
- Generates monthly Court Visitor payment forms
- Selects which month to bill for
- Automatically opens generated Word document for printing

**How to use:**
1. Click Step 13
2. Select month from dropdown (or type month number)
3. Click "Generate"
4. Word document opens automatically
5. Review and print for submission

**What to watch for:**
- Only visits with dates in the selected month are included
- Form opens automatically for immediate review/printing

### Step 14: Mileage Reimbursement

**What this step does:**
- Generates monthly mileage reimbursement forms
- Calculates total miles traveled
- Opens for printing and submission

**How to use:**
1. Click Step 14
2. Select month
3. Review mileage calculations
4. Print and submit

---

## Important Validation & Security Notices

### Microsoft Office Security Prompts

When opening CVR documents or running automation, you may see security warnings:

**Example warning:**
```
⚠️ Microsoft Office Security Notice
"Do you trust Court Visitor Report.docx?"
"This file might contain harmful content."
```

**YOUR RESPONSE SHOULD BE: YES**

**Why this appears:**
- Word documents contain content controls and macros for automation
- Microsoft shows this warning for ANY document with active content
- This is a security feature, not an indication of actual danger

**What to do:**
1. Click "Yes" or "Enable Content"
2. Check "Trust documents from this publisher" if available
3. This allows the automation to function properly

**Why it's safe:**
- These are YOUR documents created by YOUR automation scripts
- Content controls are used to fill in form fields
- No harmful macros or scripts are present
- This is standard for any Word form automation

### OAuth Re-authorization

**Normal behavior:** Google may occasionally ask you to re-authorize the app's access to Gmail, Calendar, or Drive.

**What you'll see:**
```
🔐 Authorization Required
"Court Visitor App needs permission to access your Google Calendar"
```

**YOUR RESPONSE SHOULD BE: YES / ALLOW / APPROVE**

**When this happens:**
- After 7-30 days of inactivity (normal Google security)
- After updating the app
- After changing your Google password
- Randomly as a security measure

**What to do:**
1. Click "Allow" or "Approve"
2. Check the box "Trust this application"
3. Complete authorization
4. App will continue functioning normally

**Why this is safe:**
- OAuth is Google's secure authorization system
- App only accesses YOUR account
- Permissions are limited to specific functions (email, calendar)
- You can revoke access anytime from Google Account settings

**Security tip:**
- Never approve authorization requests from unknown applications
- Only approve when YOU initiated an action in the Court Visitor App

---

## Troubleshooting Common Issues

### Excel File Issues

**Problem:** "Cannot write to Excel file"
- **Solution:** Close the Excel file completely before running OCR (Step 1)

**Problem:** "Row is incomplete after OCR"
- **Solution:** Verify both ORDER and ARP files are present with matching numbers

**Problem:** "OCR extracted wrong data"
- **Solution:** Manually correct in Excel and report the issue for improvement

### PDF File Issues

**Problem:** "Cannot move PDF files"
- **Solution:** Close all PDF viewers and the New Files folder before Step 2

**Problem:** "Files in Unmatched folder"
- **Solution:** Manually move to correct ward folder in New Clients directory

### Email Issues

**Problem:** "Gmail drafts not appearing"
- **Solution:** Check Gmail authorization is complete; refresh Gmail in browser

**Problem:** "Wrong supervisor email"
- **Solution:** Run Step 12, edit email field in preview dialog, click Send

**Problem:** "Duplicate emails being sent"
- **Solution:** Check `emailsent` or `confirmemail` columns aren't cleared accidentally

### Calendar Issues

**Problem:** "Calendar event not created"
- **Solution:** Complete Google Calendar OAuth authorization when prompted

**Problem:** "Documents not attached to calendar"
- **Solution:** Verify ORDER and ARP files exist in case folder

### CVR Issues

**Problem:** "CVR template broken"
- **Solution:** Do NOT edit template yourself - report issue for proper fix

**Problem:** "Google Form data not filling CVR"
- **Solution:** Verify guardian filled out form; check cause number matches

**Problem:** "Need to regenerate CVR"
- **Solution:** Clear `CVR created?` column in Excel, then re-run Step 8

### Folder Issues

**Problem:** "Case folder not found"
- **Solution:** Verify folder name format: `Last, First Middle - CauseNumber`
- Check folder is in New Clients directory, not Completed

**Problem:** "Cannot move folder to Completed"
- **Solution:** Close any open files from that folder; file locks prevent moving

---

## Best Practices

### Daily Workflow

1. **Start of day:**
   - Close Excel file
   - Run Step 1 to process any new cases
   - Verify OCR results immediately

2. **After OCR:**
   - Run Step 2 to organize files
   - Check Unmatched folder
   - Fix any filing issues

3. **Before visits:**
   - Run Steps 4-7 to communicate with guardians
   - Review Gmail drafts before sending
   - Check calendar events created correctly

4. **After visits:**
   - Run Step 8 to generate CVRs
   - Wait for Google Form responses (Step 10)
   - Complete any remaining CVR fields manually

5. **End of process:**
   - Run Step 12 to email CVRs to supervisor
   - Verify folders moved to Completed
   - Run Steps 13-14 for monthly billing

### Data Accuracy

- **Always verify OCR results** - Don't trust automation blindly
- **Keep Excel updated** - Correct any errors immediately
- **Check status columns** - Understand why steps skip cases
- **Report all issues** - Help improve the automation

### File Management

- **Keep New Files folder clean** - Move processed PDFs promptly
- **Name folders correctly** - Exact format prevents automation failures
- **Back up Excel regularly** - Status columns are critical
- **Don't manually edit CVR template** - Report issues instead

### Security & Privacy

- **Use work Gmail account** - Separate from personal email
- **Approve OAuth carefully** - Only when YOU initiated action
- **Trust Office prompts** - Required for automation
- **Keep credentials secure** - Never share API keys or passwords

---

## Getting Help

### In-App Help

- **Live Tech Support** - AI-powered help with Claude, ChatGPT, or Gemini
  - Click "🆘 Live Tech Support" in sidebar
  - Provides complete context to AI assistant
  - Get instant answers to technical questions

- **User Manual** - Complete reference documentation
  - Click "📖 Manual" in sidebar
  - Opens PDF with detailed instructions

- **Bug Reports** - Report issues
  - Click "🐛 Report Bug" in sidebar
  - Template provided for easy reporting

- **Feature Requests** - Suggest improvements
  - Click "💡 Request Feature" in sidebar
  - Share ideas for new functionality

### Contact

- Questions about workflow: [Your contact info]
- Technical issues: Use in-app bug report
- Feature ideas: Use in-app feature request

---

## Quick Reference

### File Locations

| What | Location | Quick Access |
|------|----------|--------------|
| Excel Database | `C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx` | 📊 Excel File button |
| New Files (PDFs) | `C:\GoogleSync\GuardianShip_App\New Files` | Open manually |
| Case Folders | `C:\GoogleSync\GuardianShip_App\New Clients` | 📁 Guardian Folders button |
| Completed Cases | `C:\GoogleSync\GuardianShip_App\Completed` | Open manually |
| CVR Template | `C:\GoogleSync\GuardianShip_App\Templates\Court_Visitor_Report_Template.docx` | Open manually |
| Unmatched Files | `C:\GoogleSync\GuardianShip_App\New Files\Unmatched` | Check after Step 2 |

### Status Column Quick Reference

| Clear This Column | To Do This | Used By Step |
|-------------------|------------|--------------|
| emailsent | Resend meeting request | Step 4 |
| confirmemail | Resend confirmation | Step 6 |
| foldercreated | Recreate folder | Step 2 |
| CVR created? | Regenerate CVR | Step 8 |
| datesubmitted | Re-email CVR | Step 12 |
| followupsent | Send another follow-up | Step 11 |

---

## 📞 Getting Technical Support

The Court Visitor App provides multiple support resources to help you succeed:

### Built-In Support Tools

**🤖 Interactive Chatbot** (Click "🤖 Ask Chatbot" button)
- AI-powered assistant with app knowledge
- Answers questions about workflow steps
- Provides troubleshooting guidance
- Suggests solutions to common problems
- Available 24/7 right in the app

**❓ Help Dialog** (Click "❓ Help" button)
- Quick reference for common tasks
- Basic troubleshooting steps
- Getting started checklist

**📖 Complete Manual** (Click "📖 Manual" button)
- Comprehensive technical documentation
- Advanced configuration options
- Detailed API setup guides
- In-depth troubleshooting section

### Live Support Options

**🆘 Live Tech Support** (Click "🆘 Live Tech Support" button)
- Connects you to AI assistance for complex issues
- Can analyze error messages
- Provides code-level troubleshooting
- Best for technical problems

**🐛 Report Bug** (Click "🐛 Report Bug" button)
- Submit bug reports directly from the app
- Helps improve the software for everyone
- Your feedback shapes future updates

**💡 Request Feature** (Click "💡 Request Feature" button)
- Suggest new capabilities
- Share workflow improvement ideas
- Influence the app's development roadmap

### Before Contacting Support

To get the fastest help, please:

1. **Try the Chatbot First** - It can solve most common issues instantly
2. **Check the Manual** - Many questions are answered in detail there
3. **Note the Error Message** - Copy the exact error text if there is one
4. **Note What Step Failed** - Know which workflow step you're on
5. **Check Excel Status** - Verify the Excel file isn't locked/open

### Support Best Practices

**For Technical Issues:**
- Use the chatbot for instant guidance
- Check if Excel/PDFs are closed
- Verify Google API setup is complete
- Try clearing Excel status columns

**For Feature Questions:**
- Check the Manual first for detailed explanations
- Use chatbot for "how do I..." questions
- Watch for tooltips and help text in dialogs

**Remember:** This is BETA software - your reports help everyone! Even small issues are worth reporting.

---

**Last Updated:** October 2025
**App Version:** BETA

For the most current information, check the in-app manual or use Live Tech Support.
