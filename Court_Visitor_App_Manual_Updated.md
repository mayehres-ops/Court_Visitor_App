# Court Visitor App - User Manual

**Complete guide for the Court Visitor App - Your 14-step workflow automation tool**

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [14-Step Workflow](#14-step-workflow)
3. [Quick Access Tools](#quick-access-tools)
4. [Interactive Chatbot Assistant](#interactive-chatbot-assistant)
5. [API Setup Wizards](#api-setup-wizards)
6. [Troubleshooting](#troubleshooting)
7. [Tips & Best Practices](#tips--best-practices)

---

## Getting Started

### How to Launch the App

**Option 1: Desktop Shortcut (Recommended)**
- Double-click `Court Visitor App` shortcut on your desktop
- App launches silently in the background

**Option 2: From App Folder**
- Navigate to: `C:\GoogleSync\GuardianShip_App\`
- Double-click `Launch Court Visitor App.vbs`

**Option 3: Command Line**
```
cd C:\GoogleSync\GuardianShip_App
python guardianship_app.py
```

### First Time Setup

**NEW! Getting Started Dialog**
- Click `📖 Getting Started` button in sidebar
- View comprehensive quick-start guide
- Access README PDF and Manual
- Learn about all features

**Initial Requirements:**
1. Ensure all your automation scripts are in place
2. Verify the Excel database exists: `App Data\ward_guardian_info.xlsx`
3. Check that the "New Files" folder exists for PDF processing
4. Set up Google APIs (use API Setup Wizards in sidebar!)

### NEW Features in Latest Version

**🤖 Interactive Chatbot**
- Purple button prominently displayed at position #2 in sidebar
- AI-powered assistant with personality
- Answers questions about workflow steps
- Available 24/7 with Matt Rife charm and Monday.com sass!

**3-Column Layout**
- See all 14 steps at once!
- Column 1: Input Phase (Steps 1-5)
- Column 2: Communication (Steps 6-10)
- Column 3: Wrap-Up (Steps 11-15) - Room for future expansion!

**Scrollable Sidebar**
- All buttons organized for easy access
- Most-used features at top
- API Setup at bottom (one-time use)
- Smooth mousewheel scrolling

---

## 14-Step Workflow

### Column 1: Input Phase (Steps 1-5)

#### Step 1: 📧 Process New Cases

**What it does:** Extracts data from ORDER.pdf and ARP.pdf files using OCR and updates the Excel database.

**How to use:**
1. Place ORDER.pdf and ARP.pdf files in: `C:\GoogleSync\GuardianShip_App\New Files\`
2. Click "Process PDFs" button
3. Wait for processing to complete (watch the output window)
4. Check Excel file for extracted data

**Script:** `guardian_extractor_claudecode20251023_bestever_11pm.py`

**What gets extracted:**
- Ward information (name, DOB, address)
- Guardian information (name, DOB, phone, email, address)
- Court information (cause number, court number, judge)
- Dates (ARP filed date, order signed date)

**Exit Codes:** Now properly reports [OK] or [FAIL] status!

---

#### Step 2: 🗂️ Organize Case Files

**What it does:** Creates folder structure for each case and moves PDFs to appropriate folders.

**How to use:**
1. Ensure Step 1 is complete
2. Click "Organize Files" button
3. Verify folders were created in "New Clients" and "Completed" directories

**Script:** `Automation/Create CV report_move to folder/Scripts/build_cvr_from_excel_cc_working.py`

**Exit Codes:** Reports success/failure with proper error messages

---

#### Step 3: 🗺️ Generate Route Map

**What it does:** Creates a map showing ward locations for visit planning.

**How to use:**
1. Ensure ward addresses are in Excel
2. Click "Create Map" button
3. Map will be generated showing all ward locations

**Script:** `Automation/Build Map Sheet/Scripts/build_map_sheet.py`

**Requires:** Google Maps API (optional - works with limited features without it)

---

#### Step 4: 📧 Send Meeting Requests

**What it does:** Sends initial meeting request emails to guardians.

**How to use:**
1. Review guardian email addresses in Excel
2. Click "Send Requests" button
3. Email templates will be used to send meeting requests

**Script:** `Automation/Email Meeting Request/scripts/send_guardian_emails.py`

**Requires:** Gmail API setup (use Setup Gmail API wizard in sidebar)

---

#### Step 5: 👥 Add Contacts

**What it does:** Adds guardians and wards to your Google contact list.

**How to use:**
1. Ensure guardian/ward information is in Excel
2. Click "Add Contacts" button
3. Contacts will be added to your system

**Script:** `Automation/Contacts - Guardians/scripts/add_guardians_to_contacts.py`

**Requires:** People API setup (use Setup People/Calendar wizard in sidebar)

---

### Column 2: Communication (Steps 6-10)

#### Step 6: 📅 Confirm Appointment

**What it does:** Sends appointment confirmation emails to guardians.

**How to use:**
1. After scheduling appointments
2. Click "Send Confirmation" button
3. Confirmation emails will be sent

**Script:** `Automation/Appt Email Confirm/scripts/send_confirmation_email.py`

**Requires:** Gmail API

---

#### Step 7: 📅 Schedule Calendar

**What it does:** Adds appointments to Google Calendar.

**How to use:**
1. Ensure appointment dates/times are set
2. Click "Schedule" button
3. Calendar events will be created

**Script:** `Automation/Calendar appt send email conf/scripts/create_calendar_event.py`

**Requires:** Calendar API setup (use Setup People/Calendar wizard in sidebar)

---

#### Step 8: 📋 Generate CVR

**What it does:** Creates Court Visitor Report (CVR) document templates for each case.

**How to use:**
1. Click "Generate CVR" button
2. CVR templates will be created in case folders

**Script:** `Automation/Create CV report_move to folder/Scripts/build_cvr_from_excel_cc_working.py`

---

#### Step 9: 📄 Visit Summary

**What it does:** Creates one-page visit summary sheets.

**How to use:**
1. After completing visits
2. Click "Generate Summary" button
3. Summary sheets will be created

**Script:** `Automation/Court Visitor Summary/build_court_visitor_summary.py`

---

#### Step 10: ✅ Complete CVR

**What it does:** Fills CVR documents with data from Google Forms.

**How to use:**
1. Ensure Google Form data is available
2. Click "Complete CVR" button
3. CVR documents will be populated

**Script:** `google_sheets_cvr_integration_fixed.py`

**Requires:** Google Sheets API

---

### Column 3: Wrap-Up (Steps 11-15)

#### Step 11: 📧 Send Follow-up

**What it does:** Sends follow-up emails to guardians (thank you notes, additional requests).

**How to use:**
1. Select cases needing follow-up
2. Click "Send Follow-up" button
3. Choose recipients and send emails

**Script:** `Automation/TX email to guardian/send_followups_picker.py`

---

#### Step 12: 📧 Email CVR

**What it does:** Emails completed CVR documents to your supervisor.

**How to use:**
1. Ensure CVRs are complete
2. Click "Email CVR" button
3. CVRs will be sent to supervisor

**Script:** `email_cvr_to_supervisor.py`

**Requires:** Gmail API

---

#### Step 13: 💵 Payment Form

**What it does:** Generates monthly payment reimbursement forms.

**How to use:**
1. Click "Generate Form" button
2. Payment forms will be created for the month

**Script:** `Automation/CV Payment Form Script/scripts/build_payment_forms_sdt.py`

---

#### Step 14: 🚗 Mileage Log

**What it does:** Generates monthly mileage reimbursement logs.

**How to use:**
1. Click "Generate Mileage" button
2. Mileage logs will be created for the month

**Script:** `Automation/Mileage Reimbursement Script/scripts/build_mileage_forms.py`

---

#### Step 15: [Reserved for Future]

**Room for expansion!**
- Space reserved for your next automation feature
- Easy to add new functionality

---

## Quick Access Tools

The right sidebar provides quick access to common tools (organized for efficiency!):

### Essential Tools (Always Visible)

**📖 Getting Started**
- Opens Getting Started guide
- View quick setup instructions
- Access comprehensive documentation

**🤖 Ask Chatbot** (FEATURED with purple styling!)
- Interactive AI assistant
- Answers workflow questions
- Troubleshooting help
- Fun personality with Matt Rife charm!
- Available 24/7

### File Access

**📊 Excel File**
- Opens the ward_guardian_info.xlsx database
- View/edit all extracted data

**📁 New Clients**
- Opens the New Clients directory
- Access active case folders

**📂 New Files**
- Opens New Files folder
- Where you place new PDFs for processing

### Help & Support

**💬 HELP & SUPPORT Section**

**❓ Quick Help**
- Shows quick help dialog
- Common troubleshooting tips

**📖 Manual**
- Opens this user manual
- Complete documentation

**🆘 Live Tech Support**
- Opens AI help assistant
- Copy context for Claude/ChatGPT/Gemini
- Get specialized help

### Additional Tools

**👥 Contacts**
- Opens Windows Contacts/People app
- Quick access to contact management

**📧 Email**
- Opens your default email client
- Quick access to email

**🐛 Report Bug**
- Submit bug reports
- Help improve the app

**💡 Request Feature**
- Suggest new features
- Share ideas

### API Setup (One-time - At Bottom)

**⚙️ API SETUP (One-time) Section**

These setup wizards are at the bottom since you only need them once during initial setup:

**🔧 Setup Vision API**
- Step-by-step wizard for Google Vision API
- Required for PDF OCR (Step 1)

**🗺️ Setup Maps API**
- Google Maps API setup wizard
- Optional but enhances route mapping (Step 3)

**📧 Setup Gmail API**
- Gmail API configuration wizard
- Required for email steps (4, 6, 11, 12)

**👥 Setup People/Calendar**
- People and Calendar API setup wizard
- Required for contacts (Step 5) and calendar (Step 7)

---

## Interactive Chatbot Assistant

### 🤖 NEW Feature: Your Sassy, Helpful AI Friend!

The Court Visitor App now includes an interactive chatbot with personality!

**Where to find it:**
- Look for the **purple button** at position #2 in the sidebar
- Says "🤖 Ask Chatbot"
- Can't miss it - it's prominently displayed!

**What it can do:**
- Answer questions about any of the 14 workflow steps
- Provide troubleshooting help (especially for Excel issues!)
- Offer encouragement for volunteers
- Tell guardian and court visitor jokes
- Provide "real talk" appreciation for your volunteer work

**Personality:**
- **Sassy** like Monday.com
- **Charming** like Matt Rife
- **Appreciative** of your volunteer work
- **Helpful** with technical issues
- **Fun** - tracks your visits and celebrates your dedication!

**Quick Questions Buttons:**
- 📋 What are the steps?
- 😫 Excel is locked AGAIN
- 🤷 Step 1 hates me
- 📧 Google is mad at me
- 😂 Tell me a joke
- 🎉 Random fun fact

**Try these phrases:**
- "joke" - Get uplifting court visitor & guardian jokes
- "tired" or "hard day" - Get volunteer encouragement
- "guardian" or "volunteer" - Appreciation messages
- "thanks" - Get genuine appreciation (chatbot flips it back to you!)
- "hello" - Get a sassy greeting

**Visit Tracking:**
- Chatbot remembers how many times you've visited
- First visit: Extra welcoming with compliments
- Return visits: Celebrates your dedication
- Many visits: "Real talk" about your heroic volunteer work

---

## API Setup Wizards

### NEW! Step-by-Step Setup Guides

The app now includes interactive wizards for setting up Google APIs. Find them in the sidebar under "⚙️ API SETUP (One-time)".

#### Why APIs Are Needed

- **Vision API** - Required for PDF OCR text extraction (Step 1)
- **Maps API** - Optional for enhanced route mapping (Step 3)
- **Gmail API** - Required for sending emails (Steps 4, 6, 11, 12)
- **People API** - Required for adding contacts (Step 5)
- **Calendar API** - Required for calendar events (Step 7)

#### Setup Process

Each wizard provides:
1. **What you need** - Requirements and prerequisites
2. **Step-by-step instructions** - Clear, detailed setup guide
3. **Where to save files** - Exact folder locations
4. **Troubleshooting** - Common issues and solutions
5. **Testing** - How to verify setup worked

#### API Status Indicators

Each step shows its API status:
- ✅ **Ready** - API is configured and ready
- ⚠️ **Partial** - Works with limited features (e.g., Maps)
- 🔒 **Missing** - API needs to be set up

---

## Troubleshooting

### App Won't Start

**Problem:** Double-clicking the launcher does nothing

**Solutions:**
1. Check that Python is installed: `python --version`
2. Right-click the VBS file → "Open with" → "Microsoft Windows Based Script Host"
3. Try running from command line to see errors
4. Verify you're in the correct directory: `C:\GoogleSync\GuardianShip_App`

---

### Step Fails to Run

**Problem:** Clicking a step button shows an error

**Solutions:**
1. Check the output window for specific [FAIL] messages
2. Verify the automation script exists at the specified path
3. Ensure required input files are present (Excel, PDFs, etc.)
4. Check that Python dependencies are installed for that script
5. Try running the script directly to isolate the issue
6. **NEW: Check API status indicators** - 🔒 means API needs setup

---

### No PDF Files Found (Step 1)

**Problem:** "No PDF files found" message when processing

**Solutions:**
1. Check that PDFs are in: `C:\GoogleSync\GuardianShip_App\New Files\`
2. Verify files are named ORDER.pdf and ARP.pdf (or contain those words)
3. Make sure files are actual PDFs, not shortcuts or other file types
4. Close any open PDF files

---

### Excel File Locked

**Problem:** "Excel file is locked" error

**Ask the Chatbot!** It has sassy, helpful responses for this exact issue.

**Solutions:**
1. **Close Excel** (Yes, really. Just close it.)
2. **Wait 30 seconds** (Excel holds onto files like a toddler with a toy)
3. **Check Task Manager** (Ctrl+Shift+Esc → Find Excel → End Task)
4. **Restart if desperate** (Works 60% of the time, every time)
5. **Clear status columns** in Excel to re-run steps

---

### Excel File Not Found

**Problem:** "Excel file not found" error

**Solutions:**
1. Run Step 1 to create the database
2. Check that folder exists: `C:\GoogleSync\GuardianShip_App\App Data\`
3. Verify Excel file hasn't been moved or renamed

---

### Script Output Window is Blank

**Problem:** Processing window opens but shows no output

**Solutions:**
1. Wait - some scripts take time to start
2. Check that Python can run the script
3. Look for console errors
4. Verify script permissions
5. Check if Excel is open (common cause!)

---

### Email/Calendar Steps Fail

**Problem:** Email or calendar automation doesn't work

**Solutions:**
1. **Use API Setup Wizards!** Click the API setup buttons in sidebar
2. Check Google API credentials are set up correctly
3. Verify internet connection
4. Check that email addresses are valid in Excel
5. Look for authentication errors in output window
6. Re-authorize if you see "token expired" errors

---

### Chatbot Button Has No Text

**Problem:** Purple button in sidebar is blank or has no emoji

**Solutions:**
1. Close and reopen the app
2. Check that the app fully loaded (wait a few seconds)
3. Try clicking the blank purple button - it may still work!
4. Restart the app using the Desktop shortcut

---

### Sidebar Scrollbar Not Visible

**Problem:** Can't see scrollbar to access API setup buttons

**Solutions:**
1. Make window wider (drag edge)
2. Use mousewheel to scroll sidebar
3. Window should be 1400px wide automatically
4. Most-used buttons are at top - only need to scroll for one-time API setup

---

## Tips & Best Practices

### Daily Workflow

1. Start each day: Process new PDFs (Step 1)
2. Follow the sequence: Steps are numbered in order for a reason
3. Check after each step: Verify output before moving to next step
4. Use Quick Access: Sidebar buttons save time for common tasks
5. **NEW: Ask the Chatbot** when you're stuck or need encouragement!

### Data Quality

1. Review extracted data: Always check Excel after Step 1
2. Correct errors early: Fix typos in Excel before running other steps
3. Validate addresses: Ensure addresses are complete for mapping
4. Check email addresses: Verify guardian emails before sending
5. **Use 3-column layout**: See all steps at once to track progress

### File Organization

1. Clear New Files folder: After processing, PDFs auto-move to case folders
2. Backup Excel regularly: Save copies of your database
3. Keep scripts untouched: The app calls your scripts - don't modify them
4. Organize by case: Use New Clients and Completed folders structure
5. Archive old cases: Move completed cases periodically

### Performance

1. One step at a time: Don't run multiple steps simultaneously
2. Watch the output: Processing windows show progress and errors
3. Be patient: OCR and document processing takes time
4. Close completed windows: Keep workspace clean
5. **Look for [OK] or [FAIL]**: New exit codes show clear status

### Error Prevention

1. Check file paths: Ensure all automation scripts are in correct locations
2. Verify permissions: Make sure Python can read/write files
3. Test after updates: If you update a script, test it
4. Keep backups: Save working versions of scripts
5. **Set up APIs once**: Use the setup wizards and you're done forever!

### Monthly Tasks

- **End of month:** Run Step 13 (Payment Form) and Step 14 (Mileage Log)
- **Archive completed cases:** Move old cases to archive folder
- **Backup Excel database:** Save monthly snapshots
- **Review automation logs:** Check for recurring errors

### Using the Chatbot Effectively

1. **Start here for questions** - It knows the workflow inside-out
2. **Type naturally** - Just ask your question in plain English
3. **Try keywords** - "joke", "tired", "volunteer", "guardian"
4. **Use Quick Questions** - Buttons for common issues
5. **Get encouragement** - Chatbot celebrates your volunteer work!

---

## Keyboard Shortcuts

- **Alt+F4** - Close current window
- **Esc** - Cancel current operation (when available)
- **F1** - Open help (in processing windows)
- **Mousewheel** - Scroll sidebar up/down

---

## File Locations

### Main Application

- **App Folder:** `C:\GoogleSync\GuardianShip_App\`
- **Launcher:** `C:\GoogleSync\GuardianShip_App\Launch Court Visitor App.vbs`
- **Main Script:** `C:\GoogleSync\GuardianShip_App\guardianship_app.py`
- **Desktop Shortcut:** `%USERPROFILE%\Desktop\Court Visitor App.lnk`

### Data Files

- **Excel Database:** `C:\GoogleSync\GuardianShip_App\App Data\ward_guardian_info.xlsx`
- **New PDFs:** `C:\GoogleSync\GuardianShip_App\New Files\`
- **Active Cases:** `C:\GoogleSync\GuardianShip_App\New Clients\`
- **Completed Cases:** `C:\GoogleSync\GuardianShip_App\Completed\`

### Configuration

- **API Credentials:** `C:\GoogleSync\GuardianShip_App\Config\API\`
- **API Keys:** `C:\GoogleSync\GuardianShip_App\Config\Keys\`
- **Chatbot Data:** `C:\GoogleSync\GuardianShip_App\App Data\chatbot_stats.json`

### Automation Scripts

- **All Automation:** `C:\GoogleSync\GuardianShip_App\Automation\`
- **Main Extractor:** `C:\GoogleSync\GuardianShip_App\guardian_extractor_claudecode20251023_bestever_11pm.py`

### Documentation

- **This Manual:** `C:\GoogleSync\GuardianShip_App\Court_Visitor_App_Manual.pdf`
- **README:** `C:\GoogleSync\GuardianShip_App\README_FIRST.pdf`
- **Getting Started:** Click button in app sidebar

---

## Support

### Getting Help

1. **🤖 Ask the Chatbot** - Click the purple button! Available 24/7
2. Click the **"❓ Quick Help"** button in the sidebar
3. Review this manual (click **"📖 Manual"** button)
4. Check the output window for specific error messages with [FAIL] tags
5. Use **"🆘 Live Tech Support"** for AI-powered help
6. Review the README_FIRST.pdf file

### Reporting Issues

When reporting problems via **"🐛 Report Bug"**, include:
- Which step was running
- Error messages from output window (especially [FAIL] messages)
- What you were trying to do
- What files were being processed
- API status indicators (✅⚠️🔒)

### Feature Requests

Have ideas for Step 15 or other improvements?
- Click **"💡 Request Feature"** button
- Describe what you'd like to see
- Explain how it would help your workflow

---

## Version Information

### Court Visitor App v2.0 (October 2024)

**Major Features:**
- 14-step automated workflow (with room for Step 15!)
- **NEW: 3-column layout** - see all steps at once
- **NEW: Interactive chatbot** with personality
- **NEW: API setup wizards** - step-by-step guides
- **NEW: Scrollable sidebar** with organized buttons
- **NEW: Exit code reporting** - clear [OK]/[FAIL] status
- All automation scripts connected
- No modifications to existing scripts
- Real-time output monitoring
- Integrated help system

**What's Different from v1.0:**
- Layout changed from 2 columns to 3 columns
- Chatbot assistant added (purple button #2)
- API setup wizards added to sidebar
- Getting Started dialog added
- Sidebar made scrollable with reorganized buttons
- Window size increased to 1400x850px
- All 14 steps now report proper exit codes
- Desktop shortcut uses VBS for silent launch

---

## Important Notes

### About Your Scripts

- **NO SCRIPTS ARE MODIFIED** - The app only calls your existing scripts
- All your working scripts remain untouched in their original locations
- The app is completely separate in the GuardianShip_App folder
- Each step verifies the script exists before running
- Script output is displayed in real-time for monitoring
- Exit codes now clearly show [OK] or [FAIL] status

### About the 3-Column Layout

- **Column 1 (Left):** Input Phase - Steps 1-5 get data ready
- **Column 2 (Middle):** Communication - Steps 6-10 handle correspondence
- **Column 3 (Right):** Wrap-Up - Steps 11-14 finish up (plus room for Step 15!)
- All 14 steps visible without scrolling
- Same card style and functionality as before
- Just rearranged for better visibility

### About the Chatbot

- Runs locally - no internet required for basic features
- Tracks your visits to provide personalized greetings
- Has personality but stays professional and respectful
- Celebrates your volunteer work as a court visitor
- Knowledge base includes all 14 workflow steps
- Can't access external information (by design)
- Data stored locally in App Data folder

---

**Last Updated:** October 31, 2024

**App Version:** 2.0

**For Support:** Use the 🤖 Chatbot, ❓ Quick Help, or 🆘 Live Tech Support buttons in the app!

---

*Thank you for being a dedicated court visitor volunteer! Your work makes a real difference in the lives of guardians and wards.* 💜
