# Step 10 Setup Guide: Google Form to CVR Auto-Fill

## Overview
Step 10 will automatically fill CVR documents with responses from the Google Form that guardians complete after receiving the meeting confirmation email (Step 6).

## Current Status
- ✅ Step 8 (Generate CVR) is working - fills basic info from Excel
- ✅ Google Form created with 29 questions
- ✅ Mapping configuration created: `Config/cvr_google_form_mapping.json`
- ⚠️ Google Sheets API needs to be enabled
- ⚠️ CVR template content controls need to be named
- ⚠️ Google Form needs cause number field for matching

---

## What YOU Need To Do

### Task 1: Enable Google Sheets API (REQUIRED)

**Option A: Using Service Account** (Recommended - no user interaction needed)

1. Go to: https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=172649392933
2. Click "Enable" button
3. Wait 2-3 minutes for activation
4. Share the Google Sheet with your service account email:
   - Open the response spreadsheet
   - Click "Share" button
   - Add the email from your service account JSON file (format: `xxxxx@xxxxx.iam.gserviceaccount.com`)
   - Give "Viewer" permission

**Option B: Using OAuth** (Alternative - requires user login)

- We can set up OAuth credentials instead
- Requires one-time browser authentication
- Contact me if you prefer this method

### Task 2: Add Cause Number to Google Form (CRITICAL)

The form needs a way to match responses to specific CVR files. Add this question at the **TOP** of the form:

**New Question #1:**
- Question text: "Case/Cause Number (found on court documents)"
- Type: Short answer
- Required: Yes
- Example/Help text: "Example: 00-074302 or 05-082942"

This ensures we can match form responses to the correct CVR file.

### Task 3: Name Content Controls in CVR Template

The CVR template has 96 unnamed content controls that need to be named. Here's how:

**Instructions:**
1. Open: `C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx`
2. Go to Developer tab (if not visible: File → Options → Customize Ribbon → Check "Developer")
3. For each content control listed below, click on it and set the **Title** property

**Required Control Names:**

| Control Name | Google Form Question | Type |
|-------------|---------------------|------|
| `ward_description` | Would you please tell me a little about the person under your care? | Text |
| `guardian_present` | Will you be present during the meeting? | Checkbox/Text |
| `other_attendees` | If no, please provide the name and relationship... | Text |
| `visit_info` | Please provide any information needed to visit... | Text |
| `livetogether` | Do you live together? | Checkbox/Text |
| `ownbed` | Does he/she have their own bed? | Checkbox/Text |
| `hotwater` | Is hot water available? | Checkbox/Text |
| `hvac` | Does the residence have air conditioning/heating | Checkbox/Text |
| `accessible` | Are most areas accessible to the person under your care? | Checkbox/Text |
| `accessible_explain` | If no, explain which areas are not accessible... | Text |
| `access_items` | Does the individual have access to the following | Text |
| `living_situation` | Does the person under your care live | Text |
| `relative_info` | If Ward resides in Other Relatives Home... | Text |
| `facility_name` | If [a] facility, the name of the place... | Text |
| `activities` | Does the person you care for participate in... | Text |
| `has_visitors` | Does the individual have any visitors? | Checkbox/Text |
| `exercises` | Does the individual exercise regularly? | Checkbox/Text |
| `transportation` | If the individual wants to go on trips... | Checkbox/Text |
| `fire_extinguisher` | Are fire extinguishers available? | Checkbox/Text |
| `understands_fire` | Is the person under your care able to understand... | Checkbox/Text |
| `knows_fire_location` | Does the individual know where the fire extinguisher... | Checkbox/Text |
| `safe_alone` | Is the person under your care safe in the home alone? | Checkbox/Text |
| `can_read` | Can the individual read | Checkbox/Text |
| `can_write` | Can the individual write | Checkbox/Text |
| `oriented` | Is the individual oriented in time and space? | Checkbox/Text |
| `needs_help` | Does the person in your care need help with... | Text |
| `conditions` | Does the person in your care have any of the following? | Text |

**Tips:**
- The CVR template already has BuildingBlock checkboxes (the "☐" symbols)
- You can convert these to modern Checkbox content controls if desired
- For multi-select questions (like "access to telephone/TV/radio"), you can either:
  - Create separate controls for each option (e.g., `access_telephone`, `access_tv`, etc.)
  - OR use one text control and display comma-separated values

### Task 4: Test with Sample Data

Once the above is complete:
1. Fill out the Google Form with test data for one of your cases
2. Run Step 10 from the GuardianShip App
3. Verify the CVR is filled correctly

---

## What I Will Do Next

Once you complete Tasks 1-3 above, I will:

1. **Update Step 10 script** to:
   - Read form responses from Google Sheet (using Spreadsheet ID: `1O9Sv5M8SEdD_bbxew28QScKOazCYivUTvTxpMfZl1HI`)
   - Match responses to CVR files using cause number
   - Auto-fill ONLY the fields NOT already filled by Step 8
   - Handle Yes/No checkboxes properly
   - Handle multi-select checkboxes (comma-separated text)
   - Skip fields that are already populated from Excel

2. **Add Step 10 wizard** to the main app:
   - Button to trigger auto-fill
   - Progress indicator showing which CVRs are being filled
   - Error handling for missing data or API issues

3. **Create end-user documentation** explaining the complete workflow:
   - Step 1-7: Case setup, OCR, folders, emails
   - Step 8: Generate CVR from Excel (basic info)
   - Guardian fills Google Form (via link in email)
   - Step 10: Auto-fill CVR with Google Form responses (additional details)
   - Steps 11-14: Final review, submission, tracking

---

## Notes

- **No data loss**: Step 10 will NOT overwrite fields already filled by Step 8
- **Flexible checkboxes**: Works with both actual checkbox controls and text fields showing "Yes"/"No"
- **Partial responses**: If a guardian doesn't complete all fields, only available data will be filled
- **Multiple responses**: If there are multiple form responses for the same cause number, the most recent one will be used

---

## Questions?

- **Q: Can I change the Google Form questions?**
  - A: Yes, but you'll need to update the mapping file (`Config/cvr_google_form_mapping.json`) to match

- **Q: What if I don't want to name all 96 controls?**
  - A: You can name just the important ones. Unmapped form questions will be shown to the user for manual entry

- **Q: Can I use a different form/spreadsheet?**
  - A: Yes, just provide the new Spreadsheet ID and I'll update the script

---

## Files Created

1. **`Config/cvr_google_form_mapping.json`** - Mapping between Google Form questions and CVR controls
2. **`Scripts/analyze_cvr_template.py`** - Helper script to analyze template controls
3. **`Scripts/cvr_content_control_utils.py`** - Reusable utility for filling Word content controls
4. **This guide** - `STEP_10_SETUP_GUIDE.md`

---

## Ready to Proceed?

Let me know when you've completed:
- ☐ Task 1: Google Sheets API enabled
- ☐ Task 2: Cause number added to Google Form
- ☐ Task 3: Content controls named in CVR template

Then I'll implement the Step 10 auto-fill functionality!
