# CVR Template: Content Control Naming Instructions

## How to Name Content Controls in Word

### Step 1: Open the CVR Template
1. Open: `C:\GoogleSync\GuardianShip_App\App Data\Templates\Court Visitor Report fillable new.docx`
2. **Enable Developer Tab** (if not visible):
   - Click **File** → **Options** → **Customize Ribbon**
   - Check the box next to **"Developer"**
   - Click **OK**

### Step 2: Name Each Content Control

For each content control you want to fill from Google Form:

1. Click on the content control (the field where you want data to appear)
2. Click the **Developer** tab
3. Click **"Properties"** button
4. In the **"Title"** field, enter the exact control name from the list below
5. Click **OK**
6. Repeat for each control

---

## Content Controls to Name (27 Total)

### Section: Home Information

| Where to Find in CVR | Control Name to Enter | Google Form Question |
|---------------------|----------------------|---------------------|
| "Ward description" area | `ward_description` | Would you please tell me a little about the person under your care? |
| "Visit information needed" area | `visit_info` | Please provide any information needed to visit... |
| "Do you live together?" checkbox | `livetogether` | Do you live together? |
| "Does ward have own bed?" checkbox | `ownbed` | Does he/she have their own bed? |
| "Hot water available?" checkbox | `hotwater` | Is hot water available? |
| "Air conditioning/heating?" checkbox | `hvac` | Does the residence have air conditioning/heating |
| "Areas accessible?" checkbox | `accessible` | Are most areas accessible to the person under your care? |
| "Areas not accessible explanation" field | `accessible_explain` | If no, explain which areas are not accessible and why. |
| "Access to telephone/TV/radio" field | `access_items` | Does the individual have access to the following: |
| "Living situation" field | `living_situation` | Does the person under your care live: |
| "Other relatives home info" field | `relative_info` | If Ward resides in Other Relatives Home state the name and relationship |
| "Facility name" field | `facility_name` | If [a] facility, the name of the place the ward lives. |

### Section: Guardian Information

| Where to Find in CVR | Control Name to Enter | Google Form Question |
|---------------------|----------------------|---------------------|
| "Guardian present at meeting?" checkbox | `guardian_present` | Will you be present during the meeting? |
| "Other attendees" field | `other_attendees` | If no, please provide the name and relationship of any other individuals... |

### Section: Activities

| Where to Find in CVR | Control Name to Enter | Google Form Question |
|---------------------|----------------------|---------------------|
| "Participates in activities" field | `activities` | Does the person you care for participate in any of the following |
| "Has visitors?" checkbox | `has_visitors` | Does the individual have any visitors? |
| "Exercises regularly?" checkbox | `exercises` | Does the individual exercise regularly? |
| "Transportation provided?" checkbox | `transportation` | If the individual wants to go on trips, is transportation provided? |

### Section: Safety

| Where to Find in CVR | Control Name to Enter | Google Form Question |
|---------------------|----------------------|---------------------|
| "Fire extinguishers available?" checkbox | `fire_extinguisher` | Are fire extinguishers available? |
| "Understands fire safety?" checkbox | `understands_fire` | Is the person under your care able to understand how to use a fire... |
| "Knows fire extinguisher location?" checkbox | `knows_fire_location` | Does the individual know where the fire extinguisher is located? |
| "Safe alone?" checkbox | `safe_alone` | Is the person under your care safe in the home alone? |

### Section: Health & Education

| Where to Find in CVR | Control Name to Enter | Google Form Question |
|---------------------|----------------------|---------------------|
| "Can read?" checkbox | `can_read` | Can the individual read: |
| "Can write?" checkbox | `can_write` | Can the individual write: |
| "Oriented in time/space?" checkbox | `oriented` | Is the individual oriented in time and space? |
| "Needs help with..." field | `needs_help` | Does the person in your care need help with any of the following: |
| "Has conditions..." field | `conditions` | Does the person in your care have any of the following? |

---

## Tips

1. **Start with the most important fields** - You don't have to name all 27 at once. Start with the fields you use most.

2. **Look for checkbox symbols** - The CVR template already has many checkboxes (☐). These are the ones that should get names like `ownbed`, `hotwater`, `hvac`, etc.

3. **Text areas** - Longer text fields should get names like `ward_description`, `visit_info`, `accessible_explain`, etc.

4. **Exact spelling matters** - Make sure the control name matches EXACTLY what's in the table above (case-sensitive).

5. **Don't rename existing controls** - The following controls are already named and used by Step 8. **Don't change these**:
   - causeno
   - wardfirst, wardmiddle, wardlast
   - visitdate, visittime
   - waddress
   - wfirst, wlast
   - datefiled

---

## After Naming Controls

Once you've named the controls:

1. **Save the template**
2. Run the verification script to check your work:
   ```
   python test_step10_setup.py
   ```
3. If verification passes, test Step 10:
   ```
   python google_sheets_cvr_autofill.py
   ```

---

## Need Help?

- Run `python test_step10_setup.py` to see which controls still need naming
- The script will show you which ones are found vs. missing
