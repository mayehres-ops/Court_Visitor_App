# Proposed Folder Structure for Distribution

**Date:** November 5, 2024
**Purpose:** Clean, professional structure for end users

---

## Current Problem

### Root Directory is Messy:
```
C:\GoogleSync\GuardianShip_App\
├── guardianship_app.py                        ← Main app ✅
├── guardian_extractor_*.py                    ← Should be in Automation ⚠️
├── google_sheets_cvr_integration_fixed.py     ← Should be in Automation ⚠️
├── email_cvr_to_supervisor.py                 ← Should be in Automation ⚠️
├── auto_updater.py                            ← Core file ✅
├── setup_wizard.py                            ← Core file ✅
├── (20+ other files)                          ← TOO MUCH IN ROOT ❌
```

**Issues:**
- Root is cluttered
- Hard to find main app
- Confusing for end users
- Scripts exposed that should be hidden

---

## Proposed Clean Structure

### Option A: Hide Automation Completely (RECOMMENDED)

```
C:\CourtVisitorApp\                            ← End user installs here
│
├── 📄 guardianship_app.py                     ← Main app (only .py in root)
├── 📄 Launch Court Visitor App.vbs            ← Launcher (double-click this)
├── 📄 EULA.txt                                ← License agreement
├── 📄 README.txt                              ← Quick start
├── 📄 User_Manual.pdf                         ← Full manual
│
├── 📁 Config/                                 ← User configuration
│   ├── 📁 API/                                ← Google credentials (user adds)
│   │   ├── credentials.json                   ← User provides
│   │   ├── token_gmail.json                   ← Generated on first use
│   │   └── README.txt                         ← Instructions
│   └── app_settings.json                      ← App settings (auto-created)
│
├── 📁 App Data/                               ← User's data
│   ├── ward_guardian_info.xlsx                ← User's database
│   ├── 📁 Backup/                             ← Auto backups of database
│   ├── 📁 Inbox/                              ← Email downloads (Step 5)
│   ├── 📁 Staging/                            ← Temp files (auto-cleaned)
│   └── 📁 Templates/                          ← Word templates
│       ├── Court Visitor Report fillable new.docx
│       ├── Court Visitor Payment Form TEMPLATE.docx
│       ├── MILEAGE LOG CV Visitors template.docx
│       └── Ward Map Sheet.docx
│
├── 📁 New Files/                              ← User drops PDFs here (Step 1)
├── 📁 New Clients/                            ← Case folders (created by Step 2)
├── 📁 Completed/                              ← Finished cases (manual move)
│
├── 📁 _Internal/                              ← HIDDEN FOLDER (all scripts)
│   │   ├── attrib +h "_Internal"              ← Make folder hidden on Windows
│   │
│   ├── 📁 Core/                               ← Core processing scripts
│   │   ├── guardian_extractor.py              ← OCR script (Step 1)
│   │   ├── google_sheets_integration.py       ← Sheets autofill (Step 10)
│   │   ├── email_cvr_to_supervisor.py         ← Email CVR (Step 6)
│   │   └── auto_updater.py                    ← Update checker
│   │
│   ├── 📁 Utils/                              ← Utility scripts
│   │   ├── app_paths.py                       ← Path management
│   │   ├── app_config_manager.py              ← Settings manager
│   │   └── cvr_content_control_utils.py       ← Word utils
│   │
│   └── 📁 Automation/                         ← Step automation scripts
│       ├── 📁 Step_02_Create_Folders/
│       │   └── cvr_folder_builder.py
│       ├── 📁 Step_07_Build_Map_Sheet/
│       │   └── build_map_sheet.py
│       ├── 📁 Step_08_Generate_CVR/
│       │   └── build_cvr_from_excel.py
│       ├── 📁 Step_09_Email_Meeting_Request/
│       │   └── send_guardian_emails.py
│       ├── 📁 Step_10_Autofill_Google_CVR/
│       │   └── (uses Core/google_sheets_integration.py)
│       ├── 📁 Step_11_Email_Confirmation/
│       │   └── send_confirmation_email.py
│       ├── 📁 Step_12_Create_Calendar_Event/
│       │   └── create_calendar_event.py
│       ├── 📁 Step_13_Add_Contacts/
│       │   └── add_guardians_to_contacts.py
│       ├── 📁 Step_14_Payment_Forms/
│       │   └── build_payment_forms.py
│       └── 📁 Step_15_Mileage_Log/
│           └── build_mileage_forms.py
```

**Pros:**
- ✅ Clean root directory
- ✅ All scripts hidden from user
- ✅ Professional appearance
- ✅ User only sees what they need
- ✅ Easy to navigate
- ✅ Scripts organized by purpose

**Cons:**
- ⚠️ Requires reorganizing current structure
- ⚠️ Need to update all script paths
- ⚠️ Testing needed after reorganization

---

### Option B: Keep Current Structure (Minimal Changes)

```
C:\CourtVisitorApp\                            ← End user installs here
│
├── guardianship_app.py                        ← Main app
├── guardian_extractor_*.py                    ← OCR (stays in root)
├── google_sheets_cvr_integration_fixed.py     ← Sheets (stays in root)
├── email_cvr_to_supervisor.py                 ← Email (stays in root)
├── auto_updater.py                            ← Core
├── setup_wizard.py                            ← Core
├── Launch Court Visitor App.vbs               ← Launcher
│
├── Config/                                    ← Same as current
├── App Data/                                  ← Same as current
├── Scripts/                                   ← Same as current
├── Automation/                                ← Same as current
├── New Files/                                 ← Same as current
├── New Clients/                               ← Same as current
└── Completed/                                 ← Same as current
```

**Pros:**
- ✅ No reorganization needed
- ✅ Less risk of breaking things
- ✅ Faster to distribute

**Cons:**
- ❌ Root directory cluttered
- ❌ Scripts visible to users
- ❌ Less professional
- ❌ Users might accidentally modify scripts

---

## Recommendation: Hybrid Approach (Best of Both)

### What to Do NOW for v1.0:

**Keep current structure for distribution** but:
1. ✅ Hide the `Automation` folder from users (Windows folder attribute)
2. ✅ Hide the `Scripts` folder from users
3. ✅ Add clear README.txt explaining what each folder is for
4. ✅ Fix hardcoded paths to work from any install location

**What to Do LATER for v2.0:**
1. Clean up root directory
2. Move core scripts to `_Internal` folder
3. Reorganize Automation by step number

### Why This Approach?

**For v1.0 (Now):**
- Less risk - don't move files during path fixes
- Focus on making it WORK first
- Get it distributed sooner

**For v2.0 (Later):**
- Reorganize when you have users testing
- Based on feedback
- After path system is proven stable

---

## Installation Process

### What Installer Creates:

```python
# Installer creates this structure:
folders_to_create = [
    "Config/API",
    "App Data/Backup",
    "App Data/Inbox",
    "App Data/Staging",
    "App Data/Templates",
    "New Files",
    "New Clients",
    "Completed",
]

# All other folders come from distribution package:
# - Scripts/ (included in ZIP)
# - Automation/ (included in ZIP)
```

### Installation Steps:

1. **User downloads:** `CourtVisitorApp_v1.0.0.zip`

2. **User extracts to chosen location** (default: `C:\CourtVisitorApp\`)

3. **Installer runs** (setup_wizard.py):
   ```python
   # Auto-detects where user extracted to
   install_dir = detect_installation_directory()

   # Creates necessary folders
   create_folder_structure(install_dir)

   # Installs Python dependencies
   install_dependencies()

   # Prompts for Court Visitor name
   setup_user_settings()

   # EULA acceptance
   show_eula()

   # License key entry
   activate_license()

   # Creates desktop shortcut (points to install_dir)
   create_shortcut(install_dir)
   ```

4. **User ready to use!**

---

## Hiding Folders from End Users

### Windows: Make Folder Hidden

```python
# In installer or first run:
import os
import subprocess

def hide_folder(folder_path):
    """Make folder hidden on Windows."""
    if os.name == 'nt':  # Windows
        subprocess.run(['attrib', '+h', folder_path], shell=True)

# Hide technical folders
hide_folder("C:\\CourtVisitorApp\\Scripts")
hide_folder("C:\\CourtVisitorApp\\Automation")
```

**Result:** Folders don't show in File Explorer unless user enables "Show Hidden Files"

### Alternative: Underscore Prefix

```
C:\CourtVisitorApp\
├── _Scripts/          ← Underscore indicates "internal"
└── _Automation/       ← Visual cue "don't touch"
```

**Less technical but clear to users.**

---

## Scripts That Must Stay in Root

### Keep in Root (User-Facing):
- ✅ `guardianship_app.py` - Main application
- ✅ `Launch Court Visitor App.vbs` - Launcher
- ✅ `setup_wizard.py` - First-run setup
- ✅ `EULA.txt` - License
- ✅ `README.txt` - Quick start
- ✅ `User_Manual.pdf` - Documentation

### Move to Hidden Folder (Internal):
- ⚠️ `guardian_extractor_*.py` - OCR processing
- ⚠️ `google_sheets_cvr_integration_fixed.py` - Sheets integration
- ⚠️ `email_cvr_to_supervisor.py` - Email CVR
- ⚠️ `auto_updater.py` - Update checker
- ⚠️ All Scripts/ files
- ⚠️ All Automation/ files

### How Main App Calls Hidden Scripts:

```python
# In guardianship_app.py:
from _Internal.Core.guardian_extractor import process_pdfs
from _Internal.Core.google_sheets_integration import autofill_google_cvr
from _Internal.Utils.app_paths import get_app_paths

# Or keep current structure for v1.0:
import guardian_extractor_claudecode20251023_bestever_11pm as ocr
from Scripts.app_paths import get_app_paths
```

---

## Your Decisions Needed

### Question 1: Installation Directory
**Should end users choose where to install?**

- **Option A:** Force `C:\CourtVisitorApp\` only (simpler)
- **Option B:** Let user choose (more professional) ← RECOMMENDED

**My recommendation:** Option B - We already built `app_paths.py` for this!

### Question 2: Folder Structure
**Clean up root directory now or later?**

- **Option A:** Keep current structure for v1.0, clean up for v2.0 ← RECOMMENDED
- **Option B:** Reorganize everything now before distribution

**My recommendation:** Option A - Less risk during path fixes

### Question 3: Hide Technical Folders
**How to hide Scripts and Automation folders?**

- **Option A:** Windows hidden attribute (`attrib +h`)
- **Option B:** Underscore prefix (`_Scripts`, `_Automation`)
- **Option C:** Move to `_Internal` folder
- **Option D:** Leave visible but add README warnings

**My recommendation:** Option A for v1.0, Option C for v2.0

### Question 4: Scripts in Root
**What to do about scripts currently in root directory?**

- **Option A:** Move to `_Internal/Core/` now
- **Option B:** Leave in root for v1.0, move later ← RECOMMENDED
- **Option C:** Move to `Scripts/` folder

**My recommendation:** Option B - Focus on path fixes first

---

## Implementation Plan

### For v1.0 Distribution (Next 2-3 Weeks):

**Phase 1: Fix Paths (Keep Current Structure)**
- Don't move any files yet
- Fix hardcoded paths to use `app_paths.py`
- Test from different install locations
- Verify all 14 steps work

**Phase 2: Hide Technical Folders**
- Add code to hide `Scripts/` folder
- Add code to hide `Automation/` folder
- Add clear README.txt in root
- Test on clean Windows machine

**Phase 3: Installer**
- Update `setup_wizard.py` to:
  - Let user choose install location (default: C:\CourtVisitorApp\)
  - Create all necessary folders
  - Hide technical folders
  - Prompt for settings

### For v2.0 (Future - After User Feedback):

**Phase 4: Reorganize Structure**
- Create `_Internal/` folder structure
- Move core scripts from root
- Reorganize Automation by step number
- Update all imports
- Test thoroughly

---

## Immediate Next Steps

1. **Don't reorganize folders yet** - Too risky during path fixes

2. **Fix paths first** using current structure

3. **After paths work** - Then decide on reorganization

4. **For now** - Focus on:
   - Creating backup
   - Fixing hardcoded paths
   - Making it work from any install location

---

## Summary

**Your Questions Answered:**

1. **Do users choose install location?**
   - YES (recommended) - app_paths.py handles this
   - Installer suggests `C:\CourtVisitorApp\` but allows choice

2. **Does installation create folders?**
   - YES - setup_wizard.py creates all necessary folders
   - Config/, App Data/, New Files/, etc.

3. **Scripts in root directory?**
   - KEEP FOR NOW - Don't move during path fixes
   - Move to hidden folder in v2.0

4. **Hide Automation folder?**
   - YES - Use Windows hidden attribute for v1.0
   - Move to _Internal/ for v2.0

**Bottom Line:** Don't reorganize folders now. Fix paths using current structure. Clean up organization in v2.0 after it's proven stable.

---

**Ready to proceed with path fixes using current folder structure?**
