# GitHub Quick Start - Exact Commands to Run

**Ready to push your code to GitHub? Follow these steps exactly.**

---

## Important First Steps

### 1. Clear Templates of Personal Information

Before pushing to GitHub, make sure your templates are blank:

1. Open `Templates/Mileage_Reimbursement_Form.xlsx`
2. Clear cells B8-B11 (your name, vendor #, etc.)
3. Save and close

Check the other templates too (Payment Form, CVR).

### 2. Verify No Sensitive Data

Make sure these folders exist but are empty (or contain no real data):
- `App Data/` - Should not contain real ward data
- `New Files/` - Should be empty
- `Config/API/` - Should not contain real API tokens

The `.gitignore` file will prevent these from being uploaded anyway, but double-check!

---

## Step-by-Step Commands

### Step 1: Open PowerShell

1. Press `Windows + X`
2. Click "Windows PowerShell" or "Terminal"
3. Navigate to your project:

```powershell
cd C:\GoogleSync\GuardianShip_App
```

### Step 2: Configure Git (First Time Only)

```powershell
# Set your name (for commit history)
git config user.name "Your Name"

# Set your email (use support@guardianshipeasy.com)
git config user.email "support@guardianshipeasy.com"
```

### Step 3: Initialize Git Repository

```powershell
# Initialize Git in this folder
git init

# Set main branch name
git branch -M main
```

### Step 4: Add Files to Git

```powershell
# See what files will be added (.gitignore protects sensitive files)
git status

# Add all files (except those in .gitignore)
git add .

# Check what was added
git status
```

**Expected output:**
You should see files like:
- `guardianship_app.py`
- `Scripts/`
- `Automation/`
- `Templates/`
- `Documentation/`
- `EULA.txt`
- `CHANGELOG.md`
- `README.md`

You should NOT see:
- `App Data/ward_guardian_info.xlsx`
- `Config/API/`
- `New Clients/`
- Any `.json` files with tokens

### Step 5: Create First Commit

```powershell
# Create the first commit
git commit -m "Initial commit: Court Visitor App v1.0.0"
```

### Step 6: Connect to GitHub

```powershell
# Add GitHub as remote (use YOUR GitHub URL)
git remote add origin https://github.com/mayehres-ops/Court_Visitor_App.git

# Verify it was added
git remote -v
```

### Step 7: Push to GitHub

```powershell
# Push code to GitHub
git push -u origin main
```

**You'll be prompted for credentials:**

- **Username**: Your GitHub username
- **Password**: Use a **Personal Access Token** (NOT your password!)

### Step 8: Get Personal Access Token (If Needed)

If you don't have a token yet:

1. **Go to GitHub**:
   - Click your profile picture → Settings
   - Scroll down → Developer settings
   - Personal access tokens → Tokens (classic)

2. **Generate New Token**:
   - Click "Generate new token (classic)"
   - Note: "Court Visitor App"
   - Expiration: 90 days
   - Scopes: Check `repo` (full control)

3. **Copy Token**:
   - Copy the token (starts with `ghp_...`)
   - **Save it** - you won't see it again!

4. **Use as Password**:
   - When `git push` asks for password, paste the token

---

## Verify Upload Successful

1. **Go to GitHub**:
   - Navigate to `https://github.com/mayehres-ops/Court_Visitor_App`

2. **Check Files**:
   - You should see all your files
   - README.md should be displayed on the main page

3. **Verify No Sensitive Data**:
   - Check that `App Data/` folder is NOT there
   - Check that `Config/API/` is NOT there
   - Look for any `.json` files - there should be NONE

---

## Common Errors and Fixes

### Error: "Authentication failed"

**Cause:** Using password instead of token

**Fix:** Generate a Personal Access Token and use that instead

### Error: "Repository not found"

**Cause:** Wrong URL or repo doesn't exist

**Fix:** Double-check the URL:
```powershell
git remote -v
```

If wrong, fix it:
```powershell
git remote remove origin
git remote add origin https://github.com/mayehres-ops/Court_Visitor_App.git
```

### Error: "fatal: not a git repository"

**Cause:** You're not in the right folder

**Fix:**
```powershell
cd C:\GoogleSync\GuardianShip_App
git init
```

### Warning: "LF will be replaced by CRLF"

**Not an error** - This is normal on Windows, you can ignore it.

---

## After Successful Push

### What You Can Do Now:

1. **View Your Code on GitHub**:
   - Go to `https://github.com/mayehres-ops/Court_Visitor_App`
   - Browse the files
   - Read the README

2. **Build the Executable**:
   - Follow [PyInstaller Build Guide](PYINSTALLER_BUILD_GUIDE.md)
   - Build `CourtVisitorApp.exe`

3. **Create First Release**:
   - See instructions below

---

## Creating Your First Release

After building the .exe:

### Step 1: Go to Releases

1. Navigate to `https://github.com/mayehres-ops/Court_Visitor_App`
2. Click "Releases" in the right sidebar
3. Click "Create a new release"

### Step 2: Fill Out Release Form

- **Choose a tag**: `v1.0.0` (create new tag)
- **Release title**: `Court Visitor App v1.0.0`
- **Description**: Copy from CHANGELOG.md or write:

```markdown
## Court Visitor App v1.0.0

Initial release of the Court Visitor workflow automation software.

### Features

- Complete 14-step workflow automation
- OCR data extraction from PDFs
- Automatic form generation (Mileage, Payment, CVR)
- Google API integration (Gmail, Calendar, Sheets, Drive, Maps)
- Court Visitor configuration system
- User data backup functionality
- Desktop shortcut creation
- Legal protections (EULA, copyright notices)

### Installation

1. Download `CourtVisitorApp_v1.0.zip`
2. Extract to your desired location
3. Run `CourtVisitorApp.exe`
4. Follow first-run setup wizard

### Requirements

- Windows 10 or 11
- Microsoft Office (Word and Excel)
- Internet connection for Google APIs

For detailed instructions, see the Installation Guide in the Documentation folder.
```

### Step 3: Attach Files

- Click "Attach binaries"
- Upload `CourtVisitorApp_v1.0.zip` (your distribution package)

### Step 4: Publish

- Check "This is a pre-release" if you want to test first
- Click "Publish release"

### Step 5: Share with Users

Send users this URL:
```
https://github.com/mayehres-ops/Court_Visitor_App/releases/latest
```

They'll see a download button for your ZIP file.

---

## Updating the App (Pushing v1.1, v1.2, etc.)

When you make changes:

```powershell
# Navigate to project
cd C:\GoogleSync\GuardianShip_App

# Check what changed
git status

# Add all changes
git add .

# Commit with descriptive message
git commit -m "Add custom app icon and auto-updater"

# Push to GitHub
git push
```

Then create a new Release on GitHub with the new version number.

---

## Getting Help

**If you get stuck:**

1. Check error message carefully
2. Search GitHub Docs: https://docs.github.com
3. Review [GITHUB_SETUP_GUIDE.md](GITHUB_SETUP_GUIDE.md) for detailed explanations
4. Email support@guardianshipeasy.com with:
   - What command you ran
   - What error you got
   - Screenshot if possible

---

## Security Reminder

**Before EVERY git push:**

```powershell
# See what you're about to upload
git status

# See the actual changes
git diff
```

Make sure NO sensitive data is being uploaded:
- ✗ No API tokens
- ✗ No personal ward/guardian data
- ✗ No real configuration files
- ✗ No Excel files with real data

The `.gitignore` file protects you, but always double-check!

---

## Next Steps

1. ✓ Push code to GitHub (you're doing this now!)
2. ☐ Build executable with PyInstaller
3. ☐ Test the .exe thoroughly
4. ☐ Create distribution package (ZIP)
5. ☐ Create first GitHub Release
6. ☐ Share download link with authorized users

---

**Ready? Let's push to GitHub!**

Just copy and paste the commands above into PowerShell, one section at a time.

---

**Document Version:** 1.0
**Last Updated:** November 6, 2024
