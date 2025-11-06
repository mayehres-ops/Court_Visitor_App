# GitHub Setup & Distribution Guide

Complete guide for using GitHub to distribute the Court Visitor App.

---

## Understanding Your GitHub Repository

### What You're Looking At

Your screenshot shows a **new, empty GitHub repository** at:
```
https://github.com/mayehres-ops/Court_Visitor_App
```

This is a private repository, which is perfect for your needs.

### What is GitHub?

Think of GitHub as:
- **Cloud storage** for your code
- **Version control** - tracks all changes over time
- **Distribution platform** - users can download releases
- **Backup system** - your code is safely stored in the cloud

---

## Important Clarifications

### Does GitHub Auto-Update User Installations?

**NO** - GitHub does NOT automatically push updates to users.

Here's how it works:

1. **You upload a new version** → Create a new release on GitHub
2. **Users must manually download** → They download and install the new .exe
3. **No automatic updates** → Users won't get updates unless they check GitHub

### Auto-Updates (Future Feature)

To get automatic updates, you would need to implement an **auto-updater** in your app (v1.1 feature):
- App checks GitHub for new versions on startup
- Prompts user: "New version available! Download now?"
- Downloads and installs automatically (with user permission)

For v1.0, users will need to manually check for updates.

---

## What to Upload to GitHub

### ⚠️ CRITICAL: What NOT to Upload

**NEVER upload these to GitHub (even private repo):**

- ✗ API credentials (`credentials.json`, `token_gmail.json`)
- ✗ OAuth tokens (anything in `Config/API/`)
- ✗ Personal data (ward info, Excel files with real data)
- ✗ User-specific configuration files
- ✗ Backup files with real user data

**Why?** Even private repos can be accidentally made public, and you'd expose sensitive HIPAA-protected data.

### ✓ What TO Upload

**Source Code:**
- ✓ All `.py` files
- ✓ `guardianship_app.py`
- ✓ `Scripts/` folder (all Python scripts)
- ✓ `Automation/` folder

**Templates (CLEARED of personal info):**
- ✓ `Templates/Mileage_Reimbursement_Form.xlsx` (blank)
- ✓ `Templates/Court_Visitor_Payment_Invoice.docx` (blank)
- ✓ `Templates/Court Visitor Report fillable new.docx` (blank)

**Documentation:**
- ✓ `Documentation/` folder (all guides)
- ✓ `EULA.txt`
- ✓ `CHANGELOG.md`
- ✓ `README.md` (create this - see below)

**Built Executable (for releases):**
- ✓ `CourtVisitorApp.exe` (upload as a Release, not in code)

---

## Step-by-Step Setup

### Step 1: Initialize Git in Your Project

Open PowerShell in your project folder:

```powershell
cd C:\GoogleSync\GuardianShip_App

# Initialize git (if not already done)
git init

# Set your identity
git config user.name "Your Name"
git config user.email "support@guardianshipeasy.com"
```

### Step 2: Create .gitignore File

This tells Git to IGNORE sensitive files:

```powershell
# Create .gitignore file (see below for contents)
```

I'll create this file for you in a moment.

### Step 3: Add Files to Git

```powershell
# Add all files (except those in .gitignore)
git add .

# Create first commit
git commit -m "Initial commit: Court Visitor App v1.0"
```

### Step 4: Connect to GitHub

```powershell
# Add GitHub as remote origin (CHANGE URL TO YOURS)
git remote add origin https://github.com/mayehres-ops/Court_Visitor_App.git

# Set main branch
git branch -M main

# Push to GitHub
git push -u origin main
```

**Note:** You'll be prompted for GitHub credentials. Use a **Personal Access Token** (not password).

---

## Creating a Personal Access Token (PAT)

GitHub requires a token for authentication:

1. **Go to GitHub Settings**:
   - Click your profile picture → Settings
   - Scroll down → Developer settings
   - Personal access tokens → Tokens (classic)

2. **Generate New Token**:
   - Click "Generate new token (classic)"
   - Note: "Court Visitor App Upload"
   - Expiration: 90 days (or custom)
   - Scopes: Check `repo` (full control of private repositories)

3. **Copy Token**:
   - Copy the token (starts with `ghp_...`)
   - **SAVE IT** - you won't see it again!

4. **Use Token as Password**:
   - When `git push` asks for password, paste the token

---

## Distributing the App

### Option 1: GitHub Releases (RECOMMENDED)

This is the professional way to distribute software:

#### After Building Your .exe:

1. **Go to GitHub Repository**:
   - Navigate to `https://github.com/mayehres-ops/Court_Visitor_App`

2. **Click "Releases"** (right sidebar):
   - Click "Create a new release"

3. **Create Release**:
   - **Tag version**: `v1.0.0`
   - **Release title**: `Court Visitor App v1.0.0`
   - **Description**: Copy from CHANGELOG.md
   - **Attach files**: Upload `CourtVisitorApp_v1.0.zip` (the distribution package)

4. **Publish Release**:
   - Click "Publish release"
   - Users can now download from Releases page!

#### Distribution URL

Share this URL with authorized users:
```
https://github.com/mayehres-ops/Court_Visitor_App/releases/latest
```

They'll see a download button for `CourtVisitorApp_v1.0.zip`.

### Option 2: Direct File Upload (NOT RECOMMENDED)

You could upload the .exe directly to the repository, but this is NOT recommended because:
- Large binary files slow down Git
- Every version stays in history (bloats repo size)
- Not the standard way to distribute software

**Use Releases instead.**

---

## Updating the App (Pushing v1.1, v1.2, etc.)

### When You Make Changes:

1. **Update version number** in `guardianship_app.py`:
   ```python
   __version__ = "1.1.0"
   ```

2. **Update CHANGELOG.md** with new features

3. **Commit changes**:
   ```powershell
   git add .
   git commit -m "Release v1.1.0: Added auto-updater and custom icon"
   git push
   ```

4. **Build new .exe** with PyInstaller

5. **Create new GitHub Release**:
   - Tag: `v1.1.0`
   - Upload new `CourtVisitorApp_v1.1.zip`

### Users Get Updates By:

1. **Checking Releases page** on GitHub
2. **Downloading new version** manually
3. **Installing over old version** (or deleting old folder first)

**In v1.1+**, you can add an auto-updater that:
- Checks GitHub for new releases on app startup
- Shows "Update available" message
- Downloads and installs automatically

---

## Private vs Public Repository

### Current Setup: PRIVATE

Your repo is private, which means:
- ✓ Only you and invited collaborators can see it
- ✓ Code is not public
- ✓ Releases are only visible to authorized users

### Inviting Users to Private Repo

If you want specific users to download from your private repo:

1. **Go to Settings** → Collaborators
2. **Click "Add people"**
3. **Enter GitHub username or email**
4. They'll receive an invitation to access the repo

### Making it Public (NOT RECOMMENDED)

If you made it public:
- ✗ Anyone can see your code
- ✗ Anyone can download the app
- ✗ Competitors could copy your software

**Keep it private** unless you want to open-source it.

---

## Understanding GitHub for Your Use Case

### Your Distribution Model

You have a **closed, authorized user base**:
- Travis County Court Visitors only
- Must be added as Test Users in Google Cloud Console
- Should accept EULA before using

### Best Approach

1. **Keep GitHub repo PRIVATE**
2. **Invite authorized users** as collaborators
3. **Share Release download link** via email
4. **Users download manually** from Releases page

### Alternative: Direct Distribution

Instead of GitHub, you could:
- **Email the ZIP file** directly to users
- **Use Google Drive** shared folder (authorized users only)
- **Use Dropbox** or OneDrive with access control

GitHub is more professional, but direct distribution might be simpler for a small user base.

---

## GitHub Repository Structure

After pushing to GitHub, your repo will look like:

```
Court_Visitor_App/
├── .gitignore                          ← Ignores sensitive files
├── README.md                           ← Project description (for GitHub)
├── CHANGELOG.md                        ← Version history
├── EULA.txt                            ← License agreement
├── guardianship_app.py                 ← Main app
├── Scripts/                            ← All Python modules
│   ├── backup_manager.py
│   ├── cv_info_manager.py
│   ├── desktop_shortcut.py
│   └── ...
├── Automation/                         ← Automation scripts
│   ├── Mileage Reimbursement Script/
│   ├── CV Payment Form Script/
│   └── ...
├── Templates/                          ← Blank templates
│   ├── Mileage_Reimbursement_Form.xlsx
│   ├── Court_Visitor_Payment_Invoice.docx
│   └── Court Visitor Report fillable new.docx
└── Documentation/                      ← User guides
    ├── INSTALLATION_GUIDE.md
    ├── PYINSTALLER_BUILD_GUIDE.md
    └── ...
```

**NOT in GitHub:**
- `App Data/` (user data)
- `Config/API/` (API tokens)
- `New Files/` (case PDFs)
- `*.pyc`, `__pycache__/` (Python bytecode)

---

## Security Best Practices

### 1. Use .gitignore (CRITICAL)

Prevents sensitive files from being uploaded.

### 2. Never Commit Secrets

If you accidentally commit API keys:
1. **Revoke the keys immediately** in Google Cloud Console
2. **Generate new keys**
3. **Remove from Git history** (complicated - ask for help if needed)

### 3. Review Before Pushing

Always check what you're uploading:
```powershell
# See what files will be uploaded
git status

# See what changes will be uploaded
git diff
```

### 4. Use Private Repo

Keep the repo private until you're ready to open-source (if ever).

---

## Quick Command Reference

### First Time Setup

```powershell
cd C:\GoogleSync\GuardianShip_App
git init
git config user.name "Your Name"
git config user.email "support@guardianshipeasy.com"
git add .
git commit -m "Initial commit: Court Visitor App v1.0"
git remote add origin https://github.com/mayehres-ops/Court_Visitor_App.git
git branch -M main
git push -u origin main
```

### Pushing Updates

```powershell
cd C:\GoogleSync\GuardianShip_App
git add .
git commit -m "Description of changes"
git push
```

### Checking Status

```powershell
git status          # See what files have changed
git log             # See commit history
git diff            # See what changed in files
```

---

## Next Steps

1. **I'll create .gitignore file** to protect sensitive data
2. **I'll create README.md** for GitHub repository page
3. **You push code to GitHub** using commands above
4. **Build the .exe** with PyInstaller (default icon)
5. **Create first Release** on GitHub (v1.0.0)
6. **Share download link** with authorized users

---

## Questions Answered

### "It will also push new updates to anyone who installs from there?"

**No** - GitHub does NOT automatically update user installations. Users must:
1. Check the Releases page
2. Download the new version manually
3. Install it (replace old version)

To get auto-updates, you'd need to implement an auto-updater feature (planned for v1.1).

### "How I put my application there for download?"

1. Push code to GitHub (git push)
2. Build the .exe with PyInstaller
3. Create a Release on GitHub
4. Upload the .exe (or ZIP package) to the Release
5. Share the Release URL with users

### "Use default icon, we can change later"

Perfect! We'll build with the default Python icon for v1.0, and you can add a custom icon in v1.1.

---

## Support Resources

- **GitHub Docs**: https://docs.github.com
- **Git Tutorial**: https://git-scm.com/book/en/v2
- **GitHub Desktop** (GUI alternative): https://desktop.github.com/

---

**Ready to push to GitHub?**

Let me create the `.gitignore` and `README.md` files first, then I'll give you the exact commands to run!

---

**Document Version:** 1.0
**Last Updated:** November 6, 2024
**Status:** Ready to Push to GitHub
