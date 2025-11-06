# Court Visitor App - Distribution & Update Guide

## Overview

This guide explains how to package, distribute, and auto-update the Court Visitor application for end users.

---

## Table of Contents

1. [Distribution Options](#distribution-options)
2. [Recommended Approach](#recommended-approach)
3. [Setup Instructions](#setup-instructions)
4. [Auto-Update System](#auto-update-system)
5. [Cost Breakdown](#cost-breakdown)
6. [Deployment Checklist](#deployment-checklist)

---

## Distribution Options

### Option 1: GitHub Releases (FREE - Recommended for initial release)

**Pros:**
- FREE hosting
- Fast CDN (Content Delivery Network)
- Version control built-in
- Easy rollback to previous versions
- Can be private or public

**Cons:**
- Requires GitHub account
- 2GB file size limit per release
- Need to manage releases manually

### Option 2: Your Website

**Pros:**
- Full control
- Your branding
- Can track downloads
- No third-party dependencies

**Cons:**
- Web hosting costs (~$5-20/month)
- Need to implement download tracking
- Bandwidth costs for large downloads

### Option 3: Microsoft Store

**Pros:**
- Professional distribution
- Automatic Windows updates
- Built-in payment processing
- Trusted by users

**Cons:**
- $99 one-time developer fee
- App review process (2-5 days per update)
- Must follow Microsoft Store policies
- 15-30% revenue share if selling app

---

## Recommended Approach

### Phase 1: Initial Release (Months 1-3)
**Distribution:** GitHub Releases (private repo)
**Package:** Single `.exe` file (PyInstaller)
**Updates:** Auto-check with manual download
**Cost:** FREE

### Phase 2: Professional Release (Months 4-12)
**Distribution:** Your website + GitHub backup
**Package:** Signed installer (Inno Setup + code signing certificate)
**Updates:** Auto-download and install
**Cost:** ~$200/year (code signing only)

### Phase 3: Enterprise Release (Year 2+)
**Distribution:** Microsoft Store (optional)
**Package:** MSIX installer
**Updates:** Automatic via Windows Store
**Cost:** $99 one-time + $200/year for cert

---

## Setup Instructions

### Step 1: Prepare for Building

1. Install PyInstaller:
   ```bash
   pip install pyinstaller packaging requests
   ```

2. Create an application icon (optional):
   - Size: 256x256 pixels
   - Format: `.ico` file
   - Place at: `C:\GoogleSync\GuardianShip_App\App Data\icon.ico`

### Step 2: Build the Executable

Run the build script:
```bash
cd C:\GoogleSync\GuardianShip_App
python build_executable.py
```

This creates:
- `dist/CourtVisitorApp.exe` (standalone executable)
- `build/` folder (can be deleted)

**Expected size:** 100-200MB (includes Python + all dependencies)

### Step 3: Test the Executable

1. Copy `CourtVisitorApp.exe` to a test location
2. Copy required folders:
   - `Config/` (API credentials)
   - `Automation/` (all automation scripts)
   - `Scripts/` (utility scripts)
   - `App Data/Templates/` (document templates)
3. Run the `.exe` and test all 14 steps

### Step 4: Create GitHub Repository

1. Create a private GitHub repository:
   - Name: `court-visitor-app`
   - Private: ✅

2. Create a new release:
   - Go to: Releases → Draft a new release
   - Tag version: `v1.0.0`
   - Title: `Court Visitor App v1.0.0`
   - Description: Release notes
   - Upload: `CourtVisitorApp.exe`
   - Publish release

3. Update `guardianship_app.py`:
   ```python
   github_repo="your-username/court-visitor-app"
   ```

### Step 5: Setup Auto-Updates

The app now automatically:
1. Checks GitHub for updates on startup
2. Shows update dialog if new version available
3. Opens browser to download new version

**To release an update:**
1. Increment version in `guardianship_app.py`:
   ```python
   __version__ = "1.0.1"
   ```
2. Rebuild executable: `python build_executable.py`
3. Create new GitHub release with new `.exe`

---

## Auto-Update System

### How It Works

```
App Startup
    ↓
Check GitHub API (background thread)
    ↓
Compare versions
    ↓
If update available → Show dialog
    ↓
User clicks "Yes" → Open download page
```

### Update Flow

1. **App starts** → Waits 500ms for UI to load
2. **Background thread** → Checks GitHub API
3. **If new version** → Shows dialog with release notes
4. **User confirms** → Opens browser to download
5. **User downloads** → Replaces old `.exe` with new one
6. **Restart app** → Now running latest version

### Manual Update Check

You can add a "Check for Updates" button in the GUI:
```python
def check_updates_manually(self):
    updater = AutoUpdater(__version__, "your-username/court-visitor-app")
    updater.prompt_update(parent=self.root)
```

---

## Cost Breakdown

### FREE Option (Phase 1)
| Item | Cost |
|------|------|
| GitHub Hosting | FREE |
| PyInstaller | FREE |
| SSL Certificate (GitHub provides) | FREE |
| **Total** | **$0/year** |

### Professional Option (Phase 2)
| Item | Cost |
|------|------|
| Code Signing Certificate | $200/year |
| Web Hosting (optional) | $60/year |
| Domain Name (optional) | $15/year |
| **Total** | **$275/year** |

### Enterprise Option (Phase 3)
| Item | Cost |
|------|------|
| Microsoft Developer Account | $99 one-time |
| Code Signing Certificate | $200/year |
| **Total** | **$299 first year, $200/year after** |

---

## Code Signing Certificate

**Why you need it:**
- Windows Defender won't flag your app
- Users won't see "Unknown Publisher" warning
- Professional appearance

**Where to get it:**
- **DigiCert** (~$474/year) - Most trusted
- **Sectigo** (~$200/year) - Good value
- **SSL.com** (~$249/year) - Mid-range

**Process:**
1. Verify your business identity (1-3 days)
2. Install certificate on your PC
3. Sign your `.exe` with SignTool

---

## Deployment Checklist

### Before First Release

- [ ] Update version number in `guardianship_app.py`
- [ ] Create application icon
- [ ] Build executable with PyInstaller
- [ ] Test on clean Windows machine
- [ ] Create GitHub repository
- [ ] Update GitHub repo URL in `guardianship_app.py`
- [ ] Create first GitHub release
- [ ] Test auto-update checker
- [ ] Write user documentation
- [ ] Create installation guide

### For Each Update

- [ ] Increment version number
- [ ] Test all functionality
- [ ] Rebuild executable
- [ ] Write release notes
- [ ] Create GitHub release
- [ ] Upload new `.exe`
- [ ] Tag release with version
- [ ] Test update notification

### Distribution

- [ ] Send download link to users
- [ ] Provide installation instructions
- [ ] Document Google API setup process
- [ ] Create video tutorial (optional)
- [ ] Setup support email/system

---

## File Structure for Distribution

```
CourtVisitorApp/
├── CourtVisitorApp.exe          (Main application)
├── Config/
│   ├── API/                     (API credentials - user provides)
│   └── cvr_google_form_mapping.json
├── Automation/                  (All automation scripts)
├── Scripts/                     (Utility scripts)
├── App Data/
│   ├── Templates/               (Word document templates)
│   └── ward_guardian_info.xlsx  (User's data)
├── New Clients/                 (Created by app)
├── New Files/                   (Created by app)
└── README_FIRST.md              (Setup instructions)
```

---

## Alternative: Full Installer

For a more professional installation experience, create an installer using **Inno Setup**:

**Benefits:**
- Creates Start Menu shortcut
- Adds Uninstall option
- Can set file associations
- Professional appearance

**Inno Setup script example:**
```ini
[Setup]
AppName=Court Visitor App
AppVersion=1.0.0
DefaultDirName={autopf}\CourtVisitorApp
DefaultGroupName=Court Visitor App
OutputDir=installer
OutputBaseFilename=CourtVisitorApp_Setup_v1.0.0

[Files]
Source: "dist\CourtVisitorApp.exe"; DestDir: "{app}"
Source: "Config\*"; DestDir: "{app}\Config"; Flags: recursesubdirs
Source: "Automation\*"; DestDir: "{app}\Automation"; Flags: recursesubdirs

[Icons]
Name: "{group}\Court Visitor App"; Filename: "{app}\CourtVisitorApp.exe"
Name: "{autodesktop}\Court Visitor App"; Filename: "{app}\CourtVisitorApp.exe"
```

---

## Support & Maintenance

### User Support

1. **Documentation**: Provide comprehensive README
2. **Video tutorials**: Screen recordings of each step
3. **Support email**: Dedicated support contact
4. **FAQ**: Common issues and solutions

### Monitoring

Consider adding:
- **Error reporting**: Send crash logs to you
- **Usage analytics**: Track which features are used
- **Update metrics**: Monitor update adoption rate

### Backup Strategy

- Keep all releases on GitHub (unlimited history)
- Backup user data (Excel files, configs)
- Version control all code changes

---

## Next Steps

1. ✅ Auto-update system integrated
2. ⏳ Build first executable
3. ⏳ Create GitHub repository
4. ⏳ Test on clean Windows machine
5. ⏳ Write end-user documentation
6. ⏳ Release v1.0.0

---

## Questions?

Contact: [Your support email]
Repository: https://github.com/your-username/court-visitor-app
Documentation: See README_FIRST.md
