# PyInstaller Build Guide

Complete guide for building the Court Visitor App executable.

---

## Prerequisites

### 1. Install PyInstaller

```bash
pip install pyinstaller
```

### 2. Install Required Dependencies

Make sure all dependencies are installed:

```bash
pip install -r requirements.txt
```

If `requirements.txt` doesn't exist, install these manually:
```bash
pip install openpyxl python-docx google-api-python-client google-auth-httplib2 google-auth-oauthlib pillow winshell pywin32
```

---

## App Icon (.ico file)

### Option 1: Create Custom Icon

If you want a custom "cute" icon for the app:

1. **Design the icon** - Use a graphic design tool:
   - Adobe Illustrator
   - Figma
   - Canva
   - Or hire a designer on Fiverr ($5-20)

2. **Convert to .ico format**:
   - Use online converter: https://convertio.co/png-ico/
   - Or use tool like IcoFX
   - **Requirements**:
     - Multiple sizes: 16x16, 32x32, 48x48, 256x256
     - Format: .ico

3. **Save as**: `C:\GoogleSync\GuardianShip_App\icon.ico`

### Option 2: Use Default Python Icon

PyInstaller will use the default Python icon if no custom icon is provided. This is acceptable for v1.0.

### Option 3: Create Simple Icon with Python

```python
from PIL import Image, ImageDraw, ImageFont

# Create 256x256 icon
size = 256
img = Image.new('RGB', (size, size), color='#667eea')  # Purple background

# Add text
draw = ImageDraw.Draw(img)
# Add "CV" or scales of justice symbol

# Save as .ico (requires pillow)
img.save('icon.ico', format='ICO', sizes=[(256, 256)])
```

---

## Build Configuration

### Build Script: `build.spec`

Create a PyInstaller spec file for better control:

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Data files to include
datas = [
    ('Templates', 'Templates'),
    ('Documentation', 'Documentation'),
    ('EULA.txt', '.'),
    ('Config/API', 'Config/API'),
]

# Hidden imports (sometimes PyInstaller misses these)
hiddenimports = [
    'openpyxl.cell._writer',
    'google.oauth2.credentials',
    'googleapiclient.discovery',
    'win32com.client',
    'winshell',
]

a = Analysis(
    ['guardianship_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CourtVisitorApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window (GUI app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Add your icon here (or remove this line for default)
)
```

---

## Build Process

### Simple One-File Build

```bash
cd C:\GoogleSync\GuardianShip_App

# Basic build (no icon)
pyinstaller --name "CourtVisitorApp" --onefile --windowed guardianship_app.py

# With icon (if icon.ico exists)
pyinstaller --name "CourtVisitorApp" --onefile --windowed --icon=icon.ico guardianship_app.py

# With data files
pyinstaller --name "CourtVisitorApp" --onefile --windowed --icon=icon.ico ^
  --add-data "Templates;Templates" ^
  --add-data "Documentation;Documentation" ^
  --add-data "EULA.txt;." ^
  guardianship_app.py
```

### Build with Spec File

If you created `build.spec`:

```bash
pyinstaller build.spec
```

---

## Build Output

After building, you'll find:

```
GuardianShip_App/
├── dist/
│   └── CourtVisitorApp.exe    ← The built executable
├── build/                      ← Temporary build files (can delete)
└── CourtVisitorApp.spec        ← Build configuration (keep)
```

---

## Testing the Build

### 1. Test on Your Machine

```bash
cd dist
.\CourtVisitorApp.exe
```

**Check:**
- [ ] App launches without errors
- [ ] EULA shows on first launch (if fresh install)
- [ ] CV Settings dialog works
- [ ] All 14 automation steps run correctly
- [ ] Forms generate with CV info filled
- [ ] Google API authentication works
- [ ] Backup button works
- [ ] Desktop shortcut is created
- [ ] Help menu works (About, License)

### 2. Test on Fresh Windows Installation

Use a Virtual Machine (VirtualBox or VMware) with fresh Windows 10/11:

- [ ] Copy `CourtVisitorApp.exe` to VM
- [ ] Run and verify Windows SmartScreen warning appears
- [ ] Click "More info" → "Run anyway"
- [ ] Verify all functionality works

---

## Creating Distribution Package

### Structure for Distribution

```
CourtVisitorApp_v1.0/
├── CourtVisitorApp.exe
├── Templates/
│   ├── Mileage_Reimbursement_Form.xlsx
│   ├── Court_Visitor_Payment_Invoice.docx
│   └── Court Visitor Report fillable new.docx
├── Documentation/
│   ├── INSTALLATION_GUIDE.md
│   ├── LEGAL_PROTECTIONS_SUMMARY.md
│   └── (other docs)
├── EULA.txt
└── README.txt
```

### Create Distribution ZIP

```bash
# PowerShell
Compress-Archive -Path "CourtVisitorApp_v1.0" -DestinationPath "CourtVisitorApp_v1.0.zip"
```

---

## Common Build Issues

### Issue 1: Missing Modules

**Error:** `ModuleNotFoundError: No module named 'xyz'`

**Fix:** Add to `hiddenimports` in spec file or use:
```bash
pyinstaller --hidden-import xyz guardianship_app.py
```

### Issue 2: Missing Data Files

**Error:** Template files not found

**Fix:** Use `--add-data` flag:
```bash
--add-data "Templates;Templates"
```

### Issue 3: Large File Size

**Problem:** .exe is 100MB+

**Solutions:**
- Use `--onefile` (already doing)
- Use UPX compression: `--upx-dir=C:\path\to\upx`
- Exclude unnecessary packages: `--exclude-module matplotlib`

### Issue 4: Slow Startup

**Problem:** App takes 5-10 seconds to start

**Explanation:** Normal for PyInstaller one-file builds (unpacking temp files)

**Solutions for future:**
- Use `--onedir` instead (faster startup, but multiple files)
- Add splash screen: `--splash splash.png`

---

## Code Signing (Optional - v1.1+)

### Why Sign Code?

- Removes Windows SmartScreen warning
- Users trust "Verified Publisher"
- Less likely to be flagged by antivirus

### How to Get Certificate

1. **Purchase Certificate** ($200-400/year)
   - DigiCert
   - Sectigo
   - GlobalSign

2. **Verify Your Identity**
   - Business verification
   - Takes 3-7 days

3. **Sign the Executable**
   ```bash
   signtool sign /f certificate.pfx /p password /tr http://timestamp.digicert.com /td sha256 /fd sha256 CourtVisitorApp.exe
   ```

### Is It Worth It?

**For v1.0:** NO
- Installation Guide explains warnings
- Limited distribution (Travis County CVs)
- Expensive for small user base

**For v1.1+:** CONSIDER
- If expanding to other counties
- If users report antivirus issues
- If you want professional appearance

---

## Build Checklist

Before building:
- [ ] All Python scripts tested and working
- [ ] Version number updated in `guardianship_app.py` (`__version__ = "1.0.0"`)
- [ ] CHANGELOG.md updated
- [ ] Templates cleared of personal info
- [ ] icon.ico created (optional)
- [ ] All dependencies installed
- [ ] PyInstaller installed

Build steps:
- [ ] Run PyInstaller command
- [ ] Test .exe on your machine
- [ ] Test on fresh Windows VM
- [ ] Test all 14 automation steps
- [ ] Test Google API authentication
- [ ] Verify desktop shortcut creation
- [ ] Verify backup functionality

After build:
- [ ] Create distribution folder structure
- [ ] Write README.txt for users
- [ ] Create distribution ZIP
- [ ] Test ZIP extraction and installation
- [ ] Add test users to Google Cloud Console

---

## Next Steps After Building

1. **Test Distribution Package**
   - Extract ZIP on different machine
   - Run `CourtVisitorApp.exe`
   - Verify all features work

2. **Prepare User Onboarding**
   - Send registration emails (see `NEW_USER_EMAIL_TEMPLATE.md`)
   - Add users as Test Users in Google Cloud Console
   - Prepare support email (support@guardianshipeasy.com)

3. **Create Support Resources**
   - FAQ document
   - Troubleshooting guide
   - Video walkthrough (optional)

4. **Monitor Initial Rollout**
   - Gather user feedback
   - Fix any critical bugs
   - Plan v1.1 features

---

## Questions to Answer Before Building

1. **Do you have an icon file (.ico)?**
   - If NO: Build without icon (use default)
   - If YES: Place at `C:\GoogleSync\GuardianShip_App\icon.ico`

2. **One-file or one-folder build?**
   - **Recommended:** One-file (single .exe, easier distribution)
   - Alternative: One-folder (faster startup)

3. **Do you want a splash screen?**
   - Shown during startup loading
   - Requires splash.png image
   - **Recommendation:** Skip for v1.0

---

**Ready to build?**

If all pre-build items are checked in `FINAL_CHECKLIST.md`, run:

```bash
cd C:\GoogleSync\GuardianShip_App
pyinstaller --name "CourtVisitorApp" --onefile --windowed --add-data "Templates;Templates" --add-data "Documentation;Documentation" --add-data "EULA.txt;." guardianship_app.py
```

Then test the executable in `dist\CourtVisitorApp.exe`!

---

**Document Version:** 1.0
**Last Updated:** November 6, 2024
**Status:** Ready to Build
