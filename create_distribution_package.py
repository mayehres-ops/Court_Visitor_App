#!/usr/bin/env python3
"""
Create distribution package for Court Visitor App

This script creates a ZIP file containing everything an end user needs:
- All Python scripts
- All automation scripts
- Templates
- Config folders (empty - user adds credentials)
- Launcher
- Documentation
"""

import shutil
import zipfile
from pathlib import Path
from datetime import datetime

# Version from main app
VERSION = "1.0.0"

# Base directories
SOURCE_DIR = Path(r"C:\GoogleSync\GuardianShip_App")
DIST_DIR = Path(r"C:\GoogleSync\GuardianShip_App\Distribution")
PACKAGE_NAME = f"CourtVisitorApp_v{VERSION}_{datetime.now().strftime('%Y%m%d')}"
PACKAGE_DIR = DIST_DIR / PACKAGE_NAME

print(f"Creating distribution package: {PACKAGE_NAME}")
print("="*60)

# Clean and create distribution directory
if DIST_DIR.exists():
    shutil.rmtree(DIST_DIR)
PACKAGE_DIR.mkdir(parents=True, exist_ok=True)

# Files and folders to include
INCLUDE = {
    # Main application files
    "guardianship_app.py": "guardianship_app.py",
    "auto_updater.py": "auto_updater.py",
    "setup_wizard.py": "setup_wizard.py",
    "Launch Court Visitor App.vbs": "Launch Court Visitor App.vbs",

    # OCR and main scripts
    "guardian_extractor_claudecode20251023_bestever_11pm.py": "guardian_extractor_claudecode20251023_bestever_11pm.py",
    "google_sheets_cvr_integration_fixed.py": "google_sheets_cvr_integration_fixed.py",
    "email_cvr_to_supervisor.py": "email_cvr_to_supervisor.py",

    # Documentation
    "README_FIRST.md": "README_FIRST.md",
    "END_USER_INSTALLATION_GUIDE.md": "INSTALLATION_GUIDE.md",

    # Installers
    "COMPREHENSIVE_INSTALLER.bat": "INSTALL.bat",

    # Entire folders (recursively)
    "Automation": "Automation",
    "Scripts": "Scripts",
}

# Copy files
print("\n1. Copying application files...")
for src, dst in INCLUDE.items():
    src_path = SOURCE_DIR / src
    dst_path = PACKAGE_DIR / dst

    if src_path.is_file():
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dst_path)
        print(f"   ✓ {src}")
    elif src_path.is_dir():
        shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
        print(f"   ✓ {src}/ (folder)")
    else:
        print(f"   ⚠ Skipped: {src} (not found)")

# Create empty structure folders
print("\n2. Creating folder structure...")
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

for folder in folders_to_create:
    folder_path = PACKAGE_DIR / folder
    folder_path.mkdir(parents=True, exist_ok=True)

    # Create .gitkeep to preserve empty folders in git
    (folder_path / ".gitkeep").touch()
    print(f"   ✓ {folder}/")

# Create requirements.txt
print("\n3. Creating requirements.txt...")
requirements = """# Court Visitor App - Python Dependencies

# Core dependencies
openpyxl>=3.1.2
pandas>=2.0.0
pytesseract>=0.3.10
pdf2image>=1.16.3
pdfplumber>=0.10.0
Pillow>=10.0.0

# Google API
google-auth>=2.23.0
google-auth-oauthlib>=1.1.0
google-auth-httplib2>=0.1.1
google-api-python-client>=2.100.0

# Windows COM (Word automation)
pywin32>=306

# Auto-updater
requests>=2.31.0
packaging>=23.0

# UI
tkinter  # Usually included with Python

# Note: Tesseract OCR and Poppler must be installed separately
# See INSTALLATION_GUIDE.md for details
"""

(PACKAGE_DIR / "requirements.txt").write_text(requirements, encoding='utf-8')
print("   ✓ requirements.txt")

# Create INSTALL.bat for easy setup
print("\n4. Creating installation batch file...")
install_bat = f"""@echo off
REM Court Visitor App - Installation Script v{VERSION}

echo ========================================
echo Court Visitor App Installation
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed!
    echo.
    echo Please install Python 3.10 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Install dependencies
echo Installing Python dependencies...
echo This may take a few minutes...
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ========================================
echo Installation Complete!
echo ========================================
echo.
echo Next steps:
echo 1. Read INSTALLATION_GUIDE.md for Google API setup
echo 2. Add your Google credentials to Config\\API\\
echo 3. Place Excel database in App Data\\
echo 4. Double-click "Launch Court Visitor App.vbs"
echo.
pause
"""

(PACKAGE_DIR / "INSTALL.bat").write_text(install_bat, encoding='utf-8')
print("   ✓ INSTALL.bat")

# Create README for the package
print("\n5. Creating README...")
readme = f"""# Court Visitor App v{VERSION}

## Quick Start

1. **Extract this folder** to `C:\\CourtVisitorApp\\`
   - Final location: `C:\\CourtVisitorApp\\guardianship_app.py`

2. **Run the Setup Wizard**
   - OPTION A (Recommended): Double-click `setup_wizard.py` for GUI installer
   - OPTION B: Double-click `INSTALL.bat` for command-line installer
   - OR manually run: `pip install -r requirements.txt`

3. **Setup Google API**
   - Follow instructions in `INSTALLATION_GUIDE.md`
   - Place credentials in `Config\\API\\`

4. **Launch the app**
   - Double-click `Launch Court Visitor App.vbs`
   - OR run: `python guardianship_app.py`

## What's Included

- ✅ All application files
- ✅ All automation scripts
- ✅ Folder structure
- ✅ Documentation
- ✅ Installation script

## What's NOT Included (You provide)

- ❌ Python installation (download from python.org)
- ❌ Google API credentials (follow setup guide)
- ❌ Excel database (use your data)
- ❌ Word templates (provided separately)
- ❌ Tesseract OCR (download from GitHub)
- ❌ Poppler (download from GitHub)

## Full Documentation

See `INSTALLATION_GUIDE.md` for complete setup instructions.

## Support

Email: [Your support email]
Version: {VERSION}
Release Date: {datetime.now().strftime('%Y-%m-%d')}

---

Copyright © 2024 GuardianShip Easy, LLC. All rights reserved.
"""

(PACKAGE_DIR / "README.txt").write_text(readme, encoding='utf-8')
print("   ✓ README.txt")

# Create ZIP file
print("\n6. Creating ZIP archive...")
zip_path = DIST_DIR / f"{PACKAGE_NAME}.zip"

with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for file_path in PACKAGE_DIR.rglob('*'):
        if file_path.is_file():
            arcname = file_path.relative_to(DIST_DIR)
            zipf.write(file_path, arcname)

print(f"   ✓ {zip_path.name}")

# Calculate size
zip_size_mb = zip_path.stat().st_size / (1024 * 1024)

print("\n" + "="*60)
print("✅ Distribution package created successfully!")
print("="*60)
print(f"\nPackage location: {zip_path}")
print(f"Package size: {zip_size_mb:.1f} MB")
print(f"\nUpload this ZIP file to:")
print("  - GitHub Releases")
print("  - Your website")
print("  - Google Drive / Dropbox")
print("\nUsers download and extract to C:\\CourtVisitorApp\\")
print("="*60)
