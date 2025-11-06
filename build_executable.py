#!/usr/bin/env python3
"""
Build script for Court Visitor App
Creates a standalone executable using PyInstaller
"""

import PyInstaller.__main__
import os
from pathlib import Path

# Paths
BASE_DIR = Path(__file__).parent
MAIN_SCRIPT = BASE_DIR / "guardianship_app.py"
ICON_PATH = BASE_DIR / "App Data" / "icon.ico"  # Add your icon here

# PyInstaller arguments
PyInstaller.__main__.run([
    str(MAIN_SCRIPT),
    '--name=CourtVisitorApp',
    '--onefile',  # Single executable
    '--windowed',  # No console window
    f'--icon={ICON_PATH}' if ICON_PATH.exists() else '',
    '--add-data=Config;Config',  # Include config files
    '--add-data=Automation;Automation',  # Include automation scripts
    '--add-data=Scripts;Scripts',  # Include utility scripts
    '--hidden-import=win32com',
    '--hidden-import=win32com.client',
    '--hidden-import=pywintypes',
    '--hidden-import=openpyxl',
    '--hidden-import=pytesseract',
    '--hidden-import=pdf2image',
    '--hidden-import=pdfplumber',
    '--hidden-import=PIL',
    '--hidden-import=google.oauth2',
    '--hidden-import=googleapiclient',
    '--collect-all=tkinter',
    '--noconfirm',  # Overwrite without asking
])

print("\n" + "="*60)
print("Build complete!")
print(f"Executable location: {BASE_DIR / 'dist' / 'CourtVisitorApp.exe'}")
print("="*60)
