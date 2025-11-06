@echo off
REM Court Visitor App - Comprehensive Installation Script
REM Checks and installs ALL dependencies including external tools

setlocal enabledelayedexpansion

echo.
echo ========================================================
echo          Court Visitor App - Installation
echo ========================================================
echo.
echo This installer will check and install:
echo   - Python (if missing)
echo   - Python packages (openpyxl, pandas, etc.)
echo   - Tesseract OCR (for PDF text extraction)
echo   - Poppler (for PDF to image conversion)
echo   - Microsoft Word (check only - manual install needed)
echo.
pause
echo.

REM ======================
REM 1. CHECK PYTHON
REM ======================
echo [1/5] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] Python is not installed!
    echo.
    echo Please install Python 3.10 or higher from:
    echo https://www.python.org/downloads/windows/
    echo.
    echo IMPORTANT: During installation, check the box:
    echo   [X] Add Python to PATH
    echo.
    echo After installing Python, run this script again.
    pause
    exit /b 1
) else (
    for /f "tokens=2" %%i in ('python --version') do set PYTHON_VERSION=%%i
    echo [OK] Python !PYTHON_VERSION! found
)
echo.

REM ======================
REM 2. CHECK PYTHON VERSION
REM ======================
echo [2/5] Checking Python version...
python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" 2>nul
if errorlevel 1 (
    echo [WARNING] Python version is too old
    echo           You need Python 3.10 or higher
    echo           Current version: !PYTHON_VERSION!
    echo.
    echo Please update Python from:
    echo https://www.python.org/downloads/windows/
    pause
    exit /b 1
) else (
    echo [OK] Python version is compatible
)
echo.

REM ======================
REM 3. INSTALL PYTHON PACKAGES
REM ======================
echo [3/5] Installing Python packages...
echo      This may take 2-5 minutes...
echo.

python -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo [WARNING] Could not upgrade pip
)

python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install Python packages
    echo.
    echo Try running manually:
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
) else (
    echo [OK] Python packages installed successfully
)
echo.

REM ======================
REM 4. CHECK TESSERACT OCR
REM ======================
echo [4/5] Checking Tesseract OCR...

REM Check if tesseract is in PATH
tesseract --version >nul 2>&1
if errorlevel 1 (
    REM Check common installation locations
    set TESSERACT_FOUND=0

    if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
        set TESSERACT_FOUND=1
        set TESSERACT_PATH=C:\Program Files\Tesseract-OCR
    )
    if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" (
        set TESSERACT_FOUND=1
        set TESSERACT_PATH=C:\Program Files (x86)\Tesseract-OCR
    )

    if !TESSERACT_FOUND! == 1 (
        echo [WARNING] Tesseract found at: !TESSERACT_PATH!
        echo           But not in PATH. Adding to PATH...
        setx PATH "!TESSERACT_PATH!;%PATH%" >nul 2>&1
        echo [OK] Tesseract added to PATH (restart required)
    ) else (
        echo [MISSING] Tesseract OCR is not installed
        echo.
        echo Tesseract is REQUIRED for Step 1 (PDF text extraction)
        echo.
        echo Would you like to download Tesseract installer? (Y/N)
        set /p INSTALL_TESSERACT="> "

        if /i "!INSTALL_TESSERACT!" == "Y" (
            echo.
            echo Opening Tesseract download page...
            start https://github.com/UB-Mannheim/tesseract/wiki
            echo.
            echo INSTRUCTIONS:
            echo 1. Download "tesseract-ocr-w64-setup-5.x.x.exe"
            echo 2. Run the installer
            echo 3. Use default installation path
            echo 4. Run this installer again after Tesseract is installed
            echo.
            pause
            exit /b 1
        ) else (
            echo [SKIPPED] Tesseract installation skipped
            echo           You can install it later if needed
        )
    )
) else (
    echo [OK] Tesseract OCR is installed
)
echo.

REM ======================
REM 5. CHECK POPPLER
REM ======================
echo [5/5] Checking Poppler (PDF tools)...

REM Check if pdftoppm is in PATH
pdftoppm -v >nul 2>&1
if errorlevel 1 (
    REM Check common locations
    set POPPLER_FOUND=0

    if exist "C:\Program Files\poppler\Library\bin\pdftoppm.exe" (
        set POPPLER_FOUND=1
        set POPPLER_PATH=C:\Program Files\poppler\Library\bin
    )
    if exist "C:\poppler\Library\bin\pdftoppm.exe" (
        set POPPLER_FOUND=1
        set POPPLER_PATH=C:\poppler\Library\bin
    )

    if !POPPLER_FOUND! == 1 (
        echo [WARNING] Poppler found at: !POPPLER_PATH!
        echo           But not in PATH. Adding to PATH...
        setx PATH "!POPPLER_PATH!;%PATH%" >nul 2>&1
        echo [OK] Poppler added to PATH (restart required)
    ) else (
        echo [MISSING] Poppler is not installed
        echo.
        echo Poppler is REQUIRED for Step 1 (PDF to image conversion)
        echo.
        echo Would you like instructions to install Poppler? (Y/N)
        set /p INSTALL_POPPLER="> "

        if /i "!INSTALL_POPPLER!" == "Y" (
            echo.
            echo Opening Poppler download page...
            start https://github.com/oschwartz10612/poppler-windows/releases
            echo.
            echo INSTRUCTIONS:
            echo 1. Download "Release-XX.XX.X-X.zip"
            echo 2. Extract to C:\poppler\
            echo 3. Add C:\poppler\Library\bin to PATH
            echo 4. Run this installer again
            echo.
            echo OR use the automated installer:
            echo   https://github.com/oschwartz10612/poppler-windows
            echo.
            pause
            exit /b 1
        ) else (
            echo [SKIPPED] Poppler installation skipped
            echo           You can install it later if needed
        )
    )
) else (
    echo [OK] Poppler is installed
)
echo.

REM ======================
REM 6. CHECK MICROSOFT WORD
REM ======================
echo.
echo ========================================================
echo           Additional Requirements Check
echo ========================================================
echo.
echo Checking Microsoft Word...

python -c "import win32com.client; word = win32com.client.Dispatch('Word.Application'); word.Quit()" 2>nul
if errorlevel 1 (
    echo [WARNING] Microsoft Word not found or not accessible
    echo.
    echo Microsoft Word is REQUIRED for:
    echo   - Step 8: Generate CVR documents
    echo   - Step 13: Generate Payment Forms
    echo   - Step 14: Generate Mileage Log
    echo.
    echo Please install Microsoft Office if you haven't already.
) else (
    echo [OK] Microsoft Word is installed and accessible
)
echo.

REM ======================
REM FINAL SUMMARY
REM ======================
echo.
echo ========================================================
echo              Installation Summary
echo ========================================================
echo.

REM Create a results file for reference
set RESULTS_FILE=installation_results.txt
echo Installation Results - %date% %time% > %RESULTS_FILE%
echo. >> %RESULTS_FILE%

python --version >nul 2>&1
if errorlevel 1 (
    echo [X] Python: NOT INSTALLED
    echo [X] Python: NOT INSTALLED >> %RESULTS_FILE%
    set INSTALL_SUCCESS=0
) else (
    echo [OK] Python: Installed
    echo [OK] Python: Installed >> %RESULTS_FILE%
)

python -m pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo [X] Python Packages: FAILED
    echo [X] Python Packages: FAILED >> %RESULTS_FILE%
    set INSTALL_SUCCESS=0
) else (
    echo [OK] Python Packages: Installed
    echo [OK] Python Packages: Installed >> %RESULTS_FILE%
)

tesseract --version >nul 2>&1
if errorlevel 1 (
    echo [!] Tesseract OCR: MISSING (required for Step 1)
    echo [!] Tesseract OCR: MISSING >> %RESULTS_FILE%
) else (
    echo [OK] Tesseract OCR: Installed
    echo [OK] Tesseract OCR: Installed >> %RESULTS_FILE%
)

pdftoppm -v >nul 2>&1
if errorlevel 1 (
    echo [!] Poppler: MISSING (required for Step 1)
    echo [!] Poppler: MISSING >> %RESULTS_FILE%
) else (
    echo [OK] Poppler: Installed
    echo [OK] Poppler: Installed >> %RESULTS_FILE%
)

python -c "import win32com.client; word = win32com.client.Dispatch('Word.Application'); word.Quit()" 2>nul
if errorlevel 1 (
    echo [!] Microsoft Word: NOT FOUND (required for Steps 8, 13, 14)
    echo [!] Microsoft Word: NOT FOUND >> %RESULTS_FILE%
) else (
    echo [OK] Microsoft Word: Installed
    echo [OK] Microsoft Word: Installed >> %RESULTS_FILE%
)

echo.
echo Results saved to: %RESULTS_FILE%
echo.

REM Check if critical dependencies are missing
tesseract --version >nul 2>&1
if errorlevel 1 (
    echo ========================================================
    echo                   ACTION REQUIRED
    echo ========================================================
    echo.
    echo Some dependencies are missing. The app will work but
    echo some features may not function until you install:
    echo.
    echo   - Tesseract OCR: Required for Step 1
    echo   - Poppler: Required for Step 1
    echo.
    echo Run this installer again after installing them.
    echo ========================================================
    echo.
) else (
    echo ========================================================
    echo          Installation Complete - Ready to Use!
    echo ========================================================
    echo.
    echo Next steps:
    echo   1. Read: INSTALLATION_GUIDE.md
    echo   2. Setup Google API credentials
    echo   3. Place credentials in Config\API\
    echo   4. Add your Excel database to App Data\
    echo   5. Launch: Double-click "Launch Court Visitor App.vbs"
    echo.
    echo ========================================================
)

echo.
pause
