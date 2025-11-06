@echo off
setlocal
pushd "%~dp0"

REM Ensure the script exists next to this .bat
if not exist "build_court_visitor_summary.py" (
  echo ERROR: build_court_visitor_summary.py not found in "%~dp0"
  pause
  exit /b 1
)

REM Print as you go: each doc is printed right after it's created.
REM Requires Word and:  py -3.13 -m pip install pywin32
py -3.13 "%~dp0build_court_visitor_summary.py" --print --open

if errorlevel 1 (
  echo.
  echo If printing failed:
  echo  - Make sure Microsoft Word is installed
  echo  - Set a default printer in Windows
  echo  - Install pywin32:  py -3.13 -m pip install pywin32
)

echo.
pause
popd
