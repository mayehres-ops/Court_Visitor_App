@echo off
setlocal
pushd "%~dp0"

REM Ensure the script exists next to this .bat
if not exist "build_court_visitor_summary.py" (
  echo ERROR: build_court_visitor_summary.py not found in "%~dp0"
  pause
  exit /b 1
)

REM Picker pops up. You select rows, it saves docs and opens the folder.
py -3.13 "%~dp0build_court_visitor_summary.py" --open

echo.
echo Done. If the folder didn't open, files are in:
echo   C:\GoogleSync\Automation\Court Visitor Summary
echo.
pause
popd
