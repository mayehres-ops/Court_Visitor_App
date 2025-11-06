@echo off
setlocal
set DOCSTRANGE_API_KEY_FILE=C:\configlocal\API\docstrange.key.txt
py -3.12 "C:\GoogleSync\Automation\GuardianAutomation\scripts\docstrange_cloud_pull.py"
if errorlevel 1 (
  echo ❌ Script failed with error %ERRORLEVEL%
  exit /b %ERRORLEVEL%
)
echo ✅ Done. CSV written to C:\GoogleSync\Guardianship Files\Extracted\docstrange_fields.csv
