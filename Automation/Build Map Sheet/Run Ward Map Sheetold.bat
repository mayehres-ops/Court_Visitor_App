@echo off
setlocal

rem ========= Config =========
set "PY=py -3"
set "SCRIPT=C:\GoogleSync\Automation\Build Map Sheet\Scripts\build_map_sheet.py"
set "OUTPUT=C:\GoogleSync\Automation\Build Map Sheet\Ward_Map_Sheet.docx"

rem Auto-print toggle (0 = ask first, 1 = print automatically)
set "AUTO_PRINT=0"

rem ===== Optional: set your key only for this run =====
rem set GOOGLE_MAPS_API_KEY=PASTE_YOUR_ACTUAL_KEY_HERE

echo.
if defined GOOGLE_MAPS_API_KEY (
  echo [init] GOOGLE_MAPS_API_KEY detected in this session/user env.
) else (
  echo [init] No GOOGLE_MAPS_API_KEY detected. Script will use the free fallback geocoder.
)

echo.
echo Running map sheet script...
"%PY%" "%SCRIPT%"
set "RC=%ERRORLEVEL%"

echo.
if %RC% NEQ 0 (
  echo Script failed with exit code %RC%.
  echo (Scroll up for errors; fix them and re-run.)
  pause
  exit /b %RC%
)

if not exist "%OUTPUT%" (
  echo Script reported success but the output file was not found:
  echo   %OUTPUT%
  echo Please check console messages above.
  pause
  exit /b 1
)

echo [done] Output created: "%OUTPUT%"

rem ========= Print logic =========
if "%AUTO_PRINT%"=="1" goto do_print

echo.
choice /M "Print the map sheet now?"
if errorlevel 2 goto maybe_open
if errorlevel 1 goto do_print

:do_print
echo Sending to default printer...
powershell -NoProfile -Command "Start-Process -FilePath '%OUTPUT%' -Verb Print"
rem small wait so Word can grab the file before we try to open it
timeout /t 2 >nul

:maybe_open
echo.
choice /M "Open the map sheet now?"
if errorlevel 2 goto end
if errorlevel 1 start "" "%OUTPUT%"

:end
exit /b 0
