@echo off
set "PY=C:\Users\may\AppData\Local\Programs\Python\Python313\python.exe"
set "SCRIPT=C:\GoogleSync\Automation\Contacts - Guardians\scripts\add_guardians_to_contacts.py"

echo You are about to create/update contacts for ALL eligible rows and mark Contact_added=Y.
set /p GO=Type YES to continue: 
if /I not "%GO%"=="YES" goto :END

"%PY%" "%SCRIPT%" --mode live
echo.
echo Exit code: %errorlevel%
echo Log: C:\GoogleSync\Automation\Contacts - Guardians\Logs\contacts.log
:END
pause
