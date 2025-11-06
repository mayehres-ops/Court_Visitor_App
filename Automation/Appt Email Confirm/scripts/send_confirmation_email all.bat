@echo off
set "SCRIPT=C:\GoogleSync\Automation\Appt Email Confirm\scripts\send_confirmation_email.py"
echo You are about to send confirmation emails and mark Appt_confirmed=Y for ALL eligible rows.
choice /m "Continue?"
if errorlevel 2 goto :eof
echo Make sure Excel is CLOSED so the script can write "Y".
py -3 "%SCRIPT%" --mode live
echo.
echo Log: C:\GoogleSync\Automation\Appt Email Confirm\Logs\appt_confirm.log
pause
