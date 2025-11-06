@echo off
set "SCRIPT=C:\GoogleSync\Automation\Appt Email Confirm\scripts\send_confirmation_email.py"
echo ===============================
echo  Appt Email Confirm - MENU
echo ===============================
echo 1) Test - Dry Run (no send, no write)
echo 2) Test - Send One (marks Y)
echo 3) Live - All Eligible (marks Y)
echo.
choice /c 123 /m "Select an option"
if errorlevel 3 goto LIVE
if errorlevel 2 goto TESTONE
if errorlevel 1 goto DRYRUN

:DRYRUN
py -3 "%SCRIPT%" --mode test_last_row --dry-run
goto END

:TESTONE
echo Make sure Excel is CLOSED so the script can write "Y".
py -3 "%SCRIPT%" --mode test_last_row
goto END

:LIVE
echo You are about to send confirmation emails and mark Appt_confirmed=Y for ALL eligible rows.
choice /m "Continue?"
if errorlevel 2 goto END
echo Make sure Excel is CLOSED so the script can write "Y".
py -3 "%SCRIPT%" --mode live
goto END

:END
echo.
echo Log: C:\GoogleSync\Automation\Appt Email Confirm\Logs\appt_confirm.log
pause
