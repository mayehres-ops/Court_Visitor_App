@echo off
REM Preview run (renders HTML previews; no email is sent)
py "C:\GoogleSync\Automation\GuardianComms\comm_module.py" preview --workbook "C:\GoogleSync\Automation\GuardianComms\Clients.xlsx" --template simple_reminder --subject "Friendly Reminder"
pause
