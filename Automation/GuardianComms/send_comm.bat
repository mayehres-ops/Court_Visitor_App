@echo off
REM Send for real via Gmail SMTP. Requires env vars set once:
REM   setx EMAIL_USER "yourname@gmail.com"
REM   setx EMAIL_APP_PASSWORD "your_16_char_app_password"
py "C:\GoogleSync\Automation\GuardianComms\comm_module.py" send --workbook "C:\GoogleSync\Automation\GuardianComms\Clients.xlsx" --template simple_reminder --subject "Friendly Reminder"
pause
