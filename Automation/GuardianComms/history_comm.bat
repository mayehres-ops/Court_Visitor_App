@echo off
REM Show recent messages for a client (by CauseNo or ClientID substring)
set /p CLIENTKEY=Enter CauseNo or ClientID piece: 
py "C:\GoogleSync\Automation\GuardianComms\comm_module.py" history --workbook "C:\GoogleSync\Automation\GuardianComms\Clients.xlsx" --client "%CLIENTKEY%"
pause
