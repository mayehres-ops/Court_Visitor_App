@echo off
rem Run from the same folder as this BAT
cd /d "%~dp0"
python build_payment_forms_sdt.py
pause
