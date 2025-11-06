@echo off
set "OUTPUT=C:\GoogleSync\Automation\Build Map Sheet\Ward_Map_Sheet.docx"
if not exist "%OUTPUT%" (
  echo Can't find: %OUTPUT%
  exit /b 1
)
powershell -NoProfile -Command "Start-Process -FilePath '%OUTPUT%' -Verb Print"
