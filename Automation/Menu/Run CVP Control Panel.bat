@echo on
setlocal
REM === Court Visitor Control Panel launcher (diagnostic & self-locating) ===

set "HERE=%~dp0"
set "MENUPY=%HERE%cvp_menu.py"
if not exist "%MENUPY%" set "MENUPY=%HERE%cvp_menu.py.txt"

echo Looking for: "%MENUPY%"
if not exist "%MENUPY%" (
  echo [!] cvp_menu.py not found in "%HERE%".
  dir /b "%HERE%"
  pause
  exit /b 1
)

echo Python launchers on PATH:
where py
where python

set "LAUNCHLOG=%HERE%launch_stdout_stderr.txt"
echo ===== %DATE% %TIME% ===== >> "%LAUNCHLOG%"
echo Launching: "%MENUPY%" >> "%LAUNCHLOG%"

REM Try Python 3.13, then fallback to py -3, then python
py -3.13 "%MENUPY%" 1>>"%LAUNCHLOG%" 2>&1 ^
  || py -3 "%MENUPY%" 1>>"%LAUNCHLOG%" 2>&1 ^
  || python "%MENUPY%" 1>>"%LAUNCHLOG%" 2>&1

echo ExitCode=%ERRORLEVEL%
echo Log written to: "%LAUNCHLOG%"
echo (If the GUI didn't appear, open that log to see the error)
pause
endlocal
