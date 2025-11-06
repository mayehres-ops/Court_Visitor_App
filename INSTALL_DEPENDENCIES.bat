@echo off
REM Install all required Python libraries for GuardianShip App
echo ========================================
echo GuardianShip App - Dependency Installer
echo ========================================
echo.
echo This will install all required Python libraries.
echo This may take 5-10 minutes depending on your internet speed.
echo.
pause

echo.
echo Installing dependencies...
echo.

pip install -r requirements.txt

echo.
echo ========================================
echo Installation Complete!
echo ========================================
echo.
echo You can now run the app using launch_app.bat
echo.
pause
