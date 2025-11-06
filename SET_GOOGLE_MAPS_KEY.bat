@echo off
REM Set Google Maps API Key for GuardianShip App
REM This sets the environment variable for the current session only
REM For permanent setup, add to System Environment Variables

echo ========================================
echo Google Maps API Key Setup
echo ========================================
echo.
echo This script will set your Google Maps API key
echo for the current session.
echo.
echo To make it permanent:
echo 1. Right-click 'This PC' ^> Properties
echo 2. Advanced System Settings ^> Environment Variables
echo 3. Add new SYSTEM variable:
echo    Name: GOOGLE_MAPS_API_KEY
echo    Value: YOUR_API_KEY_HERE
echo.
echo ========================================
echo.

set /p APIKEY="Enter your Google Maps API Key: "

if "%APIKEY%"=="" (
    echo Error: No API key entered!
    pause
    exit /b 1
)

setx GOOGLE_MAPS_API_KEY "%APIKEY%"

echo.
echo ========================================
echo SUCCESS!
echo ========================================
echo.
echo Google Maps API Key has been set!
echo.
echo IMPORTANT: You need to RESTART your computer
echo or close and reopen the app for the change to take effect.
echo.
pause
