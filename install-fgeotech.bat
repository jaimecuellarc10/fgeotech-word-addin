@echo off
title FGeotech - Add-in Installer
cd /d "%~dp0"

echo ============================================
echo   FGeotech Word Add-in Installer
echo   ACT Geotechnical Engineers
echo ============================================
echo.
echo This will register the FGeotech add-in in Word.
echo.

:: Set registry key pointing to manifest.xml in the same folder as this script
set MANIFEST_PATH=%~dp0manifest.xml

:: Abort if running from a Temp folder (e.g. inside a zip)
echo %MANIFEST_PATH% | findstr /i "\\Temp\\" >nul
if %errorlevel% equ 0 (
    echo ERROR: Do not run this from inside a zip file.
    echo Extract the folder first, then run install-fgeotech.bat from there.
    echo.
    pause
    exit /b 1
)

if not exist "%MANIFEST_PATH%" (
    echo ERROR: manifest.xml not found next to this script.
    echo Make sure install-fgeotech.bat and manifest.xml are in the same folder.
    echo.
    pause
    exit /b 1
)

reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\Developer" /v "FGeotech" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

if %errorlevel% equ 0 (
    echo SUCCESS: FGeotech add-in has been registered.
    echo.
    echo Next steps:
    echo   1. Open or restart Microsoft Word
    echo   2. Look for the FGeotech button in the Home ribbon
    echo.
) else (
    echo ERROR: Could not write to the registry.
    echo Try right-clicking this file and selecting "Run as administrator".
    echo.
)

pause
