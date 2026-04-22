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
