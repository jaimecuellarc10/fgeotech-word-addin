@echo off
title FGeotech - Word Add-in Server
cd /d "%~dp0"

where node >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js is not installed.
    echo Please download and install it from https://nodejs.org
    pause
    exit /b 1
)

echo ============================================
echo   FGeotech Word Add-in Server
echo   ACT Geotechnical Engineers
echo ============================================
echo.
echo Starting server at https://localhost:3000
echo Keep this window open while using Word.
echo Close this window to stop the server.
echo.

node server.js --https

echo.
echo Server stopped.
pause
