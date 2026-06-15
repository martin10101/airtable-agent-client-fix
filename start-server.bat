@echo off
REM =============================================================
REM  Airtable Document Agent — one-click launcher
REM  Double-click this file to start the local server.
REM =============================================================

cd /d "%~dp0"

REM Required for corporate networks that do SSL inspection
set NODE_TLS_REJECT_UNAUTHORIZED=0

if not exist ".env" (
    echo.
    echo [!] .env file is missing.
    echo     Copy .env.example to .env and fill in your API keys, then re-run.
    echo.
    pause
    exit /b 1
)

if not exist "node_modules" (
    echo.
    echo [!] node_modules folder is missing.
    echo     Run: npm install
    echo     in this folder, then try again.
    echo.
    pause
    exit /b 1
)

echo.
echo  Starting Airtable Document Agent...
echo  Keep this window open. Close it to stop the server.
echo.

node server.js

echo.
echo [!] Server stopped.
pause
