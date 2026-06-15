@echo off
REM =============================================================
REM  Airtable Document Agent - Pre-flight diagnostic
REM
REM  Run this FIRST, before first-time-setup.bat.
REM  It tests whether this PC can reach every service the agent
REM  needs and whether Node.js is installed correctly.
REM  Nothing is installed or modified.
REM =============================================================

cd /d "%~dp0"
set NODE_TLS_REJECT_UNAUTHORIZED=0

echo.
echo ======================================================
echo   Airtable Agent - Environment Diagnostic
echo ======================================================

REM -- Step 1: Node.js present?
where node >nul 2>&1
if errorlevel 1 (
    echo.
    echo   [FAIL] Node.js is not installed or not in PATH.
    echo.
    echo   Fix: Install Node.js 22 LTS from https://nodejs.org
    echo         Default options are fine. After installing,
    echo         OPEN A NEW Command Prompt window and re-run
    echo         this file.
    echo.
    pause
    exit /b 1
)

REM -- Step 2: npm callable?
where npm >nul 2>&1
if errorlevel 1 (
    echo.
    echo   [FAIL] npm is not callable from this shell.
    echo          (This usually means Node.js was installed
    echo          but the PATH was not refreshed — open a
    echo          NEW Command Prompt window and try again.)
    echo.
    pause
    exit /b 1
)

REM -- Step 3: run the Node-based connectivity + write-permission check
node check-connectivity.js
set EXITCODE=%ERRORLEVEL%

echo ======================================================
if %EXITCODE% EQU 0 (
    echo   OK to proceed. Next step: double-click first-time-setup.bat
) else (
    echo   Fix the items marked [FAIL] above, then re-run this diagnostic.
    echo   Send the output to whoever set up this agent if you need help.
)
echo ======================================================
echo.
pause
exit /b %EXITCODE%
