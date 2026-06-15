@echo off
REM =============================================================
REM  Airtable Document Agent - First-time setup
REM
REM  What this does (safe to re-run):
REM    1. Verifies Node.js / npm are callable
REM    2. Runs npm install using a project-local .npmrc that
REM       tolerates corporate SSL inspection
REM    3. Copies .env.example to .env if .env doesn't exist yet
REM =============================================================

cd /d "%~dp0"
set NODE_TLS_REJECT_UNAUTHORIZED=0

echo.
echo ======================================================
echo   Airtable Agent - First-time Setup
echo ======================================================
echo.

REM -- Verify Node.js / npm are available
where node >nul 2>&1
if errorlevel 1 (
    echo   [FAIL] Node.js is not installed.
    echo   Install Node.js 22 LTS from https://nodejs.org, then re-run.
    echo.
    pause
    exit /b 1
)
where npm >nul 2>&1
if errorlevel 1 (
    echo   [FAIL] npm is not callable. Open a NEW Command Prompt and try again.
    echo.
    pause
    exit /b 1
)

for /f "delims=" %%v in ('node --version') do set NODE_VERSION=%%v
echo   Node.js: %NODE_VERSION%
for /f "delims=" %%v in ('npm --version') do set NPM_VERSION=%%v
echo   npm:     %NPM_VERSION%
echo.

REM -- Install dependencies
echo   Installing dependencies (this can take 1-3 minutes)...
echo   ---------------------------------------------------
echo.
call npm install
set INSTALL_EXIT=%ERRORLEVEL%
echo.
echo   ---------------------------------------------------

if not %INSTALL_EXIT% EQU 0 (
    echo.
    echo   [FAIL] npm install failed.
    echo.
    echo   Common fixes:
    echo     * "unable to verify the first certificate" or similar SSL error:
    echo         Corporate SSL inspection is blocking npm. Ask IT to
    echo         whitelist registry.npmjs.org, OR run check-environment.bat
    echo         and send the output to IT.
    echo     * "ETIMEDOUT" or "ENOTFOUND":
    echo         The firewall or proxy is blocking the npm registry.
    echo         Run check-environment.bat to see which endpoints fail.
    echo     * "EACCES" or permission denied:
    echo         Move this folder to a location where you have write
    echo         access (e.g. your Documents folder).
    echo.
    pause
    exit /b 1
)

REM -- Create .env from template if missing
if not exist ".env" (
    if exist ".env.example" (
        copy ".env.example" ".env" >nul
        echo.
        echo   Created .env from template.
        echo   *** Open .env in Notepad and fill in your API keys. ***
    ) else (
        echo.
        echo   [WARN] .env.example is missing. You will need to create .env manually.
    )
) else (
    echo.
    echo   .env already exists — leaving it alone.
)

echo.
echo ======================================================
echo   Setup complete.
echo ======================================================
echo.
echo   Next steps:
echo     1. Open .env in Notepad and fill in:
echo          ANTHROPIC_API_KEY=...
echo          AIRTABLE_API_KEY=...
echo          AIRTABLE_BASE_ID=...
echo          AIRTABLE_TABLE_NAME=...
echo     2. Double-click start-server.bat to launch the agent.
echo     3. Open http://localhost:3000 in your browser.
echo.
pause
