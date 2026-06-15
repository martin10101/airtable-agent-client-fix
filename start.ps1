# =============================================================
#  Airtable Document Agent - Launcher
#  Right-click -> Run with PowerShell to start the local server.
# =============================================================

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot
$env:NODE_TLS_REJECT_UNAUTHORIZED = "0"

if (-not (Test-Path ".env")) {
    Write-Host ""
    Write-Host "[!] .env file is missing." -ForegroundColor Red
    Write-Host "    Run install.ps1 first, then fill in your API keys in .env."
    Read-Host "Press Enter to close"
    exit 1
}

if (-not (Test-Path "node_modules")) {
    Write-Host ""
    Write-Host "[!] node_modules is missing." -ForegroundColor Red
    Write-Host "    Run install.ps1 first."
    Read-Host "Press Enter to close"
    exit 1
}

Write-Host ""
Write-Host "Starting Airtable Document Agent..." -ForegroundColor Cyan
Write-Host "Open http://localhost:3000 in your browser." -ForegroundColor Yellow
Write-Host "Keep this window open. Close it to stop the server." -ForegroundColor Gray
Write-Host ""

& node server.js

Write-Host ""
Write-Host "[!] Server stopped." -ForegroundColor Red
Read-Host "Press Enter to close"
