# =============================================================
#  Airtable Document Agent - Disable auto-start at login
#  Removes the shortcut from the Startup folder.
# =============================================================

$ErrorActionPreference = "Stop"

$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath  = Join-Path $startupFolder "Airtable Document Agent.lnk"

Write-Host ""
if (Test-Path $shortcutPath) {
    Remove-Item $shortcutPath -Force
    Write-Host "  [OK] Auto-start disabled. Shortcut removed." -ForegroundColor Green
    Write-Host ""
    Write-Host "  The agent will no longer launch at login."
    Write-Host "  Start manually with start.ps1 when needed."
} else {
    Write-Host "  Auto-start was not enabled (no shortcut found)." -ForegroundColor Yellow
}
Write-Host ""
Read-Host "Press Enter to close"
