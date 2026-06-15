# =============================================================
#  Airtable Document Agent - Auto-start at login
#
#  Drops a shortcut to start.ps1 into the Windows Startup folder
#  so the agent boots automatically every time you log in.
#
#  Run once. Re-runs are safe (just overwrites the shortcut).
#  To disable: run disable-autostart.ps1 (or delete the shortcut
#  manually from the Startup folder).
# =============================================================

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Write-Ok($msg)   { Write-Host "  [OK] $msg"   -ForegroundColor Green }
function Write-Fail($msg) { Write-Host "  [FAIL] $msg" -ForegroundColor Red }
function Write-Info($msg) { Write-Host "  $msg"        -ForegroundColor Gray }

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Enable Auto-start at Login"                 -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$startPs1 = Join-Path $PSScriptRoot "start.ps1"
if (-not (Test-Path $startPs1)) {
    Write-Fail "Cannot find start.ps1 in this folder."
    Write-Info "This script must live next to start.ps1."
    Read-Host "Press Enter to close"
    exit 1
}

$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath  = Join-Path $startupFolder "Airtable Document Agent.lnk"

# Remove any prior shortcut so re-runs don't pile up
if (Test-Path $shortcutPath) {
    Remove-Item $shortcutPath -Force
    Write-Info "Replacing existing shortcut..."
}

$ws = New-Object -ComObject WScript.Shell
$shortcut = $ws.CreateShortcut($shortcutPath)
$shortcut.TargetPath       = "powershell.exe"
$shortcut.Arguments        = "-ExecutionPolicy Bypass -WindowStyle Minimized -File `"$startPs1`""
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.Description      = "Starts the Airtable Document Agent local server"
$shortcut.IconLocation     = "powershell.exe,0"
$shortcut.Save()

Write-Ok "Shortcut created at:"
Write-Info "  $shortcutPath"
Write-Host ""
Write-Host "What happens now:" -ForegroundColor Yellow
Write-Host "  - Every time you log in to Windows, the agent starts automatically."
Write-Host "  - The PowerShell window opens MINIMIZED (you'll see it in the taskbar)."
Write-Host "  - localhost:3000 is reachable as soon as the window finishes loading (~5s)."
Write-Host ""
Write-Host "To stop the agent right now: close the minimized PowerShell window."
Write-Host "To disable auto-start later: run disable-autostart.ps1"
Write-Host ""
Read-Host "Press Enter to close"
