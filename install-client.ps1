param(
  [string]$InstallDir = "$env:USERPROFILE\airtable-agent",
  [string]$RepoZipUrl = "https://github.com/martin10101/airtable-agent-client-fix/archive/refs/heads/main.zip"
)

$ErrorActionPreference = "Stop"

function Write-Step($message) {
  Write-Host "[airtable-agent] $message"
}

Write-Step "Installing to $InstallDir"

$tmpRoot = Join-Path $env:TEMP ("airtable-agent-install-" + [Guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tmpRoot "package.zip"
$extractDir = Join-Path $tmpRoot "extract"
New-Item -ItemType Directory -Force -Path $tmpRoot, $extractDir | Out-Null

Write-Step "Downloading package"
Invoke-WebRequest -Uri $RepoZipUrl -OutFile $zipPath -UseBasicParsing
Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force
$sourceDir = Get-ChildItem -LiteralPath $extractDir -Directory | Select-Object -First 1
if (-not $sourceDir) { throw "Could not find extracted package folder." }

$listeners = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue |
  Where-Object { $_.OwningProcess -and $_.OwningProcess -gt 0 } |
  Select-Object -ExpandProperty OwningProcess -Unique

if ($listeners) {
  Write-Step "Stopping existing server on port 3000"
  foreach ($listenerPid in $listeners) {
    Stop-Process -Id $listenerPid -Force -ErrorAction SilentlyContinue
  }
  for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 500
    $stillListening = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue |
      Where-Object { $_.OwningProcess -and $_.OwningProcess -gt 0 }
    if (-not $stillListening) { break }
  }
}

$remainingListeners = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue |
  Where-Object { $_.OwningProcess -and $_.OwningProcess -gt 0 }
if ($remainingListeners) {
  $pids = ($remainingListeners | Select-Object -ExpandProperty OwningProcess -Unique) -join ", "
  throw "Could not stop the existing server on port 3000. Still listening PID(s): $pids"
}

if (Test-Path -LiteralPath $InstallDir) {
  $backup = Join-Path $InstallDir ("backup-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
  Write-Step "Backing up current files to $backup"
  New-Item -ItemType Directory -Force -Path $backup | Out-Null
  Get-ChildItem -LiteralPath $InstallDir -Force |
    Where-Object { $_.Name -notin @("node_modules", "logs", "work") -and $_.Name -notlike "backup-*" } |
    ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $backup -Recurse -Force }
} else {
  New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
}

Write-Step "Copying updated app files"
Get-ChildItem -LiteralPath $sourceDir.FullName -Force |
  Where-Object { $_.Name -notin @(".git", ".env", "node_modules", "logs", "work") } |
  ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $InstallDir -Recurse -Force }

$serverPath = Join-Path $InstallDir "server.js"
$versionLine = Select-String -Path $serverPath -Pattern "APP_VERSION" | Select-Object -First 1
if ($versionLine) {
  Write-Step ("Installed " + $versionLine.Line.Trim())
}

$envPath = Join-Path $InstallDir ".env"
if (-not (Test-Path -LiteralPath $envPath)) {
  Write-Step "No .env found; creating one from .env.example. Fill in API keys before running."
  Copy-Item -LiteralPath (Join-Path $InstallDir ".env.example") -Destination $envPath -Force
}

Write-Step "Installing npm dependencies"
Push-Location $InstallDir
npm install
Pop-Location

Write-Step "Starting server"
Start-Process -FilePath "node" -ArgumentList "server.js" -WorkingDirectory $InstallDir -WindowStyle Hidden

for ($i = 0; $i -lt 20; $i++) {
  Start-Sleep -Milliseconds 500
  try {
    $health = Invoke-RestMethod "http://localhost:3000/health" -TimeoutSec 2
    if ($health.version) {
      Write-Step ("Running version " + $health.version)
      break
    }
  } catch {}
}

Write-Step "Done. Open http://localhost:3000"
