# ============================================================
#  run_app.ps1  -  Install uv, sync deps, launch Flask app
# ============================================================

$ErrorActionPreference = "Stop"

function Write-Step { param($msg) Write-Host "`n[>>] $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "[ OK] $msg"  -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "[!]  $msg"   -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "[X]  $msg"   -ForegroundColor Red; Read-Host "Press Enter to exit"; exit 1 }

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir
Write-Step "Working directory: $ScriptDir"

if (-not (Test-Path "main.py")) {
    Write-Fail "main.py not found. Place run_app.ps1 next to main.py and try again."
}
Write-OK "main.py found"

# ── 1. Install uv if not already present ─────────────────────
Write-Step "Checking for uv..."

$uvCmd = $null
try {
    $uvVer = & uv --version 2>&1
    if ($uvVer -match "uv \d+\.\d+") {
        $uvCmd = "uv"
        Write-OK "Found: $uvVer"
    }
} catch {}

if (-not $uvCmd) {
    Write-Warn "uv not found. Installing via official installer..."
    try {
        powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    } catch {
        Write-Fail "uv installation failed: $_"
    }

    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = $machinePath + ";" + $userPath

    $uvLocalBin = Join-Path $env:USERPROFILE ".local\bin"
    if (Test-Path $uvLocalBin) {
        $env:Path = $uvLocalBin + ";" + $env:Path
    }

    try {
        $uvVer = & uv --version 2>&1
        if ($uvVer -match "uv \d+\.\d+") {
            $uvCmd = "uv"
            Write-OK "uv installed: $uvVer"
        }
    } catch {}

    if (-not $uvCmd) {
        Write-Fail "uv was installed but could not be found on PATH. Please restart PowerShell and run the script again."
    }
}

# ── 2. Sync dependencies from pyproject.toml ─────────────────
Write-Step "Syncing dependencies with uv..."

if (-not (Test-Path (Join-Path $ScriptDir "pyproject.toml"))) {
    Write-Fail "pyproject.toml not found. uv requires a pyproject.toml to manage dependencies."
}

& uv sync
if ($LASTEXITCODE -ne 0) { Write-Fail "uv sync failed." }
Write-OK "Dependencies ready"

# ── 3. Open browser after a short delay ──────────────────────
Write-Step "Starting Flask app (main.py)..."
Write-Host ""
Write-Host "  Opening http://localhost:80 in your browser..." -ForegroundColor DarkGray
Write-Host "  Press Ctrl+C to stop the server." -ForegroundColor DarkGray
Write-Host ""

# Launch browser in background after 2s to give Flask time to start
Start-Job -ScriptBlock {
    Start-Sleep -Seconds 2
    Start-Process "http://localhost:80"
} | Out-Null

# ── 4. Launch the Flask app via uv run ───────────────────────
$env:FLASK_APP = "main.py"
$env:FLASK_ENV = "development"

& uv run main.py