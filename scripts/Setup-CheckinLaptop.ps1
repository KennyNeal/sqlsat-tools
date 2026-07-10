<#
.SYNOPSIS
    One-time setup for a second check-in laptop: installs dependencies,
    configures the secret vault, and gets event.config.json/event.db in place.
.DESCRIPTION
    Run this once on a laptop that will run Start-Checkin.bat as an
    independent check-in desk (its own local copy of event.db — not a live
    shared database with any other laptop; see issue for multi-desk support).

    Assumes the sqlsat-tools repo has already been cloned or copied onto this
    laptop and this script is being run from inside it.
.EXAMPLE
    .\scripts\Setup-CheckinLaptop.ps1
#>
param(
    [switch]$SkipEventbriteToken
)

$ErrorActionPreference = 'Stop'
$repoRoot = Join-Path $PSScriptRoot ".."

function Write-Step($msg) { Write-Host ""; Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    OK: $msg" -ForegroundColor Green }
function Write-Warn2($msg) { Write-Host "    NOTE: $msg" -ForegroundColor Yellow }

# ── PowerShell version ────────────────────────────────────────────────────

Write-Step "Checking PowerShell version"
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warn2 "You're on PowerShell $($PSVersionTable.PSVersion). PowerShell 7+ is required."
    Write-Warn2 "Install it with: winget install --id Microsoft.PowerShell -e"
    Write-Warn2 "Then re-run this script from a pwsh (not powershell.exe) prompt."
    exit 1
}
Write-Ok "PowerShell $($PSVersionTable.PSVersion)"

# ── winget-installed tools ────────────────────────────────────────────────

Write-Step "Checking SumatraPDF (silent printing)"
$sumatra = Get-Command SumatraPDF.exe -ErrorAction SilentlyContinue
if (-not $sumatra) {
    $sumatra = @("$env:LOCALAPPDATA\SumatraPDF\SumatraPDF.exe", "${env:ProgramFiles}\SumatraPDF\SumatraPDF.exe") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $sumatra) {
    Write-Warn2 "Not found — installing via winget..."
    winget install --id SumatraPDF.SumatraPDF -e --accept-source-agreements --accept-package-agreements
} else {
    Write-Ok "Found at $sumatra"
}

Write-Step "Checking Microsoft Edge (HTML-to-PDF rendering)"
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) {
    Write-Warn2 "Not found — installing via winget..."
    winget install --id Microsoft.Edge -e --accept-source-agreements --accept-package-agreements
} else {
    Write-Ok "Found at $edge"
}

# ── PowerShell modules ────────────────────────────────────────────────────

Write-Step "Checking required PowerShell modules"
foreach ($mod in @('PSSQLite', 'Microsoft.PowerShell.SecretManagement', 'Microsoft.PowerShell.SecretStore')) {
    if (Get-Module -ListAvailable -Name $mod) {
        Write-Ok "$mod already installed"
    } else {
        Write-Warn2 "Installing $mod..."
        Install-Module -Name $mod -Scope CurrentUser -Force -Repository PSGallery
        Write-Ok "$mod installed"
    }
}

# ── Secret vault (no master password — matches the primary laptop) ───────

Write-Step "Configuring the secret vault"
if (-not (Get-SecretVault -Name 'SecretStore' -ErrorAction SilentlyContinue)) {
    Register-SecretVault -Name 'SecretStore' -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
    Write-Ok "Registered SecretStore as the default vault"
}
Set-SecretStoreConfiguration -Authentication None -Confirm:$false | Out-Null
Write-Ok "Vault set to no master password (same as the primary laptop)"

# ── Eventbrite token (only needed if this laptop will run option 3) ──────

if (-not $SkipEventbriteToken) {
    $configPath = Join-Path $repoRoot "event.config.json"
    if (Test-Path $configPath) {
        $config = Get-Content -Raw $configPath | ConvertFrom-Json
        $secretName = $config.eventbrite.secretName
        Write-Step "Eventbrite token ($secretName)"
        if (Get-SecretInfo -Name $secretName -ErrorAction SilentlyContinue) {
            Write-Ok "Already set"
        } else {
            Write-Warn2 "Not set. This laptop only needs it if it will use menu option 3"
            Write-Warn2 "(Sync new registrations from Eventbrite). Skip if it's walk-ins only."
            $answer = Read-Host "  Set it now? (y/N)"
            if ($answer -match '^[Yy]') {
                $token = Read-Host "  Paste the Eventbrite API token" -AsSecureString
                Set-Secret -Name $secretName -SecureStringSecret $token
                Write-Ok "Saved"
            }
        }
    }
}

# ── event.config.json / event.db / QRCoder.dll ────────────────────────────

Write-Step "Checking event.config.json"
$configPath = Join-Path $repoRoot "event.config.json"
if (Test-Path $configPath) {
    Write-Ok "Found"
} else {
    Write-Warn2 "Missing. Copy event.config.json from the primary laptop into the repo root"
    Write-Warn2 "(event.config.template.json shows the shape if you need to recreate it)."
}

Write-Step "Checking event.db"
$config = if (Test-Path $configPath) { Get-Content -Raw $configPath | ConvertFrom-Json } else { $null }
$dbPath = if ($config) { Join-Path $repoRoot $config.database.path } else { Join-Path $repoRoot "event.db" }
if (Test-Path $dbPath) {
    Write-Ok "Found at $dbPath"
    Write-Warn2 "This is an INDEPENDENT copy — it will not see badges printed on any other"
    Write-Warn2 "laptop, and vice versa. Copy the latest event.db from the primary laptop"
    Write-Warn2 "right before the event so attendee data is current."
} elseif ($config) {
    Write-Warn2 "Not found. Copy event.db from the primary laptop into the repo root, or run:"
    Write-Warn2 "  .\scripts\Initialize-Database.ps1 -Config `$config"
    Write-Warn2 "to create an empty one (only useful if this laptop is walk-ins only)."
}

Write-Step "Checking lib\QRCoder.dll"
$libPath = Join-Path $repoRoot "lib\QRCoder.dll"
if (Test-Path $libPath) {
    Write-Ok "Found"
} else {
    Write-Warn2 "Missing — should have come with the repo clone/copy. Copy lib\QRCoder.dll"
    Write-Warn2 "from the primary laptop's repo folder."
}

# ── Printer sanity check ──────────────────────────────────────────────────

Write-Step "Checking the label printer"
$printerName = if ($config -and $config.PSObject.Properties['badge'] -and $config.badge.walkinPrinter) {
    $config.badge.walkinPrinter
} else {
    "Brother QL-820NWB"
}
if (Get-Printer -Name $printerName -ErrorAction SilentlyContinue) {
    Write-Ok "'$printerName' is installed on this laptop"
} else {
    Write-Warn2 "'$printerName' not found. Install its driver / connect it before check-in day."
}

Write-Host ""
Write-Host "Setup complete. Double-click Start-Checkin.bat to launch the check-in menu." -ForegroundColor Cyan
