<#
.SYNOPSIS
    One-time setup for a second check-in laptop: installs dependencies,
    configures the secret vault, and gets event.config.json/event.db in place.
.DESCRIPTION
    Run this once on a laptop that will run Start-Checkin.bat as a check-in
    desk. If azure.enabled is true in event.config.json, this desk shares
    live state (check-ins, badge prints) with every other desk via Azure
    SQL — its local event.db is just a warm cache that self-populates on
    first run and queues writes during any network drop. If azure.enabled
    is false, this laptop runs the older local-only, single-desk model.

    Assumes the sqlsat-tools repo has already been cloned or copied onto this
    laptop and this script is being run from inside it.
.EXAMPLE
    .\scripts\setup\Setup-CheckinLaptop.ps1
#>
param(
    [switch]$SkipEventbriteToken
)

$ErrorActionPreference = 'Stop'
$repoRoot = Join-Path $PSScriptRoot ".." ".."

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
$azureEnabled = $config -and $config.PSObject.Properties['azure'] -and $config.azure.enabled

if (Test-Path $dbPath) {
    Write-Ok "Found at $dbPath"
    if (-not $azureEnabled) {
        Write-Warn2 "azure.enabled is false, so this is an INDEPENDENT copy — it will not see"
        Write-Warn2 "badges printed on any other laptop, and vice versa. Copy the latest event.db"
        Write-Warn2 "from the primary laptop right before the event so attendee data is current."
    }
} elseif ($config) {
    Write-Warn2 "Not found. Run:"
    Write-Warn2 "  .\scripts\setup\Initialize-Database.ps1 -Config `$config"
    if ($azureEnabled) {
        Write-Warn2 "It'll be empty at first but self-populates from Azure SQL on the first"
        Write-Warn2 "Checkin-Menu.ps1 run — no need to copy event.db from another laptop."
    } else {
        Write-Warn2 "Then copy event.db from the primary laptop, or run it here as walk-ins only."
    }
}

if ($azureEnabled) {
    Write-Step "Checking Azure SQL connectivity"
    if (-not (Get-SecretInfo -Name $config.azure.authSecretName -ErrorAction SilentlyContinue)) {
        Write-Warn2 "Azure SQL password secret '$($config.azure.authSecretName)' isn't set on this laptop."
        $answer = Read-Host "  Set it now? (y/N)"
        if ($answer -match '^[Yy]') {
            $azurePassword = Read-Host "  Paste the Azure SQL password" -AsSecureString
            Set-Secret -Name $config.azure.authSecretName -SecureStringSecret $azurePassword
            Write-Ok "Saved"
        }
    } else {
        Write-Ok "Secret '$($config.azure.authSecretName)' is set"
    }

    if ((Get-SecretInfo -Name $config.azure.authSecretName -ErrorAction SilentlyContinue)) {
        try {
            . (Join-Path $repoRoot "scripts\internal\Data-Access.ps1")
            $ctx = New-DataContext -Config $config
            $detail = Test-DatabaseReadiness -DataContext $ctx
            Write-Ok "Azure SQL reachable — $detail"
        } catch {
            Write-Warn2 "Could not confirm Azure SQL is reachable: $($_.Exception.Message)"
            Write-Warn2 "This desk will still work — it'll queue writes locally and catch up once"
            Write-Warn2 "the connection is sorted out. Worth fixing before event day, though."
        }
    }
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
