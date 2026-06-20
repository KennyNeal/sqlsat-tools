<#
.SYNOPSIS
    Master pipeline: import attendees, generate SpeedPasses, send emails.
.DESCRIPTION
    Runs the three event pipeline steps in sequence. Use -WhatIf to preview
    without making any changes. Use -StepThrough to confirm each step before it runs.

    For one-off tasks (stamp game, schedule) use the individual scripts in scripts\.
.PARAMETER ConfigPath
    Path to event.config.json. Defaults to event.config.json next to this script.
.PARAMETER StepThrough
    Pause and prompt before each pipeline step.
.PARAMETER ImportOnly
    Run only the EventBrite import step.
.PARAMETER GenerateOnly
    Run only the SpeedPass generation step.
.PARAMETER EmailOnly
    Run only the email delivery step.
.EXAMPLE
    .\Update-Event.ps1
    Full pipeline run (import → generate → email).
.EXAMPLE
    .\Update-Event.ps1 -WhatIf
    Preview what each step would do without making changes.
.EXAMPLE
    .\Update-Event.ps1 -StepThrough
    Confirm each step before it runs.
.EXAMPLE
    .\Update-Event.ps1 -EmailOnly
    Only send emails to attendees with generated SpeedPasses not yet emailed.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ConfigPath = "$PSScriptRoot\event.config.json",
    [switch]$StepThrough,
    [switch]$ImportOnly,
    [switch]$GenerateOnly,
    [switch]$EmailOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $ConfigPath)) {
    throw "Config not found: $ConfigPath`nCopy event.config.template.json to event.config.json and fill in your credentials."
}
$config = Get-Content $ConfigPath | ConvertFrom-Json

. "$PSScriptRoot\scripts\Resolve-EventConfig.ps1"
$config = Resolve-EventConfig -Config $config

function Confirm-Step {
    param([string]$Description)
    if ($StepThrough) {
        Write-Host "`n$Description" -ForegroundColor Cyan
        $ans = Read-Host "Proceed? [Y/N]"
        return $ans -eq 'Y'
    }
    return $true
}

Write-Host "`n=== SQL Saturday Event Pipeline ===" -ForegroundColor Cyan
Write-Host "Event : $($config.event.name)"
Write-Host "Config: $ConfigPath"
if ($WhatIfPreference) { Write-Host "WHATIF MODE — no changes will be made`n" -ForegroundColor Yellow }

# ── Step 1: Import ────────────────────────────────────────────────────────────
if (-not $EmailOnly -and -not $GenerateOnly) {
    if (Confirm-Step "Step 1/3: Import attendees from EventBrite into SQLite") {
        Write-Host "`n[1/3] Importing attendees from EventBrite..." -ForegroundColor Green
        if (-not $WhatIfPreference) {
            & "$PSScriptRoot\scripts\Import-Attendees.ps1" -Config $config
        } else {
            Write-Host "  WHATIF: Would fetch attendees from EventBrite event $($config.eventbrite.eventId) and upsert to database."
        }
    } else {
        Write-Host "  Skipped import." -ForegroundColor DarkGray
    }
}

# ── Step 2: Generate SpeedPasses ─────────────────────────────────────────────
if (-not $ImportOnly -and -not $EmailOnly) {
    if (Confirm-Step "Step 2/3: Generate SpeedPass PDFs for new attendees") {
        Write-Host "`n[2/3] Generating SpeedPasses..." -ForegroundColor Green
        if (-not $WhatIfPreference) {
            & "$PSScriptRoot\scripts\Generate-SpeedPasses.ps1" -Config $config
        } else {
            Write-Host "  WHATIF: Would generate SpeedPass PDFs for attendees without one."
        }
    } else {
        Write-Host "  Skipped SpeedPass generation." -ForegroundColor DarkGray
    }
}

# ── Step 3: Send Emails ───────────────────────────────────────────────────────
if (-not $ImportOnly -and -not $GenerateOnly) {
    if (Confirm-Step "Step 3/3: Email SpeedPasses to attendees who haven't received one") {
        Write-Host "`n[3/3] Sending emails..." -ForegroundColor Green
        if ($PSCmdlet.ShouldProcess("attendees with unsent SpeedPasses", "Send emails")) {
            & "$PSScriptRoot\scripts\Send-SpeedPasses.ps1" -Config $config
        }
    } else {
        Write-Host "  Skipped email." -ForegroundColor DarkGray
    }
}

Write-Host "`n=== Pipeline complete ===" -ForegroundColor Green
