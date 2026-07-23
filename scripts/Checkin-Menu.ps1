<#
.SYNOPSIS
    Menu-driven check-in desk console app — the one thing to run on event day.
.DESCRIPTION
    Wraps Print-WalkinBadge.ps1's lookup/print/quick-add flow, Import-Attendees.ps1,
    and List-UnsyncedWalkins.ps1 behind a big numbered menu with plain-language
    prompts and colored feedback, so someone who's never touched PowerShell
    (e.g. a kid helping at the desk) can run check-in unsupervised.
.PARAMETER ConfigPath
    Path to event.config.json. Defaults to the repo-root copy.
.EXAMPLE
    .\Checkin-Menu.ps1
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot ".." "event.config.json")
)

$ErrorActionPreference = 'Stop'

function Write-Banner {
    try { Clear-Host } catch { }
    Write-Host ""
    Write-Host "  ============================" -ForegroundColor Cyan
    Write-Host "     SQL Saturday Check-In" -ForegroundColor Cyan
    Write-Host "  ============================" -ForegroundColor Cyan
    Write-Host ""
}

function Wait-ForEnter {
    Write-Host ""
    Read-Host "Press Enter to go back to the menu" | Out-Null
}

# ── Startup ────────────────────────────────────────────────────────────────

try {
    if (-not (Test-Path $ConfigPath)) {
        throw "Can't find the event settings file at $ConfigPath."
    }
    $config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json

    $scriptsDir = $PSScriptRoot
    $repoRoot   = Join-Path $scriptsDir ".."
    $dbPath     = Join-Path $repoRoot $config.database.path
    $outputDir  = Join-Path $repoRoot "output"
    $libPath    = Join-Path $scriptsDir "..\lib\QRCoder.dll"

    if (-not (Test-Path $dbPath)) {
        throw "Can't find the event database at $dbPath. Ask a grown-up to run scripts\setup\Initialize-Database.ps1 first."
    }
    if (-not (Test-Path $libPath)) {
        throw "Can't find QRCoder.dll at $libPath. Ask a grown-up — a file is missing from the project."
    }
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
    Add-Type -Path $libPath

    . (Join-Path $scriptsDir "internal" "Badge-Helpers.ps1")
    . (Join-Path $scriptsDir "internal" "Data-Access.ps1")
    . (Join-Path $scriptsDir "internal" "Checkin-Core.ps1")

    $dataContext = New-DataContext -Config $config

    if ($dataContext.AzureEnabled) {
        Write-Host "  Checking Azure SQL and pulling the latest shared data..." -ForegroundColor DarkGray
        if (Sync-FromAzure -DataContext $dataContext) {
            Write-Host "  Up to date with Azure." -ForegroundColor DarkGray
        } else {
            Write-Host "  Azure unreachable — starting with the local cache. This desk will queue" -ForegroundColor Yellow
            Write-Host "  writes and catch up automatically once the connection is back." -ForegroundColor Yellow
        }
    }

    # Cheap insurance: a hot backup at startup, then opportunistically every
    # few minutes in the main loop (see below). Restores the point Data-Access's
    # auto-recovery falls back to if event.db corrupts again this desk hasn't
    # synced to Azure yet.
    Backup-LocalDatabase -DataContext $dataContext | Out-Null

    $printerName = if ($config.PSObject.Properties['badge'] -and $config.badge.walkinPrinter) {
        $config.badge.walkinPrinter
    } else {
        "Brother QL-820NWB"
    }
} catch {
    Write-Host ""
    Write-Host "  Something's wrong before check-in can start:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Press Enter to close." -ForegroundColor Yellow
    Read-Host | Out-Null
    return
}

# ── Shared attendee lookup + quick-add flow ──────────────────────────────────

function Read-RequiredText {
    param([string]$Prompt)
    $value = ""
    while (-not $value) { $value = Read-Host $Prompt }
    return $value
}

function New-WalkinInteractive {
    param([switch]$PracticeMode)

    Write-Host ""
    Write-Host "  No match found. Add them as a walk-in:" -ForegroundColor Yellow
    $firstName = Read-RequiredText "  First name"
    $lastName  = Read-RequiredText "  Last name"
    $walkinEmail = Read-RequiredText "  Email"
    $company   = Read-Host "  Company (optional)"
    $jobTitle  = Read-Host "  Job title (optional)"

    if ($PracticeMode) {
        Write-Host "  (practice) Not actually saved — this is just for the preview." -ForegroundColor DarkYellow
        return [PSCustomObject]@{
            Barcode = "PRACTICE-$([guid]::NewGuid())"; OrderId = 'WALKIN'
            FirstName = $firstName; LastName = $lastName; Email = $walkinEmail
            Company = $company; JobTitle = $jobTitle; LunchType = $null; PrintedAt = $null
        }
    }

    $attendee = Add-Attendee -DataContext $dataContext -FirstName $firstName -LastName $lastName `
        -Email $walkinEmail -Company $company -JobTitle $jobTitle

    Write-Host "  Added! (This is a local walk-in — a grown-up will need to add a real" -ForegroundColor DarkYellow
    Write-Host "  order in Eventbrite for them later. Use menu option 4 to see the list.)" -ForegroundColor DarkYellow
    return $attendee
}

function Select-AttendeeFromMatches {
    param([array]$Matches, [switch]$PracticeMode)

    if ($Matches.Count -eq 0) { return New-WalkinInteractive -PracticeMode:$PracticeMode }
    if ($Matches.Count -eq 1) { return $Matches[0] }

    Write-Host ""
    Write-Host "  Found more than one person on that order:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Matches.Count; $i++) {
        $m = $Matches[$i]
        $printed = if ($m.PrintedAt) { "already printed" } else { "not printed yet" }
        Write-Host "    [$($i + 1)] $($m.FirstName) $($m.LastName) <$($m.Email)> — $printed"
    }
    $choice = 0
    while ($choice -lt 1 -or $choice -gt $Matches.Count) {
        $raw = Read-Host "  Type a number (1-$($Matches.Count))"
        [void][int]::TryParse($raw, [ref]$choice)
    }
    return $Matches[$choice - 1]
}

function Invoke-CheckinLoop {
    param([switch]$PracticeMode)

    Write-Banner
    if ($PracticeMode) {
        Write-Host "  PRACTICE MODE — nothing will print for real, no records saved as printed." -ForegroundColor Magenta
    }
    Write-Host "  Scan or type an order number or email address." -ForegroundColor White
    Write-Host "  Leave it blank and press Enter to go back to the menu." -ForegroundColor DarkGray
    Write-Host ""

    while ($true) {
        $lookup = Read-Host "Order # or email"
        if (-not $lookup) { return }

        try {
            $matches = if ($lookup -match '@') {
                Get-AttendeesByOrderOrEmail -DataContext $dataContext -Email $lookup
            } else {
                Get-AttendeesByOrderOrEmail -DataContext $dataContext -OrderId $lookup
            }
            $attendee = Select-AttendeeFromMatches -Matches $matches -PracticeMode:$PracticeMode

            if ($PracticeMode) {
                $path = New-BadgePreview -Attendee $attendee -OutputDir $outputDir -Format 'Html'
                Write-Host "  (practice) Opening what $($attendee.FirstName)'s badge would look like..." -ForegroundColor Cyan
                Invoke-Item $path
                continue
            }

            if ($attendee.PrintedAt) {
                Write-Host "  Heads up: already printed for $($attendee.FirstName) $($attendee.LastName) at $($attendee.PrintedAt)." -ForegroundColor Yellow
                $confirm = Read-Host "  Print again anyway? (y/N)"
                if ($confirm -notmatch '^[Yy]') {
                    Write-Host "  Skipped." -ForegroundColor DarkGray
                    continue
                }
            }

            Write-Host "  Printing..." -ForegroundColor Cyan
            Send-BadgeToPrinter -Attendee $attendee -DataContext $dataContext -OutputDir $outputDir -PrinterName $printerName
            Write-Host "  Printed badge for $($attendee.FirstName) $($attendee.LastName)." -ForegroundColor Green
        } catch {
            Write-Host "  Something went wrong: $($_.Exception.Message)" -ForegroundColor Red
        }
        Write-Host ""
    }
}

function Invoke-EventbriteSync {
    Write-Banner
    Write-Host "  Pulling the latest registrations from Eventbrite. Needs the internet," -ForegroundColor White
    Write-Host "  can take a little while..." -ForegroundColor White
    Write-Host ""

    $maxAttempts = 3
    $succeeded = $false
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            & (Join-Path $scriptsDir "Import-Attendees.ps1") -Config $config
            $succeeded = $true
            break
        } catch {
            if ($attempt -lt $maxAttempts) {
                Write-Host "  Connection hiccup: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "  Trying again ($($attempt + 1) of $maxAttempts)..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
            } else {
                Write-Host "  Sync failed after $maxAttempts tries: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "  Check the wifi and try again from the menu." -ForegroundColor Yellow
            }
        }
    }

    # The import itself only touched local SQLite (fast) and queued the same
    # rows for Azure — push that queue now so other desks pick up the new
    # registrations without waiting for someone to notice and hit option 5.
    if ($succeeded -and $dataContext.AzureEnabled) {
        if (Test-AzureReachable -DataContext $dataContext) {
            Write-Host "  Sending the new registrations to Azure..." -ForegroundColor Cyan
            $drained = Sync-PendingWrites -DataContext $dataContext
            Write-Host "  Sent $drained update(s) to Azure." -ForegroundColor Green
        } else {
            Write-Host "  Azure's unreachable right now — the new registrations are queued and" -ForegroundColor Yellow
            Write-Host "  will send automatically once the connection is back." -ForegroundColor Yellow
        }
    }

    Wait-ForEnter
}

function Show-UnsyncedWalkins {
    Write-Banner
    try {
        & (Join-Path $scriptsDir "List-UnsyncedWalkins.ps1") -Config $config
    } catch {
        Write-Host "  Couldn't load the list: $($_.Exception.Message)" -ForegroundColor Red
    }
    Wait-ForEnter
}

function Invoke-AzureSyncNow {
    Write-Banner
    if (-not $dataContext.AzureEnabled) {
        Write-Host "  Azure sync isn't turned on for this event (azure.enabled is false)." -ForegroundColor Yellow
        Wait-ForEnter
        return
    }

    Write-Host "  Checking Azure SQL..." -ForegroundColor Cyan
    if (-not (Test-AzureReachable -DataContext $dataContext)) {
        Write-Host "  Still can't reach Azure. Anything printed/added here is safely queued and" -ForegroundColor Yellow
        Write-Host "  will sync automatically once the connection is back." -ForegroundColor Yellow
        Wait-ForEnter
        return
    }

    $drained = Sync-PendingWrites -DataContext $dataContext
    if ($drained -gt 0) { Write-Host "  Sent $drained queued update(s) to Azure." -ForegroundColor Green }

    if (Sync-FromAzure -DataContext $dataContext) {
        Write-Host "  Pulled the latest from Azure — this desk now matches every other desk." -ForegroundColor Green
    } else {
        Write-Host "  Connected briefly but the pull failed — try again in a moment." -ForegroundColor Yellow
    }
    Wait-ForEnter
}

# ── Main menu ────────────────────────────────────────────────────────────────

while ($true) {
    # Opportunistic: if Azure just came back, quietly catch this desk up
    # before showing the menu, so a volunteer doesn't have to remember to.
    if ($dataContext.AzureEnabled -and -not $dataContext.AzureReachable -and (Test-AzureReachable -DataContext $dataContext)) {
        Sync-PendingWrites -DataContext $dataContext | Out-Null
        Sync-FromAzure -DataContext $dataContext | Out-Null
    }
    Backup-LocalDatabaseIfDue -DataContext $dataContext -IntervalMinutes 10 | Out-Null

    Write-Banner
    Write-Host "  [1] Check in an attendee" -ForegroundColor White
    Write-Host "  [2] Practice mode (no printing)" -ForegroundColor White
    Write-Host "  [3] Sync new registrations from Eventbrite" -ForegroundColor White
    Write-Host "  [4] Show walk-ins not yet in Eventbrite" -ForegroundColor White
    Write-Host "  [5] Sync now with Azure (other desks)" -ForegroundColor White
    Write-Host "  [Q] Quit" -ForegroundColor White
    Write-Host ""
    $choice = Read-Host "Pick an option"

    switch ($choice.Trim().ToUpper()) {
        '1' { Invoke-CheckinLoop }
        '2' { Invoke-CheckinLoop -PracticeMode }
        '3' { Invoke-EventbriteSync }
        '4' { Show-UnsyncedWalkins }
        '5' { Invoke-AzureSyncNow }
        'Q' { Write-Host ""; Write-Host "  Bye! Thanks for helping with check-in." -ForegroundColor Cyan; return }
        default { }
    }
}
