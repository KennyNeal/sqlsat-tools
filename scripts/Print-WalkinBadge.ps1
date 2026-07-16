<#
.SYNOPSIS
    Looks up an attendee by order number or email and prints their badge info
    directly to a Brother label printer, for day-of registration and walk-ins.
.DESCRIPTION
    Prints a single 2.4"x3.9" label (name, company/title, vCard QR — no
    background art) sized for continuous label tape, meant to be stuck onto a
    pre-printed blank badge template. Tracks what's already been printed in the
    PrintedBadges table so staff can see reprint status at the desk.

    Lunch type is printed when the attendee's registration came from an
    Eventbrite order (LunchType is set). Walk-ins registered on the spot at
    the desk don't collect lunch type, so nothing is shown for them.

    If no -OrderId/-Email is given, runs as an interactive loop: prompts for a
    lookup value, prints, and asks for the next one, so one launch can serve
    the whole registration desk. Pass -OrderId/-Email for a single one-off run.

    Unrecognized order numbers/emails trigger a quick-add prompt so true
    walk-ins (never in the Eventbrite import) can be registered on the spot.
.PARAMETER Config
    Parsed event.config.json object.
.PARAMETER OrderId
    Look up by Eventbrite order number. An order can cover multiple attendees;
    if so, you'll be prompted to pick one.
.PARAMETER Email
    Look up by attendee email address (unique per attendee).
.PARAMETER Printer
    Windows printer name. Defaults to Config.badge.walkinPrinter, or
    "Brother QL-820NWB".
.PARAMETER Force
    Reprint without prompting for confirmation, even if already printed.
.PARAMETER Preview
    Generate the label as HTML or PDF and open it instead of printing. Does
    not touch the PrintedBadges table. Useful for testing label layout
    without a printer or a real attendee record.
.PARAMETER PreviewFormat
    Format to generate when -Preview is set: "Html" (default) or "Pdf".
.EXAMPLE
    .\Print-WalkinBadge.ps1 -Config $config
.EXAMPLE
    .\Print-WalkinBadge.ps1 -Config $config -Email "jane.doe@example.com"
.EXAMPLE
    .\Print-WalkinBadge.ps1 -Config $config -Email "jane.doe@example.com" -Preview
.EXAMPLE
    .\Print-WalkinBadge.ps1 -Config $config -Preview -PreviewFormat Pdf
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config,
    [string]$OrderId,
    [string]$Email,
    [string]$Printer,
    [switch]$Force,
    [switch]$Preview,
    [ValidateSet('Html', 'Pdf')]
    [string]$PreviewFormat = 'Html'
)

$outputDir   = Join-Path $PSScriptRoot "..\output"
$printerName = if ($Printer) {
    $Printer
} elseif ($Config.PSObject.Properties['badge'] -and $Config.badge.walkinPrinter) {
    $Config.badge.walkinPrinter
} else {
    "Brother QL-820NWB"
}

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

. (Join-Path $PSScriptRoot "internal" "Badge-Helpers.ps1")
. (Join-Path $PSScriptRoot "internal" "Data-Access.ps1")
. (Join-Path $PSScriptRoot "internal" "Checkin-Core.ps1")

$dataContext = New-DataContext -Config $Config

# ── Interactive quick-add / multi-match picker ────────────────────────────
# New-LabelHtml lives in Checkin-Core.ps1; attendee lookup/insert live in
# Data-Access.ps1 so Checkin-Menu.ps1 can reuse them without Read-Host.

function New-WalkinAttendeeInteractive {
    Write-Host "  No matching attendee found. Quick-add a walk-in:" -ForegroundColor Yellow
    $firstName = ""
    while (-not $firstName) { $firstName = Read-Host "  First name" }
    $lastName = ""
    while (-not $lastName) { $lastName = Read-Host "  Last name" }
    $walkinEmail = ""
    while (-not $walkinEmail) { $walkinEmail = Read-Host "  Email" }
    $company  = Read-Host "  Company (optional)"
    $jobTitle = Read-Host "  Job title (optional)"

    $attendee = Add-Attendee -DataContext $dataContext -FirstName $firstName -LastName $lastName `
        -Email $walkinEmail -Company $company -JobTitle $jobTitle

    Write-Host "  Added locally only — Eventbrite has no API to create a real registration." -ForegroundColor Yellow
    Write-Host "  Add a free/comp order in Eventbrite (dashboard or Box Office app) when you get a chance," -ForegroundColor Yellow
    Write-Host "  or run .\scripts\List-UnsyncedWalkins.ps1 later to see everyone still pending." -ForegroundColor Yellow

    return $attendee
}

function Select-Attendee {
    param([array]$Matches)

    if ($Matches.Count -eq 0) { return New-WalkinAttendeeInteractive }
    if ($Matches.Count -eq 1) { return $Matches[0] }

    Write-Host "  Multiple attendees on this order:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $Matches.Count; $i++) {
        $m = $Matches[$i]
        $printed = if ($m.PrintedAt) { "printed $($m.PrintedAt)" } else { "not printed" }
        Write-Host "    [$($i + 1)] $($m.FirstName) $($m.LastName) <$($m.Email)> — $printed"
    }
    $choice = 0
    while ($choice -lt 1 -or $choice -gt $Matches.Count) {
        $choice = [int](Read-Host "  Pick a number (1-$($Matches.Count))")
    }
    return $Matches[$choice - 1]
}

function Invoke-PrintOne {
    param([string]$OrderIdArg, [string]$EmailArg)

    $matches = Get-AttendeesByOrderOrEmail -DataContext $dataContext -OrderId $OrderIdArg -Email $EmailArg
    $attendee = Select-Attendee -Matches $matches

    if ($Preview) {
        $path = New-BadgePreview -Attendee $attendee -OutputDir $outputDir -Format $PreviewFormat
        Write-Host "  Preview $PreviewFormat for $($attendee.FirstName) $($attendee.LastName): $path" -ForegroundColor Cyan
        Invoke-Item $path
        return
    }

    if ($attendee.PrintedAt -and -not $Force) {
        Write-Host "  Already printed for $($attendee.FirstName) $($attendee.LastName) at $($attendee.PrintedAt)." -ForegroundColor Yellow
        $confirm = Read-Host "  Reprint anyway? (y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Host "  Skipped." -ForegroundColor DarkGray
            return
        }
    }

    Send-BadgeToPrinter -Attendee $attendee -DataContext $dataContext -OutputDir $outputDir -PrinterName $printerName

    Write-Host "  Printed badge for $($attendee.FirstName) $($attendee.LastName)." -ForegroundColor Green
}

# ── Main ──────────────────────────────────────────────────────────────────────

if ($OrderId -or $Email) {
    Invoke-PrintOne -OrderIdArg $OrderId -EmailArg $Email
    return
}

Write-Host "Walk-in badge printing (printer: $printerName). Press Ctrl+C or enter blank to stop." -ForegroundColor Cyan
while ($true) {
    $lookup = Read-Host "`nOrder # or email"
    if (-not $lookup) { break }
    try {
        if ($lookup -match '@') {
            Invoke-PrintOne -EmailArg $lookup
        } else {
            Invoke-PrintOne -OrderIdArg $lookup
        }
    } catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
