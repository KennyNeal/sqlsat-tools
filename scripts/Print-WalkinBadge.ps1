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

Import-Module PSSQLite

$dbPath      = Join-Path $PSScriptRoot ".." $Config.database.path
$libPath     = Join-Path $PSScriptRoot "..\lib\QRCoder.dll"
$outputDir   = Join-Path $PSScriptRoot "..\output"
$printerName = if ($Printer) {
    $Printer
} elseif ($Config.PSObject.Properties['badge'] -and $Config.badge.walkinPrinter) {
    $Config.badge.walkinPrinter
} else {
    "Brother QL-820NWB"
}

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not (Test-Path $libPath)) { throw "QRCoder.dll not found at $libPath." }
Add-Type -Path $libPath

. (Join-Path $PSScriptRoot "Badge-Helpers.ps1")

# ── Label HTML builder (2.4in x 3.9in landscape, no background art) ──────────

function New-LabelHtml {
    param($Attendee)

    $vcard = New-VCard -FirstName $Attendee.FirstName -LastName $Attendee.LastName `
                       -Email $Attendee.Email -Company $Attendee.Company -JobTitle $Attendee.JobTitle
    $qrB64 = New-QRBase64 -Data $vcard

    $titleHtml   = if ($Attendee.JobTitle) { "<div class=`"job-title`">$([System.Web.HttpUtility]::HtmlEncode($Attendee.JobTitle))</div>" } else { "" }
    $companyHtml = if ($Attendee.Company)  { "<div class=`"company`">$([System.Web.HttpUtility]::HtmlEncode($Attendee.Company))</div>" }  else { "" }
    # Only Eventbrite orders carry a LunchType; walk-ins registered on the spot don't collect one.
    $lunchHtml   = if ($Attendee.OrderId -ne 'WALKIN' -and $Attendee.LunchType) { "<div class=`"lunch-type`">$([System.Web.HttpUtility]::HtmlEncode($Attendee.LunchType))</div>" } else { "" }

    return @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
@page { size: 3.9in 2.4in; margin: 0; }
body { width: 3.9in; height: 2.4in; }
.label {
    width: 3.9in;
    height: 2.4in;
    position: relative;
    font-family: Arial, sans-serif;
    overflow: hidden;
}
.info-col {
    position: absolute;
    top: 0.15in; left: 0.2in; right: 1.55in; bottom: 0.15in;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    min-width: 0;
}
.first-name {
    font-size: 30pt;
    font-weight: bold;
    line-height: 1.0;
    color: #000;
    white-space: nowrap;
    overflow: hidden;
}
.last-name {
    font-size: 18pt;
    font-weight: bold;
    line-height: 1.1;
    color: #000;
    margin-top: 0.03in;
}
.job-title { font-size: 10pt; color: #333; margin-top: 0.08in; line-height: 1.2; }
.company   { font-size: 10pt; color: #555; line-height: 1.2; }
.lunch-type {
    font-size: 10pt;
    color: #cc2200;
    margin-top: 0.08in;
    font-weight: 500;
}
.qr {
    position: absolute;
    top: 0.15in; right: 0.15in;
    width: 1.4in; height: 1.4in;
}
</style>
</head>
<body>
<div class="label">
  <div class="info-col">
    <div class="first-name">$([System.Web.HttpUtility]::HtmlEncode($Attendee.FirstName))</div>
    <div class="last-name">$([System.Web.HttpUtility]::HtmlEncode($Attendee.LastName))</div>
    $titleHtml
    $companyHtml
    $lunchHtml
  </div>
  <img class="qr" src="data:image/png;base64,$qrB64"/>
</div>
<script>
window.addEventListener('DOMContentLoaded', function() {
  var el = document.querySelector('.first-name');
  var fs = 30;
  el.style.fontSize = fs + 'pt';
  while (el.scrollWidth > el.offsetWidth && fs > 12) {
    fs -= 0.5;
    el.style.fontSize = fs + 'pt';
  }
});
</script>
</body>
</html>
"@
}

# ── Lookup ─────────────────────────────────────────────────────────────────

function Find-Attendees {
    param([string]$OrderId, [string]$Email)

    $query = @"
SELECT a.Barcode, a.OrderId, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle, a.LunchType, p.PrintedAt
FROM   Attendees a
LEFT JOIN PrintedBadges p ON p.Barcode = a.Barcode
WHERE  ($(if ($OrderId) { "a.OrderId = @OrderId" } else { "a.Email = @Email" }))
ORDER  BY a.LastName, a.FirstName
"@
    $params = if ($OrderId) { @{ OrderId = $OrderId } } else { @{ Email = $Email } }
    Invoke-SqliteQuery -DataSource $dbPath -Query $query -SqlParameters $params
}

function New-WalkinAttendee {
    Write-Host "  No matching attendee found. Quick-add a walk-in:" -ForegroundColor Yellow
    $firstName = ""
    while (-not $firstName) { $firstName = Read-Host "  First name" }
    $lastName = ""
    while (-not $lastName) { $lastName = Read-Host "  Last name" }
    $walkinEmail = ""
    while (-not $walkinEmail) { $walkinEmail = Read-Host "  Email" }
    $company  = Read-Host "  Company (optional)"
    $jobTitle = Read-Host "  Job title (optional)"

    $barcode = "WALKIN-$([guid]::NewGuid().ToString())"
    Invoke-SqliteQuery -DataSource $dbPath -Query @"
INSERT INTO Attendees
    (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle, TicketType, AttendeeStatus)
VALUES
    (@Barcode, 'WALKIN', datetime('now'), @FirstName, @LastName, @Email, @Company, @JobTitle, 'Walk-in', 'attending')
"@ -SqlParameters @{
        Barcode   = $barcode
        FirstName = $firstName
        LastName  = $lastName
        Email     = $walkinEmail
        Company   = $company
        JobTitle  = $jobTitle
    }

    Write-Host "  Added locally only — Eventbrite has no API to create a real registration." -ForegroundColor Yellow
    Write-Host "  Add a free/comp order in Eventbrite (dashboard or Box Office app) when you get a chance," -ForegroundColor Yellow
    Write-Host "  or run .\scripts\List-UnsyncedWalkins.ps1 later to see everyone still pending." -ForegroundColor Yellow

    return [PSCustomObject]@{
        Barcode   = $barcode
        OrderId   = 'WALKIN'
        FirstName = $firstName
        LastName  = $lastName
        Email     = $walkinEmail
        Company   = $company
        JobTitle  = $jobTitle
        LunchType = $null
        PrintedAt = $null
    }
}

function Select-Attendee {
    param([array]$Matches)

    if ($Matches.Count -eq 0) { return New-WalkinAttendee }
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

    $matches = Find-Attendees -OrderId $OrderIdArg -Email $EmailArg
    $attendee = Select-Attendee -Matches $matches

    if ($Preview) {
        $htmlPath = Join-Path $outputDir "walkin-label-preview.html"
        $html = New-LabelHtml -Attendee $attendee
        Set-Content -Path $htmlPath -Value $html -Encoding UTF8

        if ($PreviewFormat -eq 'Pdf') {
            $pdfPath = Join-Path $outputDir "walkin-label-preview.pdf"
            ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $pdfPath
            Remove-Item $htmlPath -ErrorAction SilentlyContinue
            Write-Host "  Preview PDF for $($attendee.FirstName) $($attendee.LastName): $pdfPath" -ForegroundColor Cyan
            Invoke-Item $pdfPath
        } else {
            Write-Host "  Preview HTML for $($attendee.FirstName) $($attendee.LastName): $htmlPath" -ForegroundColor Cyan
            Invoke-Item $htmlPath
        }
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

    $htmlPath = Join-Path $outputDir "walkin-label.html"
    $pdfPath  = Join-Path $outputDir "walkin-label.pdf"
    $html = New-LabelHtml -Attendee $attendee
    Set-Content -Path $htmlPath -Value $html -Encoding UTF8

    ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $pdfPath
    Remove-Item $htmlPath -ErrorAction SilentlyContinue

    Send-ToPrinter -PdfPath $pdfPath -PrinterName $printerName

    Invoke-SqliteQuery -DataSource $dbPath -Query @"
INSERT OR REPLACE INTO PrintedBadges (Barcode, PrintedAt, PrintedBy)
VALUES (@Barcode, datetime('now'), @PrintedBy)
"@ -SqlParameters @{ Barcode = $attendee.Barcode; PrintedBy = $env:USERNAME }

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
