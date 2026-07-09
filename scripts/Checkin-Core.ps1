<#
.SYNOPSIS
    Shared, non-interactive check-in logic used by Print-WalkinBadge.ps1 and
    Checkin-Menu.ps1: attendee lookup, walk-in creation, label rendering, and
    printing. No Read-Host/Write-Host prompts live here — callers own all
    interaction so this file can be reused by any front end.
.DESCRIPTION
    Requires Badge-Helpers.ps1 (New-VCard, New-QRBase64, ConvertTo-PdfViaEdge,
    Send-ToPrinter) and the PSSQLite module to already be loaded by the caller.
#>

function Find-Attendees {
    param([Parameter(Mandatory)][string]$DbPath, [string]$OrderId, [string]$Email)

    $query = @"
SELECT a.Barcode, a.OrderId, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle, a.LunchType, p.PrintedAt
FROM   Attendees a
LEFT JOIN PrintedBadges p ON p.Barcode = a.Barcode
WHERE  ($(if ($OrderId) { "a.OrderId = @OrderId" } else { "a.Email = @Email" }))
ORDER  BY a.LastName, a.FirstName
"@
    $params = if ($OrderId) { @{ OrderId = $OrderId } } else { @{ Email = $Email } }
    Invoke-SqliteQuery -DataSource $DbPath -Query $query -SqlParameters $params
}

function New-WalkinRecord {
    <#
    .SYNOPSIS
        Inserts a quick-add walk-in attendee. Pure function — caller has
        already collected and validated FirstName/LastName/Email.
    #>
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][string]$Email,
        [string]$Company,
        [string]$JobTitle
    )

    $barcode = "WALKIN-$([guid]::NewGuid().ToString())"
    Invoke-SqliteQuery -DataSource $DbPath -Query @"
INSERT INTO Attendees
    (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle, TicketType, AttendeeStatus)
VALUES
    (@Barcode, 'WALKIN', datetime('now'), @FirstName, @LastName, @Email, @Company, @JobTitle, 'Walk-in', 'attending')
"@ -SqlParameters @{
        Barcode   = $barcode
        FirstName = $FirstName
        LastName  = $LastName
        Email     = $Email
        Company   = $Company
        JobTitle  = $JobTitle
    }

    return [PSCustomObject]@{
        Barcode   = $barcode
        OrderId   = 'WALKIN'
        FirstName = $FirstName
        LastName  = $LastName
        Email     = $Email
        Company   = $Company
        JobTitle  = $JobTitle
        LunchType = $null
        PrintedAt = $null
    }
}

# ── Label HTML builder (2.4in x 3.9in landscape, no background art) ──────────

function New-LabelHtml {
    param($Attendee)

    $vcard = New-VCard -FirstName $Attendee.FirstName -LastName $Attendee.LastName `
                       -Email $Attendee.Email -Company $Attendee.Company -JobTitle $Attendee.JobTitle
    $qrB64 = New-QRBase64 -Data $vcard

    $titleHtml   = if ($Attendee.JobTitle) { "<div class=`"job-title`">$([System.Web.HttpUtility]::HtmlEncode($Attendee.JobTitle))</div>" } else { "" }
    $companyHtml = if ($Attendee.Company)  { "<div class=`"company`">$([System.Web.HttpUtility]::HtmlEncode($Attendee.Company))</div>"  } else { "" }
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

function Send-BadgeToPrinter {
    <#
    .SYNOPSIS
        Renders the label, prints it, and records the print. No reprint
        confirmation here — caller decides whether to call this at all.
    #>
    param(
        [Parameter(Mandatory)]$Attendee,
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$OutputDir,
        [Parameter(Mandatory)][string]$PrinterName
    )

    $htmlPath = Join-Path $OutputDir "walkin-label.html"
    $pdfPath  = Join-Path $OutputDir "walkin-label.pdf"
    $html = New-LabelHtml -Attendee $Attendee
    Set-Content -Path $htmlPath -Value $html -Encoding UTF8

    ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $pdfPath
    Remove-Item $htmlPath -ErrorAction SilentlyContinue

    Send-ToPrinter -PdfPath $pdfPath -PrinterName $PrinterName

    Invoke-SqliteQuery -DataSource $DbPath -Query @"
INSERT OR REPLACE INTO PrintedBadges (Barcode, PrintedAt, PrintedBy)
VALUES (@Barcode, datetime('now'), @PrintedBy)
"@ -SqlParameters @{ Barcode = $Attendee.Barcode; PrintedBy = $env:USERNAME }
}

function New-BadgePreview {
    <#
    .SYNOPSIS
        Renders the label to HTML or PDF for practice/testing. Does not touch
        the PrintedBadges table and does not print. Returns the file path.
    #>
    param(
        [Parameter(Mandatory)]$Attendee,
        [Parameter(Mandatory)][string]$OutputDir,
        [ValidateSet('Html', 'Pdf')]
        [string]$Format = 'Html'
    )

    $htmlPath = Join-Path $OutputDir "walkin-label-preview.html"
    $html = New-LabelHtml -Attendee $Attendee
    Set-Content -Path $htmlPath -Value $html -Encoding UTF8

    if ($Format -eq 'Pdf') {
        $pdfPath = Join-Path $OutputDir "walkin-label-preview.pdf"
        ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $pdfPath
        Remove-Item $htmlPath -ErrorAction SilentlyContinue
        return $pdfPath
    }

    return $htmlPath
}
