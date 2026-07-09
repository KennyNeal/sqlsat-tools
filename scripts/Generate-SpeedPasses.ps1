<#
.SYNOPSIS
    Generates personalized SpeedPass PDFs for attendees.
.DESCRIPTION
    Creates a PDF per attendee containing:
      - One admission ticket with order barcode QR
      - One raffle slip per tiered sponsor (each with vCard QR + sponsor logo)
      - One name tag with vCard QR

    Sponsor logos are fetched live from the website repo. Only attendees without
    an existing SpeedPass are processed unless -Force is specified.
.PARAMETER Config
    Parsed event.config.json object (passed by Update-Event.ps1).
.PARAMETER Force
    Regenerate SpeedPasses for all attendees, not just new ones.
.PARAMETER Email
    Generate SpeedPass for a single attendee by email address.
.EXAMPLE
    .\Generate-SpeedPasses.ps1 -Config $config
.EXAMPLE
    .\Generate-SpeedPasses.ps1 -Config $config -Email "jane.doe@example.com"
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config,
    [switch]$Force,
    [string]$Email
)

Import-Module PSSQLite

. "$PSScriptRoot\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config
. "$PSScriptRoot\Web-Helpers.ps1"
. "$PSScriptRoot\Badge-Helpers.ps1"

$dbPath      = Join-Path $PSScriptRoot ".." $Config.database.path
$outputDir   = Join-Path $PSScriptRoot ".." $Config.speedpass.outputDir

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

Import-QRCoder

# ── Sponsor logo fetch ────────────────────────────────────────────────────────

function Get-SponsorLogos {
    param([PSCustomObject]$Config, [string[]]$Tiers)

    $rawBase = Get-RawBase $Config
    $logos   = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($group in (Get-SponsorGroups -Config $Config)) {
        if ($group.tier -notin $Tiers) { continue }
        foreach ($sponsor in $group.sponsors) {
            Write-Host "    Downloading: $($sponsor.name)" -ForegroundColor DarkGray
            try {
                $img = Get-WebImage -Url "$rawBase/static/$($sponsor.logo)"
                $logos.Add(@{ Name = $sponsor.name; Base64 = $img.Base64; Mime = $img.Mime; Fit = $sponsor.logoFit })
            } catch {
                Write-Host "    Warning: could not download logo for $($sponsor.name): $_" -ForegroundColor Yellow
            }
        }
    }
    return $logos
}

# ── SpeedPass HTML builder ────────────────────────────────────────────────────

function New-SpeedPassHtml {
    param($Attendee, $SponsorLogos, $EventConfig)

    $fullName     = "$($Attendee.FirstName) $($Attendee.LastName)"
    $nameLastFirst = "$($Attendee.LastName), $($Attendee.FirstName)"
    $vcard        = New-VCard -FirstName $Attendee.FirstName -LastName $Attendee.LastName `
                              -Email $Attendee.Email -Company $Attendee.Company `
                              -JobTitle $Attendee.JobTitle -Website $Attendee.Website `
                              -TwitterHandle $Attendee.TwitterHandle
    $vCardQR  = New-QRBase64 -Data $vcard -PixelSize 30
    $orderQR  = New-QRBase64 -Data $Attendee.Barcode -PixelSize 30
    $hashtag  = $EventConfig.event.hashtag
    $ename    = $EventConfig.event.name

    $css = @"
<style>
body{margin:0;padding:0;font-family:Arial,sans-serif}
@page{size:Letter;margin:.35in}
.sheet{display:grid;grid-template-columns:repeat(2,3.5in);grid-template-rows:repeat(5,2in);padding:.25in}
.card{width:3.5in;height:2in;border:1px dashed #ccc;box-sizing:border-box;display:flex;
      flex-direction:row;align-items:stretch;position:relative;padding:.5in .25in .2in .25in;
      font-size:9pt;break-inside:avoid}
.left{display:flex;flex-direction:column;justify-content:flex-end;flex:1 1 0;
      padding-right:.2in;min-width:1.5in}
.ticket-banner{position:absolute;top:0;left:0;width:100%;text-align:center;
               font-size:10pt;font-weight:bold;padding-top:.05in}
.logo{width:1.5in;height:.6in;object-fit:contain;margin-bottom:.05in}
.qr-block{display:flex;flex-direction:column;align-items:center;width:1.4in;min-width:1.2in}
.qr{width:1.2in;height:1.2in;object-fit:contain;margin-bottom:.05in}
.raffle-name{font-weight:bold;margin-top:.05in;font-size:10pt;word-break:break-word;line-height:1.1}
.email-text{font-size:9pt;word-break:break-all;margin-top:.02in;line-height:1.1}
.card.nametag{flex-direction:column;justify-content:flex-start;align-items:center;
              text-align:center;font-size:10pt;padding:.1in}
.nametag-top{display:flex;justify-content:space-between;align-items:center;width:100%;margin-bottom:.1in}
.nametag .qr{width:1in;height:1in}
.fit-text{width:100%;height:1.2in;overflow:hidden;line-height:1.2;word-wrap:break-word}
</style>
"@

    $cards = ""

    # Admission ticket
    $cards += @"
<div class="card">
  <div class="ticket-banner">$hashtag</div>
  <div class="left">
    <strong>Admission Ticket</strong><br/>
    <strong>$nameLastFirst</strong><br/>
    Lunch: $($Attendee.LunchType)
    <div style="margin-top:auto;font-size:9pt;text-align:center">$ename</div>
  </div>
  <img src="data:image/png;base64,$orderQR" class="qr"/>
</div>
"@

    # Raffle slips — one per tiered sponsor
    foreach ($logo in $SponsorLogos) {
        $cards += @"
<div class="card">
  <div class="ticket-banner">$hashtag — Raffle Ticket</div>
  <div class="left">
    <img src="data:$($logo.Mime);base64,$($logo.Base64)" class="logo"/>
    <div class="raffle-name">$fullName</div>
    <div class="email-text">$($Attendee.Email)</div>
  </div>
  <div class="qr-block">
    <img src="data:image/png;base64,$vCardQR" class="qr"/>
  </div>
</div>
"@
    }

    # Name tag
    $cards += @"
<div class="card nametag">
  <div class="nametag-top">
    <span style="font-size:11pt;font-weight:bold">$hashtag</span>
    <img src="data:image/png;base64,$vCardQR" class="qr"/>
  </div>
  <div class="fit-text"><strong>$fullName</strong><br/>$($Attendee.JobTitle)<br/>$($Attendee.Company)</div>
</div>
"@

    return @"
<html><head>$css</head><body>
<div class="sheet">$cards</div>
<script>
window.addEventListener('DOMContentLoaded',()=>{
  document.querySelectorAll('.fit-text').forEach(el=>{
    let fs=20; el.style.fontSize=fs+'pt';
    while((el.scrollHeight>el.clientHeight)&&fs>6){fs-=.5;el.style.fontSize=fs+'pt';}
  });
});
</script>
</body></html>
"@
}

# ── Main ──────────────────────────────────────────────────────────────────────

Write-Host "Loading sponsor logos for raffle tiers: $($Config.speedpass.raffleTiers -join ', ')..." -ForegroundColor Cyan
$sponsorLogos = Get-SponsorLogos -Config $Config -Tiers $Config.speedpass.raffleTiers

# Query attendees
$emailFilter = if ($Email) { "a.Email = @Email" } else { $null }
$nullFilter  = if (-not $Force) { "p.SpeedPassGeneratedAt IS NULL" } else { $null }
$conditions  = @($emailFilter, $nullFilter) | Where-Object { $_ }
$whereClause = if ($conditions) { "WHERE " + ($conditions -join " AND ") } else { "" }

$query = @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle,
       a.LunchType, a.TwitterHandle, a.Website
FROM   Attendees a
LEFT   JOIN ProcessedAttendees p ON a.Barcode = p.Barcode
$whereClause
ORDER  BY a.LastName, a.FirstName
"@

$sqlParams = if ($Email) { @{ Email = $Email } } else { @{} }
$attendees = Invoke-SqliteQuery -DataSource $dbPath -Query $query -SqlParameters $sqlParams

if ($attendees.Count -eq 0) {
    Write-Host "No attendees need SpeedPass generation. Use -Force to regenerate all." -ForegroundColor Yellow
    return
}

Write-Host "Generating SpeedPasses for $($attendees.Count) attendee(s)..." -ForegroundColor Cyan
$generated = 0

foreach ($attendee in $attendees) {
    $safeName = "$($attendee.LastName)_$($attendee.FirstName)" -replace '[^\w]', ''
    $htmlPath = Join-Path $outputDir "$safeName.html"
    $pdfPath  = Join-Path $outputDir "$safeName.pdf"

    $html = New-SpeedPassHtml -Attendee $attendee -SponsorLogos $sponsorLogos -EventConfig $Config
    Set-Content -Path $htmlPath -Value $html -Encoding UTF8
    ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $pdfPath
    Remove-Item $htmlPath -ErrorAction SilentlyContinue

    Invoke-SqliteQuery -DataSource $dbPath -Query @"
INSERT INTO ProcessedAttendees (Barcode, SpeedPassPath, SpeedPassGeneratedAt)
VALUES (@Barcode, @Path, datetime('now'))
ON CONFLICT(Barcode) DO UPDATE SET SpeedPassPath=@Path, SpeedPassGeneratedAt=datetime('now')
"@ -SqlParameters @{ Barcode = $attendee.Barcode; Path = $pdfPath }

    Write-Host "  Generated: $safeName.pdf" -ForegroundColor Green
    $generated++
}

Write-Host "`nSpeedPass generation complete. Generated: $generated" -ForegroundColor Green
