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
if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
    Write-Host "Installing powershell-yaml module..." -ForegroundColor Cyan
    Install-Module -Name powershell-yaml -Scope CurrentUser -Force
}
Import-Module powershell-yaml

$dbPath      = Join-Path $PSScriptRoot ".." $Config.database.path
$outputDir   = Join-Path $PSScriptRoot ".." $Config.speedpass.outputDir
$libPath     = Join-Path $PSScriptRoot "..\lib\QRCoder.dll"
$edgePaths   = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
)

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# Load QRCoder
if (-not (Test-Path $libPath)) { throw "QRCoder.dll not found at $libPath. Copy it from the old repo's lib\ folder." }
Add-Type -Path $libPath

# ── Sponsor logo fetch ────────────────────────────────────────────────────────

function Get-SponsorLogos {
    param([PSCustomObject]$Config, [string[]]$Tiers)

    $repo   = $Config.websiteRepo
    $rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
    $yamlUrl = "$rawBase/data/sponsors/$($repo.sponsorDataFile).yaml"

    Write-Host "  Fetching sponsor data from $yamlUrl" -ForegroundColor Cyan
    $yaml = (Invoke-RestMethod -Uri $yamlUrl -Method Get)
    $data = ConvertFrom-Yaml $yaml

    $logos = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($group in $data.groups) {
        if ($group.tier -notin $Tiers) { continue }
        foreach ($sponsor in $group.sponsors) {
            $logoUrl = "$rawBase/static/$($sponsor.logo)"
            Write-Host "    Downloading: $($sponsor.name)" -ForegroundColor DarkGray
            try {
                $bytes    = (Invoke-WebRequest -Uri $logoUrl -UseBasicParsing).Content
                $b64      = [Convert]::ToBase64String($bytes)
                $ext      = [System.IO.Path]::GetExtension($sponsor.logo).TrimStart('.')
                $mimeType = switch ($ext) {
                    'svg'  { 'image/svg+xml' }
                    'jpg'  { 'image/jpeg' }
                    'jpeg' { 'image/jpeg' }
                    default { "image/$ext" }
                }
                $logos.Add(@{ Name = $sponsor.name; Base64 = $b64; Mime = $mimeType; Fit = $sponsor.logoFit })
            } catch {
                Write-Host "    Warning: could not download logo for $($sponsor.name): $_" -ForegroundColor Yellow
            }
        }
    }
    return $logos
}

# ── vCard builder ─────────────────────────────────────────────────────────────

function New-VCard {
    param($FirstName, $LastName, $Email, $Company, $JobTitle, $Website, $TwitterHandle)
    $twitter = if ($TwitterHandle -and -not $TwitterHandle.StartsWith('@')) { "@$TwitterHandle" } else { $TwitterHandle }
    $lines   = @(
        "BEGIN:VCARD",
        "VERSION:3.0",
        "N:$LastName;$FirstName",
        "FN:$FirstName $LastName"
    )
    if ($Company)     { $lines += "ORG:$Company" }
    if ($JobTitle)    { $lines += "TITLE:$JobTitle" }
    if ($Email)       { $lines += "EMAIL:$Email" }
    if ($Website)     { $lines += "URL:$Website" }
    if ($twitter)     { $lines += "X-SOCIALPROFILE;TYPE=twitter:$twitter" }
    $lines += "END:VCARD"
    return $lines -join "`r`n"
}

# ── QR code generator ─────────────────────────────────────────────────────────

function New-QRBase64 {
    param([string]$Data, [int]$PixelSize = 30)
    $gen    = New-Object QRCoder.QRCodeGenerator
    $qrData = $gen.CreateQrCode($Data, [QRCoder.QRCodeGenerator+ECCLevel]::L)
    $qr     = New-Object QRCoder.QRCode($qrData)
    $bmp    = $qr.GetGraphic($PixelSize)
    $ms     = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $b64    = [Convert]::ToBase64String($ms.ToArray())
    $ms.Dispose(); $bmp.Dispose(); $qr.Dispose(); $gen.Dispose()
    return $b64
}

# ── Edge PDF renderer ─────────────────────────────────────────────────────────

function ConvertTo-Pdf {
    param([string]$HtmlPath, [string]$PdfPath)
    $edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $edge) { throw "Microsoft Edge not found. Install Edge or update \$edgePaths." }
    $null = & $edge --headless=new --print-to-pdf="$PdfPath" --no-margins "file:///$HtmlPath" --disable-gpu --disable-extensions --no-pdf-header-footer 2>&1
    Start-Sleep -Seconds 4
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
    $vCardQR  = New-QRBase64 -Data $vcard
    $orderQR  = New-QRBase64 -Data $Attendee.Barcode
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
$whereClause = if ($Email) {
    "WHERE a.Email = @Email AND p.SpeedPassGeneratedAt IS NULL"
} elseif ($Force) {
    ""
} else {
    "WHERE p.SpeedPassGeneratedAt IS NULL"
}

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
    ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $pdfPath
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
