<#
.SYNOPSIS
    Generates Avery 5392-compatible (4"×3") ID badge sheets as a single PDF.
.DESCRIPTION
    Lays out 6 badges per letter sheet (2 columns × 3 rows), each badge using
    the event background image with attendee name, company, title, lunch type,
    and a vCard QR code overlaid. Output is a single PDF for batch printing.
.PARAMETER Config
    Parsed event.config.json object.
.PARAMETER Force
    Regenerate badges even for attendees already processed.
.PARAMETER Email
    Generate a badge sheet for a single attendee by email address.
.PARAMETER BackgroundImage
    Override path to the badge background PNG (default: Config.badge.backgroundImage).
.EXAMPLE
    .\Generate-NameTag.ps1 -Config $config
.EXAMPLE
    .\Generate-NameTag.ps1 -Config $config -Email "jane.doe@example.com"
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config,
    [switch]$Force,
    [string]$Email,
    [string]$BackgroundImage = ""
)

Import-Module PSSQLite

$dbPath    = Join-Path $PSScriptRoot ".." $Config.database.path
$outputFile = if ($Config.PSObject.Properties['badge'] -and $Config.badge.outputFile) {
    Join-Path $PSScriptRoot ".." $Config.badge.outputFile
} else {
    Join-Path $PSScriptRoot "..\output\badges.pdf"
}
$outputDir = Split-Path $outputFile -Parent
$libPath   = Join-Path $PSScriptRoot "..\lib\QRCoder.dll"
$edgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
)

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Background image ──────────────────────────────────────────────────────────

$bgImagePath = if ($BackgroundImage) {
    $BackgroundImage
} elseif ($Config.PSObject.Properties['badge'] -and $Config.badge.backgroundImage) {
    Join-Path $PSScriptRoot ".." $Config.badge.backgroundImage
} else {
    Join-Path $PSScriptRoot "..\assets\badge-background.png"
}

if (-not (Test-Path $bgImagePath)) {
    throw "Badge background image not found: $bgImagePath`nPlace your blank badge template at that path (PNG or JPG)."
}

$bgBytes = [System.IO.File]::ReadAllBytes($bgImagePath)
$bgB64   = [Convert]::ToBase64String($bgBytes)
$bgExt   = [System.IO.Path]::GetExtension($bgImagePath).TrimStart('.')
$bgMime  = switch ($bgExt) {
    'jpg'  { 'image/jpeg' }
    'jpeg' { 'image/jpeg' }
    default { 'image/png' }
}
Write-Host "  Loaded badge background: $bgImagePath" -ForegroundColor DarkGray

# ── QRCoder ───────────────────────────────────────────────────────────────────

if (-not (Test-Path $libPath)) { throw "QRCoder.dll not found at $libPath." }
Add-Type -Path $libPath

function New-VCard {
    param($FirstName, $LastName, $Email, $Company, $JobTitle)
    $lines = @(
        "BEGIN:VCARD",
        "VERSION:3.0",
        "N:$LastName;$FirstName",
        "FN:$FirstName $LastName"
    )
    if ($Company)  { $lines += "ORG:$Company" }
    if ($JobTitle) { $lines += "TITLE:$JobTitle" }
    if ($Email)    { $lines += "EMAIL:$Email" }
    $lines += "END:VCARD"
    return $lines -join "`r`n"
}

function New-QRBase64 {
    param([string]$Data, [int]$PixelSize = 20)
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
    if (-not $edge) { throw "Microsoft Edge not found." }
    $null = & $edge --headless=new --print-to-pdf="$PdfPath" --no-margins "file:///$HtmlPath" --disable-gpu --disable-extensions --no-pdf-header-footer 2>&1
    Start-Sleep -Seconds 5
}

# ── Badge HTML builder ────────────────────────────────────────────────────────

function New-BadgeHtml {
    param([array]$Attendees, [string]$BgDataUri)

    $css = @"
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { background: white; }
@page { size: 8.5in 11in; margin: 0; }

.sheet {
    width: 8.5in;
    height: 11in;
    display: grid;
    grid-template-columns: 4in 4in;
    grid-template-rows: 3in 3in 3in;
    column-gap: 0.2in;
    row-gap: 0.25in;
    padding: 0.75in 0.15in;
    break-after: page;
}
.sheet:last-child { break-after: avoid; }

.badge {
    position: relative;
    width: 4in;
    height: 3in;
    overflow: hidden;
    break-inside: avoid;
}
.badge-bg {
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 100%;
    display: block;
}
.badge-body {
    position: absolute;
    top: 0.52in;
    left: 0.18in;
    right: 0.18in;
    bottom: 0.46in;
    display: flex;
    flex-direction: row;
    align-items: stretch;
    gap: 0.1in;
    font-family: Arial, sans-serif;
}
.info-col {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    min-width: 0;
}
.first-name {
    font-size: 32pt;
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
    margin-top: 0.02in;
}
.job-title {
    font-size: 10pt;
    color: #333;
    margin-top: 0.06in;
    line-height: 1.2;
}
.company {
    font-size: 10pt;
    color: #555;
    line-height: 1.2;
}
.lunch-type {
    font-size: 10pt;
    color: #cc2200;
    margin-top: auto;
    padding-bottom: 0.04in;
    font-weight: 500;
}
.qr-col {
    width: 1.05in;
    display: flex;
    flex-direction: column;
    justify-content: flex-end;
    flex-shrink: 0;
}
.qr {
    width: 1.05in;
    height: 1.05in;
    display: block;
}
</style>
"@

    $sheets = ""
    $chunks = [System.Collections.Generic.List[object]]::new()
    $chunk  = [System.Collections.Generic.List[object]]::new()

    foreach ($a in $Attendees) {
        $chunk.Add($a)
        if ($chunk.Count -eq 6) {
            $chunks.Add($chunk.ToArray())
            $chunk = [System.Collections.Generic.List[object]]::new()
        }
    }
    if ($chunk.Count -gt 0) { $chunks.Add($chunk.ToArray()) }

    foreach ($batch in $chunks) {
        $cards = ""
        foreach ($a in $batch) {
            $vcard = New-VCard -FirstName $a.FirstName -LastName $a.LastName `
                               -Email $a.Email -Company $a.Company -JobTitle $a.JobTitle
            $qrB64 = New-QRBase64 -Data $vcard

            $titleHtml   = if ($a.JobTitle) { "<div class=`"job-title`">$([System.Web.HttpUtility]::HtmlEncode($a.JobTitle))</div>" } else { "" }
            $companyHtml = if ($a.Company)  { "<div class=`"company`">$([System.Web.HttpUtility]::HtmlEncode($a.Company))</div>" }  else { "" }
            $lunchHtml   = if ($a.LunchType) { "<div class=`"lunch-type`">$([System.Web.HttpUtility]::HtmlEncode($a.LunchType))</div>" } else { "" }

            $cards += @"
<div class="badge">
  <img class="badge-bg" src="$BgDataUri"/>
  <div class="badge-body">
    <div class="info-col">
      <div class="first-name">$([System.Web.HttpUtility]::HtmlEncode($a.FirstName))</div>
      <div class="last-name">$([System.Web.HttpUtility]::HtmlEncode($a.LastName))</div>
      $titleHtml
      $companyHtml
      $lunchHtml
    </div>
    <div class="qr-col">
      <img class="qr" src="data:image/png;base64,$qrB64"/>
    </div>
  </div>
</div>
"@
        }
        $sheets += "<div class=`"sheet`">$cards</div>`n"
    }

    return @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
$css
</head>
<body>
$sheets
<script>
window.addEventListener('DOMContentLoaded', function() {
  document.querySelectorAll('.first-name').forEach(function(el) {
    var fs = 32;
    el.style.fontSize = fs + 'pt';
    while (el.scrollWidth > el.offsetWidth && fs > 12) {
      fs -= 0.5;
      el.style.fontSize = fs + 'pt';
    }
  });
});
</script>
</body>
</html>
"@
}

# ── Main ──────────────────────────────────────────────────────────────────────

# Query attendees
$emailFilter = if ($Email) { "Email = @Email" } else { $null }
$conditions  = @($emailFilter) | Where-Object { $_ }
$whereClause = if ($conditions) { "WHERE " + ($conditions -join " AND ") } else { "" }

$query = @"
SELECT FirstName, LastName, Email, Company, JobTitle, LunchType
FROM   Attendees
$whereClause
ORDER  BY LastName, FirstName
"@

$sqlParams = if ($Email) { @{ Email = $Email } } else { @{} }
$attendees = Invoke-SqliteQuery -DataSource $dbPath -Query $query -SqlParameters $sqlParams

if ($attendees.Count -eq 0) {
    Write-Host "No attendees found." -ForegroundColor Yellow
    return
}

Write-Host "Generating badges for $($attendees.Count) attendee(s)..." -ForegroundColor Cyan

$bgDataUri = "data:$bgMime;base64,$bgB64"
$html      = New-BadgeHtml -Attendees $attendees -BgDataUri $bgDataUri

$htmlPath = [System.IO.Path]::ChangeExtension($outputFile, '.html')
Set-Content -Path $htmlPath -Value $html -Encoding UTF8

ConvertTo-Pdf -HtmlPath $htmlPath -PdfPath $outputFile
Remove-Item $htmlPath -ErrorAction SilentlyContinue

Write-Host "Badges written to: $outputFile" -ForegroundColor Green
Write-Host "  $([math]::Ceiling($attendees.Count / 6)) sheet(s), $($attendees.Count) badge(s) total" -ForegroundColor DarkGray
