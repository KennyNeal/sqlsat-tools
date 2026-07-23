<#
.SYNOPSIS
    Generates the printable sponsor Stamp Game sheet as a PDF.
.DESCRIPTION
    Fetches sponsor logos from the website repo for the configured tiers and builds
    a grid of logo boxes attendees take around to get stamped at each sponsor booth.
    Generates a landscape PDF via Edge headless.
.PARAMETER Config
    Parsed event.config.json object.
.PARAMETER GridColumns
    Number of columns in the logo grid. Overrides the value in config.
.EXAMPLE
    .\Generate-StampGame.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config,
    [int]$GridColumns = 0
)

. "$PSScriptRoot\internal\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config
. "$PSScriptRoot\internal\Web-Helpers.ps1"
. "$PSScriptRoot\internal\Badge-Helpers.ps1"

$cols       = if ($GridColumns -gt 0) { $GridColumns } else { $Config.stampGame.gridColumns }
$outputFile = Join-Path $PSScriptRoot ".." $Config.stampGame.outputFile
$outputDir  = Split-Path $outputFile
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Fetch sponsor logos ───────────────────────────────────────────────────────

$rawBase = Get-RawBase $Config

$logos = [System.Collections.Generic.List[hashtable]]::new()
foreach ($group in (Get-SponsorGroups -Config $Config)) {
    if ($group.tier -notin $Config.sponsors.tableTiers) { continue }
    foreach ($sponsor in $group.sponsors) {
        if ($sponsor.name -in $Config.sponsors.tableExcludeSponsors) {
            Write-Host "  Skipping (excluded): $($sponsor.name)" -ForegroundColor DarkGray
            continue
        }
        try {
            $img = Get-WebImage -Url "$rawBase/static/$($sponsor.logo)"
            $logos.Add(@{ Name = $sponsor.name; Base64 = $img.Base64; Mime = $img.Mime })
            Write-Host "  $($sponsor.name)" -ForegroundColor DarkGray
        } catch {
            Write-Host "  Warning: could not load logo for $($sponsor.name)" -ForegroundColor Yellow
        }
    }
}

# ── Resolve free-space logo ───────────────────────────────────────────────────

. "$PSScriptRoot\internal\Get-EventLogo.ps1"
$eventLogo = Get-EventLogo -Config $Config -Override $Config.stampGame.freeSpaceLogoFile

# ── Build HTML ────────────────────────────────────────────────────────────────

$cellList = [System.Collections.Generic.List[string]]::new()
foreach ($logo in $logos) {
    $cellList.Add(@"
<div class="cell">
  <img src="data:$($logo.Mime);base64,$($logo.Base64)" alt="$($logo.Name)"/>
  <div class="stamp-area"></div>
</div>
"@)
}

$n      = $cellList.Count
$sqrtN     = [int][Math]::Round([Math]::Sqrt($n))
$emptyCell = '<div class="cell empty"><div class="stamp-area" style="width:85%;flex:1;margin:.04in 0"></div></div>'
$totalCells = [int]([Math]::Ceiling([double]$n / $cols) * $cols)

if ($sqrtN * $sqrtN -eq $n) {
    # Already a perfect square — snap cols to match, no padding needed
    $cols = $sqrtN
    $totalCells = $n
} elseif ($totalCells - $n -eq 1) {
    # Exactly one empty slot — place the free space in the center instead
    $freeInner = if ($eventLogo) {
        '<img src="data:' + $eventLogo.Mime + ';base64,' + $eventLogo.Base64 + '" alt="' + $Config.event.name + '"/>'
    } else { "" }
    $freeCell = '<div class="cell free-space">' + $freeInner + '<div class="free-label">FREE</div></div>'
    $cellList.Insert([Math]::Floor($totalCells / 2), $freeCell)
} else {
    # Multiple empty slots — pad end with blank stamp squares
    while ($cellList.Count -lt $totalCells) { $cellList.Add($emptyCell) }
}

$gridCells = $cellList -join ""

# Compute row height to fill available card height without overflowing.
# Page is 8in tall (landscape Letter minus .25in margins each side).
# Header area (h2 + instructions + name-row) takes ~0.75in; gaps between rows = (rows-1)*0.1in.
$totalRows   = [int][Math]::Ceiling([double]$cellList.Count / $cols)
$gridGapsIn  = ($totalRows - 1) * 0.1
$rowHeight   = [Math]::Round((8.0 - 0.75 - $gridGapsIn) / $totalRows, 2)
if ($rowHeight -gt 1.4) { $rowHeight = 1.4 }
$rowHeightStr = "${rowHeight}in"

$eventName = $Config.event.name

$cardHtml = @"
<div class="card">
  <h2>$eventName — Sponsor Stamp Game</h2>
  <div class="instructions">Visit booths to get a stamp. Turn in for raffle entry!</div>
  <div class="name-row">Name: <span class="name-line">&nbsp;</span></div>
  <div class="grid">$gridCells</div>
</div>
"@

$html = @"
<html><head>
<style>
  @page { size: Letter landscape; margin: .25in }
  body  { font-family: Arial, sans-serif; margin: 0; display: flex; flex-direction: row; height: 8in }
  .card {
    flex: 1;
    overflow: hidden;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
  }
  .cut-line {
    width: 0;
    border: none;
    border-left: 1px dashed #bbb;
    align-self: stretch;
    flex-shrink: 0;
  }
  h2    { text-align: center; margin: .05in 0 .06in; font-size: 10pt }
  .instructions { text-align: center; font-size: 8.5pt; margin-bottom: 0; color: #444 }
  .name-row { text-align: center; font-size: 8.5pt; color: #444; margin-top: .12in; margin-bottom: .1in; }
  .name-line { border-bottom: 1px solid #333; width: 3in; display: inline-block; margin-left: .15in }
  .grid {
    display: grid;
    grid-template-columns: repeat($cols, 1fr);
    grid-auto-rows: $rowHeightStr;
    align-content: start;
    gap: .1in;
  }
  .cell {
    border: 2px solid #333;
    border-radius: 4px;
    padding: .08in;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;
  }
  .cell img {
    max-width: 100%;
    max-height: .65in;
    object-fit: contain;
  }
  .stamp-area {
    width: .85in;
    height: .35in;
    border: 1px dashed #aaa;
    margin-top: .04in;
    border-radius: 3px;
  }
  .cell.free-space { background: #eef2ff; border-color: #013169; }
  .free-label { font-size: 9pt; font-weight: bold; color: #013169; text-align: center; margin-top: .04in; }
  .cell.empty { border-style: dashed; border-color: #ccc; }
</style>
</head><body>
$cardHtml
<div class="cut-line"></div>
$cardHtml
</body></html>
"@

$htmlPath = [System.IO.Path]::ChangeExtension($outputFile, ".html")
Set-Content -Path $htmlPath -Value $html -Encoding UTF8

ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $outputFile
Remove-Item $htmlPath -ErrorAction SilentlyContinue

Write-Host "Stamp game PDF: $outputFile" -ForegroundColor Green
