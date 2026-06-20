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

if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
    Install-Module -Name powershell-yaml -Scope CurrentUser -Force
}
Import-Module powershell-yaml

$cols       = if ($GridColumns -gt 0) { $GridColumns } else { $Config.stampGame.gridColumns }
$outputFile = Join-Path $PSScriptRoot ".." $Config.stampGame.outputFile
$outputDir  = Split-Path $outputFile
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$edgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
)

# ── Fetch sponsor logos ───────────────────────────────────────────────────────

$repo    = $Config.websiteRepo
$rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
$yamlUrl = "$rawBase/data/sponsors/$($repo.sponsorDataFile).yaml"

Write-Host "Fetching sponsor data..." -ForegroundColor Cyan
$yaml = (Invoke-RestMethod -Uri $yamlUrl -Method Get)
$data = ConvertFrom-Yaml $yaml

$logos = [System.Collections.Generic.List[hashtable]]::new()
foreach ($group in $data.groups) {
    if ($group.tier -notin $Config.stampGame.tiers) { continue }
    foreach ($sponsor in $group.sponsors) {
        $logoUrl = "$rawBase/static/$($sponsor.logo)"
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
            $logos.Add(@{ Name = $sponsor.name; Base64 = $b64; Mime = $mimeType })
            Write-Host "  $($sponsor.name)" -ForegroundColor DarkGray
        } catch {
            Write-Host "  Warning: could not load logo for $($sponsor.name)" -ForegroundColor Yellow
        }
    }
}

# ── Build HTML ────────────────────────────────────────────────────────────────

$gridCells = ""
foreach ($logo in $logos) {
    $gridCells += @"
<div class="cell">
  <img src="data:$($logo.Mime);base64,$($logo.Base64)" alt="$($logo.Name)"/>
  <div class="stamp-area"></div>
</div>
"@
}

$eventName = $Config.event.name

$html = @"
<html><head>
<style>
  @page { size: Letter landscape; margin: .4in }
  body  { font-family: Arial, sans-serif; margin: 0 }
  h2    { text-align: center; margin: 0 0 .15in; font-size: 14pt }
  .instructions { text-align: center; font-size: 9pt; margin-bottom: .2in; color: #444 }
  .name-line { border-bottom: 1px solid #333; width: 3in; display: inline-block; margin-left: .2in }
  .grid {
    display: grid;
    grid-template-columns: repeat($cols, 1fr);
    gap: .15in;
  }
  .cell {
    border: 2px solid #333;
    border-radius: 4px;
    padding: .1in;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;
    min-height: 1.3in;
  }
  .cell img {
    max-width: 100%;
    max-height: .8in;
    object-fit: contain;
  }
  .stamp-area {
    width: .9in;
    height: .4in;
    border: 1px dashed #aaa;
    margin-top: .05in;
    border-radius: 3px;
  }
</style>
</head><body>
<h2>$eventName — Sponsor Stamp Game</h2>
<div class="instructions">
  Visit each sponsor booth to get your card stamped. Turn in for raffle entry!
  &nbsp;&nbsp;&nbsp; Name: <span class="name-line">&nbsp;</span>
</div>
<div class="grid">$gridCells</div>
</body></html>
"@

$htmlPath = [System.IO.Path]::ChangeExtension($outputFile, ".html")
Set-Content -Path $htmlPath -Value $html -Encoding UTF8

$edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw "Microsoft Edge not found." }

$null = & $edge --headless=new --print-to-pdf="$outputFile" --no-margins "file:///$htmlPath" --disable-gpu --disable-extensions --no-pdf-header-footer 2>&1
Start-Sleep -Seconds 4
Remove-Item $htmlPath -ErrorAction SilentlyContinue

Write-Host "Stamp game PDF: $outputFile" -ForegroundColor Green
