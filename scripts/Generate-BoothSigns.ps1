<#
.SYNOPSIS
    Generates a printable sponsor booth sign for each vendor as a PDF.
.DESCRIPTION
    Fetches sponsor logos from the website repo for the configured tiers and builds
    one full-page sign per sponsor (logo, tier, and name) for vendors to display at
    their booth table. Same sponsor set as the Stamp Game by default. Generates a
    portrait PDF via Edge headless, one page per sponsor.
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\Generate-BoothSigns.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
    Install-Module -Name powershell-yaml -Scope CurrentUser -Force
}
Import-Module powershell-yaml

. "$PSScriptRoot\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config

$outputFile = Join-Path $PSScriptRoot ".." $Config.boothSigns.outputFile
$outputDir  = Split-Path $outputFile
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$edgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
)

# ── Fetch sponsor logos ───────────────────────────────────────────────────────

$repo    = $Config.websiteRepo
$rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
$yamlUrl = "$rawBase/content/events/$($repo.eventKey)/sponsors.yaml"

Write-Host "Fetching sponsor data..." -ForegroundColor Cyan
$yaml = (Invoke-RestMethod -Uri $yamlUrl -Method Get)
$data = ConvertFrom-Yaml $yaml

$sponsors = [System.Collections.Generic.List[hashtable]]::new()
foreach ($group in $data.groups) {
    if ($group.tier -notin $Config.boothSigns.tiers) { continue }
    foreach ($sponsor in $group.sponsors) {
        if ($sponsor.name -in $Config.boothSigns.excludeSponsors) {
            Write-Host "  Skipping (excluded): $($sponsor.name)" -ForegroundColor DarkGray
            continue
        }
        $logoUrl = "$rawBase/static/$($sponsor.logo)"
        try {
            $raw      = (Invoke-WebRequest -Uri $logoUrl -UseBasicParsing).Content
            $bytes    = if ($raw -is [string]) { [System.Text.Encoding]::UTF8.GetBytes($raw) } else { $raw }
            $b64      = [Convert]::ToBase64String($bytes)
            $ext      = [System.IO.Path]::GetExtension($sponsor.logo).TrimStart('.')
            $mimeType = switch ($ext) {
                'svg'  { 'image/svg+xml' }
                'jpg'  { 'image/jpeg' }
                'jpeg' { 'image/jpeg' }
                default { "image/$ext" }
            }
            $sponsors.Add(@{ Name = $sponsor.name; Base64 = $b64; Mime = $mimeType; Tier = $group.tier })
            Write-Host "  $($sponsor.name)" -ForegroundColor DarkGray
        } catch {
            Write-Host "  Warning: could not load logo for $($sponsor.name)" -ForegroundColor Yellow
        }
    }
}

if ($sponsors.Count -eq 0) { throw "No sponsors found for configured tiers." }

# ── Build HTML ────────────────────────────────────────────────────────────────

$signList = [System.Collections.Generic.List[string]]::new()
foreach ($s in $sponsors) {
    $tierLabel = (Get-Culture).TextInfo.ToTitleCase($s.Tier) + " Sponsor"
    $signList.Add(@"
<div class="sign">
  <div class="tier">$tierLabel</div>
  <div class="logo-wrap"><img src="data:$($s.Mime);base64,$($s.Base64)" alt="$($s.Name)"/></div>
  <div class="name">$($s.Name)</div>
</div>
"@)
}

$eventName = $Config.event.name

$html = @"
<html><head>
<style>
  @page { size: Letter portrait; margin: .5in }
  body  { font-family: Arial, sans-serif; margin: 0; }
  .sign {
    width: 7.5in;
    height: 10in;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    page-break-after: always;
    text-align: center;
  }
  .sign:last-child { page-break-after: auto; }
  .tier {
    font-size: 22pt;
    font-weight: bold;
    color: #013169;
    text-transform: uppercase;
    letter-spacing: .04in;
    margin-bottom: .6in;
  }
  .logo-wrap {
    width: 6in;
    height: 5in;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .logo-wrap img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
  }
  .name {
    font-size: 30pt;
    font-weight: bold;
    color: #222;
    margin-top: .6in;
  }
  .footer {
    font-size: 10pt;
    color: #888;
    margin-top: .3in;
  }
</style>
</head><body>
$($signList -join "`n")
</body></html>
"@

$htmlPath = [System.IO.Path]::ChangeExtension($outputFile, ".html")
Set-Content -Path $htmlPath -Value $html -Encoding UTF8

$edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw "Microsoft Edge not found." }

$null = & $edge --headless=new --print-to-pdf="$outputFile" --no-margins "file:///$htmlPath" --disable-gpu --disable-extensions --no-pdf-header-footer 2>&1
Start-Sleep -Seconds 4
Remove-Item $htmlPath -ErrorAction SilentlyContinue

Write-Host "Booth signs PDF: $outputFile ($($sponsors.Count) sponsors)" -ForegroundColor Green
