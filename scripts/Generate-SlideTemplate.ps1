<#
.SYNOPSIS
    Generates the presenter slide-deck template (.potx) for the event.
.DESCRIPTION
    Builds a 3-slide PowerPoint template: a blank title slide, a "Thank You,
    Sponsors!" slide with every sponsor logo pulled live from the website
    repo, and a session-evaluation slide with an empty QR-code placeholder
    for each presenter to drop their own Sessionize code into.

    Uses the event's logo (from content/events/{eventKey}/_index.md) as a
    faded background watermark on the title and sponsor slides, matching
    previous years' decks. Colors come from event.config.json's
    slideTemplate section (defaults to the Day of Data navy/gold brand).

    Requires Python 3 with the python-pptx and Pillow packages (installed
    automatically on first run) and Microsoft Edge (used to rasterize any
    SVG sponsor logos to PNG, since PowerPoint templates can't embed SVG
    reliably).
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\Generate-SlideTemplate.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
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
. "$PSScriptRoot\Get-EventLogo.ps1"

$slideCfg      = $Config.slideTemplate
$outputFile    = Join-Path $PSScriptRoot ".." $slideCfg.outputFile
$footerText    = if ($slideCfg.PSObject.Properties['footerText']) { $slideCfg.footerText } else { "NO FOOD OR DRINKS IN THE CLASSROOMS" }
$primaryColor  = if ($slideCfg.PSObject.Properties['primaryColor']) { $slideCfg.primaryColor } else { "013169" }
$secondaryColor = if ($slideCfg.PSObject.Properties['secondaryColor']) { $slideCfg.secondaryColor } else { "F7C15D" }

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) "sqlsat-slidetemplate-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $workDir | Out-Null

$edgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
)
$edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

function ConvertTo-Png {
    param([string]$SvgPath, [string]$PngPath)
    if (-not $edge) { throw "Microsoft Edge not found (needed to rasterize SVG sponsor logos)." }
    & $edge --headless=new --disable-gpu --disable-extensions `
        --window-size=1000,1000 --force-device-scale-factor=2 `
        --default-background-color=00000000 `
        --screenshot="$PngPath" "file:///$SvgPath" 2>&1 | Out-Null

    # Edge writes the screenshot asynchronously as the process shuts down;
    # poll until the file appears and its size stops growing.
    $deadline = (Get-Date).AddSeconds(10)
    $lastSize = -1
    while ((Get-Date) -lt $deadline) {
        if (Test-Path $PngPath) {
            $size = (Get-Item $PngPath).Length
            if ($size -gt 0 -and $size -eq $lastSize) { break }
            $lastSize = $size
        }
        Start-Sleep -Milliseconds 300
    }
}

# ── Event logo (watermark) ──────────────────────────────────────────────────

Write-Host "Fetching event logo..." -ForegroundColor Cyan
$eventLogo = Get-EventLogo -Config $Config
if (-not $eventLogo) { throw "Could not load the event logo — check websiteRepo.eventKey in event.config.json." }
$logoPath = Join-Path $workDir "event-logo.png"
[IO.File]::WriteAllBytes($logoPath, [Convert]::FromBase64String($eventLogo.Base64))

# ── Sponsor logos (all tiers) ───────────────────────────────────────────────

$repo    = $Config.websiteRepo
$rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
$yamlUrl = "$rawBase/content/events/$($repo.eventKey)/sponsors.yaml"

Write-Host "Fetching sponsor data..." -ForegroundColor Cyan
$yaml = (Invoke-RestMethod -Uri $yamlUrl -Method Get)
$data = ConvertFrom-Yaml $yaml

$sponsors = [System.Collections.Generic.List[hashtable]]::new()
$i = 0
foreach ($group in $data.groups) {
    foreach ($sponsor in $group.sponsors) {
        $i++
        $ext = [System.IO.Path]::GetExtension($sponsor.logo)
        $rawLogoPath = Join-Path $workDir "sponsor$i$ext"
        try {
            Invoke-WebRequest -Uri "$rawBase/static/$($sponsor.logo)" -OutFile $rawLogoPath -UseBasicParsing
        } catch {
            Write-Host "  Warning: could not download logo for $($sponsor.name)" -ForegroundColor Yellow
            continue
        }

        $finalLogoPath = $rawLogoPath
        if ($ext -eq ".svg") {
            $pngPath = Join-Path $workDir "sponsor$i.png"
            ConvertTo-Png -SvgPath $rawLogoPath -PngPath $pngPath
            if (Test-Path $pngPath) { $finalLogoPath = $pngPath }
        }

        $sponsors.Add(@{ name = $sponsor.name; logoPath = $finalLogoPath })
        Write-Host "  $($sponsor.name)" -ForegroundColor DarkGray
    }
}

# ── Python dependencies ─────────────────────────────────────────────────────

$pythonOk = $false
try {
    & python -c "import pptx, PIL" 2>$null
    if ($LASTEXITCODE -eq 0) { $pythonOk = $true }
} catch { }
if (-not $pythonOk) {
    Write-Host "Installing python-pptx and Pillow..." -ForegroundColor Cyan
    & python -m pip install --quiet python-pptx Pillow
}

# ── Build manifest and hand off to Python ───────────────────────────────────

$brugLogoPath = Join-Path $PSScriptRoot ".." "assets" "brug-logo.png"

$manifest = @{
    eventName       = $Config.event.name
    footerText      = $footerText
    primaryColor    = $primaryColor
    secondaryColor  = $secondaryColor
    logoPath        = $logoPath
    brugLogoPath    = (Resolve-Path $brugLogoPath).Path
    sponsors        = $sponsors
    outputPath      = (New-Object System.IO.FileInfo($outputFile)).FullName
}
$manifestPath = Join-Path $workDir "manifest.json"
$manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding UTF8

python "$PSScriptRoot\generate_slide_template.py" $manifestPath
if ($LASTEXITCODE -ne 0) { throw "generate_slide_template.py failed." }

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Slide template: $outputFile" -ForegroundColor Green
