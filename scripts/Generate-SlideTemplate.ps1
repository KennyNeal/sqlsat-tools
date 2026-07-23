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

. "$PSScriptRoot\internal\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config
. "$PSScriptRoot\internal\Get-EventLogo.ps1"
. "$PSScriptRoot\internal\Badge-Helpers.ps1"
. "$PSScriptRoot\internal\Slide-Common.ps1"

$slideCfg      = $Config.slideTemplate
$outputFile    = Join-Path $PSScriptRoot ".." $slideCfg.outputFile
$footerText    = if ($slideCfg.PSObject.Properties['footerText']) { $slideCfg.footerText } else { "NO FOOD OR DRINKS IN THE CLASSROOMS" }
$primaryColor  = if ($slideCfg.PSObject.Properties['primaryColor']) { $slideCfg.primaryColor } else { "013169" }
$secondaryColor = if ($slideCfg.PSObject.Properties['secondaryColor']) { $slideCfg.secondaryColor } else { "F7C15D" }

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) "sqlsat-slidetemplate-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $workDir | Out-Null

# ── Event logo (watermark) ──────────────────────────────────────────────────

Write-Host "Fetching event logo..." -ForegroundColor Cyan
$eventLogo = Get-EventLogo -Config $Config
if (-not $eventLogo) { throw "Could not load the event logo — check websiteRepo.eventKey in event.config.json." }
$logoPath = Join-Path $workDir "event-logo.png"
[IO.File]::WriteAllBytes($logoPath, [Convert]::FromBase64String($eventLogo.Base64))

# ── Sponsor logos (all tiers) ───────────────────────────────────────────────

$groups   = Get-SponsorLogoFiles -Config $Config -WorkDir $workDir
$sponsors = @($groups | ForEach-Object { $_.sponsors })

# ── Python dependencies ─────────────────────────────────────────────────────

Install-PythonSlideDeps

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

python "$PSScriptRoot\internal\generate_slide_template.py" $manifestPath
if ($LASTEXITCODE -ne 0) { throw "generate_slide_template.py failed." }

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Slide template: $outputFile" -ForegroundColor Green
