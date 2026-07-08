<#
.SYNOPSIS
    Generates the end-of-day raffle deck (.pptx) for the event.
.DESCRIPTION
    Builds a single deck with two real PowerPoint sections, each also saved
    as a same-named custom show (Slide Show > Custom Slide Show >
    Recognition / Raffle):
      1. "Recognition" -- self-playing (each slide auto-advances after
         raffleDeck.loopAdvanceSeconds): one slide per platinum/global
         sponsor, one or more grid slides for gold, one or more grid slides
         for silver & bronze, then a QR-code slide linking to the
         Sessionize app for session evaluations. Run as a custom show, it
         loops continuously until Esc, so it can play itself before the
         raffle starts.
      2. "Raffle" -- starts with a "Raffle Time!" slide, then one
         manually-advanced hero slide per raffle-eligible sponsor
         (speedpass.raffleTiers) to display while that sponsor does their
         drawing, then a second copy of the evaluation QR slide. Sponsors
         listed in raffleDeck.excludeSponsors are skipped here (they still
         get their Recognition slide) -- useful since which sponsors are
         actually doing a drawing often isn't confirmed until the day
         before, and the Global sponsor typically opts out.
    The deck's default "Show slides" range is left at "All" rather than
    pinned to Recognition, since pinning it breaks Shift+F5 ("Show From
    Current Slide") for every Raffle slide. So launch each section by name
    from Slide Show > Custom Slide Show; when it's time, Esc out of the
    looping Recognition show and launch Raffle the same way.

    Sponsor tiers/logos come live from the website repo's sponsors.yaml, the
    same source used by Generate-SlideTemplate.ps1, Generate-StampGame.ps1,
    and Generate-SpeedPasses.ps1. The eval QR code is generated with
    QRCoder against schedule.appUrl. Colors and footer text come from
    event.config.json's slideTemplate section (shared with the presenter
    template) so both decks stay on-brand together.

    Requires Python 3 with python-pptx and Pillow (installed automatically
    on first run), Microsoft Edge (to rasterize SVG sponsor logos), and
    lib\QRCoder.dll.
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\Generate-RaffleDeck.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
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

$raffleCfg      = $Config.raffleDeck
$outputFile     = Join-Path $PSScriptRoot ".." $raffleCfg.outputFile
$loopSeconds    = if ($raffleCfg.PSObject.Properties['loopAdvanceSeconds']) { $raffleCfg.loopAdvanceSeconds } else { 8 }
$maxPerGrid     = if ($raffleCfg.PSObject.Properties['maxPerGridSlide']) { $raffleCfg.maxPerGridSlide } else { 8 }
$footerText     = if ($raffleCfg.PSObject.Properties['footerText']) { $raffleCfg.footerText } else { $Config.event.hashtag }

# @() wraps the whole if/else, not just a branch -- if/else output goes
# through pipeline enumeration, so a branch that's an array with exactly one
# element silently collapses to a bare scalar on assignment otherwise.
$individualTiers = @(if ($raffleCfg.PSObject.Properties['individualTiers']) { $raffleCfg.individualTiers } else { @('global', 'platinum') })
$gridGroups      = @(if ($raffleCfg.PSObject.Properties['gridGroups']) { $raffleCfg.gridGroups } else { @(@('gold'), @('silver', 'bronze')) })
$excludeSponsors = @(if ($raffleCfg.PSObject.Properties['excludeSponsors']) { $raffleCfg.excludeSponsors } else { @() })

$slideCfg       = $Config.slideTemplate
$primaryColor   = if ($slideCfg.PSObject.Properties['primaryColor']) { $slideCfg.primaryColor } else { "013169" }
$secondaryColor = if ($slideCfg.PSObject.Properties['secondaryColor']) { $slideCfg.secondaryColor } else { "F7C15D" }

$raffleTiers = $Config.speedpass.raffleTiers
if (-not $raffleTiers) { throw "speedpass.raffleTiers is not set in event.config.json — needed to know which sponsors get a raffle hero slide." }

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) "sqlsat-raffledeck-$([guid]::NewGuid())"
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

# ── Sponsor logos (all tiers — the loop deck decides which ones to use) ────

$repo    = $Config.websiteRepo
$rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
$yamlUrl = "$rawBase/content/events/$($repo.eventKey)/sponsors.yaml"

Write-Host "Fetching sponsor data..." -ForegroundColor Cyan
$yaml = (Invoke-RestMethod -Uri $yamlUrl -Method Get)
$data = ConvertFrom-Yaml $yaml

$groups = [System.Collections.Generic.List[hashtable]]::new()
$i = 0
foreach ($group in $data.groups) {
    $sponsors = [System.Collections.Generic.List[hashtable]]::new()
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
        Write-Host "  [$($group.tier)] $($sponsor.name)" -ForegroundColor DarkGray
    }
    $groups.Add(@{ tier = $group.tier; title = $group.title; sponsors = $sponsors })
}

# ── Session-evaluation QR code ──────────────────────────────────────────────

$evalUrl = $Config.schedule.appUrl
if (-not $evalUrl) { throw "schedule.appUrl is not set in event.config.json — needed for the evaluation QR code." }

$libPath = Join-Path $PSScriptRoot ".." "lib" "QRCoder.dll"
if (-not (Test-Path $libPath)) { throw "QRCoder.dll not found at $libPath." }
Add-Type -Path $libPath

Write-Host "Generating evaluation QR code..." -ForegroundColor Cyan
$gen    = New-Object QRCoder.QRCodeGenerator
$qrData = $gen.CreateQrCode($evalUrl, [QRCoder.QRCodeGenerator+ECCLevel]::M)
$qr     = New-Object QRCoder.QRCode($qrData)
$bmp    = $qr.GetGraphic(20)
$qrPath = Join-Path $workDir "eval-qr.png"
$bmp.Save($qrPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose(); $qr.Dispose(); $gen.Dispose()

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

$manifest = @{
    eventName        = $Config.event.name
    footerText       = $footerText
    primaryColor     = $primaryColor
    secondaryColor   = $secondaryColor
    logoPath         = $logoPath
    qrPath           = $qrPath
    evalUrl          = $evalUrl
    loopAdvanceSeconds = $loopSeconds
    maxPerGridSlide  = $maxPerGrid
    individualTiers  = $individualTiers
    gridGroups       = $gridGroups
    heroTiers        = $raffleTiers
    excludeSponsors  = $excludeSponsors
    groups           = $groups
    outputPath       = (New-Object System.IO.FileInfo($outputFile)).FullName
}
$manifestPath = Join-Path $workDir "manifest.json"
$manifest | ConvertTo-Json -Depth 6 | Set-Content -Path $manifestPath -Encoding UTF8

python "$PSScriptRoot\generate_raffle_deck.py" $manifestPath
if ($LASTEXITCODE -ne 0) { throw "generate_raffle_deck.py failed." }

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Raffle deck: $outputFile" -ForegroundColor Green
