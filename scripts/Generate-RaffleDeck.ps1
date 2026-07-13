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
         (raffleDeck.heroTiers, defaulting to speedpass.raffleTiers) to
         display while that sponsor does their drawing, then any
         raffleDeck.extraHeroSlides (non-sponsor drawings, e.g. the local
         user group's own raffle, with a locally-supplied logo file), then
         a second copy of the evaluation QR slide. Sponsors listed in
         raffleDeck.excludeSponsors are skipped here (they still get their
         Recognition slide) -- useful since which sponsors are actually
         doing a drawing often isn't confirmed until the day before, or a
         sponsor tier (e.g. Global) typically opts out but this year didn't.
         Sponsors listed in raffleDeck.heroLast are moved to the end of the
         sponsor hero slides (still before extraHeroSlides), e.g. to run a
         sponsor's drawing right before the user group's.
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

. "$PSScriptRoot\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config
. "$PSScriptRoot\Get-EventLogo.ps1"
. "$PSScriptRoot\Badge-Helpers.ps1"
. "$PSScriptRoot\Slide-Common.ps1"

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

# Sponsor names in heroLast are pulled out of their normal tier position and
# raffled last, in the order listed -- e.g. a sponsor whose drawing the
# presenter wants to run right before a closing/non-sponsor hero slide.
$heroLast        = @(if ($raffleCfg.PSObject.Properties['heroLast']) { $raffleCfg.heroLast } else { @() })

$slideCfg       = $Config.slideTemplate
$primaryColor   = if ($slideCfg.PSObject.Properties['primaryColor']) { $slideCfg.primaryColor } else { "013169" }
$secondaryColor = if ($slideCfg.PSObject.Properties['secondaryColor']) { $slideCfg.secondaryColor } else { "F7C15D" }

$raffleTiers = $Config.speedpass.raffleTiers
if (-not $raffleTiers) { throw "speedpass.raffleTiers is not set in event.config.json — needed to know which sponsors get a raffle hero slide." }

# heroTiers defaults to speedpass.raffleTiers but can be overridden independently --
# e.g. to add "global" for a year where the Global sponsor opts in to the drawing
# without also handing them a raffle-ticket stamp on the speedpass card.
$heroTiers = @(if ($raffleCfg.PSObject.Properties['heroTiers']) { $raffleCfg.heroTiers } else { $raffleTiers })

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) "sqlsat-raffledeck-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $workDir | Out-Null

# ── Event logo (watermark) ──────────────────────────────────────────────────

Write-Host "Fetching event logo..." -ForegroundColor Cyan
$eventLogo = Get-EventLogo -Config $Config
if (-not $eventLogo) { throw "Could not load the event logo — check websiteRepo.eventKey in event.config.json." }
$logoPath = Join-Path $workDir "event-logo.png"
[IO.File]::WriteAllBytes($logoPath, [Convert]::FromBase64String($eventLogo.Base64))

# ── Sponsor logos (all tiers — the loop deck decides which ones to use) ────

$groups = Get-SponsorLogoFiles -Config $Config -WorkDir $workDir

# ── Extra raffle-only hero slides (not sourced from sponsors.yaml) ─────────
# e.g. the local user group's own drawing, tacked on after every sponsor.

$extraHeroSlides = [System.Collections.Generic.List[hashtable]]::new()
if ($raffleCfg.PSObject.Properties['extraHeroSlides']) {
    $i = 0
    foreach ($extra in $raffleCfg.extraHeroSlides) {
        $i++
        $srcLogoPath = Join-Path $PSScriptRoot ".." $extra.logoPath
        if (-not (Test-Path $srcLogoPath)) { throw "extraHeroSlides logo not found: $srcLogoPath" }
        $ext = [System.IO.Path]::GetExtension($srcLogoPath)
        $destLogoPath = Join-Path $workDir "extra-hero-$i$ext"
        Copy-Item $srcLogoPath $destLogoPath
        if ($ext -eq ".svg") {
            $pngPath = Join-Path $workDir "extra-hero-$i.png"
            ConvertTo-PngFromSvg -SvgPath $destLogoPath -PngPath $pngPath
            if (Test-Path $pngPath) { $destLogoPath = $pngPath }
        }
        $extraHeroSlides.Add(@{
            name     = $extra.name
            kicker   = if ($extra.PSObject.Properties['kicker']) { $extra.kicker } else { $extra.name }
            logoPath = $destLogoPath
        })
    }
}

# ── Session-evaluation QR code ──────────────────────────────────────────────

$evalUrl = $Config.schedule.appUrl
if (-not $evalUrl) { throw "schedule.appUrl is not set in event.config.json — needed for the evaluation QR code." }

Import-QRCoder

Write-Host "Generating evaluation QR code..." -ForegroundColor Cyan
$qrB64  = New-QRBase64 -Data $evalUrl -PixelSize 20 -EccLevel M
$qrPath = Join-Path $workDir "eval-qr.png"
[IO.File]::WriteAllBytes($qrPath, [Convert]::FromBase64String($qrB64))

# ── Python dependencies ─────────────────────────────────────────────────────

Install-PythonSlideDeps

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
    heroTiers        = $heroTiers
    heroLast         = $heroLast
    excludeSponsors  = $excludeSponsors
    extraHeroSlides  = $extraHeroSlides
    groups           = $groups
    outputPath       = (New-Object System.IO.FileInfo($outputFile)).FullName
}
$manifestPath = Join-Path $workDir "manifest.json"
$manifest | ConvertTo-Json -Depth 6 | Set-Content -Path $manifestPath -Encoding UTF8

python "$PSScriptRoot\generate_raffle_deck.py" $manifestPath
if ($LASTEXITCODE -ne 0) { throw "generate_raffle_deck.py failed." }

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Raffle deck: $outputFile" -ForegroundColor Green
