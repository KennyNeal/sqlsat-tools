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

. "$PSScriptRoot\internal\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config
. "$PSScriptRoot\internal\Web-Helpers.ps1"
. "$PSScriptRoot\internal\Badge-Helpers.ps1"

$outputFile = Join-Path $PSScriptRoot ".." $Config.boothSigns.outputFile
$outputDir  = Split-Path $outputFile
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Fetch sponsor logos ───────────────────────────────────────────────────────

$rawBase = Get-RawBase $Config

$sponsors = [System.Collections.Generic.List[hashtable]]::new()
foreach ($group in (Get-SponsorGroups -Config $Config)) {
    if ($group.tier -notin $Config.sponsors.tableTiers) { continue }
    foreach ($sponsor in $group.sponsors) {
        if ($sponsor.name -in $Config.sponsors.tableExcludeSponsors) {
            Write-Host "  Skipping (excluded): $($sponsor.name)" -ForegroundColor DarkGray
            continue
        }
        try {
            $img = Get-WebImage -Url "$rawBase/static/$($sponsor.logo)"
            $sponsors.Add(@{ Name = $sponsor.name; Base64 = $img.Base64; Mime = $img.Mime; Tier = $group.tier })
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

ConvertTo-PdfViaEdge -HtmlPath $htmlPath -PdfPath $outputFile
Remove-Item $htmlPath -ErrorAction SilentlyContinue

Write-Host "Booth signs PDF: $outputFile ($($sponsors.Count) sponsors)" -ForegroundColor Green
