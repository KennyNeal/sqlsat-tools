<#
.SYNOPSIS
    Preflight check: validates config, secrets, data sources, and local
    tools before an event run.
.DESCRIPTION
    Runs every assumption the tools make and reports PASS/WARN/FAIL for
    each check, so problems surface all at once instead of one at a time
    mid-run on whichever script hits them first. Run it the day before the
    event and again the morning of.

    FAIL means a pipeline script will break. WARN means something is only
    needed later (day-of printing, first-run auto-installs) or is worth
    eyeballing (no published sessions yet).

    Exit code is non-zero when anything FAILs, so it can gate a scheduled
    Update-Event.ps1 run.
.PARAMETER ConfigPath
    Path to event.config.json. Defaults to event.config.json in the repo root.
.EXAMPLE
    .\scripts\Test-EventReadiness.ps1
#>
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "..\event.config.json")
)

. "$PSScriptRoot\Web-Helpers.ps1"
. "$PSScriptRoot\Badge-Helpers.ps1"
. "$PSScriptRoot\Get-EventLogo.ps1"
. "$PSScriptRoot\Data-Access.ps1"

$script:results = [System.Collections.Generic.List[object]]::new()

function Test-Check {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Check,
        [switch]$WarnOnly
    )
    try {
        $detail = & $Check
        $script:results.Add([PSCustomObject]@{ Name = $Name; Status = 'PASS'; Detail = "$detail" })
        Write-Host ("  [PASS] {0}{1}" -f $Name, $(if ("$detail") { " — $detail" } else { "" })) -ForegroundColor Green
    } catch {
        $status = if ($WarnOnly) { 'WARN' } else { 'FAIL' }
        $color  = if ($WarnOnly) { 'Yellow' } else { 'Red' }
        $script:results.Add([PSCustomObject]@{ Name = $Name; Status = $status; Detail = $_.Exception.Message })
        Write-Host ("  [{0}] {1} — {2}" -f $status, $Name, $_.Exception.Message) -ForegroundColor $color
    }
}

function Get-ConfigValue {
    param($Object, [string]$Path)
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $Object -or -not $Object.PSObject.Properties[$part]) { return $null }
        $Object = $Object.$part
    }
    return $Object
}

# ── Config ────────────────────────────────────────────────────────────────────

Write-Host "`nConfig" -ForegroundColor Cyan

$script:config = $null
Test-Check "event.config.json parses" {
    if (-not (Test-Path $ConfigPath)) {
        throw "not found at $ConfigPath — copy event.config.template.json and fill it in"
    }
    $script:config = Get-Content $ConfigPath | ConvertFrom-Json
    (Resolve-Path $ConfigPath).Path
}

if (-not $script:config) {
    Write-Host "`nCannot continue without a config file." -ForegroundColor Red
    exit 1
}
$config = $script:config

Test-Check "Required config keys" {
    $required = @(
        'event.hashtag',
        'websiteRepo.owner', 'websiteRepo.name', 'websiteRepo.branch', 'websiteRepo.eventKey',
        'eventbrite.eventId', 'eventbrite.secretName',
        'email.secretName', 'email.subject', 'email.batchSize', 'email.delaySeconds',
        'database.path',
        'speedpass.raffleTiers', 'speedpass.outputDir',
        'stampGame.tiers', 'stampGame.gridColumns', 'stampGame.outputFile',
        'schedule.outputFile', 'schedule.appUrl',
        'sessionize.eventId'
    )
    $missing = @($required | Where-Object { $null -eq (Get-ConfigValue $config $_) })
    if ($missing) { throw "missing: $($missing -join ', ')" }
    "$($required.Count) keys present"
}

Test-Check "Deck output config (slideTemplate/raffleDeck)" -WarnOnly {
    $missing = @('slideTemplate.outputFile', 'raffleDeck.outputFile' |
        Where-Object { $null -eq (Get-ConfigValue $config $_) })
    if ($missing) { throw "missing: $($missing -join ', ') — deck generators will fail without them (other keys have defaults)" }
    "present"
}

# ── Local tools ───────────────────────────────────────────────────────────────

Write-Host "`nLocal tools" -ForegroundColor Cyan

Test-Check "PowerShell 7+" -WarnOnly {
    if ($PSVersionTable.PSVersion.Major -lt 7) { throw "running $($PSVersionTable.PSVersion) — 7+ recommended" }
    "$($PSVersionTable.PSVersion)"
}

Test-Check "Microsoft Edge" { Get-EdgePath }

Test-Check "QRCoder.dll loads" { Import-QRCoder; "loaded" }

Test-Check "PSSQLite module" -WarnOnly {
    if (-not (Get-Module -ListAvailable -Name PSSQLite)) { throw "not installed — Initialize-Database.ps1 installs it on first run" }
    "installed"
}

Test-Check "powershell-yaml module" -WarnOnly {
    if (-not (Get-Module -ListAvailable -Name powershell-yaml)) { throw "not installed — installed automatically on first sponsor fetch" }
    "installed"
}

Test-Check "Python + python-pptx/Pillow" -WarnOnly {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) { throw "python not on PATH — needed for Generate-SlideTemplate/Generate-RaffleDeck" }
    & python -c "import pptx, PIL" 2>$null
    if ($LASTEXITCODE -ne 0) { throw "python-pptx/Pillow not installed — installed automatically on first deck run" }
    "ready"
}

Test-Check "SumatraPDF (day-of badge printing)" -WarnOnly {
    $cmd = Get-Command SumatraPDF.exe -ErrorAction SilentlyContinue
    $found = if ($cmd) { $cmd.Source } else {
        @("$env:LOCALAPPDATA\SumatraPDF\SumatraPDF.exe", "${env:ProgramFiles}\SumatraPDF\SumatraPDF.exe") |
            Where-Object { Test-Path $_ } | Select-Object -First 1
    }
    if (-not $found) { throw "not found — install with: winget install --id SumatraPDF.SumatraPDF -e" }
    $found
}

Test-Check "Walk-in label printer (day-of)" -WarnOnly {
    $printerName = if ($config.PSObject.Properties['badge'] -and $config.badge.walkinPrinter) { $config.badge.walkinPrinter } else { "Brother QL-820NWB" }
    if (-not (Get-Command Get-Printer -ErrorAction SilentlyContinue)) { throw "Get-Printer unavailable on this system" }
    if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) { throw "printer '$printerName' not visible to Windows" }
    $printerName
}

Test-Check "Badge background image" -WarnOnly {
    $bg = if ($config.PSObject.Properties['badge'] -and $config.badge.backgroundImage) {
        Join-Path $PSScriptRoot ".." $config.badge.backgroundImage
    } else {
        Join-Path $PSScriptRoot "..\assets\badge-background.png"
    }
    if (-not (Test-Path $bg)) { throw "not found at $bg — Generate-NameTag.ps1 needs it" }
    $bg
}

# ── Secrets ───────────────────────────────────────────────────────────────────

Write-Host "`nSecrets" -ForegroundColor Cyan

Test-Check "Eventbrite token secret" {
    if (-not (Get-Command Get-Secret -ErrorAction SilentlyContinue)) {
        throw "SecretManagement not installed — see the README's install one-liner"
    }
    $null = Get-Secret -Name $config.eventbrite.secretName -ErrorAction Stop
    "'$($config.eventbrite.secretName)' retrievable"
}

Test-Check "Gmail credential secret" {
    if (-not (Get-Command Get-Secret -ErrorAction SilentlyContinue)) {
        throw "SecretManagement not installed — see the README's install one-liner"
    }
    $cred = Get-Secret -Name $config.email.secretName -ErrorAction Stop
    if ($cred -isnot [pscredential]) { throw "'$($config.email.secretName)' is not a PSCredential — store it with Set-Secret -Secret (Get-Credential)" }
    "'$($config.email.secretName)' retrievable ($($cred.UserName))"
}

# ── Data sources ──────────────────────────────────────────────────────────────

Write-Host "`nData sources" -ForegroundColor Cyan

Test-Check "Website repo _index.md" {
    $url = "$(Get-RawBase $config)/content/events/$($config.websiteRepo.eventKey)/_index.md"
    $md  = (Invoke-WebRequest -Uri $url -UseBasicParsing).Content
    if ($md -notmatch '(?m)^title:\s*(.+)$') { throw "fetched but has no title: field" }
    $Matches[1].Trim()
}

Test-Check "Event logo" {
    $logo = Get-EventLogo -Config $config
    if (-not $logo) { throw "could not resolve — deck generators require it" }
    $logo.Path
}

$script:sponsorGroups = $null
Test-Check "sponsors.yaml parses" {
    $script:sponsorGroups = Get-SponsorGroups -Config $config
    $n = ($script:sponsorGroups | ForEach-Object { $_.sponsors.Count } | Measure-Object -Sum).Sum
    "$($script:sponsorGroups.Count) tiers, $n sponsors"
}

Test-Check "All sponsor logos download" {
    if (-not $script:sponsorGroups) { throw "skipped — sponsors.yaml did not parse" }
    $rawBase = Get-RawBase $config
    $failed = @()
    foreach ($group in $script:sponsorGroups) {
        foreach ($sponsor in $group.sponsors) {
            try { $null = Get-WebImage -Url "$rawBase/static/$($sponsor.logo)" }
            catch { $failed += $sponsor.name }
        }
    }
    if ($failed) { throw "failed: $($failed -join ', ')" }
    "all downloaded"
}

$script:sessionizeJson = $null
Test-Check "Sessionize endpoint returns JSON" {
    $url  = "https://sessionize.com/api/v2/$($config.sessionize.eventId)/view/All"
    $resp = Invoke-WebRequest -Uri $url -Method Get
    $body = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
    if ($body.TrimStart() -notmatch '^[\{\[]') {
        throw "endpoint '$($config.sessionize.eventId)' returned JS embed code, not JSON — create a JSON-format endpoint under Embed & API in Sessionize"
    }
    $script:sessionizeJson = $body | ConvertFrom-Json
    "endpoint '$($config.sessionize.eventId)'"
}

Test-Check "Sessionize has published sessions" -WarnOnly {
    if (-not $script:sessionizeJson) { throw "skipped — endpoint check failed" }
    $n = @($script:sessionizeJson.sessions).Count
    if ($n -eq 0) { throw "0 sessions — schedule not published/accepted yet?" }
    "$n sessions"
}

Test-Check "Eventbrite API accepts token" {
    $token = Get-Secret -Name $config.eventbrite.secretName -ErrorAction Stop | ConvertFrom-SecureString -AsPlainText
    $resp  = Invoke-RestMethod -Uri "https://www.eventbriteapi.com/v3/events/$($config.eventbrite.eventId)/" `
                -Headers @{ Authorization = "Bearer $token" } -Method Get
    "$($resp.name.text) ($($resp.status))"
}

# ── Database ──────────────────────────────────────────────────────────────────

Write-Host "`nDatabase" -ForegroundColor Cyan

Test-Check "event.db initialized" -WarnOnly {
    $dbPath = Join-Path $PSScriptRoot ".." $config.database.path
    if (-not (Test-Path $dbPath)) { throw "not found at $dbPath — run Initialize-Database.ps1" }
    Import-Module PSSQLite -ErrorAction Stop
    $tables = (Invoke-SqliteQuery -DataSource $dbPath -Query "SELECT name FROM sqlite_master WHERE type='table'").name
    $missing = @('Attendees', 'ProcessedAttendees', 'PrintedBadges' | Where-Object { $_ -notin $tables })
    if ($missing) { throw "missing table(s): $($missing -join ', ') — run Initialize-Database.ps1" }
    $count = (Invoke-SqliteQuery -DataSource $dbPath -Query "SELECT COUNT(*) AS c FROM Attendees").c
    "$count attendees"
}

Test-Check "event.db outside a cloud-synced folder" -WarnOnly {
    $dbFullPath = Join-Path $PSScriptRoot ".." $config.database.path
    if (Test-Path $dbFullPath) { $dbFullPath = (Resolve-Path $dbFullPath).Path }
    $flagged = @()
    if ($env:OneDrive -and $dbFullPath.StartsWith($env:OneDrive, [StringComparison]::OrdinalIgnoreCase)) { $flagged += "OneDrive ($($env:OneDrive))" }
    if ($dbFullPath -match '\\Dropbox\\') { $flagged += "Dropbox" }
    if ($flagged) {
        throw "repo is inside $($flagged -join ', ') — a sync client can lock/rewrite event.db while it's open and corrupt it (this happened at a past event). Move the repo to a local-only folder before event day."
    }
    "not inside a known sync folder"
}

Test-Check "Azure SQL (shared multi-desk store)" -WarnOnly {
    if (-not $config.PSObject.Properties['azure'] -or -not $config.azure.enabled) {
        throw "azure.enabled is false — running local-only, single-desk mode"
    }
    $ctx = New-DataContext -Config $config
    Test-DatabaseReadiness -DataContext $ctx
}

# ── Summary ───────────────────────────────────────────────────────────────────

$fails = @($script:results | Where-Object Status -eq 'FAIL')
$warns = @($script:results | Where-Object Status -eq 'WARN')

Write-Host ""
Write-Host ("=" * 60)
Write-Host ("{0} checks: {1} passed, {2} warnings, {3} failed" -f $script:results.Count,
    @($script:results | Where-Object Status -eq 'PASS').Count, $warns.Count, $fails.Count) `
    -ForegroundColor $(if ($fails) { 'Red' } elseif ($warns) { 'Yellow' } else { 'Green' })

if ($fails) {
    Write-Host "`nFix before running the pipeline:" -ForegroundColor Red
    $fails | ForEach-Object { Write-Host "  - $($_.Name): $($_.Detail)" -ForegroundColor Red }
}
if ($warns) {
    Write-Host "`nWarnings (day-of items or first-run auto-installs):" -ForegroundColor Yellow
    $warns | ForEach-Object { Write-Host "  - $($_.Name): $($_.Detail)" -ForegroundColor Yellow }
}
if (-not $fails -and -not $warns) {
    Write-Host "`nAll clear — ready for the event." -ForegroundColor Green
}

exit $(if ($fails) { 1 } else { 0 })
