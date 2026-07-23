<#
.SYNOPSIS
    Lists accepted Sessionize speakers/co-presenters who have no matching
    Eventbrite registration yet.
.DESCRIPTION
    Fetches the accepted session list from the Sessionize public API and
    collects every speaker and co-presenter across those sessions, then
    compares each by name (normalized, case/diacritic/punctuation-insensitive)
    against the local Attendees table. The Sessionize public API does not
    expose speaker emails, so matching is name-based — review near-misses
    (nicknames, married names, typos) by hand before assuming someone truly
    isn't registered.
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\List-UnregisteredSpeakers.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

Import-Module PSSQLite

. "$PSScriptRoot\internal\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config

$dbPath = Join-Path $PSScriptRoot ".." $Config.database.path
$sessionizeId = $Config.sessionize.eventId

Write-Host "Fetching accepted sessions from Sessionize ($sessionizeId)..." -ForegroundColor Cyan
$apiUrl = "https://sessionize.com/api/v2/$sessionizeId/view/All"
$resp = Invoke-WebRequest -Uri $apiUrl -Method Get
$body = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())

if ($body.TrimStart() -notmatch '^[\{\[]') {
    throw "Sessionize endpoint '$sessionizeId' returned JavaScript embed code, not JSON. In Sessionize go to 'Embed & API', create (or edit) an endpoint with format 'JSON', and put that endpoint ID in event.config.json (sessionize.eventId)."
}
$data = $body | ConvertFrom-Json

$speakerMap = @{}
foreach ($s in $data.speakers) { $speakerMap[$s.id] = $s.fullName }

# Every speaker/co-presenter attached to a non-service session (accepted talks only)
$sessions = @($data.sessions | Where-Object { -not $_.isServiceSession })
$sessionsBySpeaker = @{}
foreach ($session in $sessions) {
    foreach ($speakerId in $session.speakers) {
        if (-not $sessionsBySpeaker.ContainsKey($speakerId)) { $sessionsBySpeaker[$speakerId] = @() }
        $sessionsBySpeaker[$speakerId] += $session.title
    }
}

function Get-NormalizedName([string]$name) {
    $formD = $name.Normalize([Text.NormalizationForm]::FormD)
    $stripped = -join ($formD.ToCharArray() | Where-Object {
        [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark
    })
    return ($stripped -replace '[^a-zA-Z ]', '').Trim().ToLowerInvariant() -replace '\s+', ' '
}

$attendees = Invoke-SqliteQuery -DataSource $dbPath -Query "SELECT FirstName, LastName, Email FROM Attendees"
$registeredNames = [System.Collections.Generic.HashSet[string]]::new()
foreach ($a in $attendees) {
    [void]$registeredNames.Add((Get-NormalizedName "$($a.FirstName) $($a.LastName)"))
}

$unregistered = @()
foreach ($speakerId in $sessionsBySpeaker.Keys) {
    $name = $speakerMap[$speakerId]
    if (-not $name) { continue }
    if (-not $registeredNames.Contains((Get-NormalizedName $name))) {
        $unregistered += [PSCustomObject]@{
            Name     = $name
            Sessions = ($sessionsBySpeaker[$speakerId] | Select-Object -Unique) -join "; "
        }
    }
}

if ($unregistered.Count -eq 0) {
    Write-Host "All accepted speakers and co-presenters have a matching registration." -ForegroundColor Green
    return
}

$unregistered = $unregistered | Sort-Object Name
Write-Host "$($unregistered.Count) speaker(s)/co-presenter(s) with no matching Eventbrite registration:" -ForegroundColor Yellow
foreach ($u in $unregistered) {
    Write-Host "  $($u.Name)" -ForegroundColor White
    Write-Host "      $($u.Sessions)" -ForegroundColor DarkGray
}
Write-Host "`nMatching is name-based (Sessionize's public API doesn't expose emails) — double-check near-misses before adding tickets." -ForegroundColor Cyan
