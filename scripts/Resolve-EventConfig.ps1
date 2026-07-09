<#
.SYNOPSIS
    Enriches a config object with live event metadata from the website repo.
.DESCRIPTION
    Fetches content/events/{eventKey}/_index.md from the website repo and merges
    event.name (from title:) into the config object at runtime, so the event
    name has a single source of truth on the website.

    Call this once at the top of any script that needs event.name.
    The config key websiteRepo.eventKey is the only required bootstrap value.

    Note: sessionize.eventId always comes from local config. The website's
    sessionizeId is a JavaScript embed endpoint, which the tools can't use —
    they need the JSON endpoint ID (see the Sessionize error in
    Generate-Schedule.ps1).
.PARAMETER Config
    Parsed event.config.json object.
#>

. "$PSScriptRoot\Web-Helpers.ps1"

function Resolve-EventConfig {
    param([Parameter(Mandatory)][PSCustomObject]$Config)

    $indexUrl = "$(Get-RawBase $Config)/content/events/$($Config.websiteRepo.eventKey)/_index.md"

    Write-Host "Fetching event metadata ($($Config.websiteRepo.eventKey))..." -ForegroundColor Cyan
    try {
        $indexMd = Get-WebText -Url $indexUrl
    } catch {
        throw "Could not fetch event metadata from $indexUrl (and no cached copy in cache\): $_"
    }

    if ($indexMd -match '(?m)^title:\s*(.+)$') {
        $Config.event | Add-Member -NotePropertyName 'name' -NotePropertyValue $Matches[1].Trim() -Force
    }

    return $Config
}
