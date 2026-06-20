<#
.SYNOPSIS
    Enriches a config object with live event metadata from the website repo.
.DESCRIPTION
    Fetches content/events/{eventKey}/_index.md from the website repo and merges
    the following fields into the config object at runtime:
      event.name         ← title
      sessionize.eventId ← sessionizeId

    Call this once at the top of any script that needs these fields.
    The config key websiteRepo.eventKey is the only required bootstrap value.
.PARAMETER Config
    Parsed event.config.json object.
#>
function Resolve-EventConfig {
    param([Parameter(Mandatory)][PSCustomObject]$Config)

    $repo    = $Config.websiteRepo
    $rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
    $indexUrl = "$rawBase/content/events/$($repo.eventKey)/_index.md"

    Write-Host "Fetching event metadata ($($repo.eventKey))..." -ForegroundColor Cyan
    try {
        $indexMd = (Invoke-WebRequest -Uri $indexUrl -UseBasicParsing).Content
    } catch {
        throw "Could not fetch event metadata from $indexUrl : $_"
    }

    function Get-Field([string]$key) {
        if ($indexMd -match "(?m)^${key}:\s*(.+)$") { return $Matches[1].Trim() }
        return $null
    }

    $title        = Get-Field 'title'
    $sessionizeId = Get-Field 'sessionizeId'

    if ($title) { $Config.event | Add-Member -NotePropertyName 'name' -NotePropertyValue $title -Force }

    if ($sessionizeId) {
        if ($Config.PSObject.Properties['sessionize']) {
            $Config.sessionize | Add-Member -NotePropertyName 'eventId' -NotePropertyValue $sessionizeId -Force
        } else {
            $Config | Add-Member -NotePropertyName 'sessionize' -NotePropertyValue ([PSCustomObject]@{ eventId = $sessionizeId })
        }
    }

    return $Config
}
