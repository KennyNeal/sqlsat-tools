<#
.SYNOPSIS
    Fetches the event logo from the website repo.
.DESCRIPTION
    Reads the logo: field from content/events/{eventKey}/_index.md,
    or uses an explicit path override. Returns a hashtable with Base64, Mime,
    and Path keys suitable for embedding in HTML, or $null on failure.
.PARAMETER Config
    Parsed event.config.json object.
.PARAMETER Override
    Explicit logo path under /static/ in the website repo. Skips _index.md lookup.
#>

. "$PSScriptRoot\Web-Helpers.ps1"

function Get-EventLogo {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,
        [string]$Override = ""
    )

    $rawBase = Get-RawBase $Config

    $logoPath = if ($Override) {
        $Override
    } else {
        try {
            $indexUrl = "$rawBase/content/events/$($Config.websiteRepo.eventKey)/_index.md"
            $indexMd  = (Invoke-WebRequest -Uri $indexUrl -UseBasicParsing).Content
            if ($indexMd -match '(?m)^logo:\s*(.+)$') { $Matches[1].Trim() } else { $null }
        } catch {
            Write-Host "  Could not read event logo from _index.md: $_" -ForegroundColor Yellow
            $null
        }
    }

    if (-not $logoPath) { return $null }

    try {
        $img = Get-WebImage -Url "$rawBase/static/$logoPath"
        Write-Host "  Loaded event logo: $logoPath" -ForegroundColor DarkGray
        return @{ Base64 = $img.Base64; Mime = $img.Mime; Path = $logoPath }
    } catch {
        Write-Host "  Warning: could not load event logo ($logoPath): $_" -ForegroundColor Yellow
        return $null
    }
}
