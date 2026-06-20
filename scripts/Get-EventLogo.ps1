<#
.SYNOPSIS
    Fetches the event logo from the website repo.
.DESCRIPTION
    Reads the logo: field from content/events/{sponsorDataFile}/_index.md,
    or uses an explicit path override. Returns a hashtable with Base64, Mime,
    and Path keys suitable for embedding in HTML, or $null on failure.
.PARAMETER Config
    Parsed event.config.json object.
.PARAMETER Override
    Explicit logo path under /static/ in the website repo. Skips _index.md lookup.
#>
function Get-EventLogo {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,
        [string]$Override = ""
    )

    $repo    = $Config.websiteRepo
    $rawBase = "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"

    $logoPath = if ($Override) {
        $Override
    } else {
        try {
            $indexUrl = "$rawBase/content/events/$($repo.eventKey)/_index.md"
            $indexMd  = (Invoke-WebRequest -Uri $indexUrl -UseBasicParsing).Content
            if ($indexMd -match '(?m)^logo:\s*(.+)$') { $Matches[1].Trim() } else { $null }
        } catch {
            Write-Host "  Could not read event logo from _index.md: $_" -ForegroundColor Yellow
            $null
        }
    }

    if (-not $logoPath) { return $null }

    try {
        $raw   = (Invoke-WebRequest -Uri "$rawBase/static/$logoPath" -UseBasicParsing).Content
        $bytes = if ($raw -is [string]) { [System.Text.Encoding]::UTF8.GetBytes($raw) } else { $raw }
        $b64   = [Convert]::ToBase64String($bytes)
        $ext   = [System.IO.Path]::GetExtension($logoPath).TrimStart('.')
        $mime  = switch ($ext) { 'svg' { 'image/svg+xml' } 'jpg' { 'image/jpeg' } 'jpeg' { 'image/jpeg' } default { "image/$ext" } }
        Write-Host "  Loaded event logo: $logoPath" -ForegroundColor DarkGray
        return @{ Base64 = $b64; Mime = $mime; Path = $logoPath }
    } catch {
        Write-Host "  Warning: could not load event logo ($logoPath): $_" -ForegroundColor Yellow
        return $null
    }
}
