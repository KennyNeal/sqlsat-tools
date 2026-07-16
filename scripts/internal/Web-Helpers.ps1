<#
.SYNOPSIS
    Shared helpers for fetching event data from the website repo: raw
    GitHub URL base, image download (bytes/base64/mime), and sponsors.yaml
    parsing. Dot-source this before calling any of the functions.

    Every successful fetch is cached under cache\ (gitignored). When a
    fetch fails — e.g. no internet at the venue — the cached copy is used
    instead, so any script that has run once with connectivity keeps
    working offline with data as fresh as the last successful run.
#>

$script:WebCacheDir = Join-Path $PSScriptRoot "..\..\cache"

function Get-WebCachePath {
    param([Parameter(Mandatory)][string]$Url)
    $sha  = [System.Security.Cryptography.SHA1]::Create()
    $hash = [BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Url))).Replace('-', '').Substring(0, 12).ToLower()
    $sha.Dispose()
    $leaf = [uri]::UnescapeDataString([System.IO.Path]::GetFileName(([uri]$Url).AbsolutePath)) -replace '[^\w.\-]', '_'
    if (-not $leaf) { $leaf = 'index' }
    return Join-Path $script:WebCacheDir "$hash-$leaf"
}

function Get-WebBytes {
    <#
    .SYNOPSIS
        Downloads a URL as raw bytes. Caches on success; falls back to the
        cached copy (with a warning) when the download fails. Throws only
        when the download fails and there is no cached copy.
    #>
    param([Parameter(Mandatory)][string]$Url)

    $cachePath = Get-WebCachePath $Url
    try {
        $raw   = (Invoke-WebRequest -Uri $Url -UseBasicParsing).Content
        $bytes = if ($raw -is [string]) { [System.Text.Encoding]::UTF8.GetBytes($raw) } else { $raw }
        if (-not (Test-Path $script:WebCacheDir)) { New-Item -ItemType Directory -Path $script:WebCacheDir | Out-Null }
        [System.IO.File]::WriteAllBytes($cachePath, $bytes)
        return $bytes
    } catch {
        if (Test-Path $cachePath) {
            Write-Host "  Offline? Using cached copy of $Url" -ForegroundColor Yellow
            return [System.IO.File]::ReadAllBytes($cachePath)
        }
        throw
    }
}

function Get-WebText {
    <# Get-WebBytes decoded as UTF-8, for text resources (_index.md, sponsors.yaml). #>
    param([Parameter(Mandatory)][string]$Url)
    return [System.Text.Encoding]::UTF8.GetString((Get-WebBytes -Url $Url))
}

function Get-RawBase {
    param([Parameter(Mandatory)][PSCustomObject]$Config)
    $repo = $Config.websiteRepo
    return "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
}

function Get-WebImage {
    <#
    .SYNOPSIS
        Downloads an image and returns @{ Bytes; Base64; Mime }. Throws on
        download failure (after the cache fallback) so callers decide
        whether it's fatal.
    #>
    param([Parameter(Mandatory)][string]$Url)

    $bytes = Get-WebBytes -Url $Url
    $ext   = [System.IO.Path]::GetExtension($Url.Split('?')[0]).TrimStart('.')
    $mime  = switch ($ext) {
        'svg'  { 'image/svg+xml' }
        'jpg'  { 'image/jpeg' }
        'jpeg' { 'image/jpeg' }
        default { "image/$ext" }
    }
    return @{ Bytes = $bytes; Base64 = [Convert]::ToBase64String($bytes); Mime = $mime }
}

function Get-SponsorGroups {
    <#
    .SYNOPSIS
        Fetches and parses sponsors.yaml from the website repo. Installs the
        powershell-yaml module on first use. Returns the raw groups list
        (tier/title/sponsors); callers do their own tier filtering.
    #>
    param([Parameter(Mandatory)][PSCustomObject]$Config)

    if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
        Write-Host "Installing powershell-yaml module..." -ForegroundColor Cyan
        Install-Module -Name powershell-yaml -Scope CurrentUser -Force
    }
    Import-Module powershell-yaml

    $yamlUrl = "$(Get-RawBase $Config)/content/events/$($Config.websiteRepo.eventKey)/sponsors.yaml"
    Write-Host "  Fetching sponsor data from $yamlUrl" -ForegroundColor Cyan
    $yaml = Get-WebText -Url $yamlUrl
    return (ConvertFrom-Yaml $yaml).groups
}
