<#
.SYNOPSIS
    Shared helpers for fetching event data from the website repo: raw
    GitHub URL base, image download (bytes/base64/mime), and sponsors.yaml
    parsing. Dot-source this before calling any of the functions.
#>

function Get-RawBase {
    param([Parameter(Mandatory)][PSCustomObject]$Config)
    $repo = $Config.websiteRepo
    return "https://raw.githubusercontent.com/$($repo.owner)/$($repo.name)/$($repo.branch)"
}

function Get-WebImage {
    <#
    .SYNOPSIS
        Downloads an image and returns @{ Bytes; Base64; Mime }. Throws on
        download failure so callers decide whether it's fatal.
    #>
    param([Parameter(Mandatory)][string]$Url)

    $raw   = (Invoke-WebRequest -Uri $Url -UseBasicParsing).Content
    $bytes = if ($raw -is [string]) { [System.Text.Encoding]::UTF8.GetBytes($raw) } else { $raw }
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
    $yaml = Invoke-RestMethod -Uri $yamlUrl -Method Get
    return (ConvertFrom-Yaml $yaml).groups
}
