<#
.SYNOPSIS
    Shared helpers for the two slide-deck builders (Generate-SlideTemplate.ps1
    and Generate-RaffleDeck.ps1): SVG rasterization via headless Edge, the
    Python dependency check, and downloading sponsor logos into a work
    directory. Requires Web-Helpers.ps1 and Badge-Helpers.ps1 to be
    dot-sourced first (uses Get-RawBase, Get-SponsorGroups, Get-EdgePath).
#>

function ConvertTo-PngFromSvg {
    param(
        [Parameter(Mandatory)][string]$SvgPath,
        [Parameter(Mandatory)][string]$PngPath
    )
    $edge = Get-EdgePath
    & $edge --headless=new --disable-gpu --disable-extensions `
        --window-size=1000,1000 --force-device-scale-factor=2 `
        --default-background-color=00000000 `
        --screenshot="$PngPath" "file:///$SvgPath" 2>&1 | Out-Null

    # Edge writes the screenshot asynchronously as the process shuts down;
    # poll until the file appears and its size stops growing.
    $deadline = (Get-Date).AddSeconds(10)
    $lastSize = -1
    while ((Get-Date) -lt $deadline) {
        if (Test-Path $PngPath) {
            $size = (Get-Item $PngPath).Length
            if ($size -gt 0 -and $size -eq $lastSize) { break }
            $lastSize = $size
        }
        Start-Sleep -Milliseconds 300
    }
}

function Install-PythonSlideDeps {
    $pythonOk = $false
    try {
        & python -c "import pptx, PIL" 2>$null
        if ($LASTEXITCODE -eq 0) { $pythonOk = $true }
    } catch { }
    if (-not $pythonOk) {
        Write-Host "Installing python-pptx and Pillow..." -ForegroundColor Cyan
        & python -m pip install --quiet python-pptx Pillow
    }
}

function Get-SponsorLogoFiles {
    <#
    .SYNOPSIS
        Downloads every sponsor logo into $WorkDir (rasterizing SVGs to PNG,
        since PowerPoint can't embed SVG reliably) and returns the tier-grouped
        list: @{ tier; title; sponsors = [@{ name; logoPath }] }.
    #>
    param(
        [Parameter(Mandatory)][PSCustomObject]$Config,
        [Parameter(Mandatory)][string]$WorkDir
    )

    $rawBase = Get-RawBase $Config
    $groups  = [System.Collections.Generic.List[hashtable]]::new()
    $i = 0
    foreach ($group in (Get-SponsorGroups -Config $Config)) {
        $sponsors = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($sponsor in $group.sponsors) {
            $i++
            $ext = [System.IO.Path]::GetExtension($sponsor.logo)
            $rawLogoPath = Join-Path $WorkDir "sponsor$i$ext"
            try {
                Invoke-WebRequest -Uri "$rawBase/static/$($sponsor.logo)" -OutFile $rawLogoPath -UseBasicParsing
            } catch {
                Write-Host "  Warning: could not download logo for $($sponsor.name)" -ForegroundColor Yellow
                continue
            }

            $finalLogoPath = $rawLogoPath
            if ($ext -eq ".svg") {
                $pngPath = Join-Path $WorkDir "sponsor$i.png"
                ConvertTo-PngFromSvg -SvgPath $rawLogoPath -PngPath $pngPath
                if (Test-Path $pngPath) { $finalLogoPath = $pngPath }
            }

            $sponsors.Add(@{ name = $sponsor.name; logoPath = $finalLogoPath })
            Write-Host "  [$($group.tier)] $($sponsor.name)" -ForegroundColor DarkGray
        }
        $groups.Add(@{ tier = $group.tier; title = $group.title; sponsors = $sponsors })
    }
    return $groups
}
