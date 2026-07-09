<#
.SYNOPSIS
    Shared helpers for badge/label/QR generation, used by every script that
    produces printable output: Edge discovery, QRCoder loading, vCard/QR
    generation, HTML-to-PDF rendering via headless Edge, and silent printing
    via SumatraPDF.
#>

function Get-EdgePath {
    $edgePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
    )
    $edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $edge) { throw "Microsoft Edge not found." }
    return $edge
}

function Import-QRCoder {
    $libPath = Join-Path $PSScriptRoot "..\lib\QRCoder.dll"
    if (-not (Test-Path $libPath)) { throw "QRCoder.dll not found at $libPath." }
    Add-Type -Path $libPath
}

function New-VCard {
    param($FirstName, $LastName, $Email, $Company, $JobTitle, $Website, $TwitterHandle)
    $twitter = if ($TwitterHandle -and -not $TwitterHandle.StartsWith('@')) { "@$TwitterHandle" } else { $TwitterHandle }
    $lines = @(
        "BEGIN:VCARD",
        "VERSION:3.0",
        "N:$LastName;$FirstName",
        "FN:$FirstName $LastName"
    )
    if ($Company)  { $lines += "ORG:$Company" }
    if ($JobTitle) { $lines += "TITLE:$JobTitle" }
    if ($Email)    { $lines += "EMAIL:$Email" }
    if ($Website)  { $lines += "URL:$Website" }
    if ($twitter)  { $lines += "X-SOCIALPROFILE;TYPE=twitter:$twitter" }
    $lines += "END:VCARD"
    return $lines -join "`r`n"
}

function New-QRBase64 {
    param([string]$Data, [int]$PixelSize = 20, [string]$EccLevel = 'L')
    $gen    = New-Object QRCoder.QRCodeGenerator
    $qrData = $gen.CreateQrCode($Data, [QRCoder.QRCodeGenerator+ECCLevel]::$EccLevel)
    $qr     = New-Object QRCoder.QRCode($qrData)
    $bmp    = $qr.GetGraphic($PixelSize)
    $ms     = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $b64    = [Convert]::ToBase64String($ms.ToArray())
    $ms.Dispose(); $bmp.Dispose(); $qr.Dispose(); $gen.Dispose()
    return $b64
}

function ConvertTo-PdfViaEdge {
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath,
        [int]$TimeoutSeconds = 20
    )
    $edge = Get-EdgePath

    if (Test-Path $PdfPath) { Remove-Item $PdfPath -Force }
    $null = & $edge --headless=new --print-to-pdf="$PdfPath" --no-margins "file:///$HtmlPath" --disable-gpu --disable-extensions --no-pdf-header-footer 2>&1

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while (-not (Test-Path $PdfPath) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    if (-not (Test-Path $PdfPath)) { throw "Edge did not produce a PDF within $TimeoutSeconds seconds: $PdfPath" }
}

function Send-ToPrinter {
    <#
    .SYNOPSIS
        Silently prints a PDF to a named Windows printer via SumatraPDF, at
        exact size (no auto-scaling). Requires SumatraPDF on PATH or in the
        default per-user install location.
    #>
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][string]$PrinterName,
        [int]$TimeoutSeconds = 15
    )

    $sumatraCmd = Get-Command SumatraPDF.exe -ErrorAction SilentlyContinue
    $sumatra = if ($sumatraCmd) {
        $sumatraCmd.Source
    } else {
        @("$env:LOCALAPPDATA\SumatraPDF\SumatraPDF.exe", "${env:ProgramFiles}\SumatraPDF\SumatraPDF.exe") |
            Where-Object { Test-Path $_ } | Select-Object -First 1
    }
    if (-not $sumatra) {
        throw "SumatraPDF not found. Install it with: winget install --id SumatraPDF.SumatraPDF -e"
    }

    if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
        throw "Printer '$PrinterName' not found on this system (SumatraPDF silently no-ops on an unknown printer name instead of erroring, so this is checked up front)."
    }

    # A prior stray SumatraPDF instance can intercept new print commands via
    # single-instance IPC instead of actually processing them — clear it first.
    Get-Process -Name SumatraPDF -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 300

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = $sumatra
    $psi.Arguments = "-print-to `"$PrinterName`" -print-settings `"noscale`" -silent -exit-when-done `"$PdfPath`""
    $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        throw "SumatraPDF did not exit within $TimeoutSeconds seconds."
    }

    Start-Sleep -Seconds 2
    $job = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*$PrinterName*" } |
        Sort-Object JobId -Descending | Select-Object -First 1
    if ($job -and $job.Status -eq 'Error') {
        throw "Printer reported an error on the job (check media type/tape loaded in $PrinterName)."
    }
}
