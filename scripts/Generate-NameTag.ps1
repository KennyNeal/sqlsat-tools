<#
.SYNOPSIS
    [STUB] Prints a name tag label on the Brother QL-820NWB for walk-in attendees.
.DESCRIPTION
    This script is a placeholder for Brother QL label printer integration.
    Intended use: look up an attendee by name or email and print a label on the
    Brother QL-820NWB at the check-in desk for walk-ins or "forgot my papers" cases.

    TODO: Set up Brother QL-820NWB templates and driver integration before implementing.
    Reference: https://support.brother.com/g/b/producttop.aspx?c=us&lang=en&prod=lpql820nwbeus

.PARAMETER Config
    Parsed event.config.json object.
.PARAMETER Email
    Look up and print by attendee email address.
.PARAMETER FirstName
    Look up by first name (use with -LastName).
.PARAMETER LastName
    Look up by last name (use with -FirstName).
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config,
    [string]$Email,
    [string]$FirstName,
    [string]$LastName
)

Write-Host "Brother QL-820NWB label printing is not yet implemented." -ForegroundColor Yellow
Write-Host ""
Write-Host "Planned workflow:"
Write-Host "  1. Look up attendee in SQLite by email or name"
Write-Host "  2. Generate vCard QR code"
Write-Host "  3. Render label HTML matching QL-820NWB DK-1201 (29x90mm) template"
Write-Host "  4. Send to printer via b-PAC SDK or Brother iPrint driver"
Write-Host ""
Write-Host "For now, generate a SpeedPass PDF and print on regular paper:"
Write-Host "  .\scripts\Generate-SpeedPasses.ps1 -Config `$config -Email '$Email' -Force"
