<#
.SYNOPSIS
    Lists walk-ins added via Print-WalkinBadge.ps1 that don't exist in Eventbrite.
.DESCRIPTION
    Eventbrite's API has no way to create a real order/attendee, so quick-added
    walk-ins only ever live in the local database (OrderId = 'WALKIN'). Run this
    periodically (or at the end of the day) to see who still needs a matching
    free/comp order created in Eventbrite, so headcounts and attendee records
    stay in sync.
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\List-UnsyncedWalkins.ps1 -Config (Get-Content .\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

. (Join-Path $PSScriptRoot "Data-Access.ps1")
$dataContext = New-DataContext -Config $Config

$walkins = Get-UnsyncedWalkins -DataContext $dataContext

if ($walkins.Count -eq 0) {
    Write-Host "No unsynced walk-ins." -ForegroundColor Green
    return
}

Write-Host "$($walkins.Count) walk-in(s) not yet registered in Eventbrite:" -ForegroundColor Yellow
foreach ($w in $walkins) {
    Write-Host "  $($w.FirstName) $($w.LastName) <$($w.Email)> — added $($w.OrderDate)"
}
Write-Host "`nCreate a matching free/comp order for each in the Eventbrite dashboard or Box Office app." -ForegroundColor Cyan
