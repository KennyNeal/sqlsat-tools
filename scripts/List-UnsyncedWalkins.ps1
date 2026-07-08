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

Import-Module PSSQLite

$dbPath = Join-Path $PSScriptRoot ".." $Config.database.path

$walkins = Invoke-SqliteQuery -DataSource $dbPath -Query @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle, a.OrderDate, p.PrintedAt
FROM   Attendees a
LEFT JOIN PrintedBadges p ON p.Barcode = a.Barcode
WHERE  a.OrderId = 'WALKIN'
ORDER  BY a.OrderDate
"@

if ($walkins.Count -eq 0) {
    Write-Host "No unsynced walk-ins." -ForegroundColor Green
    return
}

Write-Host "$($walkins.Count) walk-in(s) not yet registered in Eventbrite:" -ForegroundColor Yellow
foreach ($w in $walkins) {
    Write-Host "  $($w.FirstName) $($w.LastName) <$($w.Email)> — added $($w.OrderDate)"
}
Write-Host "`nCreate a matching free/comp order for each in the Eventbrite dashboard or Box Office app." -ForegroundColor Cyan
