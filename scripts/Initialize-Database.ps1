<#
.SYNOPSIS
    Creates the SQLite database and schema for event attendee tracking.
.DESCRIPTION
    Run once per event to initialize the database. Safe to re-run — uses
    CREATE TABLE IF NOT EXISTS so existing data is preserved.
.EXAMPLE
    .\Initialize-Database.ps1 -Config (Get-Content .\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

$dbPath = Join-Path $PSScriptRoot ".." $Config.database.path

if (-not (Get-Module -ListAvailable -Name PSSQLite)) {
    Write-Host "Installing PSSQLite module..." -ForegroundColor Cyan
    Install-Module -Name PSSQLite -Scope CurrentUser -Force
}
Import-Module PSSQLite

Invoke-SqliteQuery -DataSource $dbPath -Query @"
CREATE TABLE IF NOT EXISTS Attendees (
    Barcode         TEXT PRIMARY KEY,
    OrderId         TEXT,
    OrderDate       TEXT,
    FirstName       TEXT NOT NULL,
    LastName        TEXT NOT NULL,
    Email           TEXT NOT NULL,
    Company         TEXT,
    JobTitle        TEXT,
    LunchType       TEXT,
    TicketType      TEXT,
    AttendeeStatus  TEXT,
    IsVolunteer     INTEGER DEFAULT 0,
    TwitterHandle   TEXT,
    Website         TEXT,
    ImportedAt      TEXT DEFAULT (datetime('now')),
    UpdatedAt       TEXT DEFAULT (datetime('now'))
)
"@

Invoke-SqliteQuery -DataSource $dbPath -Query @"
CREATE TABLE IF NOT EXISTS ProcessedAttendees (
    Barcode              TEXT PRIMARY KEY,
    SpeedPassPath        TEXT,
    SpeedPassGeneratedAt TEXT,
    EmailedAt            TEXT
)
"@

Write-Host "Database ready: $dbPath" -ForegroundColor Green
