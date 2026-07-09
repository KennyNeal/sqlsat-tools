<#
.SYNOPSIS
    Imports attendees from EventBrite into the local SQLite database.
.DESCRIPTION
    Fetches all attending registrants from the EventBrite API and upserts them
    into the local database. Safe to run repeatedly — existing records are updated,
    new records are inserted. Does NOT reset email/SpeedPass tracking.
.EXAMPLE
    .\Import-Attendees.ps1 -Config (Get-Content .\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

Import-Module PSSQLite

$dbPath = Join-Path $PSScriptRoot ".." $Config.database.path

# Load EventBrite token from SecretManagement
try {
    $token = Get-Secret -Name $Config.eventbrite.secretName | ConvertFrom-SecureString -AsPlainText
} catch {
    throw "Could not load EventBrite token from SecretManagement secret '$($Config.eventbrite.secretName)'. Run: Set-Secret -Name '$($Config.eventbrite.secretName)' -Secret '<your-token>'"
}

$headers  = @{ Authorization = "Bearer $token" }
$eventId  = $Config.eventbrite.eventId
$url      = "https://www.eventbriteapi.com/v3/events/$eventId/attendees/"
$query    = "?status=attending&expand=answers"

function Get-Answer($answers, $keyword) {
    ($answers | Where-Object { $_.question -like "*$keyword*" } | Select-Object -First 1).answer
}

Write-Host "Fetching attendees from EventBrite event $eventId..." -ForegroundColor Cyan

$all = @()
do {
    $resp    = Invoke-RestMethod -Method Get -Uri ($url + $query) -Headers $headers
    $all    += $resp.attendees
    $query   = if ($resp.pagination.has_more_items) { "?continuation=$($resp.pagination.continuation)" } else { $null }
} while ($query)

Write-Host "  Fetched $($all.Count) attendees" -ForegroundColor Green

$upsertSql = @"
INSERT OR REPLACE INTO Attendees
    (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle,
     LunchType, TicketType, AttendeeStatus, IsVolunteer, TwitterHandle, Website, UpdatedAt)
VALUES
    (@Barcode, @OrderId, @OrderDate, @FirstName, @LastName, @Email, @Company, @JobTitle,
     @LunchType, @TicketType, @AttendeeStatus, @IsVolunteer, @TwitterHandle, @Website, datetime('now'))
"@

# One connection and one transaction for the whole loop — a connection (and
# implicit transaction) per row makes a few-hundred-row import take minutes.
$conn = New-SQLiteConnection -DataSource $dbPath
$imported = 0
try {
    Invoke-SqliteQuery -SQLiteConnection $conn -Query "BEGIN"
    foreach ($a in $all) {
        $profile = $a.profile
        $barcode = ($a.barcodes | Select-Object -First 1).barcode
        if (-not $barcode) { continue }

        $answers = $a.answers

        Invoke-SqliteQuery -SQLiteConnection $conn -Query $upsertSql -SqlParameters @{
            Barcode        = $barcode
            OrderId        = $a.order_id
            OrderDate      = $a.created
            FirstName      = $profile.first_name
            LastName       = $profile.last_name
            Email          = $profile.email
            Company        = $profile.company
            JobTitle       = $profile.job_title
            LunchType      = Get-Answer $answers "Lunch"
            TicketType     = $a.ticket_class_name
            AttendeeStatus = $a.status
            IsVolunteer    = if ((Get-Answer $answers "volunteer") -eq "Yes") { 1 } else { 0 }
            TwitterHandle  = Get-Answer $answers "Twitter"
            Website        = $profile.website
        }
        $imported++
    }
    Invoke-SqliteQuery -SQLiteConnection $conn -Query "COMMIT"
} catch {
    try { Invoke-SqliteQuery -SQLiteConnection $conn -Query "ROLLBACK" } catch { }
    throw
} finally {
    $conn.Close()
    $conn.Dispose()
}

Write-Host "Import complete. Upserted $imported attendees." -ForegroundColor Green
