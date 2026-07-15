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

. (Join-Path $PSScriptRoot "Data-Access.ps1")
$dataContext = New-DataContext -Config $Config

# Load EventBrite token from SecretManagement
try {
    $token = Get-Secret -Name $Config.eventbrite.secretName | ConvertFrom-SecureString -AsPlainText
} catch {
    throw "Could not load EventBrite token from SecretManagement secret '$($Config.eventbrite.secretName)'. Run: Set-Secret -Name '$($Config.eventbrite.secretName)' -Secret '<your-token>'"
}

$headers  = @{ Authorization = "Bearer $token" }
$eventId  = $Config.eventbrite.eventId
$url      = "https://www.eventbriteapi.com/v3/events/$eventId/attendees/"
$statusQs = "status=attending"
$query    = "?$statusQs&expand=answers"

function Get-Answer($answers, $keyword) {
    ($answers | Where-Object { $_.question -like "*$keyword*" } | Select-Object -First 1).answer
}

Write-Host "Fetching attendees from EventBrite event $eventId..." -ForegroundColor Cyan

$all = @()
do {
    $resp    = Invoke-RestMethod -Method Get -Uri ($url + $query) -Headers $headers
    $all    += $resp.attendees
    $query   = if ($resp.pagination.has_more_items) { "?$statusQs&continuation=$($resp.pagination.continuation)" } else { $null }
} while ($query)

Write-Host "  Fetched $($all.Count) attendees" -ForegroundColor Green

$rows = [System.Collections.Generic.List[hashtable]]::new()
foreach ($a in $all) {
    $attendeeProfile = $a.profile
    $barcode = ($a.barcodes | Select-Object -First 1).barcode
    if (-not $barcode) { continue }

    $answers = $a.answers

    $rows.Add(@{
        Barcode        = $barcode
        OrderId        = $a.order_id
        OrderDate      = $a.created
        FirstName      = $attendeeProfile.first_name
        LastName       = $attendeeProfile.last_name
        Email          = $attendeeProfile.email
        Company        = $attendeeProfile.company
        JobTitle       = $attendeeProfile.job_title
        LunchType      = Get-Answer $answers "Lunch"
        TicketType     = $a.ticket_class_name
        AttendeeStatus = $a.status
        IsVolunteer    = if ((Get-Answer $answers "volunteer") -eq "Yes") { 1 } else { 0 }
        TwitterHandle  = Get-Answer $answers "Twitter"
        Website        = $attendeeProfile.website
    })
}

$imported = Import-AttendeesFromEventbrite -DataContext $dataContext -Attendees $rows

Write-Host "Import complete. Upserted $imported attendees." -ForegroundColor Green
