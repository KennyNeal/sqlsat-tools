<#
.SYNOPSIS
    Generates a printable HTML schedule from Sessionize.
.DESCRIPTION
    Fetches the full event schedule from the Sessionize public API and renders
    a print-friendly HTML file grouped by time slot and room.
    Open the HTML in a browser and print landscape to get the paper schedule.
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\Generate-Schedule.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

$sessionizeId = $Config.sessionize.eventId
$outputFile   = Join-Path $PSScriptRoot ".." $Config.schedule.outputFile
$outputDir    = Split-Path $outputFile
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

Write-Host "Fetching schedule from Sessionize ($sessionizeId)..." -ForegroundColor Cyan
$apiUrl = "https://sessionize.com/api/v2/$sessionizeId/view/All"
$data   = Invoke-RestMethod -Uri $apiUrl -Method Get

# Build lookup tables
$speakerMap = @{}
foreach ($s in $data.speakers) { $speakerMap[$s.id] = $s.fullName }

$roomMap = @{}
foreach ($r in $data.rooms) { $roomMap[$r.id] = $r.name }

# Group sessions by start time
$bySlot = $data.sessions | Where-Object { -not $_.isServiceSession } |
    Group-Object -Property { $_.startsAt } |
    Sort-Object -Property Name

$eventName = $Config.event.name
$hashtag   = $Config.event.hashtag

$rows = ""
foreach ($slot in $bySlot) {
    $time     = [datetime]$slot.Name
    $timeStr  = $time.ToString("h:mm tt")
    $sessions = $slot.Group | Sort-Object { $roomMap[$_.roomId] }
    $cells    = ""

    foreach ($session in $sessions) {
        $room     = $roomMap[$session.roomId]
        $speakers = ($session.speakers | ForEach-Object { $speakerMap[$_] }) -join ", "
        $track    = $session.categories | Where-Object { $_.name -eq "Level" } |
                    Select-Object -ExpandProperty categoryItems -First 1 |
                    Select-Object -ExpandProperty name -First 1
        $cells += @"
<td>
  <div class="session-title">$([System.Web.HttpUtility]::HtmlEncode($session.title))</div>
  <div class="session-speaker">$([System.Web.HttpUtility]::HtmlEncode($speakers))</div>
  $(if ($track) { "<div class='session-track'>$([System.Web.HttpUtility]::HtmlEncode($track))</div>" })
  <div class="session-room">$([System.Web.HttpUtility]::HtmlEncode($room))</div>
</td>
"@
    }

    $rows += @"
<tr>
  <td class="time-col">$timeStr</td>
  $cells
</tr>
"@
}

# Collect unique rooms for header
$rooms = $data.sessions | Where-Object { -not $_.isServiceSession } |
    ForEach-Object { $roomMap[$_.roomId] } | Select-Object -Unique | Sort-Object
$headerCells = ($rooms | ForEach-Object { "<th>$_</th>" }) -join ""

Add-Type -AssemblyName System.Web

$html = @"
<!DOCTYPE html>
<html><head>
<meta charset="utf-8"/>
<title>$eventName Schedule</title>
<style>
  @page { size: Letter landscape; margin: .3in }
  body  { font-family: Arial, sans-serif; font-size: 8pt; margin: 0 }
  h2    { text-align: center; margin: 0 0 .1in; font-size: 12pt }
  table { width: 100%; border-collapse: collapse; table-layout: fixed }
  th, td { border: 1px solid #ccc; padding: .05in .07in; vertical-align: top }
  th    { background: #013169; color: white; text-align: center; font-size: 8pt }
  .time-col { width: .6in; font-weight: bold; white-space: nowrap; vertical-align: middle; text-align: center }
  .session-title   { font-weight: bold; margin-bottom: 2px }
  .session-speaker { color: #555; font-style: italic }
  .session-track   { color: #e8a33d; font-size: 7pt }
  .session-room    { color: #888; font-size: 7pt }
  tr:nth-child(even) td { background: #f9f9f9 }
</style>
</head><body>
<h2>$eventName &nbsp;|&nbsp; $hashtag</h2>
<table>
  <thead><tr><th>Time</th>$headerCells</tr></thead>
  <tbody>$rows</tbody>
</table>
</body></html>
"@

Set-Content -Path $outputFile -Value $html -Encoding UTF8
Write-Host "Schedule written to: $outputFile" -ForegroundColor Green
Write-Host "Open in a browser and print landscape to generate the paper schedule." -ForegroundColor Cyan
