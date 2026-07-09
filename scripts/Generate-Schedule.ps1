<#
.SYNOPSIS
    Generates a printable HTML schedule from Sessionize.
.DESCRIPTION
    Fetches the full event schedule from the Sessionize public API and renders
    a print-friendly HTML file in the classic layout: time slots as rows,
    rooms as columns, with the rooms split in half across two pages so the
    whole schedule prints on one double-sided letter sheet (landscape).

    Service sessions (keynote, lunch, raffle) that have a start time render as
    full-width rows. Level/Track category badges render when present.

    Optional config keys under "schedule":
      appUrl   - URL for the "Use the App" QR code in the header
      logoFile - override logo path under /static/ in the website repo;
                 by default the logo comes from the event's _index.md
.PARAMETER Config
    Parsed event.config.json object.
.EXAMPLE
    .\Generate-Schedule.ps1 -Config (Get-Content ..\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

. "$PSScriptRoot\Resolve-EventConfig.ps1"
$Config = Resolve-EventConfig -Config $Config
. "$PSScriptRoot\Badge-Helpers.ps1"

$sessionizeId = $Config.sessionize.eventId
$outputFile   = Join-Path $PSScriptRoot ".." $Config.schedule.outputFile
$outputDir    = Split-Path $outputFile
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

Write-Host "Fetching schedule from Sessionize ($sessionizeId)..." -ForegroundColor Cyan
$apiUrl = "https://sessionize.com/api/v2/$sessionizeId/view/All"
# Decode the raw bytes as UTF-8 ourselves; Invoke-RestMethod mis-decodes when
# the response omits a charset, which garbles curly quotes and accents.
$resp = Invoke-WebRequest -Uri $apiUrl -Method Get
$body = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())

if ($body.TrimStart() -notmatch '^[\{\[]') {
    throw "Sessionize endpoint '$sessionizeId' returned JavaScript embed code, not JSON. In Sessionize go to 'Embed & API', create (or edit) an endpoint with format 'JSON', and put that endpoint ID in event.config.json (sessionize.eventId)."
}
$data = $body | ConvertFrom-Json
if (-not $data.sessions -or $data.sessions.Count -eq 0) {
    throw "Sessionize endpoint '$sessionizeId' returned no sessions. Check that the schedule is published and sessions are accepted/informed in Sessionize."
}

# Build lookup tables
$speakerMap = @{}
foreach ($s in $data.speakers) { $speakerMap[$s.id] = $s.fullName }

# Only sessions with a time assigned; regular sessions also need a room
$scheduled   = @($data.sessions | Where-Object { -not $_.isServiceSession -and $_.roomId -and $_.startsAt })
$unscheduled = @($data.sessions | Where-Object { -not $_.isServiceSession -and (-not $_.roomId -or -not $_.startsAt) })
if ($unscheduled.Count -gt 0) {
    Write-Warning "Skipping $($unscheduled.Count) session(s) with no room/time assigned in Sessionize:"
    foreach ($u in $unscheduled) { Write-Warning "  - $($u.title)" }
}

# Keep only the main event day (the date with the most sessions) — drops
# pre-con/workshop days like Friday.
$byDate  = $scheduled | Group-Object { ([datetime]$_.startsAt).Date }
$mainDay = [datetime]($byDate | Sort-Object Count -Descending | Select-Object -First 1).Name
foreach ($g in $byDate | Where-Object { [datetime]$_.Name -ne $mainDay }) {
    Write-Host "Excluding $($g.Count) session(s) on $(([datetime]$g.Name).ToString('dddd M/d')) (not the main event day)." -ForegroundColor Yellow
}
$scheduled = @($scheduled | Where-Object { ([datetime]$_.startsAt).Date -eq $mainDay })

# Room lookup for service session locations
$roomMap = @{}
foreach ($r in $data.rooms) { $roomMap[$r.id] = $r.name }

# Service sessions (check-in, lunch, raffle...) are only exposed by the Grid
# views, not /view/All. Plenum ones render as full-width rows; room-scoped
# ones are placed in their room's cell like a regular session.
$gridResp = Invoke-WebRequest -Uri "https://sessionize.com/api/v2/$sessionizeId/view/GridSmart" -Method Get
$gridDays = [System.Text.Encoding]::UTF8.GetString($gridResp.RawContentStream.ToArray()) | ConvertFrom-Json

$serviceSlots   = @{}   # start time -> full-width service session
$serviceInRoom  = @()   # room-scoped service sessions, merged into the grid below
foreach ($day in $gridDays | Where-Object { ([datetime]$_.date).Date -eq $mainDay }) {
    foreach ($ts in $day.timeSlots) {
        foreach ($roomSlot in $ts.rooms) {
            $svc = $roomSlot.session
            if (-not $svc.isServiceSession) { continue }
            if ($svc.isPlenumSession) {
                $serviceSlots[([datetime]$svc.startsAt)] = [PSCustomObject]@{
                    title = $svc.title; roomName = $roomSlot.name
                    speaker = $null; endsAt = $svc.endsAt
                }
            } else {
                $serviceInRoom += [PSCustomObject]@{
                    title = $svc.title; speakers = @(); categories = @()
                    startsAt = $svc.startsAt; roomId = $roomSlot.id
                }
            }
        }
    }
}

# Time slots in chronological order; rooms in Sessionize order, only those in use
$slots = @($scheduled | ForEach-Object { [datetime]$_.startsAt }) +
         @($serviceInRoom | ForEach-Object { [datetime]$_.startsAt }) +
         @($serviceSlots.Keys) | Sort-Object -Unique
$usedRoomIds = $scheduled | Select-Object -ExpandProperty roomId -Unique
$rooms = @($data.rooms | Where-Object { $usedRoomIds -contains $_.id })

# Grid lookup: room + start time -> sessions (a room/slot cell may hold several
# short sessions, e.g. lightning talks)
$grid = @{}
foreach ($session in @($scheduled) + @($serviceInRoom)) {
    $key = "$($session.roomId)|$(([datetime]$session.startsAt).Ticks)"
    if ($grid.ContainsKey($key)) { $grid[$key] += @($session) } else { $grid[$key] = @($session) }
}

# Promote any slot with exactly one session event-wide (e.g. the keynote) to a
# full-width row instead of one lonely cell beside a band of empty ones.
foreach ($slot in $slots | Where-Object { -not $serviceSlots.ContainsKey($_) }) {
    $slotSessions = @( @($scheduled) + @($serviceInRoom) | Where-Object { [datetime]$_.startsAt -eq $slot } )
    if ($slotSessions.Count -eq 1) {
        $solo = $slotSessions[0]
        $serviceSlots[$slot] = [PSCustomObject]@{
            title    = $solo.title
            roomName = $roomMap[$solo.roomId]
            speaker  = (($solo.speakers | ForEach-Object { $speakerMap[$_] }) -join ", ")
            endsAt   = $solo.endsAt
        }
    }
}

# End time per slot (latest end among that slot's sessions) for the time column
$slotEnds = @{}
foreach ($slot in $slots) {
    $ends = @( @($scheduled) + @($serviceInRoom) | Where-Object { $_.endsAt -and [datetime]$_.startsAt -eq $slot } |
        ForEach-Object { [datetime]$_.endsAt } )
    if ($serviceSlots.ContainsKey($slot) -and $serviceSlots[$slot].endsAt) { $ends += [datetime]$serviceSlots[$slot].endsAt }
    if ($ends.Count -gt 0) { $slotEnds[$slot] = ($ends | Sort-Object | Select-Object -Last 1) }
}

$eventName = $Config.event.name
$hashtag   = $Config.event.hashtag
$dateStr   = $mainDay.ToString("MMMM d, yyyy")
$genStr    = (Get-Date).ToString("MMMM d, yyyy")
$appUrl    = $Config.schedule.appUrl

. "$PSScriptRoot\Get-EventLogo.ps1"
$eventLogo = Get-EventLogo -Config $Config -Override $Config.schedule.logoFile

Add-Type -AssemblyName System.Web
function Enc([string]$s) { [System.Web.HttpUtility]::HtmlEncode($s) }

function New-SessionCell($sessions) {
    $blocks = foreach ($session in $sessions) {
        $speakers = ($session.speakers | ForEach-Object { $speakerMap[$_] }) -join ", "
        $level = $session.categories | Where-Object { $_.name -eq "Level" } |
                 Select-Object -ExpandProperty categoryItems -First 1 |
                 Select-Object -ExpandProperty name -First 1
        $track = $session.categories | Where-Object { $_.name -eq "Track" } |
                 Select-Object -ExpandProperty categoryItems -First 1 |
                 Select-Object -ExpandProperty name -First 1
        @"
<div class="session-block">
  <div class="session-title">$(Enc $session.title)</div>
  <div class="session-speaker">$(Enc $speakers)</div>
  $(if ($level) { "<span class='session-level'>Level: $(Enc $level)</span>" })
  $(if ($track) { "<span class='session-track'>Track: $(Enc $track)</span>" })
  $(if ($sessions.Count -gt 1) { "<div class='session-time'>$(([datetime]$session.startsAt).ToString('h:mm tt'))</div>" })
</div>
"@
    }
    return "<td class='session-cell'>$($blocks -join '')</td>"
}

function New-ScheduleTable($roomGroup) {
    # Room headers split into room number + track ("BEC 1220 (Analytics)" -> two lines)
    $header = ($roomGroup | ForEach-Object {
        if ($_.name -match '^(.*?)\s*\((.+)\)$') {
            "<th class='room-column'><div class='room-name'>$(Enc $Matches[1])</div><div class='room-track'>$(Enc $Matches[2])</div></th>"
        } else {
            "<th class='room-column'><div class='room-name'>$(Enc $_.name)</div></th>"
        }
    }) -join "`n"
    $rows = ""
    foreach ($slot in $slots) {
        $timeHtml = $slot.ToString("h:mm tt")
        if ($slotEnds.ContainsKey($slot)) {
            $timeHtml += "<div class='time-end'>&ndash; $($slotEnds[$slot].ToString('h:mm tt'))</div>"
        }
        $svc = $serviceSlots[$slot]
        if ($svc) {
            $where   = if ($svc.roomName) { "<div class='service-room'>$(Enc $svc.roomName)</div>" }
            $speaker = if ($svc.speaker)  { "<div class='service-speaker'>$(Enc $svc.speaker)</div>" }
            $rows += @"
<tr class="service-row">
  <td class="time-cell">$timeHtml</td>
  <td class="service-cell" colspan="$($roomGroup.Count)">
    <div class="service-title">$(Enc $svc.title)</div>
    $speaker
    $where
  </td>
</tr>
"@
            continue
        }
        $cells = foreach ($room in $roomGroup) {
            $sessions = $grid["$($room.id)|$($slot.Ticks)"]
            if ($sessions) { New-SessionCell $sessions } else { "<td class='empty-cell'>&mdash;</td>" }
        }
        $rows += @"
<tr>
  <td class="time-cell">$timeHtml</td>
  $($cells -join "`n")
</tr>
"@
    }
    return @"
<table class="schedule-table">
  <thead><tr><th class="time-column">Time</th>$header</tr></thead>
  <tbody>$rows</tbody>
</table>
"@
}

# Split rooms in half: first half on the front page, rest on the back
$half  = [math]::Ceiling($rooms.Count / 2)
$front = $rooms[0..($half - 1)]
$back  = if ($rooms.Count -gt $half) { $rooms[$half..($rooms.Count - 1)] } else { @() }

$qrHtml = if ($appUrl) {
    # Generated locally with the bundled QRCoder.dll so the printed schedule
    # doesn't depend on an external QR web service being up at render time.
    Import-QRCoder
    $qrB64 = New-QRBase64 -Data $appUrl -PixelSize 10
    @"
<div class="header-qr">
  <img src="data:image/png;base64,$qrB64" alt="App QR Code" class="qr-code">
  <div class="qr-text">Use the App</div>
</div>
"@
} else { "" }

$logoHtml = if ($eventLogo) { "<img src=`"data:$($eventLogo.Mime);base64,$($eventLogo.Base64)`" alt=`"$(Enc $eventName) Logo`" class=`"header-logo`">" } else { "" }

$headerHtml = @"
<div class="header">
  $qrHtml
  <div class="header-content">
    <h1>$(Enc $eventName)</h1>
    <div class="date">$dateStr &nbsp;|&nbsp; $(Enc $hashtag)</div>
    <div class="fineprint">Schedule generated $genStr and is subject to change</div>
  </div>
  $logoHtml
</div>
"@

function Get-ShortRoomNames($roomGroup) {
    ($roomGroup | ForEach-Object { $_.name -replace '\s*\(.+\)$', '' }) -join ', '
}

$backTable = if ($back.Count -gt 0) {
    @"
<div class="day-section">
$headerHtml
$(New-ScheduleTable $back)
<div class='continued'>&#10148; Rooms $(Enc (Get-ShortRoomNames $front)) on other side</div>
</div>
"@
} else { "" }

$continued = if ($back.Count -gt 0) { "<div class='continued'>&#10148; Rooms $(Enc (Get-ShortRoomNames $back)) on other side</div>" } else { "" }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>$(Enc $eventName) Schedule</title>
<style>
  @page { size: Letter landscape; margin: .35in }
  body  { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 10px; line-height: 1.2; margin: 0; color: #333 }
  .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #013169; padding-bottom: 6px; margin-bottom: 8px }
  .header-content { text-align: center; flex: 1 }
  .header h1 { margin: 0; font-size: 16px; color: #013169 }
  .header .date { font-size: 10px; color: #495057; margin: 2px 0 }
  .header .fineprint { font-size: 7px; color: #666; margin-top: 3px }
  .header-qr { display: flex; flex-direction: column; align-items: center; margin-right: 10px }
  .qr-code { width: 60px; height: 60px; border: 1px solid #013169 }
  .qr-text { font-size: 7px; color: #013169; font-weight: bold; margin-top: 2px }
  .header-logo { height: 60px; width: auto; max-width: 100px }
  .day-section { page-break-after: always }
  .day-section:last-child { page-break-after: auto }
  .schedule-table { width: 100%; border-collapse: collapse; border: 2px solid #013169; table-layout: fixed; font-size: 9px }
  .schedule-table th { background: #013169; color: white; padding: 6px 4px; text-align: center; border: 1px solid #013169; line-height: 1.15 }
  .time-column { width: 55px }
  .schedule-table td { border: 1px solid #ccc; padding: 4px 3px; vertical-align: top; line-height: 1.15 }
  .schedule-table tr:nth-child(even) td { background: #f8f9fa }
  .time-cell { font-weight: bold; text-align: center; color: #013169; white-space: nowrap; font-size: 10px; vertical-align: middle !important; border-right: 2px solid #013169 }
  .time-end  { font-size: 7px; font-weight: normal; color: #666 }
  .room-name  { font-size: 9px }
  .room-track { font-size: 7px; font-weight: normal; text-transform: uppercase; letter-spacing: .03em; opacity: .85; margin-top: 1px }
  td.empty-cell { text-align: center; vertical-align: middle; color: #ccc }
  .session-title   { font-weight: bold; margin-bottom: 2px }
  .session-speaker { color: #666; font-style: italic; font-size: 8px }
  .session-level, .session-track { font-size: 7px; font-weight: bold; text-transform: uppercase; padding: 1px 4px; border-radius: 2px; display: inline-block; margin-top: 2px }
  .session-level { color: #013169; background: #01316920; border: 1px solid #013169 }
  .session-track { color: #b57714; background: #e8a33d20; border: 1px solid #e8a33d; margin-left: 4px }
  .session-time  { color: #013169; font-weight: bold; font-size: 7px; margin-top: 2px }
  .session-block { border-bottom: 1px solid #ddd; margin-bottom: 4px; padding-bottom: 4px }
  .session-block:last-child { border-bottom: none; margin-bottom: 0; padding-bottom: 0 }
  .service-row td { background: #e8a33d20 !important }
  .service-cell { border: 2px solid #013169; text-align: center }
  .service-title { font-weight: bold; color: #013169; font-size: 10px }
  .service-speaker { color: #013169; font-style: italic; font-size: 8px }
  .service-room  { color: #013169; font-size: 8px; font-weight: bold }
  .continued { text-align: center; margin-top: 8px; font-size: 8px; color: #666; font-style: italic }
</style>
</head><body>
<div class="day-section">
$headerHtml
$(New-ScheduleTable $front)
$continued
</div>
$backTable
</body></html>
"@

Set-Content -Path $outputFile -Value $html -Encoding UTF8
Write-Host "Schedule written to: $outputFile" -ForegroundColor Green
Write-Host "Rooms $(($front | Select-Object -ExpandProperty name) -join ', ') on the front; $(($back | Select-Object -ExpandProperty name) -join ', ') on the back." -ForegroundColor Cyan
Write-Host "Print landscape, double-sided (flip on long edge)." -ForegroundColor Cyan
