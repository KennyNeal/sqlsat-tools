<#
.SYNOPSIS
    Emails SpeedPass PDFs to attendees who haven't received one yet.
.DESCRIPTION
    Reads the database for attendees with a generated SpeedPass but no email sent,
    then sends personalized emails with the PDF attached. Marks each attendee as
    emailed in the database after successful delivery.

    Supports -WhatIf to preview without sending.
.PARAMETER Config
    Parsed event.config.json object (passed by Update-Event.ps1).
.PARAMETER TestEmail
    Send all emails to this address instead of the real attendee address.
.PARAMETER ShowBanner
    Prepend a re-send warning banner to the email body.
.EXAMPLE
    .\Send-SpeedPasses.ps1 -Config $config -WhatIf
.EXAMPLE
    .\Send-SpeedPasses.ps1 -Config $config -TestEmail "me@example.com"
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config,
    [string]$TestEmail,
    [switch]$ShowBanner
)

Import-Module PSSQLite

$dbPath          = Join-Path $PSScriptRoot ".." $Config.database.path
$templateFolder  = Join-Path $PSScriptRoot "..\templates"
$batchSize       = $Config.email.batchSize
$delaySeconds    = $Config.email.delaySeconds

# Load email credentials
try {
    $cred = Get-Secret -Name $Config.email.secretName -AsPlainText:$false -ErrorAction Stop
    $from = $cred.UserName
} catch {
    throw "Could not load Gmail credentials from secret '$($Config.email.secretName)'. Run: Set-Secret -Name '$($Config.email.secretName)' -Secret (Get-Credential)"
}

# Load email template
$templatePath = Join-Path $templateFolder "attendee-email.html"
if (-not (Test-Path $templatePath)) { throw "Email template not found: $templatePath" }
$bodyTemplate = Get-Content $templatePath -Raw

$bannerHtml = if ($ShowBanner) {
    '<p style="background:#fff3cd;border:1px solid #ffc107;padding:10px"><strong>Note:</strong> This is a re-send of your SpeedPass.</p>'
} else { "" }

# Find attendees ready to email
$attendees = Invoke-SqliteQuery -DataSource $dbPath -Query @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, p.SpeedPassPath
FROM   Attendees a
JOIN   ProcessedAttendees p ON a.Barcode = p.Barcode
WHERE  p.SpeedPassGeneratedAt IS NOT NULL
AND    p.EmailedAt IS NULL
ORDER  BY a.LastName, a.FirstName
"@

if ($attendees.Count -eq 0) {
    Write-Host "No attendees pending email." -ForegroundColor Yellow
    return
}

Write-Host "Found $($attendees.Count) attendees to email." -ForegroundColor Cyan
if ($TestEmail) { Write-Host "TEST MODE: all emails → $TestEmail" -ForegroundColor Yellow }

if ($WhatIfPreference) {
    Write-Host "`nWHATIF — would send:"
    $attendees | ForEach-Object { Write-Host "  $($_.Email) ← $([System.IO.Path]::GetFileName($_.SpeedPassPath))" }
    Write-Host "`nTo send for real, remove -WhatIf"
    return
}

$smtp    = "smtp.gmail.com"
$port    = 587
$subject = $Config.email.subject
$success = 0
$errors  = 0

for ($i = 0; $i -lt $attendees.Count; $i += $batchSize) {
    $batch     = $attendees[$i..[Math]::Min($i + $batchSize - 1, $attendees.Count - 1)]
    $batchNum  = [Math]::Floor($i / $batchSize) + 1
    $totalBatches = [Math]::Ceiling($attendees.Count / $batchSize)
    Write-Host "`nBatch $batchNum of $totalBatches ($($batch.Count) emails)..." -ForegroundColor Cyan

    $sentBarcodes = @()
    foreach ($a in $batch) {
        if (-not (Test-Path $a.SpeedPassPath)) {
            Write-Host "  SKIP (PDF missing): $($a.Email)" -ForegroundColor Yellow
            continue
        }

        $body    = $bodyTemplate -replace '{{FirstName}}', $a.FirstName `
                                 -replace '{{EventName}}', $Config.event.name `
                                 -replace '{{Hashtag}}',   $Config.event.hashtag `
                                 -replace '{{BANNER}}',    $bannerHtml
        $sendTo  = if ($TestEmail) { $TestEmail } else { $a.Email }

        try {
            Send-MailMessage -To $sendTo -From $from -Subject $subject -Body $body `
                -SmtpServer $smtp -Port $port -UseSsl -Credential $cred `
                -Attachments $a.SpeedPassPath -BodyAsHtml -WarningAction SilentlyContinue
            Write-Host "  Sent: $($a.Email)" -ForegroundColor Green
            $sentBarcodes += $a.Barcode
            $success++
        } catch {
            Write-Host "  ERROR $($a.Email): $_" -ForegroundColor Red
            $errors++
        }

        if ($delaySeconds -gt 0) { Start-Sleep -Seconds $delaySeconds }
    }

    # Batch-mark as emailed
    foreach ($barcode in $sentBarcodes) {
        Invoke-SqliteQuery -DataSource $dbPath -Query @"
UPDATE ProcessedAttendees SET EmailedAt = datetime('now') WHERE Barcode = @Barcode
"@ -SqlParameters @{ Barcode = $barcode }
    }

    if ($i + $batchSize -lt $attendees.Count) { Start-Sleep -Seconds ($delaySeconds * 2) }
}

Write-Host "`nEmail complete. Sent: $success  Errors: $errors" -ForegroundColor $(if ($errors -gt 0) { 'Yellow' } else { 'Green' })
