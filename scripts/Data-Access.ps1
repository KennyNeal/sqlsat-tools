<#
.SYNOPSIS
    Central data-access layer for attendee data: local SQLite cache + Azure SQL
    shared store, with automatic offline fallback and reconnect replay.
.DESCRIPTION
    Every script that used to call Invoke-SqliteQuery directly now goes through
    here instead. The model:

      - Local SQLite (event.db) is ALWAYS written first and is what every read
        in this app queries — reads never touch Azure directly, so a lookup
        during check-in never blocks on network. Local doubles as the offline
        cache.
      - Azure SQL is the shared source of truth. Every write also attempts
        Azure synchronously (short timeout) right after the local write. If
        Azure is unreachable, the write is queued in the local PendingWrites
        table and replayed later by Sync-PendingWrites.
      - Sync-FromAzure pulls the latest Attendees/PrintedBadges/ProcessedAttendees
        rows from Azure into the local cache, so a laptop picks up what other
        desks have done. Call it at menu startup and after a successful drain.
      - If azure.enabled is false in event.config.json (or Azure creds aren't
        set up), everything silently behaves exactly like the old local-only
        SQLite tool — this is the event-day safety net.

    Uses Microsoft.Data.SqlClient (via the SqlServer module's dependency) with
    real SqlParameter binding for every Azure call that carries attendee-
    supplied values (names, emails, etc.) — Invoke-Sqlcmd's -Variable
    substitution was considered but rejected because it does textual
    substitution rather than true parameterization, which is not safe for
    values that ultimately come from attendee input (Eventbrite import,
    walk-in registration).
#>

Import-Module PSSQLite -ErrorAction Stop

# ── Local database integrity / recovery ─────────────────────────────────────
#
# 2026 incident: a check-in laptop's event.db returned "file is not a
# database" partway through the event and everything local on that machine
# (walk-ins, badge-printed marks not yet synced to Azure) was gone. The likely
# cause is a cloud-sync tool (OneDrive/Dropbox) touching the file while SQLite
# had it open. The functions below make that recoverable instead of fatal:
# every context creation checks the file's integrity, and if it's corrupt,
# quarantines it (never deletes — it may be partially salvageable) and
# rebuilds, repopulating from Azure when available.

function Test-LocalDatabaseIntegrity {
    <#
    .SYNOPSIS
        Runs PRAGMA quick_check against the local db. Returns $true if the
        file is healthy or doesn't exist yet (nothing to check), $false if
        it's corrupt/unreadable ("file is not a database" etc.).
    #>
    param([Parameter(Mandatory)][string]$LocalDbPath)

    if (-not (Test-Path $LocalDbPath)) { return $true }

    try {
        $result = Invoke-SqliteQuery -DataSource $LocalDbPath -Query "PRAGMA quick_check"
        return ($result.Count -gt 0 -and $result[0].quick_check -eq 'ok')
    } catch {
        return $false
    }
}

function Repair-LocalDatabase {
    <#
    .SYNOPSIS
        Called when Test-LocalDatabaseIntegrity fails. Quarantines the
        corrupt file (and any WAL sidecar files) alongside itself with a
        timestamp suffix, then rebuilds an empty schema via
        Initialize-Database.ps1. Never deletes the corrupt file — a human
        can attempt recovery with the sqlite3 CLI's `.recover` later.
    .OUTPUTS
        Human-readable summary string.
    #>
    param([Parameter(Mandatory)][string]$LocalDbPath, [Parameter(Mandatory)][PSCustomObject]$Config)

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $quarantined = "$LocalDbPath.corrupt-$stamp"
    Move-Item -Path $LocalDbPath -Destination $quarantined -Force
    foreach ($ext in @('-wal', '-shm')) {
        $sidecar = "$LocalDbPath$ext"
        if (Test-Path $sidecar) { Move-Item -Path $sidecar -Destination "$quarantined$ext" -Force }
    }

    & (Join-Path $PSScriptRoot "Initialize-Database.ps1") -Config $Config | Out-Null

    "local database was corrupt — quarantined to $quarantined and rebuilt an empty schema"
}

# ── Context ──────────────────────────────────────────────────────────────────

function New-DataContext {
    <#
    .SYNOPSIS
        Builds the shared context object every Data-Access function takes.
        Call once per script run. Also verifies the local db is readable
        (auto-recovering + repopulating from Azure if not) and puts it in
        WAL journal mode, which tolerates crashes/power loss far better than
        the default rollback-journal mode.
    #>
    param([Parameter(Mandatory)][PSCustomObject]$Config)

    $localDbPath = Join-Path $PSScriptRoot ".." $Config.database.path

    $azureEnabled = $false
    $azureConnStr = $null
    if ($Config.PSObject.Properties['azure'] -and $Config.azure.enabled) {
        if (-not (Get-Module -ListAvailable -Name SqlServer)) {
            Write-Host "Installing SqlServer module (Azure SQL support)..." -ForegroundColor Cyan
            Install-Module -Name SqlServer -Scope CurrentUser -Force
        }
        Import-Module SqlServer -ErrorAction Stop

        try {
            $password = Get-Secret -Name $Config.azure.authSecretName -AsPlainText
            $azureConnStr = "Server=tcp:$($Config.azure.server),1433;Database=$($Config.azure.database);" +
                "User ID=$($Config.azure.username);Password=$password;Encrypt=True;" +
                "TrustServerCertificate=False;Connection Timeout=3;"
            $azureEnabled = $true
        } catch {
            Write-Host "Azure SQL is enabled in config but credentials aren't available ($_). Running local-only." -ForegroundColor Yellow
        }
    }

    $context = [PSCustomObject]@{
        LocalDbPath      = $localDbPath
        AzureEnabled     = $azureEnabled
        AzureConnStr     = $azureConnStr
        AzureReachable   = $azureEnabled   # optimistic until the first probe proves otherwise
        LastChecked      = [datetime]::MinValue
        ReachableTtlSecs = 15
    }

    if (-not (Test-LocalDatabaseIntegrity -LocalDbPath $localDbPath)) {
        Write-Host "  WARNING: local database at $localDbPath appears corrupt — quarantining and rebuilding..." -ForegroundColor Red
        $detail = Repair-LocalDatabase -LocalDbPath $localDbPath -Config $Config
        Write-Host "  $detail" -ForegroundColor Yellow
        if ($azureEnabled -and (Test-AzureReachable -DataContext $context) -and (Sync-FromAzure -DataContext $context)) {
            Write-Host "  Repopulated from Azure — this desk should now match every other desk." -ForegroundColor Green
        } else {
            Write-Host "  Could not repopulate from Azure automatically. Anything that was only on" -ForegroundColor Red
            Write-Host "  this laptop since its last sync may be lost. Run the Azure sync menu option once online." -ForegroundColor Red
        }
    }

    if (Test-Path $localDbPath) {
        try {
            Invoke-SqliteQuery -DataSource $localDbPath -Query "PRAGMA journal_mode=WAL;" | Out-Null
            Invoke-SqliteQuery -DataSource $localDbPath -Query "PRAGMA synchronous=NORMAL;" | Out-Null
        } catch { }
    }

    return $context
}

function Test-AzureReachable {
    <#
    .SYNOPSIS
        Cheap, TTL-cached reachability probe. Only actually opens a connection
        once per TTL window so a burst of check-ins during an outage doesn't
        each pay the connection timeout.
    #>
    param([Parameter(Mandatory)]$DataContext)

    if (-not $DataContext.AzureEnabled) { return $false }

    $elapsed = (Get-Date) - $DataContext.LastChecked
    if ($elapsed.TotalSeconds -lt $DataContext.ReachableTtlSecs) {
        return $DataContext.AzureReachable
    }

    try {
        $conn = [Microsoft.Data.SqlClient.SqlConnection]::new($DataContext.AzureConnStr)
        $conn.Open()
        $conn.Close()
        $DataContext.AzureReachable = $true
    } catch {
        $DataContext.AzureReachable = $false
    }
    $DataContext.LastChecked = Get-Date
    return $DataContext.AzureReachable
}

# ── Low-level Azure helpers (parameterized) ────────────────────────────────────

function Invoke-AzureNonQuery {
    <#
    .SYNOPSIS
        Runs a single parameterized Azure write. By default opens and closes
        its own connection (fine for one-off dual writes). Pass -Connection
        with an already-open SqlConnection to reuse it across a batch (e.g.
        Sync-PendingWrites draining many queued rows) instead of paying a
        fresh TLS handshake per row.
    #>
    param(
        [Parameter(Mandatory)]$DataContext,
        [Parameter(Mandatory)][string]$CommandText,
        [hashtable]$Parameters = @{},
        [Microsoft.Data.SqlClient.SqlConnection]$Connection
    )
    $conn = $Connection
    $ownsConnection = $false
    if (-not $conn) {
        $conn = [Microsoft.Data.SqlClient.SqlConnection]::new($DataContext.AzureConnStr)
        $conn.Open()
        $ownsConnection = $true
    }
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $CommandText
        foreach ($key in $Parameters.Keys) {
            $value = $Parameters[$key]
            $null = $cmd.Parameters.AddWithValue("@$key", $(if ($null -eq $value) { [DBNull]::Value } else { $value }))
        }
        $cmd.ExecuteNonQuery() | Out-Null
    } finally {
        if ($ownsConnection) { $conn.Close() }
    }
}

function Invoke-AzureQuery {
    param(
        [Parameter(Mandatory)]$DataContext,
        [Parameter(Mandatory)][string]$CommandText,
        [hashtable]$Parameters = @{}
    )
    $conn = [Microsoft.Data.SqlClient.SqlConnection]::new($DataContext.AzureConnStr)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $CommandText
        foreach ($key in $Parameters.Keys) {
            $value = $Parameters[$key]
            $null = $cmd.Parameters.AddWithValue("@$key", $(if ($null -eq $value) { [DBNull]::Value } else { $value }))
        }
        $reader  = $cmd.ExecuteReader()
        $results = [System.Collections.Generic.List[object]]::new()
        while ($reader.Read()) {
            $row = [ordered]@{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $val = $reader.GetValue($i)
                $row[$reader.GetName($i)] = if ($val -is [DBNull]) { $null } else { $val }
            }
            $results.Add([PSCustomObject]$row)
        }
        $reader.Close()
        return $results
    } finally {
        $conn.Close()
    }
}

# ── Pending-write queue ─────────────────────────────────────────────────────────

# Maps a queued Operation name to the Azure command text replayed for it.
# Kept in one place so Sync-PendingWrites and the write functions below agree.
$script:AzureOperationQueries = @{
    'InsertAttendee'      = @"
IF NOT EXISTS (SELECT 1 FROM Attendees WHERE Barcode = @Barcode)
INSERT INTO Attendees (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle, TicketType, AttendeeStatus)
VALUES (@Barcode, @OrderId, SYSUTCDATETIME(), @FirstName, @LastName, @Email, @Company, @JobTitle, 'Walk-in', 'attending')
"@
    'UpsertAttendee'      = @"
MERGE Attendees AS target
USING (SELECT @Barcode AS Barcode) AS src ON target.Barcode = src.Barcode
WHEN MATCHED THEN UPDATE SET OrderId=@OrderId, OrderDate=@OrderDate, FirstName=@FirstName, LastName=@LastName, Email=@Email,
    Company=@Company, JobTitle=@JobTitle, LunchType=@LunchType, TicketType=@TicketType, AttendeeStatus=@AttendeeStatus,
    IsVolunteer=@IsVolunteer, TwitterHandle=@TwitterHandle, Website=@Website, UpdatedAt=SYSUTCDATETIME()
WHEN NOT MATCHED THEN INSERT (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle, LunchType, TicketType, AttendeeStatus, IsVolunteer, TwitterHandle, Website, UpdatedAt)
    VALUES (@Barcode, @OrderId, @OrderDate, @FirstName, @LastName, @Email, @Company, @JobTitle, @LunchType, @TicketType, @AttendeeStatus, @IsVolunteer, @TwitterHandle, @Website, SYSUTCDATETIME());
"@
    'SetBadgePrinted'     = @"
MERGE PrintedBadges AS target
USING (SELECT @Barcode AS Barcode) AS src ON target.Barcode = src.Barcode
WHEN MATCHED THEN UPDATE SET PrintedAt = SYSUTCDATETIME(), PrintedBy = @PrintedBy
WHEN NOT MATCHED THEN INSERT (Barcode, PrintedAt, PrintedBy) VALUES (@Barcode, SYSUTCDATETIME(), @PrintedBy);
"@
    'SetSpeedPassGenerated' = @"
MERGE ProcessedAttendees AS target
USING (SELECT @Barcode AS Barcode) AS src ON target.Barcode = src.Barcode
WHEN MATCHED THEN UPDATE SET SpeedPassPath = @Path, SpeedPassGeneratedAt = SYSUTCDATETIME()
WHEN NOT MATCHED THEN INSERT (Barcode, SpeedPassPath, SpeedPassGeneratedAt) VALUES (@Barcode, @Path, SYSUTCDATETIME());
"@
    'SetSpeedPassEmailed' = @"
UPDATE ProcessedAttendees SET EmailedAt = SYSUTCDATETIME() WHERE Barcode = @Barcode
"@
}

function Add-PendingWrite {
    param([Parameter(Mandatory)]$DataContext, [Parameter(Mandatory)][string]$Operation, [Parameter(Mandatory)][string]$Barcode, [Parameter(Mandatory)][hashtable]$Parameters)

    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query @"
INSERT INTO PendingWrites (Operation, Barcode, PayloadJson) VALUES (@Operation, @Barcode, @Payload)
"@ -SqlParameters @{ Operation = $Operation; Barcode = $Barcode; Payload = ($Parameters | ConvertTo-Json -Compress) }
}

function Sync-PendingWrites {
    <#
    .SYNOPSIS
        Drains queued local writes to Azure, in order. Stops at the first
        failure to preserve ordering (don't skip ahead and risk an
        out-of-order upsert landing on the same barcode).
    .OUTPUTS
        Number of writes successfully drained.
    #>
    param([Parameter(Mandatory)]$DataContext)

    if (-not (Test-AzureReachable -DataContext $DataContext)) { return 0 }

    $pending = Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "SELECT * FROM PendingWrites ORDER BY Id"
    if ($pending.Count -eq 0) { return 0 }

    # One connection for the whole drain — a fresh TLS handshake per queued
    # row is what made the old per-row Import-Attendees Azure push slow, and
    # a queue built up during an outage can easily be a few hundred rows.
    $conn = [Microsoft.Data.SqlClient.SqlConnection]::new($DataContext.AzureConnStr)
    $drained = 0
    try {
        $conn.Open()
        foreach ($row in $pending) {
            $params = @{}
            ($row.PayloadJson | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $params[$_.Name] = $_.Value }
            $azureQuery = $script:AzureOperationQueries[$row.Operation]
            if (-not $azureQuery) {
                Write-Host "  Unknown queued operation '$($row.Operation)' (Barcode $($row.Barcode)) — skipping." -ForegroundColor Yellow
                Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "DELETE FROM PendingWrites WHERE Id = @Id" -SqlParameters @{ Id = $row.Id }
                continue
            }
            try {
                Invoke-AzureNonQuery -DataContext $DataContext -CommandText $azureQuery -Parameters $params -Connection $conn
                Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "DELETE FROM PendingWrites WHERE Id = @Id" -SqlParameters @{ Id = $row.Id }
                $drained++
            } catch {
                Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "UPDATE PendingWrites SET Attempts = Attempts + 1, LastError = @Err WHERE Id = @Id" -SqlParameters @{ Err = "$_"; Id = $row.Id }
                $DataContext.AzureReachable = $false
                break
            }
        }
    } finally {
        $conn.Close()
    }
    return $drained
}

function Sync-FromAzure {
    <#
    .SYNOPSIS
        Pulls the full Attendees/PrintedBadges/ProcessedAttendees tables from
        Azure and mirrors them into the local cache. Full-table pull each
        time — row counts here are in the hundreds, not worth watermarking.
    .OUTPUTS
        $true if the pull succeeded, $false if Azure was unreachable.
    #>
    param([Parameter(Mandatory)]$DataContext)

    if (-not (Test-AzureReachable -DataContext $DataContext)) { return $false }

    try {
        $attendees = Invoke-AzureQuery -DataContext $DataContext -CommandText "SELECT * FROM Attendees"
        foreach ($a in $attendees) {
            Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query @"
INSERT OR REPLACE INTO Attendees
    (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle, LunchType, TicketType, AttendeeStatus, IsVolunteer, TwitterHandle, Website, ImportedAt, UpdatedAt)
VALUES
    (@Barcode, @OrderId, @OrderDate, @FirstName, @LastName, @Email, @Company, @JobTitle, @LunchType, @TicketType, @AttendeeStatus, @IsVolunteer, @TwitterHandle, @Website, @ImportedAt, @UpdatedAt)
"@ -SqlParameters @{
                Barcode = $a.Barcode; OrderId = $a.OrderId; OrderDate = "$($a.OrderDate)"; FirstName = $a.FirstName
                LastName = $a.LastName; Email = $a.Email; Company = $a.Company; JobTitle = $a.JobTitle
                LunchType = $a.LunchType; TicketType = $a.TicketType; AttendeeStatus = $a.AttendeeStatus
                IsVolunteer = [int]$a.IsVolunteer; TwitterHandle = $a.TwitterHandle; Website = $a.Website
                ImportedAt = "$($a.ImportedAt)"; UpdatedAt = "$($a.UpdatedAt)"
            }
        }

        $printed = Invoke-AzureQuery -DataContext $DataContext -CommandText "SELECT * FROM PrintedBadges"
        foreach ($p in $printed) {
            Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query @"
INSERT OR REPLACE INTO PrintedBadges (Barcode, PrintedAt, PrintedBy) VALUES (@Barcode, @PrintedAt, @PrintedBy)
"@ -SqlParameters @{ Barcode = $p.Barcode; PrintedAt = "$($p.PrintedAt)"; PrintedBy = $p.PrintedBy }
        }

        $processed = Invoke-AzureQuery -DataContext $DataContext -CommandText "SELECT * FROM ProcessedAttendees"
        foreach ($p in $processed) {
            Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query @"
INSERT OR REPLACE INTO ProcessedAttendees (Barcode, SpeedPassPath, SpeedPassGeneratedAt, EmailedAt)
VALUES (@Barcode, @Path, @GeneratedAt, @EmailedAt)
"@ -SqlParameters @{ Barcode = $p.Barcode; Path = $p.SpeedPassPath; GeneratedAt = "$($p.SpeedPassGeneratedAt)"; EmailedAt = "$($p.EmailedAt)" }
        }

        return $true
    } catch {
        $DataContext.AzureReachable = $false
        return $false
    }
}

# ── Write journal (last-resort recovery aid) ────────────────────────────────

function Write-LocalJournalEntry {
    <#
    .SYNOPSIS
        Appends one line to a plain-text write journal in backup\ next to
        event.db. If SQLite itself gets corrupted, this survives — it's just
        a text file — and lets a human reconstruct what happened since the
        last Azure sync or backup. Best-effort; never throws.
    #>
    param(
        [Parameter(Mandatory)][string]$LocalDbPath,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][string]$Barcode,
        [hashtable]$Detail = @{}
    )
    try {
        $journalDir = Join-Path (Split-Path $LocalDbPath -Parent) "backup"
        if (-not (Test-Path $journalDir)) { New-Item -ItemType Directory -Path $journalDir -Force | Out-Null }
        $journalPath = Join-Path $journalDir "writes-$(Get-Date -Format 'yyyy-MM-dd').log"
        $line = "{0:o}`t{1}`t{2}`t{3}" -f (Get-Date), $Operation, $Barcode, ($Detail | ConvertTo-Json -Compress)
        Add-Content -Path $journalPath -Value $line
    } catch { }
}

# ── Local database backups ───────────────────────────────────────────────────

function Backup-LocalDatabase {
    <#
    .SYNOPSIS
        Hot online backup of the local db via VACUUM INTO (safe to run while
        other connections are open). Keeps the last $KeepCount backups in
        backup\ next to event.db, pruning older ones.
    .OUTPUTS
        Path to the new backup file, or $null if the backup failed.
    #>
    param([Parameter(Mandatory)]$DataContext, [int]$KeepCount = 10)

    $backupDir = Join-Path (Split-Path $DataContext.LocalDbPath -Parent) "backup"
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

    $stamp      = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupPath = Join-Path $backupDir "event-$stamp.db"

    try {
        $escaped = $backupPath -replace "'", "''"
        Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "VACUUM INTO '$escaped'" | Out-Null
    } catch {
        Write-Host "  Warning: local database backup failed: $_" -ForegroundColor Yellow
        return $null
    }

    Get-ChildItem -Path $backupDir -Filter "event-*.db" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip $KeepCount |
        Remove-Item -Force -ErrorAction SilentlyContinue

    return $backupPath
}

function Backup-LocalDatabaseIfDue {
    <#
    .SYNOPSIS
        Runs Backup-LocalDatabase only if the newest existing backup is older
        than $IntervalMinutes (or there isn't one yet). Cheap to call
        opportunistically in a menu loop — no extra state needed beyond the
        backup folder's own file timestamps.
    #>
    param([Parameter(Mandatory)]$DataContext, [int]$IntervalMinutes = 10)

    $backupDir = Join-Path (Split-Path $DataContext.LocalDbPath -Parent) "backup"
    $newest = if (Test-Path $backupDir) {
        Get-ChildItem -Path $backupDir -Filter "event-*.db" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    } else { $null }

    if ($newest -and ((Get-Date) - $newest.LastWriteTime).TotalMinutes -lt $IntervalMinutes) { return $null }
    return Backup-LocalDatabase -DataContext $DataContext
}

# ── Dual-store write helper ─────────────────────────────────────────────────────

function Invoke-DualWrite {
    <#
    .SYNOPSIS
        Writes to local SQLite first (always), then best-effort to Azure.
        On any Azure failure, queues the operation for later replay instead
        of failing the caller — local SQLite is the durable record.
    #>
    param(
        [Parameter(Mandatory)]$DataContext,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][string]$Barcode,
        [Parameter(Mandatory)][string]$LocalQuery,
        [Parameter(Mandatory)][hashtable]$LocalParameters,
        [Parameter(Mandatory)][hashtable]$AzureParameters
    )

    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query $LocalQuery -SqlParameters $LocalParameters | Out-Null
    Write-LocalJournalEntry -LocalDbPath $DataContext.LocalDbPath -Operation $Operation -Barcode $Barcode -Detail $LocalParameters

    if (-not $DataContext.AzureEnabled) { return }

    if (Test-AzureReachable -DataContext $DataContext) {
        try {
            Invoke-AzureNonQuery -DataContext $DataContext -CommandText $script:AzureOperationQueries[$Operation] -Parameters $AzureParameters
            return
        } catch {
            $DataContext.AzureReachable = $false
        }
    }

    Add-PendingWrite -DataContext $DataContext -Operation $Operation -Barcode $Barcode -Parameters $AzureParameters
}

# ── Attendee lookups (Checkin-Core.ps1 / Print-WalkinBadge / Checkin-Menu) ─────

function Get-AttendeesByOrderOrEmail {
    param([Parameter(Mandatory)]$DataContext, [string]$OrderId, [string]$Email)

    $query = @"
SELECT a.Barcode, a.OrderId, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle, a.LunchType, p.PrintedAt
FROM   Attendees a
LEFT JOIN PrintedBadges p ON p.Barcode = a.Barcode
WHERE  LOWER(a.AttendeeStatus) NOT IN ('cancelled', 'deleted', 'refunded', 'declined', 'not attending')
AND    $(if ($OrderId) { "a.OrderId = @OrderId" } else { "a.Email = @Email" })
ORDER  BY a.LastName, a.FirstName
"@
    $params = if ($OrderId) { @{ OrderId = $OrderId } } else { @{ Email = $Email } }
    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query $query -SqlParameters $params
}

function Add-Attendee {
    <#
    .SYNOPSIS
        Inserts a quick-add walk-in attendee. Pure function — caller has
        already collected and validated FirstName/LastName/Email.
    #>
    param(
        [Parameter(Mandatory)]$DataContext,
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][string]$Email,
        [string]$Company,
        [string]$JobTitle
    )

    $barcode = "WALKIN-$([guid]::NewGuid().ToString())"
    $params = @{ Barcode = $barcode; OrderId = 'WALKIN'; FirstName = $FirstName; LastName = $LastName; Email = $Email; Company = $Company; JobTitle = $JobTitle }

    Invoke-DualWrite -DataContext $DataContext -Operation 'InsertAttendee' -Barcode $barcode -AzureParameters $params -LocalParameters $params -LocalQuery @"
INSERT INTO Attendees
    (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle, TicketType, AttendeeStatus)
VALUES
    (@Barcode, @OrderId, datetime('now'), @FirstName, @LastName, @Email, @Company, @JobTitle, 'Walk-in', 'attending')
"@

    [PSCustomObject]@{
        Barcode = $barcode; OrderId = 'WALKIN'; FirstName = $FirstName; LastName = $LastName
        Email = $Email; Company = $Company; JobTitle = $JobTitle; LunchType = $null; PrintedAt = $null
    }
}

function Set-BadgePrinted {
    param([Parameter(Mandatory)]$DataContext, [Parameter(Mandatory)][string]$Barcode, [string]$PrintedBy)

    $params = @{ Barcode = $Barcode; PrintedBy = $PrintedBy }
    Invoke-DualWrite -DataContext $DataContext -Operation 'SetBadgePrinted' -Barcode $Barcode -AzureParameters $params -LocalParameters $params -LocalQuery @"
INSERT OR REPLACE INTO PrintedBadges (Barcode, PrintedAt, PrintedBy) VALUES (@Barcode, datetime('now'), @PrintedBy)
"@
}

function Get-UnsyncedWalkins {
    param([Parameter(Mandatory)]$DataContext)

    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle, a.OrderDate, p.PrintedAt
FROM   Attendees a
LEFT JOIN PrintedBadges p ON p.Barcode = a.Barcode
WHERE  a.OrderId = 'WALKIN'
ORDER  BY a.OrderDate
"@
}

# ── Eventbrite import (Import-Attendees.ps1) ────────────────────────────────────

function Import-AttendeesFromEventbrite {
    <#
    .SYNOPSIS
        Upserts a batch of Eventbrite attendee records into local SQLite (one
        connection/transaction for the whole batch — matches the perf note
        this replaces: a connection per row makes a few-hundred-row import
        take minutes) and queues the same batch for Azure.
    .DESCRIPTION
        The Azure side is queue-only here, not a synchronous push — opening a
        fresh connection per row against Azure (the only way to keep each
        row's own retry/ordering semantics) would turn a few-hundred-row
        import into a multi-minute wait on the volunteer running it. Queuing
        is a single local insert per row (no network), so the import stays
        fast; Sync-PendingWrites (menu option 5, or the opportunistic
        catch-up at menu startup/loop) then drains the queue over one shared
        Azure connection, off the import's critical path.
    .PARAMETER Attendees
        Array of hashtables/PSCustomObjects with Barcode, OrderId, OrderDate,
        FirstName, LastName, Email, Company, JobTitle, LunchType, TicketType,
        AttendeeStatus, IsVolunteer, TwitterHandle, Website.
    .OUTPUTS
        Count of rows upserted locally.
    #>
    param([Parameter(Mandatory)]$DataContext, [Parameter(Mandatory)][array]$Attendees)

    $localUpsertSql = @"
INSERT OR REPLACE INTO Attendees
    (Barcode, OrderId, OrderDate, FirstName, LastName, Email, Company, JobTitle,
     LunchType, TicketType, AttendeeStatus, IsVolunteer, TwitterHandle, Website, UpdatedAt)
VALUES
    (@Barcode, @OrderId, @OrderDate, @FirstName, @LastName, @Email, @Company, @JobTitle,
     @LunchType, @TicketType, @AttendeeStatus, @IsVolunteer, @TwitterHandle, @Website, datetime('now'))
"@

    $conn = New-SQLiteConnection -DataSource $DataContext.LocalDbPath
    $imported = 0
    try {
        Invoke-SqliteQuery -SQLiteConnection $conn -Query "BEGIN"
        foreach ($a in $Attendees) {
            Invoke-SqliteQuery -SQLiteConnection $conn -Query $localUpsertSql -SqlParameters $a
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

    # Queue the whole batch for Azure — no network calls here, see the
    # .DESCRIPTION above for why this isn't a synchronous push.
    if ($DataContext.AzureEnabled) {
        foreach ($a in $Attendees) {
            Add-PendingWrite -DataContext $DataContext -Operation 'UpsertAttendee' -Barcode $a.Barcode -Parameters $a
        }
    }

    return $imported
}

# ── SpeedPass generation / email (Generate-SpeedPasses.ps1 / Send-SpeedPasses.ps1) ─

function Get-AttendeesForSpeedPass {
    param([Parameter(Mandatory)]$DataContext, [string]$Email, [switch]$Force)

    $statusFilter = "LOWER(a.AttendeeStatus) NOT IN ('cancelled', 'deleted', 'refunded', 'declined', 'not attending')"
    $emailFilter  = if ($Email) { "a.Email = @Email" } else { $null }
    $nullFilter   = if (-not $Force) { "p.SpeedPassGeneratedAt IS NULL" } else { $null }
    $conditions   = @($statusFilter, $emailFilter, $nullFilter) | Where-Object { $_ }
    $whereClause  = "WHERE " + ($conditions -join " AND ")

    $query = @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle,
       a.LunchType, a.TwitterHandle, a.Website
FROM   Attendees a
LEFT   JOIN ProcessedAttendees p ON a.Barcode = p.Barcode
$whereClause
ORDER  BY a.LastName, a.FirstName
"@
    $sqlParams = if ($Email) { @{ Email = $Email } } else { @{} }
    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query $query -SqlParameters $sqlParams
}

function Set-SpeedPassGenerated {
    param([Parameter(Mandatory)]$DataContext, [Parameter(Mandatory)][string]$Barcode, [Parameter(Mandatory)][string]$Path)

    $params = @{ Barcode = $Barcode; Path = $Path }
    Invoke-DualWrite -DataContext $DataContext -Operation 'SetSpeedPassGenerated' -Barcode $Barcode -AzureParameters $params -LocalParameters $params -LocalQuery @"
INSERT INTO ProcessedAttendees (Barcode, SpeedPassPath, SpeedPassGeneratedAt)
VALUES (@Barcode, @Path, datetime('now'))
ON CONFLICT(Barcode) DO UPDATE SET SpeedPassPath=@Path, SpeedPassGeneratedAt=datetime('now')
"@
}

function Get-AttendeesForEmail {
    param([Parameter(Mandatory)]$DataContext)

    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, p.SpeedPassPath
FROM   Attendees a
JOIN   ProcessedAttendees p ON a.Barcode = p.Barcode
WHERE  p.SpeedPassGeneratedAt IS NOT NULL
AND    p.EmailedAt IS NULL
ORDER  BY a.LastName, a.FirstName
"@
}

function Set-SpeedPassEmailed {
    param([Parameter(Mandatory)]$DataContext, [Parameter(Mandatory)][string]$Barcode)

    $params = @{ Barcode = $Barcode }
    Invoke-DualWrite -DataContext $DataContext -Operation 'SetSpeedPassEmailed' -Barcode $Barcode -AzureParameters $params -LocalParameters $params -LocalQuery @"
UPDATE ProcessedAttendees SET EmailedAt = datetime('now') WHERE Barcode = @Barcode
"@
}

# ── Name tags (Generate-NameTag.ps1) ────────────────────────────────────────────

function Get-AttendeesForNameTag {
    param(
        [Parameter(Mandatory)]$DataContext,
        [string]$Email,
        [switch]$Force,
        [ValidateSet('Speaker', 'Attendee', 'All')]
        [string]$Type = 'All'
    )

    $statusFilter  = "LOWER(a.AttendeeStatus) NOT IN ('cancelled', 'deleted', 'refunded', 'declined', 'not attending')"
    $emailFilter   = if ($Email) { "a.Email = @Email" } else { $null }
    $nullFilter    = if (-not $Force -and -not $Email) { "p.PrintedAt IS NULL" } else { $null }
    $typeFilter    = switch ($Type) {
        'Speaker'  { "a.TicketType LIKE '%Speaker%'" }
        'Attendee' { "(a.TicketType IS NULL OR a.TicketType NOT LIKE '%Speaker%')" }
        default    { $null }
    }
    $conditions   = @($statusFilter, $emailFilter, $nullFilter, $typeFilter) | Where-Object { $_ }
    $whereClause  = "WHERE " + ($conditions -join " AND ")

    $query = @"
SELECT a.Barcode, a.FirstName, a.LastName, a.Email, a.Company, a.JobTitle, a.LunchType, a.TicketType
FROM   Attendees a
LEFT   JOIN PrintedBadges p ON a.Barcode = p.Barcode
$whereClause
ORDER  BY a.LastName, a.FirstName
"@
    $sqlParams = if ($Email) { @{ Email = $Email } } else { @{} }
    Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query $query -SqlParameters $sqlParams
}

# ── Readiness check (Test-EventReadiness.ps1) ───────────────────────────────────

function Test-DatabaseReadiness {
    <#
    .SYNOPSIS
        Checks local schema/row-count, and Azure connectivity + schema if
        Azure is enabled. Throws with a descriptive message on failure (call
        inside Test-Check's scriptblock, matching Test-EventReadiness's idiom).
    #>
    param([Parameter(Mandatory)]$DataContext)

    if (-not (Test-Path $DataContext.LocalDbPath)) { throw "local db not found at $($DataContext.LocalDbPath) — run Initialize-Database.ps1" }
    $tables = (Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "SELECT name FROM sqlite_master WHERE type='table'").name
    $missing = @('Attendees', 'ProcessedAttendees', 'PrintedBadges', 'PendingWrites' | Where-Object { $_ -notin $tables })
    if ($missing) { throw "local db missing table(s): $($missing -join ', ') — run Initialize-Database.ps1" }
    $count = (Invoke-SqliteQuery -DataSource $DataContext.LocalDbPath -Query "SELECT COUNT(*) AS c FROM Attendees").c

    if (-not $DataContext.AzureEnabled) {
        return "$count attendees (local only — azure.enabled is false)"
    }

    if (-not (Test-AzureReachable -DataContext $DataContext)) {
        throw "azure.enabled is true but Azure SQL is unreachable — check azure.server/database/username and the auth secret"
    }
    $azureTables = (Invoke-AzureQuery -DataContext $DataContext -CommandText "SELECT name FROM sys.tables").name
    $azureMissing = @('Attendees', 'ProcessedAttendees', 'PrintedBadges' | Where-Object { $_ -notin $azureTables })
    if ($azureMissing) { throw "Azure SQL missing table(s): $($azureMissing -join ', ') — run Initialize-AzureDatabase.ps1" }

    "$count local attendees, Azure reachable"
}
