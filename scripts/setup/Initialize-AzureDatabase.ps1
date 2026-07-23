<#
.SYNOPSIS
    Creates the Azure SQL database schema for event attendee tracking.
.DESCRIPTION
    Run once to provision the shared Azure SQL schema that all check-in
    laptops sync against. Safe to re-run — uses IF NOT EXISTS guards so
    existing data is preserved. Mirrors the local schema created by
    Initialize-Database.ps1, minus the local-only PendingWrites queue.

    Requires the SqlServer module (for Invoke-Sqlcmd) and a secret holding
    the SQL authentication password, set via:
        Set-Secret -Name '<azure.authSecretName>' -Secret '<password>'
.EXAMPLE
    .\scripts\setup\Initialize-AzureDatabase.ps1 -Config (Get-Content .\event.config.json | ConvertFrom-Json)
#>
param(
    [Parameter(Mandatory)]
    [PSCustomObject]$Config
)

if (-not $Config.azure -or -not $Config.azure.server -or -not $Config.azure.database) {
    throw "event.config.json is missing an 'azure' section with server/database. See event.config.template.json."
}

if (-not (Get-Module -ListAvailable -Name SqlServer)) {
    Write-Host "Installing SqlServer module..." -ForegroundColor Cyan
    Install-Module -Name SqlServer -Scope CurrentUser -Force
}
Import-Module SqlServer

try {
    $password = Get-Secret -Name $Config.azure.authSecretName -AsPlainText
} catch {
    throw "Could not load Azure SQL password from SecretManagement secret '$($Config.azure.authSecretName)'. Run: Set-Secret -Name '$($Config.azure.authSecretName)' -Secret '<password>'"
}
$username = $Config.azure.username
if (-not $username) {
    throw "event.config.json azure section is missing 'username' (the SQL auth login)."
}

$sqlcmdParams = @{
    ServerInstance     = $Config.azure.server
    Database           = $Config.azure.database
    Username           = $username
    Password           = $password
    ConnectionTimeout  = 30
    TrustServerCertificate = $true
}

Invoke-Sqlcmd @sqlcmdParams -Query @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Attendees')
CREATE TABLE Attendees (
    Barcode         NVARCHAR(100) PRIMARY KEY,
    OrderId         NVARCHAR(100),
    OrderDate       NVARCHAR(50),
    FirstName       NVARCHAR(255) NOT NULL,
    LastName        NVARCHAR(255) NOT NULL,
    Email           NVARCHAR(255) NOT NULL,
    Company         NVARCHAR(255),
    JobTitle        NVARCHAR(255),
    LunchType       NVARCHAR(100),
    TicketType      NVARCHAR(100),
    AttendeeStatus  NVARCHAR(50),
    IsVolunteer     BIT DEFAULT 0,
    TwitterHandle   NVARCHAR(255),
    Website         NVARCHAR(255),
    ImportedAt      DATETIME2 DEFAULT SYSUTCDATETIME(),
    UpdatedAt       DATETIME2 DEFAULT SYSUTCDATETIME()
)
"@

Invoke-Sqlcmd @sqlcmdParams -Query @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ProcessedAttendees')
CREATE TABLE ProcessedAttendees (
    Barcode              NVARCHAR(100) PRIMARY KEY,
    SpeedPassPath        NVARCHAR(500),
    SpeedPassGeneratedAt DATETIME2,
    EmailedAt            DATETIME2
)
"@

Invoke-Sqlcmd @sqlcmdParams -Query @"
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'PrintedBadges')
CREATE TABLE PrintedBadges (
    Barcode    NVARCHAR(100) PRIMARY KEY,
    PrintedAt  DATETIME2 DEFAULT SYSUTCDATETIME(),
    PrintedBy  NVARCHAR(255)
)
"@

Write-Host "Azure SQL schema ready: $($Config.azure.server)/$($Config.azure.database)" -ForegroundColor Green
