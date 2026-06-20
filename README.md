# sqlsat-tools

Reusable event tooling for Day of Data Baton Rouge (and future SQL Saturday events).
Generates attendee SpeedPasses, emails them, prints the paper schedule, and produces
the sponsor stamp game card — all driven by a single config file.

## Quick start for a new event

```
1. Copy event.config.template.json → event.config.json
2. Fill in your credentials and event IDs (see Config reference below)
3. Run: .\scripts\Initialize-Database.ps1 -Config (Get-Content event.config.json | ConvertFrom-Json)
4. Run: .\Update-Event.ps1
```

That's it. Re-run `Update-Event.ps1` daily as new registrations come in.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 7+ | Recommended |
| Microsoft Edge | Used for headless PDF generation |
| `PSSQLite` module | Installed automatically on first run |
| `powershell-yaml` module | Installed automatically on first run |
| `Microsoft.PowerShell.SecretManagement` | For storing credentials securely |

Install SecretManagement once:
```powershell
Install-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore -Scope CurrentUser
Register-SecretVault -Name LocalVault -ModuleName Microsoft.PowerShell.SecretStore
```

---

## Credential setup

**EventBrite token**
```powershell
# Get your token from https://www.eventbrite.com/platform/api-keys
Set-Secret -Name "EventbriteToken" -Secret "your-token-here"
```

**Gmail**
```powershell
# Use a Gmail App Password (not your real password)
# https://myaccount.google.com/apppasswords
Set-Secret -Name "SQLSaturday-Gmail" -Secret (Get-Credential)
```

---

## Config reference (`event.config.json`)

| Key | Description |
|---|---|
| `event.hashtag` | Social media hashtag (e.g. `#DayOfDataBR`) |
| `websiteRepo.owner` | GitHub org or user that owns the website repo |
| `websiteRepo.name` | GitHub repo name |
| `websiteRepo.branch` | Branch to read from (usually `main`) |
| `websiteRepo.eventKey` | Folder under `content/events/` in the website repo (e.g. `dodbr-2026`). Event name, Sessionize ID, sponsor data file, and event logo are all read from that folder's `_index.md` at runtime. |
| `eventbrite.eventId` | Numeric EventBrite event ID (find it in the organizer dashboard URL) |
| `eventbrite.secretName` | SecretManagement secret name for the API token |
| `email.secretName` | SecretManagement secret name for Gmail credentials |
| `email.subject` | Email subject line |
| `email.fromName` | Display name in the From header |
| `speedpass.raffleTiers` | Sponsor tiers that get a raffle slip (e.g., `["global","platinum","gold","silver"]`) |
| `stampGame.tiers` | Sponsor tiers that appear on the stamp game card |
| `stampGame.excludeSponsors` | Sponsor names to skip even if they match a listed tier |
| `stampGame.gridColumns` | Number of columns in the stamp card grid (overridden automatically when logo count is a perfect square) |

---

## Scripts

### Daily workflow

```powershell
# Full pipeline (import → generate PDFs → email)
.\Update-Event.ps1

# Preview what would happen
.\Update-Event.ps1 -WhatIf

# Step through each stage with confirmation
.\Update-Event.ps1 -StepThrough

# Individual steps
.\Update-Event.ps1 -ImportOnly
.\Update-Event.ps1 -GenerateOnly
.\Update-Event.ps1 -EmailOnly
```

### Event-day tools

```powershell
$config = Get-Content .\event.config.json | ConvertFrom-Json

# Regenerate SpeedPass for one attendee
.\scripts\Generate-SpeedPasses.ps1 -Config $config -Email "jane.doe@example.com" -Force

# Stamp game card (run once when sponsors are finalized)
.\scripts\Generate-StampGame.ps1 -Config $config

# Paper schedule (run once sessions are published on Sessionize)
.\scripts\Generate-Schedule.ps1 -Config $config

# Walk-in label printer (stub — not yet implemented)
.\scripts\Generate-NameTag.ps1 -Config $config -Email "walkin@example.com"
```

---

## Data source: website repo

Sponsor names, tiers, and logos are read live from the
[sqlsatbr-website](https://github.com/KennyNeal/sqlsatbr-website) repo at runtime.
Update `data/sponsors/{sponsorDataFile}.yaml` in that repo and the next pipeline run
picks up the changes automatically — no manual sync needed.

---

## Database

`event.db` is a local SQLite file (gitignored). Two tables:

| Table | Purpose |
|---|---|
| `Attendees` | Upserted on every import — holds all registrant data |
| `ProcessedAttendees` | Tracks `SpeedPassGeneratedAt` and `EmailedAt` per attendee |

Reset an attendee's status (to regenerate + re-email):
```sql
DELETE FROM ProcessedAttendees WHERE Barcode = 'xxx';
```

---

## Repo structure

```
sqlsat-tools/
├── Update-Event.ps1              ← run this
├── event.config.template.json    ← copy → event.config.json
├── event.config.json             ← gitignored; your live config
├── event.db                      ← gitignored; SQLite database
├── scripts/
│   ├── Initialize-Database.ps1
│   ├── Import-Attendees.ps1
│   ├── Generate-SpeedPasses.ps1
│   ├── Send-SpeedPasses.ps1
│   ├── Generate-StampGame.ps1
│   ├── Generate-Schedule.ps1
│   └── Generate-NameTag.ps1      ← stub (Brother QL-820NWB)
├── templates/
│   └── attendee-email.html
├── lib/
│   └── QRCoder.dll
└── output/                       ← gitignored; generated files
    └── speedpasses/
```
