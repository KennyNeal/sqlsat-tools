# Azure SQL infrastructure for check-in sync

Provisions the Azure SQL Database used as the shared source of truth by
`scripts/Data-Access.ps1` (see the root README / event.config.json `azure`
section). Serverless General Purpose tier, smallest vCore config, auto-pause
after 60 idle minutes — near-zero cost for the ~360 days/year it's not in use.

## Prerequisites

- [Terraform](https://developer.hashicorp.com/terraform/install) >= 1.5
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli), logged in (`az login`) to the subscription you want billed
- An Azure subscription with permission to create resource groups / SQL servers

## Usage

```powershell
cd infra
cp terraform.tfvars.example terraform.tfvars
# edit terraform.tfvars — sql_server_name must be globally unique

terraform init
terraform plan    # review what will be created before applying
terraform apply   # creates real, billed Azure resources — review the plan first
```

After apply, follow the printed `next_steps` output:

```powershell
terraform output -raw sql_admin_password | Set-Secret -Name 'AzureSqlAuth' -SecureStringSecret ...
# or, simpler:
Set-Secret -Name 'AzureSqlAuth' -Secret (terraform output -raw sql_admin_password)
```

Then fill in `event.config.json`'s `azure` section with `terraform output
sql_server_fqdn` / `sql_database_name` / `sql_admin_username`, set `enabled:
true`, and run `.\scripts\setup\Initialize-AzureDatabase.ps1` from the repo root.

## Firewall / venue network

`allowed_client_ips` is empty by default. Azure SQL firewalls by source IP,
and you generally won't know the venue WiFi's public egress IP until you're
there (or close to event day). Don't re-run `terraform apply` from an unknown
network to add it — use the Azure CLI directly, which is faster and doesn't
risk a plan picking up unrelated drift:

```powershell
az sql server firewall-rule create `
  -g rg-sqlsat-checkin -s <sql_server_name> -n venue `
  --start-ip-address <ip> --end-ip-address <ip>
```

## Teardown

This is safe to destroy after the event and recreate next year — it holds no
data that isn't also in `event.db` locally (Sync-FromAzure mirrors
everything).

```powershell
terraform destroy
```
