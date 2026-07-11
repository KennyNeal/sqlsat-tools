resource "azurerm_resource_group" "checkin" {
  name     = var.resource_group_name
  location = var.location
  tags     = var.tags
}

# Generated once; Terraform will not rotate it on subsequent applies unless
# this resource is tainted/replaced. Retrieve it via `terraform output
# sql_admin_password` and store it with Set-Secret — never put it in
# event.config.json.
resource "random_password" "sql_admin" {
  length      = 24
  special     = true
  min_upper   = 2
  min_lower   = 2
  min_numeric = 2
  min_special = 2
  # Azure SQL rejects a handful of characters in passwords; keep to a safe set.
  override_special = "!#$%&*()-_=+[]{}<>:?"
}

resource "azurerm_mssql_server" "checkin" {
  name                         = var.sql_server_name
  resource_group_name          = azurerm_resource_group.checkin.name
  location                     = azurerm_resource_group.checkin.location
  version                      = "12.0"
  administrator_login          = var.sql_admin_username
  administrator_login_password = random_password.sql_admin.result
  minimum_tls_version          = "1.2"
  tags                         = var.tags
}

resource "azurerm_mssql_database" "attendees" {
  name      = var.sql_database_name
  server_id = azurerm_mssql_server.checkin.id
  # The numeric suffix on sku_name (e.g. GP_S_Gen5_1 = 1 vCore) sets the max
  # vCore ceiling — there's no separate max_capacity argument on this
  # resource. var.max_capacity exists to document intent; if you want a
  # different ceiling, change var.sku_name's suffix to match.
  sku_name       = var.sku_name
  min_capacity   = var.min_capacity
  max_size_gb    = 2
  zone_redundant = false

  auto_pause_delay_in_minutes = var.auto_pause_delay_in_minutes

  # Trivial data volume (a few hundred rows) — locally redundant storage is
  # plenty; no need for zone/geo redundancy for a one-day event database.
  storage_account_type = "Local"

  tags = var.tags
}

resource "azurerm_mssql_firewall_rule" "azure_services" {
  count            = var.allow_azure_services ? 1 : 0
  name             = "AllowAzureServices"
  server_id        = azurerm_mssql_server.checkin.id
  start_ip_address = "0.0.0.0"
  end_ip_address   = "0.0.0.0"
}

resource "azurerm_mssql_firewall_rule" "client_ips" {
  for_each         = toset(var.allowed_client_ips)
  name             = "client-${replace(each.value, ".", "-")}"
  server_id        = azurerm_mssql_server.checkin.id
  start_ip_address = each.value
  end_ip_address   = each.value
}
