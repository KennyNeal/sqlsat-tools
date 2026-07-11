output "sql_server_fqdn" {
  description = "Fully-qualified server name for event.config.json azure.server."
  value       = azurerm_mssql_server.checkin.fully_qualified_domain_name
}

output "sql_database_name" {
  description = "Database name for event.config.json azure.database."
  value       = azurerm_mssql_database.attendees.name
}

output "sql_admin_username" {
  description = "SQL login for event.config.json azure.username."
  value       = var.sql_admin_username
}

output "sql_admin_password" {
  description = "Generated admin password. Store it with Set-Secret, then forget it — never put it in event.config.json. Run: terraform output -raw sql_admin_password"
  value       = random_password.sql_admin.result
  sensitive   = true
}

output "next_steps" {
  value = <<-EOT
    1. Store the password in the secret vault on each check-in laptop:
         Set-Secret -Name 'AzureSqlAuth' -Secret (terraform output -raw sql_admin_password)
    2. Fill in event.config.json's azure section:
         server:   ${azurerm_mssql_server.checkin.fully_qualified_domain_name}
         database: ${azurerm_mssql_database.attendees.name}
         username: ${var.sql_admin_username}
         enabled:  true
    3. From the repo root, run:
         .\scripts\Initialize-AzureDatabase.ps1 -Config (Get-Content .\event.config.json | ConvertFrom-Json)
    4. If allowed_client_ips was empty, add the venue's egress IP before event day:
         az sql server firewall-rule create -g ${var.resource_group_name} -s ${var.sql_server_name} -n venue --start-ip-address <ip> --end-ip-address <ip>
  EOT
}
