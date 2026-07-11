variable "resource_group_name" {
  description = "Name of the resource group to create for the event database. Kept separate from any other event infrastructure so it's easy to tear down after the event if desired."
  type        = string
  default     = "rg-sqlsat-checkin"
}

variable "location" {
  description = "Azure region. Pick one close to the venue for lower latency."
  type        = string
  default     = "eastus"
}

variable "sql_server_name" {
  description = "Globally unique Azure SQL logical server name (becomes <name>.database.windows.net). Lowercase letters, numbers, hyphens only."
  type        = string
  default     = "sql-sqlsat-checkin"
}

variable "sql_database_name" {
  description = "Database name. Matches event.config.json azure.database."
  type        = string
  default     = "attendees"
}

variable "sql_admin_username" {
  description = "SQL authentication login for the server admin. Matches event.config.json azure.username."
  type        = string
  default     = "checkinadmin"
}

variable "allowed_client_ips" {
  description = <<-EOT
    Public IPv4 addresses allowed through the SQL server firewall, e.g. the
    venue's WiFi egress IP(s) and your home/office IP for pre-event setup.
    Each entry gets its own firewall rule. Leave empty and add rules with
    `az sql server firewall-rule create` day-of once you know the venue's
    actual egress IP (venue WiFi is usually NAT'd behind one public IP).
  EOT
  type        = list(string)
  default     = []
}

variable "allow_azure_services" {
  description = "Allow other Azure services (e.g. a future Azure-hosted admin tool) to reach the server. Not required for the check-in laptops, which connect directly over the internet."
  type        = bool
  default     = false
}

variable "sku_name" {
  description = "Serverless General Purpose SKU. GP_S_Gen5_1 is the smallest — 1 max vCore family Gen5. Serverless auto-pauses when idle, so this is the cost lever, not the SKU choice itself."
  type        = string
  default     = "GP_S_Gen5_1"
}

variable "min_capacity" {
  description = "Minimum vCores when active (serverless allows fractional)."
  type        = number
  default     = 0.5
}

variable "max_capacity" {
  description = "Maximum vCores when active. A handful of check-in laptops doing point lookups/inserts is trivial load."
  type        = number
  default     = 1
}

variable "auto_pause_delay_in_minutes" {
  description = "Minutes of inactivity before the database auto-pauses (billing drops to storage-only). 60 min covers normal gaps during the day without pausing mid-registration-rush; it'll pause overnight and for the ~360 idle days/year."
  type        = number
  default     = 60
}

variable "tags" {
  description = "Tags applied to all resources."
  type        = map(string)
  default = {
    project = "sqlsat-tools"
    purpose = "attendee-checkin"
  }
}
