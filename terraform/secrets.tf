# Key Vault Secrets for CAPWATCH Credentials
# These secrets are required for the PowerShell functions to authenticate with CAPWATCH API

resource "azurerm_key_vault_secret" "capwatch_username" {
  name         = "capwatch-username"
  value        = var.capwatch_username
  key_vault_id = azurerm_key_vault.capwatch.id
  
  depends_on = [
    azurerm_key_vault_access_policy.current_user
  ]
  
  tags = {
    Environment = "prod"
    Purpose     = "CAPWATCH API Authentication"
    ManagedBy   = "Terraform"
  }
}

resource "azurerm_key_vault_secret" "capwatch_password" {
  name         = "capwatch-password"
  value        = var.capwatch_password
  key_vault_id = azurerm_key_vault.capwatch.id
  
  depends_on = [
    azurerm_key_vault_access_policy.current_user
  ]
  
  tags = {
    Environment = "prod"
    Purpose     = "CAPWATCH API Authentication"
    ManagedBy   = "Terraform"
  }
}
