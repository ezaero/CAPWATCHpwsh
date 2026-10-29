output "resource_group_name" {
  description = "Name of the created resource group"
  value       = azurerm_resource_group.capwatch.name
}

output "function_app_name" {
  description = "Name of the Function App"
  value       = azurerm_windows_function_app.capwatch.name
}

output "function_app_url" {
  description = "URL of the Function App"
  value       = "https://${azurerm_windows_function_app.capwatch.default_hostname}"
}

output "key_vault_name" {
  description = "Name of the Key Vault"
  value       = azurerm_key_vault.capwatch.name
}

output "key_vault_uri" {
  description = "URI of the Key Vault"
  value       = azurerm_key_vault.capwatch.vault_uri
}

output "storage_account_name" {
  description = "Name of the storage account"
  value       = azurerm_storage_account.functions.name
}


output "azure_ad_app_id" {
  description = "Azure AD Application ID"
  value       = azuread_application.capwatch.client_id
}

output "azure_ad_app_name" {
  description = "Azure AD Application Name"
  value       = azuread_application.capwatch.display_name
}

output "function_app_principal_id" {
  description = "Principal ID of the Function App's managed identity"
  value       = azurerm_windows_function_app.capwatch.identity[0].principal_id
}

output "app_service_auth_managed_identity_client_id" {
  description = "Client ID of the user-assigned managed identity used for App Service Authentication federated credentials"
  value       = azurerm_user_assigned_identity.app_service_auth.client_id
}

output "app_service_auth_managed_identity_principal_id" {
  description = "Principal ID of the user-assigned managed identity trusted by the App Service Authentication app registration"
  value       = azurerm_user_assigned_identity.app_service_auth.principal_id
}

output "app_service_auth_client_secret_setting_name" {
  description = "App Service Authentication client secret setting name to select after migration"
  value       = "OVERRIDE_USE_MI_FIC_ASSERTION_CLIENTID"
}

output "deployment_instructions" {
  description = "Post-deployment instructions"
  value       = <<-EOT

  🚀 CAPWATCH Sync Infrastructure Deployed Successfully!

  📋 NEXT STEPS:

  1. 🔐 Add CAPWATCH credentials to Key Vault:
     - Navigate to Key Vault: ${azurerm_key_vault.capwatch.name}
     - Add secret 'capwatch-username' with your CAPWATCH username
     - Add secret 'capwatch-password' with your CAPWATCH password

  2. 📱 Grant Microsoft Graph API permissions:
     - Go to Azure Portal > Azure Active Directory > App registrations
     - Find app: ${azuread_application.capwatch.display_name}
     - Go to API permissions > Grant admin consent for [your tenant]

  3. 📤 Deploy Function App code:
     - Use Azure Functions Core Tools or VS Code Azure Functions extension
     - Deploy from your local repository to: ${azurerm_windows_function_app.capwatch.name}

  4. ⚙️ Configure Exchange Online:
     - Verify Exchange organization: ${var.exchange_organization}
     - Ensure Function App managed identity has Exchange permissions

  5. 🔑 Switch App Service Authentication to managed identity:
     - In Function App > Authentication > Microsoft provider, set Client secret setting name to: OVERRIDE_USE_MI_FIC_ASSERTION_CLIENTID
     - Verify sign-in works
     - Remove the old MICROSOFT_PROVIDER_AUTHENTICATION_SECRET app setting and delete the old client secret from the auth app registration

  6. 🧪 Test deployment:
     - Run the download-extract-capwatch function first
     - Monitor logs in Application Insights (if enabled)

  📊 Resources Created:
  - Resource Group: ${azurerm_resource_group.capwatch.name}
  - Function App: ${azurerm_windows_function_app.capwatch.name}
  - Key Vault: ${azurerm_key_vault.capwatch.name}
  - Storage Account: ${azurerm_storage_account.functions.name}
  - Application Insights: (not managed by Terraform)
  - Azure AD App: ${azuread_application.capwatch.display_name}

  EOT
}
