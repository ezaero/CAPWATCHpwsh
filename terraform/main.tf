terraform {
  required_version = ">= 1.0"
  required_providers {
    azurerm = {
      source  = "hashicorp/azurerm"
      version = "~> 3.0"
    }
    azuread = {
      source  = "hashicorp/azuread"
      version = "~> 2.0"
    }
  }
}

provider "azurerm" {
  features {
    key_vault {
      purge_soft_delete_on_destroy    = true
      recover_soft_deleted_key_vaults = true
    }
  }
}

provider "azuread" {}

# Data sources
data "azurerm_client_config" "current" {}
data "azuread_client_config" "current" {}

# Local variables
locals {
  wing_designator = var.wing_designator
  resource_suffix = var.wing_designator
  # Microsoft Graph API permissions required
  graph_permissions = [
    {
      id   = "19dbc75e-c6dc-45ba-9e88-928e6c467ab0" # Directory.ReadWrite.All
      type = "Role"
    },
    {
      id   = "62a82d76-70ea-41e2-9197-370581804d09" # Group.ReadWrite.All
      type = "Role"
    },
    {
      id   = "0121dc95-1b9f-4aed-8bac-58c5ac466691" # TeamMember.ReadWrite.All
      type = "Role"
    },
    {
      id   = "df021288-bdef-4463-88db-98f22de89214" # User.Read.All
      type = "Role"
    },
    {
      id   = "741f803b-c850-494e-b5df-cde7c675a1ca" # User.ReadWrite.All
      type = "Role"
    },
    {
      id   = "b633e1c5-b582-4048-a93e-9f11b44c7e96" # Mail.Send
      type = "Role"
    }
  ]
  
  tags = {
    Environment = "prod"
    Project     = "CAPWATCHSync"
    Wing        = var.wing_designator
    ManagedBy   = "Terraform"
  }
}

# Resource Group
resource "azurerm_resource_group" "capwatch" {
  name     = "capwatch-${local.resource_suffix}-rg"
  location = var.location
  tags     = local.tags
}

# Storage Account for Azure Functions
resource "azurerm_storage_account" "functions" {
  name                     = "capwatch${lower(local.resource_suffix)}sa"
  resource_group_name      = azurerm_resource_group.capwatch.name
  location                 = azurerm_resource_group.capwatch.location
  account_tier             = "Standard"
  account_replication_type = "LRS"
  
  tags = local.tags
}


# Key Vault
resource "azurerm_key_vault" "capwatch" {
  name                        = "capwatch-${local.resource_suffix}-kv"
  location                    = azurerm_resource_group.capwatch.location
  resource_group_name         = azurerm_resource_group.capwatch.name
  enabled_for_disk_encryption = true
  tenant_id                   = data.azurerm_client_config.current.tenant_id
  soft_delete_retention_days  = 7
  purge_protection_enabled    = false
  sku_name                    = "standard"
  
  tags = local.tags
}

# Key Vault Access Policy for current user (for initial setup)
resource "azurerm_key_vault_access_policy" "current_user" {
  key_vault_id = azurerm_key_vault.capwatch.id
  tenant_id    = data.azurerm_client_config.current.tenant_id
  object_id    = data.azurerm_client_config.current.object_id
  
  secret_permissions = [
    "Get",
    "List",
    "Set",
    "Delete",
    "Recover",
    "Backup",
    "Restore"
  ]
}

# App Service Plan
resource "azurerm_service_plan" "capwatch" {
  name                = "capwatch-${local.resource_suffix}-asp"
  resource_group_name = azurerm_resource_group.capwatch.name
  location            = azurerm_resource_group.capwatch.location
  os_type             = "Windows"
  sku_name            = "Y1" # Consumption plan
  
  tags = local.tags
}

# Azure AD Application
resource "azuread_application" "capwatch" {
  display_name = "CAPWATCHSync-${var.wing_designator}"
  
  required_resource_access {
    resource_app_id = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
    dynamic "resource_access" {
      for_each = local.graph_permissions
      content {
        id   = resource_access.value.id
        type = resource_access.value.type
      }
    }
  }

  # Office 365 Exchange Online API permissions
  required_resource_access {
    resource_app_id = "00000002-0000-0ff1-ce00-000000000000" # Exchange Online
    resource_access {
      id   = "dc50a0fb-09a3-484d-be87-e023b12c6440" # Exchange.ManageAsApp
      type = "Role"
    }
    resource_access {
      id   = "64a6cdd6-aab1-4aaf-94b4-7a0a07e1e4a2" # Mail.Send (Exchange Online)
      type = "Role"
    }
  }
  
  tags = ["CAPWATCHSync", var.wing_designator]
}

# Service Principal for the Azure AD Application
resource "azuread_service_principal" "capwatch" {
  client_id                    = azuread_application.capwatch.client_id
  app_role_assignment_required = false
  tags = ["CAPWATCHSync", var.wing_designator]
}

# Function App
resource "azurerm_windows_function_app" "capwatch" {
  name                = "capwatch-${local.resource_suffix}-func"
  resource_group_name = azurerm_resource_group.capwatch.name
  location            = azurerm_resource_group.capwatch.location
  
  storage_account_name       = azurerm_storage_account.functions.name
  storage_account_access_key = azurerm_storage_account.functions.primary_access_key
  service_plan_id            = azurerm_service_plan.capwatch.id
  
  site_config {
    # Application Insights integration removed
    
    application_stack {
      powershell_core_version = "7.4"
    }
  }
  
  app_settings = {
    FUNCTIONS_WORKER_RUNTIME                = "powershell"
    FUNCTIONS_WORKER_RUNTIME_VERSION        = "7.4"
    FUNCTIONS_EXTENSION_VERSION             = "~4"
    WEBSITE_TIME_ZONE                       = var.timezone
    PSWorkerInProcConcurrencyUpperBound     = "1"
    
    # Wing-specific configuration
    WING_DESIGNATOR      = var.wing_designator
    CAPWATCH_ORGID       = var.capwatch_org_id
    KEYVAULT_NAME        = azurerm_key_vault.capwatch.name
    EXCHANGE_ORGANIZATION = var.exchange_organization
    
    # Email configuration (optional)
    LOG_EMAIL_TO_ADDRESS   = var.log_email_to_address
    LOG_EMAIL_FROM_ADDRESS = var.log_email_from_address
    
    # Application Insights (optional)
    APPLICATIONINSIGHTS_CONNECTION_STRING = azurerm_application_insights.capwatch.connection_string
  }
  
  identity {
    type = "SystemAssigned"
  }
  
  tags = local.tags
}

# Key Vault Access Policy for Function App Managed Identity
resource "azurerm_key_vault_access_policy" "function_app" {
  key_vault_id = azurerm_key_vault.capwatch.id
  tenant_id    = azurerm_windows_function_app.capwatch.identity[0].tenant_id
  object_id    = azurerm_windows_function_app.capwatch.identity[0].principal_id
  
  secret_permissions = [
    "Get",
    "List"
  ]
  
  depends_on = [azurerm_windows_function_app.capwatch]
}

# Role assignment for Microsoft Graph API permissions
# Note: This requires admin consent to be granted manually in Azure Portal
resource "azurerm_role_assignment" "graph_permissions" {
  scope                = "/subscriptions/${data.azurerm_client_config.current.subscription_id}"
  role_definition_name = "Reader"
  principal_id         = azuread_service_principal.capwatch.object_id
  
  depends_on = [azuread_service_principal.capwatch]
}

# Application Insights
resource "azurerm_application_insights" "capwatch" {
  name                = "capwatch-${local.resource_suffix}-appi"
  location            = "westus2"
  resource_group_name = azurerm_resource_group.capwatch.name
  application_type    = "web"
  retention_in_days   = 30
  tags                = local.tags
}

output "appinsights_connection_string" {
  description = "Application Insights connection string for use in Function App."
  value       = azurerm_application_insights.capwatch.connection_string
  sensitive   = true
}
