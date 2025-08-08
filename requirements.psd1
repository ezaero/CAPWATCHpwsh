# Azure Functions PowerShell Managed Dependencies
# 
# This file enables modules to be automatically managed by the Functions service.
# See https://aka.ms/functionsmanageddependency for additional information.
#
# IMPORTANT: These modules are also uploaded to Azure Storage as a fallback mechanism.
# For deployment instructions, see: DEPLOYMENT.md (Step 2: PowerShell Module Setup)
#
# The Function App uses a hybrid approach:
#   1. Primary: Azure Functions managed dependencies (this file)
#   2. Fallback: Runtime loading from Azure Storage (shared/Load-Modules.ps1)
#   3. Backup: PowerShell Gallery installation
#
@{
    'Az.Accounts' = '4.*'
    'Az.KeyVault' = '6.*'
    'ExchangeOnlineManagement' = '3.*'
    'Microsoft.Graph.Authentication' = '2.*'
    'Microsoft.Graph.Groups' = '2.*'
    'Microsoft.Graph.Teams' = '2.*'
    'Microsoft.Graph.Users' = '2.*'
}