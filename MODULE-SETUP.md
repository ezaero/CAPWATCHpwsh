# PowerShell Module Setup - Quick Reference

This directory contains the PowerShell module management system for CAPWATCHSyncPWSH.

## ⚡ Quick Start

After deploying infrastructure with Terraform:

```powershell
# 1. Connect to Azure
Connect-AzAccount

# 2. Download modules locally
./Download_Modules.ps1

# 3. Upload to Azure Storage  
./Upload-ModulesToStorage.ps1 -StorageAccountName "capwatchsyncpwsh" -ResourceGroup "CAPWATCH_Sync_PWSH"

# 4. Deploy function
func azure functionapp publish capwatchsyncpwsh --powershell
```

## 📋 Why This Setup?

Azure Functions has a 150MB deployment size limit, but our required PowerShell modules exceed this:

- `Az.Accounts` + `Az.KeyVault` ≈ 60MB
- `Microsoft.Graph.*` modules ≈ 80MB  
- `ExchangeOnlineManagement` ≈ 40MB
- **Total: ~180MB** (exceeds limit)

## 🏗️ Architecture

```
Deployment Package (53KB)     Azure Storage (180MB modules)
├── Function code             ├── modules/
├── .funcignore (excludes)    │   ├── Az.Accounts.zip
├── requirements.psd1         │   ├── Az.KeyVault.zip
├── profile.ps1               │   ├── ExchangeOnlineManagement.zip
└── shared/Load-Modules.ps1   │   └── Microsoft.Graph.*.zip
                              └── Runtime download & loading
```

## 🔄 Module Loading Strategy

1. **Primary**: Azure Functions managed dependencies (`requirements.psd1`)
2. **Fallback**: Runtime loading from Azure Storage (`shared/Load-Modules.ps1`)
3. **Backup**: PowerShell Gallery installation

## 📁 Key Files

| File | Purpose |
|------|---------|
| `Upload-ModulesToStorage.ps1` | Upload modules to Azure Storage |
| `Download_Modules.ps1` | Download modules locally |
| `shared/Load-Modules.ps1` | Runtime module loading logic |
| `requirements.psd1` | Azure Functions managed dependencies |
| `profile.ps1` | Function startup configuration |
| `.funcignore` | Excludes Modules/ from deployment |

## 📖 Complete Documentation

For detailed step-by-step instructions, see: **[DEPLOYMENT.md](DEPLOYMENT.md)**

## 🆘 Troubleshooting

**Module not found errors?**
- Verify modules uploaded to storage: Azure Portal → Storage Account → modules container
- Check Function App has Storage Blob Data Reader permissions
- Review function logs for module loading details

**Upload script fails?**
- Ensure connected to Azure: `Get-AzContext`
- Verify storage account name and resource group
- Check you have Contributor access to the resource group
