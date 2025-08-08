# CAPWATCHSyncPWSH Deployment Guide

This guide walks you through deploying the CAPWATCHSyncPWSH Azure Function App from start to finish.

## Deployment Overview

The deployment process involves several key steps:

1. **Infrastructure Deployment** - Deploy Azure resources using Terraform
2. **PowerShell Module Setup** - Upload required modules to Azure Storage ⭐ **You are here**
3. **Function App Configuration** - Configure permissions and settings
4. **Function Deployment** - Deploy the PowerShell code
5. **Testing & Validation** - Verify everything works

---

## Step 2: PowerShell Module Setup

After deploying the infrastructure with Terraform, you need to upload the required PowerShell modules to Azure Storage. This is a **critical step** that must be completed before deploying the function code.

### Why This Step is Required

The Azure Function App requires several PowerShell modules to operate:
- `Az.Accounts` - Azure PowerShell authentication
- `Az.KeyVault` - Access to Azure Key Vault secrets  
- `ExchangeOnlineManagement` - Exchange Online operations
- `Microsoft.Graph.*` - Microsoft Graph API operations

Due to Azure Functions deployment size limits (150MB), we cannot include these modules directly in the deployment package. Instead, we:

1. **Exclude modules from deployment** using `.funcignore`
2. **Upload modules to Azure Storage** for runtime access
3. **Load modules dynamically** when the function starts

### Prerequisites

Before running the module upload script, ensure you have:

✅ **Terraform infrastructure deployed** - Storage account and other resources must exist  
✅ **PowerShell 7+ installed** on your local machine  
✅ **Az PowerShell module installed** locally:
```powershell
Install-Module Az -Force -AllowClobber
```
✅ **Connected to Azure**:
```powershell
Connect-AzAccount
```

### Step-by-Step Module Upload Process

#### 1. Download Required Modules Locally

First, download all required modules to your local `./Modules` directory:

```powershell
# Run from the project root directory
./Download_Modules.ps1
```

This script will:
- Create a `./Modules` directory
- Download the latest versions of all required modules
- Organize them in the correct structure for upload

#### 2. Upload Modules to Azure Storage

Run the upload script with your Terraform-created storage account details:

```powershell
# Run from the project root directory
./Upload-ModulesToStorage.ps1 -StorageAccountName "your-storage-account-name" -ResourceGroup "your-resource-group-name"
```

**Example using typical Terraform naming:**
```powershell
./Upload-ModulesToStorage.ps1 -StorageAccountName "capwatchsyncpwsh" -ResourceGroup "CAPWATCH_Sync_PWSH"
```

#### 3. What the Upload Script Does

The `Upload-ModulesToStorage.ps1` script performs these actions:

1. **Validates Azure Connection** - Ensures you're authenticated to Azure
2. **Locates Storage Account** - Finds the storage account created by Terraform
3. **Creates Storage Container** - Creates a "modules" container if it doesn't exist
4. **Compresses Modules** - Zips each module directory for efficient transfer
5. **Uploads to Blob Storage** - Uploads each module zip file to Azure Storage
6. **Cleans Up** - Removes temporary zip files from your local machine

#### 4. Expected Output

You should see output similar to this:

```
Storage account 'capwatchsyncpwsh' found
Storage container 'modules' created/verified
Processing module: Az.Accounts
  - Compressing to: /tmp/Az.Accounts.zip
  - Uploading Az.Accounts.zip... ✓
  - Cleanup: removed /tmp/Az.Accounts.zip
Processing module: Az.KeyVault
  - Compressing to: /tmp/Az.KeyVault.zip  
  - Uploading Az.KeyVault.zip... ✓
  - Cleanup: removed /tmp/Az.KeyVault.zip
...
All modules uploaded successfully!
```

#### 5. Verify Upload Success

You can verify the modules were uploaded correctly by checking the Azure portal:

1. Navigate to your storage account
2. Go to **Containers** → **modules**
3. You should see zip files for each module:
   - `Az.Accounts.zip`
   - `Az.KeyVault.zip`
   - `ExchangeOnlineManagement.zip`
   - `Microsoft.Graph.Authentication.zip`
   - `Microsoft.Graph.Groups.zip`
   - `Microsoft.Graph.Teams.zip`
   - `Microsoft.Graph.Users.zip`

### Runtime Module Loading

Once uploaded, the Azure Function will automatically:

1. **Check for modules** during function startup (profile.ps1)
2. **Download from storage** if modules aren't available locally
3. **Extract and import** modules for use by function code
4. **Fall back to PowerShell Gallery** if storage access fails

The module loading process is handled by:
- `shared/Load-Modules.ps1` - Runtime module loading logic
- `requirements.psd1` - Azure Functions managed dependencies (primary method)
- Storage-based loading as fallback for reliability

### Troubleshooting

**Problem:** "Storage account not found"
- **Solution:** Verify the storage account name and resource group are correct
- **Check:** Ensure Terraform deployment completed successfully

**Problem:** "Access denied" during upload
- **Solution:** Verify you're connected to Azure with `Get-AzContext`
- **Check:** Ensure your Azure account has Contributor access to the resource group

**Problem:** Upload is very slow
- **Expected:** Module uploads can take 5-10 minutes depending on connection speed
- **Normal:** Az.Accounts and Microsoft.Graph modules are particularly large

**Problem:** Function still can't find modules after upload
- **Check:** Verify the Function App has Storage Blob Data Reader permissions
- **Solution:** This is configured in the next deployment step

---

## Next Steps

After successfully uploading modules to Azure Storage:

1. **Configure Function App Settings** - Set environment variables and permissions
2. **Deploy Function Code** - Use `func azure functionapp publish`
3. **Test Function Execution** - Verify modules load correctly at runtime

Continue to **Step 3: Function App Configuration** in the deployment guide.

---

## Module Upload Reference

### Required Modules List
```
Az.Accounts (v4.x)
Az.KeyVault (v6.x)  
ExchangeOnlineManagement (v3.x)
Microsoft.Graph.Authentication (v2.x)
Microsoft.Graph.Groups (v2.x)
Microsoft.Graph.Teams (v2.x)
Microsoft.Graph.Users (v2.x)
```

### Storage Structure
```
Storage Account: {your-storage-account}
├── Container: modules
    ├── Az.Accounts.zip
    ├── Az.KeyVault.zip
    ├── ExchangeOnlineManagement.zip
    ├── Microsoft.Graph.Authentication.zip
    ├── Microsoft.Graph.Groups.zip
    ├── Microsoft.Graph.Teams.zip
    └── Microsoft.Graph.Users.zip
```

### Key Files Modified for Module Management
- `.funcignore` - Excludes `Modules/` directory from deployment
- `requirements.psd1` - Specifies PowerShell dependencies for Azure Functions
- `shared/Load-Modules.ps1` - Runtime module loading logic
- `profile.ps1` - Function startup configuration
