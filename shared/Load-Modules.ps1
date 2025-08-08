# Runtime module loading for Azure Functions
# 
# This script downloads and loads PowerShell modules from Azure Storage at runtime.
# It works in conjunction with Upload-ModulesToStorage.ps1 to enable module loading
# without exceeding Azure Functions deployment size limits.
#
# Process:
#   1. Modules are uploaded to Azure Storage during deployment (see DEPLOYMENT.md)
#   2. This script downloads modules from storage at function startup
#   3. Falls back to PowerShell Gallery if storage access fails
#   4. Imports modules for use by function code
#
# Called by: profile.ps1 during function app startup
# Dependencies: Azure Storage (modules container) created by Upload-ModulesToStorage.ps1
# Fallback: PowerShell Gallery for module installation
#
# For setup instructions, see: DEPLOYMENT.md (Step 2: PowerShell Module Setup)

param(
    [string]$StorageAccountName = $env:STORAGE_ACCOUNT_NAME,
    [string]$ContainerName = "modules"
)

function Initialize-RuntimeModules {
    param(
        [string]$StorageAccountName,
        [string]$ContainerName = "modules"
    )
    
    Write-Host "Initializing runtime modules..."
    
    # Check if modules are already loaded in this session
    $requiredModules = @(
        "Az.Accounts",
        "Az.KeyVault", 
        "ExchangeOnlineManagement",
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Teams",
        "Microsoft.Graph.Users"
    )
    
    $allModulesLoaded = $true
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            $allModulesLoaded = $false
            break
        }
    }
    
    if ($allModulesLoaded) {
        Write-Host "All required modules already available."
        return
    }
    
    # Set modules path - use Azure Functions temp directory if available
    $modulesPath = if ($env:TEMP) { 
        "$env:TEMP\FunctionModules" 
    } else { 
        "/tmp/FunctionModules" 
    }
    
    if (-not (Test-Path $modulesPath)) {
        New-Item -Path $modulesPath -ItemType Directory -Force | Out-Null
    }
    
    # Add to PSModulePath if not already there
    if ($env:PSModulePath -notlike "*$modulesPath*") {
        $env:PSModulePath = "$modulesPath;$env:PSModulePath"
    }
    
    # First, try to install required modules from PowerShell Gallery if not available
    Write-Host "Checking for required modules..."
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            try {
                Write-Host "Installing module from PowerShell Gallery: $module"
                Install-Module $module -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
                Write-Host "Successfully installed: $module"
            }
            catch {
                Write-Warning "Failed to install module ${module} from PowerShell Gallery: $_"
            }
        }
    }

    # Download modules from Azure Storage if they don't exist locally (optional optimization)
    if ($StorageAccountName) {
        try {
            Write-Host "Attempting to download modules from storage account: $StorageAccountName"
            
            # Try to use Azure context if available
            $context = $null
            if (Get-Module -Name Az.Accounts -ListAvailable) {
                Import-Module Az.Accounts -Force -ErrorAction SilentlyContinue
                $context = (Get-AzContext -ErrorAction SilentlyContinue)
                if (-not $context) {
                    Write-Host "Connecting to Azure with managed identity..."
                    try {
                        Connect-AzAccount -Identity -ErrorAction Stop
                        $context = Get-AzContext
                    }
                    catch {
                        Write-Warning "Failed to connect to Azure: $_"
                    }
                }
            }
            
            
            # Download each required module if we have Azure context
            if ($context) {
                foreach ($module in $requiredModules) {
                    $localModulePath = Join-Path $modulesPath $module
                    if (-not (Test-Path $localModulePath)) {
                        Write-Host "Downloading module: $module"
                        $blobName = "$module.zip"
                        $tempDir = if ($env:TEMP) { $env:TEMP } else { "/tmp" }
                        $zipPath = Join-Path $tempDir $blobName
                        
                        try {
                            # Download from blob storage
                            $storageAccount = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $StorageAccountName } | Select-Object -First 1
                            if ($storageAccount) {
                                $ctx = $storageAccount.Context
                                Get-AzStorageBlobContent -Blob $blobName -Container $ContainerName -Destination $zipPath -Context $ctx -Force -ErrorAction Stop
                                
                                if (Test-Path $zipPath) {
                                    # Extract module
                                    Expand-Archive -Path $zipPath -DestinationPath $modulesPath -Force
                                    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
                                    Write-Host "Downloaded and extracted: $module"
                                }
                            } else {
                                Write-Warning "Storage account '$StorageAccountName' not found"
                            }
                        }
                        catch {
                            Write-Warning "Failed to download module ${module}: $_"
                        }
                    }
                }
            } else {
                Write-Warning "No Azure context available for storage download"
            }
        }
        catch {
            Write-Warning "Failed to download modules from storage: $_"
        }
    } else {
        Write-Host "STORAGE_ACCOUNT_NAME environment variable not set, skipping storage download"
    }
    
    # Import all modules with error handling
    Write-Host "Importing modules..."
    foreach ($module in $requiredModules) {
        try {
            if (-not (Get-Module -Name $module)) {
                Write-Host "Importing module: $module"
                Import-Module $module -Force -Global -ErrorAction Stop
                Write-Host "Successfully imported: $module"
            }
        }
        catch {
            Write-Warning "Failed to import module ${module}: $_"
            # Try to reinstall and import if import fails
            try {
                Write-Host "Retrying installation and import for: $module"
                Install-Module $module -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
                Import-Module $module -Force -Global -ErrorAction Stop
                Write-Host "Successfully reinstalled and imported: $module"
            }
            catch {
                Write-Error "Critical: Failed to install/import module ${module}: $_"
            }
        }
    }
    
    # Verify modules are loaded
    Write-Host "Verifying module availability..."
    foreach ($module in $requiredModules) {
        $moduleLoaded = Get-Module -Name $module
        if ($moduleLoaded) {
            Write-Host "✓ $module is loaded (Version: $($moduleLoaded.Version))"
        } else {
            Write-Warning "✗ $module is NOT loaded"
        }
    }
}

# Auto-initialize if script is called directly
if ($MyInvocation.InvocationName -ne '.') {
    Initialize-RuntimeModules -StorageAccountName $StorageAccountName -ContainerName $ContainerName
}
