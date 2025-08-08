# Upload PowerShell modules to Azure Storage for runtime loading
# Run this script locally to upload your modules to Azure Storage

param(
    [Parameter(Mandatory)]
    [string]$StorageAccountName,
    
    [Parameter(Mandatory)] 
    [string]$ResourceGroupName,
    
    [string]$ContainerName = "modules",
    [string]$ModulesPath = "./Modules"
)

# Ensure we're connected to Azure
$context = Get-AzContext -ErrorAction SilentlyContinue
if (-not $context) {
    Write-Host "Please connect to Azure first:"
    Write-Host "Connect-AzAccount"
    exit 1
}

# Get storage account
$storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -ErrorAction SilentlyContinue
if (-not $storageAccount) {
    Write-Error "Storage account '$StorageAccountName' not found in resource group '$ResourceGroupName'"
    exit 1
}

$ctx = $storageAccount.Context

# Create container if it doesn't exist
$container = Get-AzStorageContainer -Name $ContainerName -Context $ctx -ErrorAction SilentlyContinue
if (-not $container) {
    Write-Host "Creating container: $ContainerName"
    New-AzStorageContainer -Name $ContainerName -Context $ctx -Permission Off | Out-Null
}

# Get list of modules to upload
$modulesToUpload = @(
    "Az.Accounts",
    "Az.KeyVault", 
    "ExchangeOnlineManagement",
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Groups", 
    "Microsoft.Graph.Teams",
    "Microsoft.Graph.Users",
    "PackageManagement",
    "PowerShellGet"
)

Write-Host "Uploading modules to Azure Storage..."

foreach ($module in $modulesToUpload) {
    $modulePath = Join-Path $ModulesPath $module
    if (Test-Path $modulePath) {
        Write-Host "Processing module: $module"
        
        # Create zip file  
        $tempDir = if ($IsWindows) { $env:TEMP } else { $env:TMPDIR -replace '/$', '' }
        if (-not $tempDir) { $tempDir = "/tmp" }
        $zipPath = Join-Path $tempDir "$module.zip"
        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force
        }
        
        Compress-Archive -Path $modulePath -DestinationPath $zipPath -Force
        
        # Upload to blob storage
        $blobName = "$module.zip"
        Write-Host "Uploading: $blobName"
        
        Set-AzStorageBlobContent -File $zipPath -Container $ContainerName -Blob $blobName -Context $ctx -Force | Out-Null
        
        # Clean up local zip
        Remove-Item $zipPath -Force
        
        Write-Host "Uploaded: $module"
    }
    else {
        Write-Warning "Module not found: $modulePath"
    }
}

Write-Host ""
Write-Host "Module upload completed!"
Write-Host "Storage Account: $StorageAccountName"
Write-Host "Container: $ContainerName"
Write-Host ""
Write-Host "Next steps:"
Write-Host "1. Add STORAGE_ACCOUNT_NAME = '$StorageAccountName' to your Function App settings"
Write-Host "2. Ensure your Function App has Storage Blob Data Reader role on the storage account"
Write-Host "3. Update your function scripts to call: . `"`$PSScriptRoot\..\shared\Load-Modules.ps1`""
