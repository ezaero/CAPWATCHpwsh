# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Initialize runtime modules from Azure Storage as fallback
. "$PSScriptRoot\shared\Load-Modules.ps1"

# Check if the primary modules are available, if not, initialize our custom loading
$criticalModules = @("Az.Accounts", "Az.KeyVault")
$needsCustomLoading = $false
foreach ($module in $criticalModules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        $needsCustomLoading = $true
        break
    }
}

if ($needsCustomLoading) {
    Write-Host "Primary modules not available through requirements.psd1, initializing custom loading..."
    Initialize-RuntimeModules -StorageAccountName $env:STORAGE_ACCOUNT_NAME -ContainerName "modules"
} else {
    Write-Host "Modules available through requirements.psd1, skipping custom loading"
}

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
if ($env:MSI_SECRET) {
    try {
        # Import Az.Accounts if available
        if (Get-Module -Name Az.Accounts -ListAvailable) {
            Import-Module Az.Accounts -Force -ErrorAction SilentlyContinue
            Disable-AzContextAutosave -Scope Process | Out-Null
            Connect-AzAccount -Identity
        } else {
            Write-Warning "Az.Accounts module not available for Azure authentication"
        }
    } catch {
        Write-Warning "Failed to authenticate with Azure: $_"
    }
}

# Uncomment the next line to enable legacy AzureRm alias in Azure PowerShell.
# Enable-AzureRmAlias

# You can also define functions or aliases that can be referenced in any of your PowerShell functions.
