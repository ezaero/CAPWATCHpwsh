<#
.SYNOPSIS
    Downloads data from CAPWATCH and extracts it to a folder accesible to the Azure Function App
.DESCRIPTION
    This script executes on a timer as part of an Azure Function App and is responsible for
    downloading and extracting data retrieved from CAPWATCH for Colorado Wing CAP. Data
    downloaded and extracted by this script will reside in the $($env:HOME)\data\CAPWatch
    directory for use by other scripts in this Azure Function App
.NOTES
    This script pulls credentials for CAPWATCH from an Azure Key Vault specified in the $KeyVaultName variable.
    At the time of writing this script, the Key Vault (cowgcapwatch) is set to allow members of the IT staff to
    write secrets but not read them. This Function App is the only resource with permissions to retrieve secret
    values in an effort to protect the personal credentials of the user-account tied to the CAPWATCH download
    API as they are the same credentials used to log-in to eServices.
#>

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Input bindings are passed in via param block.
param($Timer)

$ErrorActionPreference = 'Stop' # Stop on error 12/13/2024 - HK

$CapwatchOrg = $env:CAPWATCH_ORGID # 423 = Broomfield (testing only)
$UnitOnly = 0
$KeyVaultName = 'cowgcapwatch'

$LocalFilePath = "$($env:HOME)\data\capwatch.zip"

class AzureKeyVaultCAPWatch {
    static hidden [string] $kvUsername
    static hidden [string] $kvPassword
    static [void] GetCredentialsFromKeyVault($KeyVaultName) {
        [AzureKeyVaultCAPWatch]::kvUsername = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name 'capwatch-username' -AsPlainText
        [AzureKeyVaultCAPWatch]::kvPassword = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name 'capwatch-password' -AsPlainText
    }
    static [hashtable] GetHeaders() {
        return @{
            'Authorization' = 'Basic ' + [Convert]::ToBase64String(
                [System.Text.Encoding]::UTF8.GetBytes(
                    ('{0}:{1}' -f [AzureKeyVaultCAPWatch]::kvUsername, [AzureKeyVaultCAPWatch]::kvPassword)
                )
            ) 
        }
    }
}

# Write-Log "DEBUGGING: $([AzureKeyVaultCAPWatch]::kvUsername)"
# Write-Log "DEBUGGING: $([AzureKeyVaultCAPWatch]::kvPassword)"

# Tell our class to retrieve credentials from Azure Key Vault
[AzureKeyVaultCAPWatch]::GetCredentialsFromKeyVault($KeyVaultName)

if (Test-Path $LocalFilePath) {
    Write-Log 'Existing CAPWATCH.ZIP file found - deleting...' -NoNewline
    Remove-Item $LocalFilePath -Force # Delete old CAPWATCH ZIP
    Write-Log 'Done'
}

Write-Log 'Downloading CAPWATCH data from CAPNHQ.GOV...' -NoNewline
Invoke-WebRequest `
    -Uri "https://www.capnhq.gov/CAP.CapWatchAPI.Web/api/cw?ORGID=$CapwatchOrg&unitOnly=$UnitOnly" `
    -Headers ([AzureKeyVaultCAPWatch]::GetHeaders()) -OutFile $LocalFilePath -ErrorAction Stop `
    -TimeoutSec 600 # Download CAPWATCH ZIP
Write-Log 'Done'

Write-Log 'Extracting archive...' -NoNewline
Expand-Archive -Path $LocalFilePath  `
    -DestinationPath ('{0}\{1}' -f ([System.IO.Directory]::GetParent($LocalFilePath).FullName), 'CAPWatch') `
    -Force # Extract archive to $($env:HOME)\data\CAPWatch folder and overwrite the file if it still exists somehow
Write-Log 'Done'

Write-Log "download-extract-capwatch Script execution completed at $(Get-Date)"