# Input bindings are passed in via param block.
param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

# Main execution block
try {
    Write-Log "Starting maintenance operations..."
    
    # Clean up old log files
    Remove-OldLogFiles -DirectoryPath "$env:HOME\logs"
    
    # Run monthly account deletion maintenance
    Write-Log "Running monthly account deletion maintenance..."
    Remove-ExpiredMemberAccounts
    
    Write-Log "Maintenance operations completed successfully."
} catch {
    Write-Log "Error during maintenance operations: $_"
    throw
}
