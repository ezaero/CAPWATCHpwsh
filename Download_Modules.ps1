<#
  .SYNOPSIS
    Downloads required modules and saves them to a local folder to reduce startup time. Run this script if you are getting errors about missing modules or need to update
#>

$Modules = @(
    @{Name = 'Az.Accounts'; RequiredVersion = '3.0.4' }
    @{Name = 'Az.KeyVault'; RequiredVersion = '6.2.0' }
    @{Name = 'Microsoft.Graph.Authentication'; RequiredVersion = '2.24.0' }
    @{Name = 'Microsoft.Graph.Users'; RequiredVersion = '2.24.0' }
    @{Name = 'ExchangeOnlineManagement'; MinimumVersion = '3.0.0' }
    @(Name = 'Microsoft.Graph'; MinimumVersion = '2.0.0' )
)

foreach ($Module in $Modules) {
    Save-Module @Module -Path "$PSScriptRoot\Modules" -Force -MinimumVersion
}