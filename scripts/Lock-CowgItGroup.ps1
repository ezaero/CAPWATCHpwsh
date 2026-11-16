[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Identity = "cowg_it@cowg.cap.gov",
    [switch]$Execute,
    [string]$Organization = $env:EXCHANGE_ORGANIZATION,
    [string]$UserPrincipalName,
    [switch]$Device,
    [switch]$ManagedIdentity
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$localExchangeModule = Join-Path $repoRoot "Modules/ExchangeOnlineManagement/3.7.0/ExchangeOnlineManagement.psd1"

if (Test-Path $localExchangeModule) {
    Import-Module $localExchangeModule -Force
} else {
    Import-Module ExchangeOnlineManagement -Force
}

$connectParams = @{
    ShowBanner = $false
}

if ($ManagedIdentity) {
    $connectParams.ManagedIdentity = $true
    if ($Organization) {
        $connectParams.Organization = $Organization
    }
} elseif ($UserPrincipalName) {
    $connectParams.UserPrincipalName = $UserPrincipalName
}

if ($Device) {
    $connectParams.Device = $true
}

Connect-ExchangeOnline @connectParams

try {
    $group = Get-UnifiedGroup -Identity $Identity -IncludeAllProperties
    $owners = @(Get-UnifiedGroupLinks -Identity $Identity -LinkType Owners)
    $members = @(Get-UnifiedGroupLinks -Identity $Identity -LinkType Members)

    [PSCustomObject]@{
        DisplayName                         = $group.DisplayName
        PrimarySmtpAddress                  = $group.PrimarySmtpAddress
        AccessType                          = $group.AccessType
        HiddenFromAddressListsEnabled       = $group.HiddenFromAddressListsEnabled
        RequireSenderAuthenticationEnabled  = $group.RequireSenderAuthenticationEnabled
        OwnerCount                          = $owners.Count
        MemberCount                         = $members.Count
    } | Format-List

    Write-Host "Owners:"
    $owners |
        Sort-Object DisplayName |
        Select-Object DisplayName, PrimarySmtpAddress |
        Format-Table -AutoSize

    Write-Host "Members:"
    $members |
        Sort-Object DisplayName |
        Select-Object DisplayName, PrimarySmtpAddress |
        Format-Table -AutoSize

    if ($group.AccessType -eq "Private") {
        Write-Host "No privacy change needed: $Identity is already private."
        return
    }

    if (-not $Execute) {
        Write-Host "Dry run only. Re-run with -Execute to set $Identity to Private."
        return
    }

    if ($PSCmdlet.ShouldProcess($Identity, "Set Microsoft 365 group access type to Private")) {
        Set-UnifiedGroup -Identity $Identity -AccessType Private
        $updatedGroup = Get-UnifiedGroup -Identity $Identity

        [PSCustomObject]@{
            DisplayName        = $updatedGroup.DisplayName
            PrimarySmtpAddress = $updatedGroup.PrimarySmtpAddress
            AccessType         = $updatedGroup.AccessType
        } | Format-List
    }
} finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
