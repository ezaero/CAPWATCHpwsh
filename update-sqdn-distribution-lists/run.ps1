<#
.SYNOPSIS
    Updates distribution lists in COWG Exchange Online
.DESCRIPTION
    Uses CAPWATCH data downloaded daily by download-extract-capwatch to update distribution lists with
    accurate member information.
.NOTES
    Author: Hunter Klein, 2d Lt.
    ---
    Version History
    2025.3.24.001a - Initial release
#>

param($Timer)

<#
Run locally:
PS C:\> $env:AZURE_FUNCTIONS_ENVIRONMENT = "DEV"
#>
switch ($env:AZURE_FUNCTIONS_ENVIRONMENT) {
    'DEV' {
        $env:HOME = "$PSScriptRoot\..\data"
        Connect-ExchangeOnline -UserPrincipalName 'hunter.klein@cowg.cap.gov' #replace with your username if running locally for development
        Connect-MgGraph -Scopes 'User.Read.All', 'Group.ReadWrite.All'
    }
    default {
        Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com
        $MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
        Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
    }
}

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

$allDistLists = Get-DistributionGroup -ResultSize Unlimited

# Get all microsoft graph users
$allMGUsers = Get-MgUser -All -Property Id, mail, employeeId, OfficeLocation

# Load CAPWATCH member list
$CSVMembers = Import-Csv .\Member.txt

## Group function
function Group-MembersIntoDistributionList {
    param(
        [ValidateSet('SENIOR', 'CADET')]
        [String]$MemberType,
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser[]]$MGUserList
        ,
        $CSVMembers
        ,
        $AllDistributionLists
    )
    $CSVGrouped = $CSVMembers.Where({ $_.Type -eq $MemberType }) | Group-Object -Property 'Unit'
    $Regex = "CO-(\d+) $MemberType" + 's'
    $AllDistributionLists | Where-Object { $_.Name -match $Regex } | ForEach-Object {
        $UNIT = ($_.Name | Select-String 'CO-(\d+)').Matches.Groups[1].Value
        $CAPID = $CSVGrouped.where({ $_.NAME -eq $UNIT }).Group.CAPID
        $MgUsers = $MGUserList.where({ $_.OfficeLocation -in $CAPID })
        [PSCustomObject]@{
            Name    = $_.Name
            Members = $MgUsers.id
            MembersEmail = $MgUsers.mail
        }
    }
}

$_GroupSharedParams = @{
    AllDistributionLists = $allDistLists
    CSVMembers           = $CSVMembers
    MGUserList           = $allMGUsers
}

$SeniorLists = Group-MembersIntoDistributionList @_GroupSharedParams -MemberType 'SENIOR'
$CadetLists = Group-MembersIntoDistributionList @_GroupSharedParams -MemberType 'CADET'

#Remove filter in production. Right now this targets just broomfield for testing.
$SeniorLists.where({$_.Name -like '*099*'}) | ForEach-Object {
    Update-DistributionGroupMember -Identity $_.Name -Members $_.Members -Confirm:$false
}

#Remove filter in production. Right now this targets just broomfield for testing.
$CadetLists.where({ $_.Name -like '*099*' }) | ForEach-Object {
    Update-DistributionGroupMember -Identity $_.Name -Members $_.Members -Confirm:$false
}

Pop-Location