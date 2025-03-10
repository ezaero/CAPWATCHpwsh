# Input bindings are passed in via param block.
param($Timer)

#region ImportClasses
<# 
    Using strongly typed classes allows us to both transform data
    (like dates) from [STRING] to their appropriate object types (like [DATETIME])
#>
. "$PSScriptRoot\classes\Member.ps1"
#endregion

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

#Abort script execution if CAPWATCH data is stale
$DownloadDate = (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours)
Write-Host "Download date is: [$DownloadDate]"
if (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours -gt 48) {
    Write-Error "CAPWATCH data in [$CAPWATCHDATADIR] is stale; aborting script execution!"
    exit 1
}

$CAPWATCH_Member = Import-Csv .\Member.txt -ErrorAction Stop

$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com

$allDistLists = Get-DistributionGroup -ResultSize Unlimited
Write-Host "Total DLs" $allDistLists.count

$allDistLists | ForEach-Object {
    Write-Host "Distribution List: $($_.DisplayName)"
    $members = Get-DistributionGroupMember -Identity $_.Identity
    Write-Host "Members:"
    $members | ForEach-Object {
        # if ($_.RecipientType -eq "UserMailbox" -or $_.RecipientType -eq "MailUser") {
        #     $user = Get-User -Identity $_.Identity
        #     Write-Host $user.DisplayName
        # } else {
            Write-Host $_.Name
        # }
    }
    Write-Host "-----------------------------"
}

# $allMgUsers = Get-MgUser -All

# $MatchedUsers = $allMgUsers.where({ $_.OfficeLocation -in $CAPWATCH_Member.CAPID })

# foreach ($MatchedUser in $MatchedUsers) {
#     $_capwatchMatchedUser = $CAPWATCH_Member.where({ $_.CAPID -eq $MatchedUser.OfficeLocation })
#     if ($_capwatchMatchedUser.count -gt 1) {
#         Write-Error "Multiple members found in CAPWATCH data for $($MatchedUser.UserPrincipalName) $($MatchedUser.OfficeLocation) - skipping"
#         continue
#     }

#     #Build the splat object for the Update-MgUser cmdlet
#     $_updateObject = @{
#         UserId      = ($MatchedUser.UserPrincipalName)
#         DisplayName = ('{0} {1}, {2}' -f $_capwatchMatchedUser.NameFirst.Trim(), $_capwatchMatchedUser.NameLast.Trim(), $_capwatchMatchedUser.Rank.Trim())
#         JobTitle    = ($_capwatchMatchedUser.Rank.Trim())
#     }

#     #Test if the user needs an update
#     $_NeedsUpdate = @(
#         ($MatchedUser.DisplayName -ne $_updateObject.DisplayName)
#         ($MatchedUser.JobTitle -ne $_updateObject.JobTitle)
#     ) -contains $true

#     #Perform updates on users that are out-of-date
#     if ($_NeedsUpdate) {
#         Write-Host "Updating user '$($_updateObject.UserId)' with DisplayName: '$($_updateObject.DisplayName)' and JobTitle '$($_updateObject.JobTitle)'"
#         Update-MgUser @_updateObject
#     }
#     else {
#         Write-Host "Skipping user $($_updateObject.UserId) - already up to date"
#     }
# }