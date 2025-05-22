# Input bindings are passed in via param block.
param($Timer)

# Include shared Functions
 . "$PSScriptRoot\..\shared\shared.ps1"

 # Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

$OrganizationFile = "$($CAPWATCHDATADIR)/Organization.txt"
# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com

function GetUnits {
    # Create a list of all Unit charter numbers and names in the Wing
    $organization_all = Import-Csv -Path $OrganizationFile
    $co_org = $organization_all | Where-Object { $_.Wing -eq "CO" } | Sort-Object Unit -Unique
    $co_org = $co_org | Select-Object Unit, Name
    # unitList will be a list of all the Distribution Groups required
    $unitList = @()
    foreach ($unit in $co_org) {
        if ($unit.Unit -ne "000" -and $unit.Unit -ne "999" -and $unit.Unit -ne "001") {
            $unitList += $unit
        }
    }
    $unitList
    # Check if the distribution group exists
}

function SquadronGroups {
    param (
        [string]$memberType,
        [array]$unitList,
        [array]$allUsers 
    )

    $memberName = ($memberType.Substring(0,1).ToUpper()) + ($memberType.Substring(1).ToLower()) + 's'
    foreach ($unit in $unitList) {
        $unitDesginator = "CO-$($unit.Unit)"
        $groupName = "CO-$($unit.Unit) $memberName"
        # Uncomment this section to create the distribution group if it doesn't exist
        # $group = Get-DistributionGroup -Identity $groupName -ErrorAction SilentlyContinue
        # # Check if the distribution group exists and create it if it doesn't
        # if (-not $group) {
        #     Write-Log "Distribution group '$groupName' does not exist. Creating..."
        #     $SMTPAddress = "CO-$($unit.Unit)-$memberName@cowg.cap.gov"
        #     $group = New-DistributionGroup -Name $groupName -DisplayName $groupName -PrimarySmtpAddress $SMTPAddress
        #     Write-Log "Distribution group '$groupName' created at $SMTPAddress."
        # } else {
        #     Write-Log "Distribution group '$groupName' already exists."
        # }
        $groupMembers = $allUsers | Where-Object { $_.companyName -eq $unitDesginator -and $_.employeeType -eq $memberType } | Select-Object -ExpandProperty mail
        if ($memberType -eq "CADET") {
            $groupMembers += $allUsers | Where-Object { $_.companyName -eq $unitDesginator -and $_.employeeType -eq "PARENT" } | Select-Object -ExpandProperty mail
            $groupMembers += $allUsers | Where-Object { $_.companyName -eq $unitDesginator -and ($_.department -like "*EX*" -or $_.department -like "*CP*") } | Select-Object -ExpandProperty mail
        }
        Update-DistributionGroupMember -Identity $groupName -Members $groupMembers -Confirm:$false
        Write-Log "Distribution group '$groupName' has $($groupMembers.count) members."
    }
}
# This function compares two arrays and returns the user IDs that are in both, only in the first array, and only in the second array.
function Compare-Arrays {
    param (
        [array]$Array1, # Full user objects from the filtered list
        [array]$Array2  # IDs of current group members
    )

    Write-Log "Inside Compare-Arrays"
    Write-Log "Array1 count: $($Array1.Count)"
    Write-Log "Array2 count: $($Array2.Count)"

    # Find user objects that are in both arrays
    $inBoth = $Array1 | Where-Object { $Array2 -contains $_.id }
    Write-Log "InBoth count: $($inBoth.Count)"

    # Find user objects that are only in Array1
    $Add = $Array1 | Where-Object { $Array2 -notcontains $_.id }
    Write-Log "Add count: $($Add.Count)"

    # Create a hash table for quick lookups of Array1 IDs
    $Array1Hash = @{}
    foreach ($user in $Array1) {
        $Array1Hash[$user.id] = $user
    }

    # Find user objects that are only in Array2
    $Remove = $Array2 | Where-Object { -not $Array1Hash.ContainsKey($_) }
    Write-Log "Remove count: $($Remove.Count)"
    # Output the results
    $result = [PSCustomObject]@{
        InBoth      = $inBoth
        Add         = $Add
        Remove      = $Remove
    }
    return $result
}

function GetGroupMemberIds {
    param (
        [string]$groupName
    )

    $group = Get-MgGroup -Filter "displayName eq '$groupName'"
    $groupId = $group.Id
    Write-Log "Group '$groupName' found. Group ID: $groupId"

    # Get all current members of the group
    $groupMembers = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members?$select=id"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $groupMembers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    # Return only the IDs of the group members
    $groupMemberIds = $groupMembers | ForEach-Object { $_.id } | Where-Object { $_ -ne $null -and $_ -ne "" } 

    return $groupMemberIds
}

function ModifyGroupMembers {
    param (
        [string]$groupName,
        [PSCustomObject]$result
    )
    Write-Log "Users in both arrays: $($result.InBoth.Count)"  
    Write-Log "Users to add: $($result.Add.Count)"
    Write-Log "Users to remove: $($result.Remove.Count)"
    Write-Log "Adding users to group '$groupName'..."
    # Add users to the group if they are not already members    
    foreach ($user in $result.Add) {
        if ($groupMemberIds -notcontains $user.id) {
            try {
                Add-DistributionGroupMember -Identity $groupName -Member $user.mail
                Write-Log "Added user: $($user.displayName) ($($user.mail)) to group '$groupName'."
            } catch {
                Write-Log "Failed to add user: $($user.displayName) ($($user.mail)) to group '$groupName'. Error: $_"
            }
        } else {
            Write-Log "User: $($user.displayName) ($($user.mail)) is already a member of group '$groupName'."
        }
    }
    # Remove users from the group if they are not in the allUsers list - decided not to do this because of the seniors - and if their account is deleted, they will be removed automatically
}

Write-Log " DLCadets script started. ------------------------------------------------"
$allUsers = GetAllUsers

# CO Wing Cadets
$groupName = "CO Wing Cadets"
#n$groupMemberIds = GetGroupMemberIds -groupName $groupName
$groupMemberIds = $allUsers | Where-Object { $_.employeeType -eq 'CADET' -or $_.jobTitle -like '*PARENT*' } | Select-Object -ExpandProperty mail
$groupMemberIds += Get-DistributionGroupMember -Identity "Cadet Programs" -ResultSize Unlimited | Select-Object -ExpandProperty PrimarySmtpAddress
$groupMemberIds += Get-DistributionGroupMember -Identity "CO Wing Commanders" -ResultSize Unlimited | Select-Object -ExpandProperty PrimarySmtpAddress

$groupUsers = $groupMemberIds | Where-Object { 
    $_ -and ($_ -is [string]) -and ($_.Trim() -ne "") 
} | Select-Object -Unique

Write-Log "Group '$groupName' has $($groupUsers.Count) members."
Update-DistributionGroupMember -Identity $groupName -Members $groupUsers -Confirm:$false
Write-Log " DLCadets script completed. ------------------------------------------------"

Write-Log " DLSeniors script started. ------------------------------------------------"
$allUsers = GetAllUsers

# CO Wing Seniors
$groupName = "CO Wing Seniors"
$groupMemberIds = $allUsers | Where-Object { $_.employeeType -eq 'SENIOR' } | Select-Object -ExpandProperty mail
$groupUsers = $groupMemberIds | Where-Object { 
    $_ -and ($_ -is [string]) -and ($_.Trim() -ne "") 
} | Select-Object -Unique

Write-Log "Group '$groupName' has $($groupUsers.Count) members."
Update-DistributionGroupMember -Identity $groupName -Members $groupUsers -Confirm:$false
Write-Log " DLSeniors script completed. ------------------------------------------------"

# CO Wing AEM List
$groupName = "CO Wing AEMList"
$groupMembers = $allUsers | Where-Object { $_.companyName -eq 'CO-004' } | Select-Object -ExpandProperty mail

Update-DistributionGroupMember -Identity $groupName -Members $groupMembers -Confirm:$false


Write-Log "Squadron Seniors/Cadets script started. ------------------------------------------------"

$unitList = GetUnits
SquadronGroups -memberType "SENIOR" -unitList $unitList -allUsers $allUsers
SquadronGroups -memberType "CADET" -unitList $unitList -allUsers $allUsers
Write-Log "Squadron Seniors/Cadets script ended. ------------------------------------------------"