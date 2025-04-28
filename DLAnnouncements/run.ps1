# Input bindings are passed in via param block.
param($Timer)

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com -ShowBanner:$false


# This function compares two arrays and returns the user IDs that are in both, only in the first array, and only in the second array.
function Compare-Arrays {
    param (
        [array]$Array1, # Full user objects from the filtered list
        [array]$Array2  # IDs of current group members
    )

    Write-Host "Inside Compare-Arrays"
    Write-Host "Array1 count: $($Array1.Count)"
    Write-Host "Array2 count: $($Array2.Count)"

    # Find user objects that are in both arrays
    $inBoth = $Array1 | Where-Object { $Array2 -contains $_.id }
    Write-Host "InBoth count: $($inBoth.Count)"

    # Find user objects that are only in Array1
    $Add = $Array1 | Where-Object { $Array2 -notcontains $_.id }
    Write-Host "Add count: $($Add.Count)"

    # Create a hash table for quick lookups of Array1 IDs
    $Array1Hash = @{}
    foreach ($user in $Array1) {
        $Array1Hash[$user.id] = $user
    }

    # Find user objects that are only in Array2
    $Remove = $Array2 | Where-Object { -not $Array1Hash.ContainsKey($_) }
    Write-Host "Remove count: $($Remove.Count)"
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
    Write-Host "Group '$groupName' found. Group ID: $groupId"

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
    Write-Host "Users in both arrays: $($result.InBoth.Count)"  
    Write-Host "Users to add: $($result.Add.Count)"
    Write-Host "Users to remove: $($result.Remove.Count)"
    Write-Host "Adding users to group '$groupName'..."
    # Add users to the group if they are not already members    
    foreach ($user in $result.Add) {
        if ($groupMemberIds -notcontains $user.id) {
            try {
                Add-DistributionGroupMember -Identity $groupName -Member $user.Id
                Write-Host "Added user: $($user.displayName) ($($user.mail)) to group '$groupName'."
            } catch {
                Write-Host "Failed to add user: $($user.displayName) ($($user.mail)) to group '$groupName'. Error: $_"
            }
        } else {
            Write-Host "User: $($user.displayName) ($($user.mail)) is already a member of group '$groupName'."
        }
    }
    # Remove users from the group if they are not in the allUsers list - decided not to do this because if their account is deleted, they will be removed automatically
}

Write-Log "DLAnnouncements script started. ------------------------------------------------"
$allUsers = GetAllUsers

# CO Wing Announcements
$groupName = "CO Wing Announcements"
$groupMemberIds = GetGroupMemberIds -groupName $groupName

# Filter users for group membership
$groupUsers = $allUsers | Where-Object {
    $_.employeeType -eq 'CADET' -or $_.jobTitle -like '*PARENT*' -or $_.employeeType -eq 'SENIOR'
}
$groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

$result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
ModifyGroupMembers -groupName $groupName -result $result
Write-Log "DLAnnouncements script completed. ------------------------------------------------"