# Input bindings are passed in via param block.
param($Timer)

$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
 . "$PSScriptRoot\..\shared\shared.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com

# Import the CSV file into an array
$achievements_all = Import-Csv "$($CAPWATCHDATADIR)\MbrAchievements.txt" -ErrorAction Stop

function Compare-Arrays {
    param (
        [array]$Array1, # Full user objects from the filtered list
        [array]$Array2  # IDs of current group members
    )

    Write-Log "Inside Compare-Arrays"
    Write-Log "Array1 count: $($Array1.Count)"
    Write-Log "Array2 count: $($Array2.Count)"

    # Ensure Array1 is unique
    $Array1 = $Array1 | Sort-Object -Property id -Unique

    # Find user objects that are in both arrays
    $inBoth = $Array1 | Where-Object { $Array2 -contains $_.id }
    Write-Log "InBoth count: $($inBoth.Count)"

    # Find user objects that are only in Array1
    $Add = @($Array1 | Where-Object { $Array2 -notcontains $_.id })
    Write-Log "Add count: $($Add.Count)"
    Write-Log "Debug: Add array content: $($Add | Format-Table -AutoSize | Out-String)"

    # Create a hash table for quick lookups of Array1 IDs
    $Array1Hash = @{}
    foreach ($user in $Array1) {
        $Array1Hash[$user.id] = $user
    }

    # Find user objects that are only in Array2
    $Remove = @($Array2 | Where-Object { -not $Array1Hash.ContainsKey($_) })
    Write-Log "Remove count: $($Remove.Count)"
    Write-Log "Debug: Remove array content: $($Remove | Format-Table -AutoSize | Out-String)"

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
    Write-Log "Debug: $($result.Add | Format-Table | Out-String)"
    Write-Log "Users to remove: $($result.Remove.Count)"
    Write-Log "Adding users to group '$groupName'..."
    # Add users to the group if they are not already members    
    foreach ($user in $result.Add) {
        if ($groupMemberIds -notcontains $user.id) {
            try {
                Add-DistributionGroupMember -Identity $groupName -Member $user.Id
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

$allUsers = GetAllUsers

Write-Log "Starting OpsQuals Distribution Group Update..."
# CO Wing Pilots
    $groupName = "Pilots"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $pilotCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '44' -and $_.Status -eq 'ACTIVE'} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $pilotCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# CO Wing ES List
    $groupName = "ESList"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $esCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '53' -and $_.Status -eq 'ACTIVE'} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $esCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Mission Check Pilots    
    $groupName = "Mission Check Pilots"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $mcpCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '90' -and $_.Status -eq 'ACTIVE'} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $mcpCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# CO Wing Aircrew List
    $groupName = "Aircrew"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $esCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '55' -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $esCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Mission Pilots
    $groupName = "Mission Pilots"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $mpCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '57' -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $mpCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Instructor Pilots
    $groupName = "Instructor Pilots"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $ipCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '59' -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $ipCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Orientation Pilots
    $groupName = "Orientation Pilots"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $opCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '91' -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $opCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Communicators
    $groupName = "Communicators"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership
    $commsCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '217' -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $commsCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Incident Command
    $groupName = "Mission Base Staff"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName
    # Define the list of AchvIDs to filter
    $ICAchvIDs = @('61', '63', '64', '65', '66', '67', '68', '75', '76', '77', '78', '79', '80')

    # Filter achievements and select unique CAPIDs
    $ICCAPIDs = $achievements_all | Where-Object {
        $_.AchvID -in $ICAchvIDs -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')
    } | Select-Object -ExpandProperty CAPID | Sort-Object -Unique
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $ICCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# sUAS
$groupName = "sUAS"
$groupMemberIds = GetGroupMemberIds -groupName $groupName
# Define the list of AchvIDs to filter
$UASAchvIDs = @('257', '258', '262', '263')

# Filter achievements and select unique CAPIDs
$UASCAPIDs = $achievements_all | Where-Object {
    $_.AchvID -in $UASAchvIDs -and ($_.Status -eq 'ACTIVE' -or $_.Status -eq 'TRAINING')
} | Select-Object -ExpandProperty CAPID | Sort-Object -Unique
$groupUsers = $allUsers | Where-Object {
    $_.officeLocation -in $UASCAPIDs
}
$groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

$result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
ModifyGroupMembers -groupName $groupName -result $result

Write-Log "OpsQuals Distribution Group Update completed."