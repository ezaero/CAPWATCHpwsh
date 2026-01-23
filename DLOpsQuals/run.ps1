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
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

# Import the CSV file into an array
$achievements_all = Import-Csv "$($CAPWATCHDATADIR)\MbrAchievements.txt" -ErrorAction Stop
$mbrTasks_all = Import-Csv "$($CAPWATCHDATADIR)\MbrTasks.txt" -ErrorAction Stop
$dutyPosition_all = Import-Csv "$($CAPWATCHDATADIR)\DutyPosition.txt" -ErrorAction Stop

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
    
    # Create a hash table for quick lookups of Array1 IDs
    $Array1Hash = @{}
    foreach ($user in $Array1) {
        $Array1Hash[$user.id] = $user
    }

    # Find user objects that are only in Array2 (filter out nulls before checking the hash)
    $Remove = @($Array2 | Where-Object { $_ -ne $null -and $_ -ne "" } | Where-Object { -not $Array1Hash.ContainsKey($_) })
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

    try {
        $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction Stop
    } catch {
        Write-Log "Graph query for group '$groupName' failed: $_"
        return @()
    }

    if (-not $group) {
        Write-Log "Graph group '$groupName' not found. Returning empty member list."
        return @()
    }

    $groupId = $group.Id
    Write-Log "Group '$groupName' found. Group ID: $groupId"

    # Get all current members of the group
    $groupMembers = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members?$select=id"
    try {
        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            $groupMembers += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)
    } catch {
        Write-Log "Failed to retrieve members for Graph group '$groupName' (ID: $groupId): $_"
        return @()
    }

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
    # Verify the Exchange distribution group exists before attempting modifications
    try {
        $dg = Get-DistributionGroup -Identity $groupName -ErrorAction Stop
    } catch {
        Write-Log "Exchange distribution group '$groupName' not found. Skipping modifications. Error: $_"
        return
    }
    # Add users to the group if they are not already members    
    foreach ($user in $result.Add) {
        if ($groupMemberIds -notcontains $user.id) {
            try {
                Add-DistributionGroupMember -Identity $groupName -Member $user.Id -ErrorAction Stop
                Write-Log "Added user: $($user.displayName) ($($user.mail)) to group '$groupName'."
            } catch {
                Write-Log "Failed to add user: $($user.displayName) ($($user.mail)) to group '$groupName'. Error: $_"
            }
        } else {
            Write-Log "User: $($user.displayName) ($($user.mail)) is already a member of group '$groupName'."
        }
    }
    # Remove users from the group if they are in the Remove array
    foreach ($userId in $result.Remove) {
        try {
            # Find the user object by userId to get the displayName
            $userObj = $allUsers | Where-Object { $_.id -eq $userId }
            $userName = if ($userObj) { $userObj.displayName } else { $userId }
#            Remove-DistributionGroupMember -Identity $groupName -Member $userId -Confirm:$false
            Write-Log "Removed user: $userName (ID: $userId) from group '$groupName'."
        } catch {
            Write-Log "Failed to remove user with ID: $userId from group '$groupName'. Error: $_"
        }
    }
}

$allUsers = GetAllUsers

Write-Log "Starting OpsQuals Distribution Group Update..."
# Wing Pilots
    $groupName = "Pilots"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter users for group membership - any user with an active Flight Review in OpsQuals
    $pilotCAPIDs = $mbrTasks_all | Where-Object { $_.TaskID -eq '69' -and $_.Status -eq 'ACTIVE'} | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $pilotCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# KAPA Pilot List
    $groupName = "KAPA Pilot List"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName
    # List of unit numbers
    $unitNumbers = @('162', '148', '157', '163', '164', '183', '186', '143', '031')
    # Build the WING-XXX format list
    $unitCompanyNames = $unitNumbers | ForEach-Object { "$($env:WING_DESIGNATOR)-$_" }

    # Filter users for group membership
    $kapaCAPIDs = $mbrTasks_all | Where-Object { $_.TaskID -eq '69' -and $_.Status -eq 'ACTIVE'} | Select-Object -ExpandProperty CAPID
    # Manual CAPIDs to include in KAPA Pilot List (persisted to KAPA_manual_capids.txt)
    $manualKapaCapIds = @('446885', '344160')

    # Merge manual CAPIDs into computed list
    $kapaCAPIDs = ($kapaCAPIDs + $manualKapaCapIds) | Sort-Object -Unique
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $kapaCAPIDs -and $unitCompanyNames -contains $_.companyName
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

# Wing ES List
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

# Group 4 ES List (subset of Wing ESList)
    $groupName = "Group 4 Emergency Services Officers"
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Squadron numbers that define Group 4
    $group4Units = @('157','162','163','173','143','148','186','183','164','031')
    # Build the WING-XXX format list for companyName comparison
    $group4CompanyNames = $group4Units | ForEach-Object { "$($env:WING_DESIGNATOR)-$_" }

    # Start from the same ES CAPIDs used for Wing ES List
    $esCAPIDs = $achievements_all | Where-Object { $_.AchvID -eq '53' -and $_.Status -eq 'ACTIVE'} | Select-Object -ExpandProperty CAPID

    # Filter users who are both on ES list and belong to one of the Group 4 squadrons
    $groupUsers = $allUsers | Where-Object {
        ($_.officeLocation -in $esCAPIDs) -and ($group4CompanyNames -contains $_.companyName)
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

# Wing Aircrew List
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

# suas
$groupName = "suas"
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

# --- Commander sync: discover commanders from DutyPosition and update 'Commanders' group ---
$SyncCommandersGroup = $true

function Get-CommanderCAPIDsFromDutyPositions {
    # Use the already-imported $dutyPosition_all
    $capids = $dutyPosition_all | Where-Object { $_.Duty -match '(?i)Commander' } | Select-Object -ExpandProperty CAPID
        # Match Commander or Chief of Staff (case-insensitive)
        $capids = $dutyPosition_all | Where-Object { $_.Lvl -eq 'WING' -and ($_.Duty -match '(?i)(Commander|Chief of Staff)') } | Select-Object -ExpandProperty CAPID
    $capids = $capids | Where-Object { $_ -ne $null -and $_ -ne '' } | Sort-Object -Unique
    return $capids
}

if ($SyncCommandersGroup) {
    Write-Log "Starting Commanders group sync..."
    $cmdCapids = Get-CommanderCAPIDsFromDutyPositions
    Write-Log "Discovered commander CAPIDs: $($cmdCapids -join ',')"

    # Resolve CAPIDs to Entra user objects (by officeLocation)
    $cmdUsers = @()
    foreach ($capid in $cmdCapids) {
        $match = $allUsers | Where-Object { $_.officeLocation -eq $capid }
        if ($match) {
            $cmdUsers += $match
        } else {
            Write-Log "No Entra user found with officeLocation = $capid"
        }
    }
    $cmdUsers = $cmdUsers | Sort-Object -Property id -Unique

    # Prepare group update using existing helpers
    $groupName = 'COWG Commanders'
    $groupMemberIds = GetGroupMemberIds -groupName $groupName

    # Filter only users with mail
    $groupUsers = $cmdUsers | Where-Object { $_.mail -ne $null }

    $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
    ModifyGroupMembers -groupName $groupName -result $result

    Write-Log "Commanders group sync complete. Added: $($result.Add.Count); Removed: $($result.Remove.Count)"
}