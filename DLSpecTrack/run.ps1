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
$specTracks = Import-Csv "$($CAPWATCHDATADIR)\SpecTrack.txt" -ErrorAction Stop
# This function compares two arrays and returns the user IDs that are in both, only in the first array, and only in the second array.
function Compare-Arrays {
    param (
        [array]$Array1, # Full user objects from the filtered list
        [array]$Array2  # IDs of current group members
    )

    Write-Log "Inside Compare-Arrays"
    Write-Log "Array1 count: $($Array1.Count)"
    Write-Log "Array2 count: $($Array2.Count)"

    # Ensure Array1 is unique and Array2 is filtered for null or empty values
    $Array1 = $Array1 | Sort-Object -Property id -Unique
    $Array2 = $Array2 | Where-Object { $_ -ne $null -and $_ -ne "" } | ForEach-Object { $_.Trim() }

    # Find user objects that are in both arrays
    $inBoth = $Array1 | Where-Object { $Array2 -contains $_.id.Trim().ToLower() }
    Write-Log "InBoth count: $($inBoth.Count)"

    # Find user objects that are only in Array1
    $Add = @($Array1 | Where-Object { $Array2 -notcontains $_.id.Trim().ToLower() })
    Write-Log "Add count: $($Add.Count)"

    # Create a hash table for quick lookups of Array1 IDs
    $Array1Hash = @{}
    foreach ($user in $Array1) {
        $Array1Hash[$user.id.Trim().ToLower()] = $user
    }

    # Find user objects that are only in Array2
    if ($Array1Hash) {
        $Remove = @($Array2 | Where-Object { -not $Array1Hash.ContainsKey($_) })
        Write-Log "Remove count: $($Remove.Count)"
    } else {
        $Remove = @()
        Write-Log "No users to remove, Array1Hash is empty."
    }

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
    if ($group) {
        Write-Host "Distribution group '$groupName' found. Group ID: $($group.Guid), $($group.Mail)"
    } else {
        Write-Host "Distribution group '$groupName' does not exist. Creating it..."
        # Sanitize the groupName to create a valid alias (mailNickname)
        $mailNickname = $groupName -replace '\s', ''

        # Create the distribution group
        $newGroup = New-DistributionGroup -Name $groupName `
            -DisplayName $groupName `
            -Alias $mailNickname `
            -PrimarySmtpAddress "$mailNickname@cowg.cap.gov" `
            -Type "Distribution"

        Write-Host "Distribution group '$groupName' created successfully. Group Alias: $($newGroup.Alias)"
        $group = $newGroup
    }

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
#    Write-Log "Debug: $($result.Add | Format-Table | Out-String)"
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

Write-Log "Starting Specialty Track Distribution Group Update..."

$allUsers = GetAllUsers
$allTracks = $specTracks | Select-Object -ExpandProperty Track | Sort-Object -Unique
Write-Host $allTracks | Format-Table -AutoSize

foreach ($track in $allTracks) {
    # Format the trackName to capitalize the first letter of each word (retain spaces)
    $track = ($track -replace '\b(\w)', { $_.Value.ToUpper() }) -replace '\B(\w)', { $_.Value.ToLower() }
    Write-Log "Processing track: $track"
    $groupMemberIds = GetGroupMemberIds -groupName $track
    # Filter users for group membership
    $groupCAPIDs = $specTracks | Where-Object { $_.Track -eq $track } | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $groupCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

        $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupMemberIds
        ModifyGroupMembers -groupName $track -result $result
    }

Write-Log "Specialty Track Distribution Group Update completed."
