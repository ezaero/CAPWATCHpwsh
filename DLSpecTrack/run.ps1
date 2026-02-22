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
$specTracks = Import-Csv "$($CAPWATCHDATADIR)\SpecTrack.txt" -ErrorAction Stop
# This function compares two arrays and returns the user IDs that are in both, only in the first array, and only in the second array.
function Compare-Arrays {
    param (
        [array]$Array1, # Full user objects from the filtered list
        [array]$Array2  # IDs of current group members
    )


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

    $groups = @(Get-MgGroup -Filter "displayName eq '$groupName'")

    if ($groups.Count -eq 0) {
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
    } elseif ($groups.Count -gt 1) {
        Write-Log "⚠️ WARNING: Multiple groups found with name '$groupName'. Found $($groups.Count) groups:"
        foreach ($g in $groups) {
            Write-Log "   - ID: $($g.Id), Email: $($g.Mail)"
        }
        # Prefer the group with the standard email format (without numbers in the alias)
        $mailNickname = $groupName -replace '\s', ''
        $preferredEmail = "$mailNickname@cowg.cap.gov"
        $group = $groups | Where-Object { $_.Mail -eq $preferredEmail } | Select-Object -First 1
        if (-not $group) {
            # If preferred not found, just take the first one
            $group = $groups[0]
            Write-Log "   Using first group: $($group.Mail)"
        } else {
            Write-Log "   Using preferred group: $($group.Mail)"
        }
    } else {
        $group = $groups[0]
        Write-Host "Distribution group '$groupName' found. Group ID: $($group.Id), $($group.Mail)"
    }

    $groupId = $group.Id
    $groupEmail = $group.Mail
    Write-Log "Group '$groupName' found. Group ID: $groupId, Email: $groupEmail"

    # Get all current members of the group
    $groupMembers = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members?$select=id"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $groupMembers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    # Return group info including email and member IDs
    $groupMemberIds = $groupMembers | ForEach-Object { $_.id } | Where-Object { $_ -ne $null -and $_ -ne "" }
    return @{
        GroupEmail = $groupEmail
        MemberIds = $groupMemberIds
    }
}

function ModifyGroupMembers {
    param (
        [string]$groupName,
        [string]$groupEmail,
        [array]$groupMemberIds,
        [PSCustomObject]$result
    )
    Write-Log "Users in both arrays: $($result.InBoth.Count)"
    Write-Log "Users to add: $($result.Add.Count)"
#    Write-Log "Debug: $($result.Add | Format-Table | Out-String)"
    Write-Log "Users to remove: $($result.Remove.Count)"
    Write-Log "Adding users to group '$groupName' ($groupEmail)..."
    # Add users to the group if they are not already members
    foreach ($user in $result.Add) {
        if ($groupMemberIds -notcontains $user.id) {
            try {
                Add-DistributionGroupMember -Identity $groupEmail -Member $user.Id
                Write-Log "Added user: $($user.displayName) ($($user.mail)) to group '$groupName'."
            } catch {
                Write-Log "Failed to add user: $($user.displayName) ($($user.mail)) to group '$groupName'. Error: $_"
            }
        } else {
            Write-Log "User: $($user.displayName) ($($user.mail)) is already a member of group '$groupName'."
        }
    }
    # Remove users from the group if they no longer match the criteria
    Write-Log "Removing users from group '$groupName' ($groupEmail)..."
    foreach ($userId in $result.Remove) {
        try {
            Remove-DistributionGroupMember -Identity $groupEmail -Member $userId -Confirm:$false
            Write-Log "Removed user: $userId from group '$groupName'."
        } catch {
            Write-Log "Failed to remove user: $userId from group '$groupName'. Error: $_"
        }
    }

}

function CreateRecruitingDistributionGroups {
    param (
        [array]$allUsers
    )

    # Get unique units from allUsers companyName (extract unit numbers)
    $uniqueUnits = $allUsers |
        Where-Object { $_.companyName -ne $null -and $_.companyName -match '\d+' } |
        ForEach-Object {
            if ($_.companyName -match '(\d+)') { $matches[1] }
        } |
        Where-Object { $_ -ne "999" -and $_ -ne "000" -and $_ -ne $null } |
        Select-Object -Unique |
        Sort-Object
    
    Write-Log "Creating recruiting distribution groups for $($uniqueUnits.Count) units..."
    
    foreach ($unit in $uniqueUnits) {
        $groupName = "CO-$unit Recruiting"
        $mailNickname = "co-$($unit)-recruiting"
        $primarySmtpAddress = "$mailNickname@cowg.cap.gov"
        
        Write-Log "Processing recruiting group for unit: $unit (Group: $groupName, Email: $primarySmtpAddress)"
        
        try {
            # Check if group exists by trying the email address first
            $existingGroup = $null
            try {
                $existingGroup = Get-DistributionGroup -Identity $primarySmtpAddress -ErrorAction Stop
            } catch {
                # If that fails, try to find by display name and filter by email
                try {
                    $allGroupsWithName = Get-DistributionGroup -Filter "DisplayName -eq '$groupName'" -ErrorAction Stop
                    if ($allGroupsWithName) {
                        $existingGroup = $allGroupsWithName | Where-Object { $_.PrimarySmtpAddress -eq $primarySmtpAddress } | Select-Object -First 1
                        if ($allGroupsWithName.Count -gt 1) {
                            Write-Log "⚠️ Found $($allGroupsWithName.Count) groups with name '$groupName', using the one with email $primarySmtpAddress"
                        }
                    }
                } catch {
                    # Group doesn't exist
                }
            }

            if (-not $existingGroup) {
                Write-Log "Creating distribution group: $groupName"
                $newGroup = New-DistributionGroup -Name $groupName `
                    -DisplayName $groupName `
                    -Alias $mailNickname `
                    -PrimarySmtpAddress $primarySmtpAddress `
                    -Type "Distribution"
                Write-Log "Distribution group created: $groupName ($primarySmtpAddress)"
                $groupMail = $primarySmtpAddress
            } else {
                Write-Log "Distribution group already exists: $groupName ($primarySmtpAddress)"
                $groupMail = $existingGroup.PrimarySmtpAddress
            }
            
            # Get recruiting specialty track members and commanders (EX department) for this unit
            $recruitingCAPIDs = $specTracks | Where-Object { $_.Track -eq 'Recruiting' } | Select-Object -ExpandProperty CAPID
            $recruitingMembers = $allUsers | Where-Object {
                $_.companyName -match $unit -and 
                ($_.officeLocation -in $recruitingCAPIDs -or $_.department -eq 'EX') -and 
                $_.mail -ne $null
            } | Sort-Object -Property id -Unique
            
            Write-Log "Found $($recruitingMembers.Count) commanders/recruiters for unit $unit"

            # Get current group members
            $groupInfo = GetGroupMemberIds -groupName $groupName

            # Compare and update group membership
            $result = Compare-Arrays -Array1 $recruitingMembers -Array2 $groupInfo.MemberIds
            ModifyGroupMembers -groupName $groupName -groupEmail $groupInfo.GroupEmail -groupMemberIds $groupInfo.MemberIds -result $result
            
        } catch {
            Write-Log "Error processing recruiting group for unit $unit : $_"
        }
    }
    
    Write-Log "Recruiting distribution groups creation completed."
}

Write-Log "Starting Specialty Track Distribution Group Update..."

$allUsers = GetAllUsers
$allTracks = $specTracks | Select-Object -ExpandProperty Track | Sort-Object -Unique
Write-Host $allTracks | Format-Table -AutoSize

foreach ($track in $allTracks) {
    # Format the trackName to capitalize the first letter of each word (retain spaces)
    $track = ($track -replace '\b(\w)', { $_.Value.ToUpper() }) -replace '\B(\w)', { $_.Value.ToLower() }
    Write-Log "Processing track: $track"
    $groupInfo = GetGroupMemberIds -groupName $track
    # Filter users for group membership
    $groupCAPIDs = $specTracks | Where-Object { $_.Track -eq $track } | Select-Object -ExpandProperty CAPID
    $groupUsers = $allUsers | Where-Object {
        $_.officeLocation -in $groupCAPIDs
    }
    $groupUsers = $groupUsers | Where-Object { $_.mail -ne $null }

        $result = Compare-Arrays -Array1 $groupUsers -Array2 $groupInfo.MemberIds
        ModifyGroupMembers -groupName $track -groupEmail $groupInfo.GroupEmail -groupMemberIds $groupInfo.MemberIds -result $result
    }

# Create recruiting distribution groups for each squadron
try {
    CreateRecruitingDistributionGroups -allUsers $allUsers
} catch {
    Write-Log "Error creating recruiting distribution groups: $_"
}

Write-Log "Specialty Track Distribution Group Update completed."
