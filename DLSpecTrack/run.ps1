# Input bindings are passed in via param block.
param($Timer)

$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
 . "$PSScriptRoot\..\shared\shared.ps1"
 . "$PSScriptRoot\DLSpecTrack.Helpers.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

# Import the CSV file into an array
$specTracks = Import-Csv "$($CAPWATCHDATADIR)\SpecTrack.txt" -ErrorAction Stop
$dutyPositions = Import-Csv "$($CAPWATCHDATADIR)\DutyPosition.txt" -ErrorAction Stop
$commanders = Import-Csv "$($CAPWATCHDATADIR)\Commanders.txt" -ErrorAction Stop

function GetGroupMemberIds {
    param (
        [string]$groupName
    )

    # First, check if the distribution group exists in Exchange using email lookup
    $mailNickname = $groupName -replace '\s', ''
    if ($groupName -match '(?i)^CO-(\d+) Recruiting$') {
        $mailNickname = Get-RecruitingGroupMailNickname -UnitCode $matches[1]
    }
    $primarySmtpAddress = "$mailNickname@cowg.cap.gov"
    
    $existingExoGroup = $null
    try {
        # Try to get by email first (most reliable)
        $existingExoGroup = Get-DistributionGroup -Identity $primarySmtpAddress -ErrorAction Stop
    } catch {
        # If email doesn't work, try by display name and filter
        try {
            $exoGroups = @(Get-DistributionGroup -Filter "DisplayName -eq '$groupName'" -ErrorAction Stop)
            if ($exoGroups.Count -gt 0) {
                # Prefer the one with the standard email
                $existingExoGroup = $exoGroups | Where-Object { $_.PrimarySmtpAddress -eq $primarySmtpAddress } | Select-Object -First 1
                if (-not $existingExoGroup -and $exoGroups.Count -gt 0) {
                    $existingExoGroup = $exoGroups[0]
                    Write-Log "⚠️ WARNING: Found $($exoGroups.Count) groups with name '$groupName', using: $($existingExoGroup.PrimarySmtpAddress)"
                }
            }
        } catch {
            # Group doesn't exist
        }
    }

    # If group doesn't exist in Exchange, create it
    if (-not $existingExoGroup) {
        try {
            Write-Log "Creating distribution group: $groupName"
            $newGroup = New-DistributionGroup -Name $groupName `
                -DisplayName $groupName `
                -Alias $mailNickname `
                -PrimarySmtpAddress $primarySmtpAddress `
                -Type "Distribution" `
                -ErrorAction Stop
            Write-Log "Distribution group created: $groupName ($primarySmtpAddress)"
            
            # Wait for replication to Graph
            Start-Sleep -Seconds 2
        } catch {
            Write-Log "ERROR creating distribution group '$groupName': $_"
            return $null
        }
    } else {
        Write-Log "Distribution group already exists: $groupName ($($existingExoGroup.PrimarySmtpAddress))"
    }

    if ($existingExoGroup -and $groupName -match '(?i)^CO-(\d+) Recruiting$') {
        $currentPrimarySmtpAddress = "$($existingExoGroup.PrimarySmtpAddress)"
        if (Test-RecruitingGroupAddressNeedsRename -CurrentAddress $currentPrimarySmtpAddress -DesiredAddress $primarySmtpAddress) {
            try {
                Write-Log "Renaming recruiting distribution group address for '$groupName' from $currentPrimarySmtpAddress to $primarySmtpAddress"
                Set-DistributionGroup -Identity $existingExoGroup.Identity `
                    -Alias $mailNickname `
                    -PrimarySmtpAddress $primarySmtpAddress `
                    -ErrorAction Stop
                $existingExoGroup = Get-DistributionGroup -Identity $primarySmtpAddress -ErrorAction Stop
                Write-Log "Recruiting distribution group address updated: $groupName ($primarySmtpAddress)"
            } catch {
                Write-Log "ERROR updating recruiting distribution group address for '$groupName' from $currentPrimarySmtpAddress to $primarySmtpAddress : $_"
            }
        }
    }

    # Now get the group from Microsoft Graph
    $graphGroups = @(Get-MgGroup -Filter "displayName eq '$groupName'")
    
    if ($graphGroups.Count -eq 0) {
        Write-Log "ERROR: Distribution group not found in Microsoft Graph: $groupName"
        return $null
    } elseif ($graphGroups.Count -gt 1) {
        Write-Log "⚠️ WARNING: Multiple groups found in Graph with name '$groupName'. Found $($graphGroups.Count) groups:"
        foreach ($g in $graphGroups) {
            Write-Log "   - ID: $($g.Id), Email: $($g.Mail)"
        }
        # Prefer the group with the standard email format
        $group = $graphGroups | Where-Object { $_.Mail -eq $primarySmtpAddress } | Select-Object -First 1
        if (-not $group) {
            $group = $graphGroups[0]
            Write-Log "   Using first group: $($group.Mail)"
        }
    } else {
        $group = $graphGroups[0]
    }

    if (-not $group) {
        Write-Log "ERROR: Could not get group information for '$groupName' from Graph"
        return $null
    }

    $groupId = $group.Id
    $groupEmail = $group.Mail
    if ($existingExoGroup -and $existingExoGroup.PrimarySmtpAddress) {
        $groupEmail = "$($existingExoGroup.PrimarySmtpAddress)"
    }
    Write-Log "Group '$groupName' found in Graph. Group ID: $groupId, Email: $groupEmail"

    # Get all current members of the group
    $groupMembers = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members?$select=id"
    do {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $groupMembers += $response.value
            $uri = $response.'@odata.nextLink'
        } catch {
            Write-Log "Error fetching group members: $_"
            break
        }
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

    $recruitingCAPIDs = Get-RecruitingCapIds -specTracks $specTracks -dutyPositions $dutyPositions -commanders $commanders -wingDesignator $env:WING_DESIGNATOR

    # Get unique units from allUsers companyName (extract unit numbers)
    $uniqueUnits = $allUsers |
        ForEach-Object { Get-UnitCodeFromCompanyName -companyName $_.companyName } |
        Where-Object { $_ -ne "999" -and $_ -ne "000" -and $_ -ne $null } |
        Select-Object -Unique |
        Sort-Object
    
    Write-Log "Creating recruiting distribution groups for $($uniqueUnits.Count) units..."
    
    foreach ($unit in $uniqueUnits) {
        $groupName = "CO-$unit Recruiting"
        
        Write-Log "Processing recruiting group for unit: $unit (Group: $groupName)"
        
        try {
            # Get recruiting specialty-track members, recruiting duty-position holders, commanders, and EX department members for this unit
            $recruitingMembers = Get-RecruitingMembersForUnit -allUsers $allUsers -unit $unit -recruitingCAPIDs $recruitingCAPIDs -commanders $commanders -wingDesignator $env:WING_DESIGNATOR
            
            Write-Log "Found $($recruitingMembers.Count) commanders/recruiters for unit $unit"

            # GetGroupMemberIds handles both checking if group exists and creating it if needed
            $groupInfo = GetGroupMemberIds -groupName $groupName
            
            if ($null -eq $groupInfo) {
                Write-Log "Error: Could not retrieve group information for '$groupName'. Skipping unit $unit"
                continue
            }

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
        Test-UserCapIdInList -User $_ -CapIds $groupCAPIDs
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
