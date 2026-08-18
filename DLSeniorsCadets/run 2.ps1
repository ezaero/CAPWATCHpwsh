# Input bindings are passed in via param block.
param($Timer)

# Include shared Functions
 . "$PSScriptRoot\..\shared\shared.ps1"
 . "$PSScriptRoot\DLSeniorsCadets.Helpers.ps1"

 # Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

$OrganizationFile = "$($CAPWATCHDATADIR)/Organization.txt"
# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

function GetUnits {
    # Create a list of all Unit charter numbers and names in the Wing
    $organization_all = Import-Csv -Path $OrganizationFile
    $wing_org = $organization_all | Where-Object { $_.Wing -eq $env:WING_DESIGNATOR } | Sort-Object Unit -Unique
    $wing_org = $wing_org | Select-Object Unit, Name
    # unitList will be a list of all the Distribution Groups required
    $unitList = @()
    foreach ($unit in $wing_org) {
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

    foreach ($unit in $unitList) {
        $unitDesginator = "$($env:WING_DESIGNATOR)-$($unit.Unit)"
        $groupMembers = @()  # Initialize as empty array
        
        if ($memberType -eq "ALL") {
            $groupName = "$($env:WING_DESIGNATOR)-$($unit.Unit) $($unit.Name)"
#            $SMTPAddress = "$($env:WING_DESIGNATOR)-$($unit.Unit)@$($env:WING_DESIGNATOR.ToLower())wg.cap.gov"
            $groupMembers = @($allUsers | Where-Object { $_.companyName -eq $unitDesginator } | Select-Object -ExpandProperty mail)
        } else {
            $memberName = ($memberType.Substring(0,1).ToUpper()) + ($memberType.Substring(1).ToLower()) + 's'
            $groupName = "$($env:WING_DESIGNATOR)-$($unit.Unit) $memberName"
            $SMTPAddress = "$($env:WING_DESIGNATOR)-$($unit.Unit)-$memberName@$($env:WING_DESIGNATOR.ToLower())wg.cap.gov"
            $groupMembers += @($allUsers | Where-Object { $_.companyName -eq $unitDesginator -and $_.employeeType -eq $memberType } | Select-Object -ExpandProperty mail)
            if ($memberType -eq "CADET") {
                $groupMembers += @($allUsers | Where-Object { $_.companyName -eq $unitDesginator -and $_.employeeType -eq "PARENT" } | Select-Object -ExpandProperty mail)
                $groupMembers += @($allUsers | Where-Object { $_.companyName -eq $unitDesginator -and ($_.department -like "*EX*" -or $_.department -like "*CP*") } | Select-Object -ExpandProperty mail)
            }
        }
        # Remove nulls, empty strings, and duplicates
        $groupMembers = $groupMembers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
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

function WingGroups {
    param (
        [string]$memberType,
        [array]$allUsers 
    )

    # Wing-level distribution groups
    $groupName = "$($env:WING_DESIGNATOR) Wing $memberType`s"
    Write-Log "Processing wing-level distribution group: '$groupName'"
    $groupMembers = @()  # Initialize as empty array
    
    # Get all users of the specified type across the entire wing
    if ($memberType -eq "CADET") {
        Write-Log "Building cadet group members..."
        $groupMembers += @($allUsers | Where-Object { $_.employeeType -eq $memberType } | Select-Object -ExpandProperty mail)
        Write-Log "  Cadets: $($groupMembers.Count)"
        
        # For wing-level cadet group, also include parents
        $parentMembers = @($allUsers | Where-Object { $_.employeeType -eq "PARENT" } | Select-Object -ExpandProperty mail)
        $groupMembers += $parentMembers
        Write-Log "  Parents added: $($parentMembers.Count)"
        
        # Include all staff assigned to cadet programs (CP in Department)
        $cpMembers = @($allUsers | Where-Object { $_.department -like "*CP*" } | Select-Object -ExpandProperty mail)
        $groupMembers += $cpMembers
        Write-Log "  Cadet Program staff (CP) added: $($cpMembers.Count)"
        
        # Include all staff with executive (EX) duties
        $exMembers = @($allUsers | Where-Object { $_.department -like "*EX*" } | Select-Object -ExpandProperty mail)
        $groupMembers += $exMembers
        Write-Log "  Executive staff (EX) added: $($exMembers.Count)"
    } elseif ($memberType -eq "SENIOR") {
        Write-Log "Building senior group members..."
        $groupMembers += @($allUsers | Where-Object { $_.employeeType -eq $memberType } | Select-Object -ExpandProperty mail)
        Write-Log "  Seniors: $($groupMembers.Count)"
    }
    
    # Remove duplicates and nulls
    $groupMembers = $groupMembers | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object -Unique
    Write-Log "Total unique members after deduplication: $($groupMembers.Count)"
    
    try {
        Update-DistributionGroupMember -Identity $groupName -Members $groupMembers -Confirm:$false
        Write-Log "Successfully updated distribution group '$groupName' with $($groupMembers.count) members."
    } catch {
        Write-Log "Failed to update wing-level distribution group '$groupName'. Error: $_"
    }
}

function EnsureRegionalDistributionGroup {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Group
    )

    try {
        $null = Get-DistributionGroup -Identity $Group.EmailAddress -ErrorAction Stop
        Write-Log "Distribution group already exists: $($Group.Name) ($($Group.EmailAddress))"
    } catch {
        try {
            $null = New-DistributionGroup -Name $Group.Name `
                -DisplayName $Group.Name `
                -Alias $Group.Alias `
                -PrimarySmtpAddress $Group.EmailAddress `
                -Type Distribution `
                -ErrorAction Stop
            Write-Log "Distribution group created: $($Group.Name) ($($Group.EmailAddress))"
        } catch {
            Write-Log "Failed to create regional distribution group '$($Group.Name)' ($($Group.EmailAddress)). Error: $_"
            throw
        }
    }
}

function RegionalGroups {
    param (
        [array]$allUsers
    )

    foreach ($group in Get-RegionalDistributionGroups) {
        Write-Log "Processing regional distribution group: '$($group.Name)'"
        EnsureRegionalDistributionGroup -Group $group

        $groupMembers = @(Get-RegionalDistributionGroupMembers -Group $group -AllUsers $allUsers)
        try {
            Update-DistributionGroupMember -Identity $group.EmailAddress -Members $groupMembers -Confirm:$false
            Write-Log "Successfully updated regional distribution group '$($group.Name)' ($($group.EmailAddress)) with $($groupMembers.Count) members."
        } catch {
            Write-Log "Failed to update regional distribution group '$($group.Name)' ($($group.EmailAddress)). Error: $_"
        }
    }
}

Write-Log "Squadron Seniors/Cadets script started. ------------------------------------------------"

$unitList = GetUnits
$allUsers = GetAllUsers
SquadronGroups -memberType "SENIOR" -unitList $unitList -allUsers $allUsers
SquadronGroups -memberType "CADET" -unitList $unitList -allUsers $allUsers
SquadronGroups -memberType "ALL" -unitList $unitList -allUsers $allUsers

# Update wing-level distribution groups
WingGroups -memberType "CADET" -allUsers $allUsers
WingGroups -memberType "SENIOR" -allUsers $allUsers
RegionalGroups -allUsers $allUsers

Write-Log "Squadron Seniors/Cadets script ended. ------------------------------------------------"
