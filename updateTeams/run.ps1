# Input bindings are passed in via param block.
param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

#Abort script execution if CAPWATCH data is stale
if (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours -gt 40) {
    Write-Error "CAPWATCH data in [$CAPWATCHDATADIR] is stale; aborting script execution!"
    exit 1
}

# Define the path to the DutyPosition.txt file
# Read the DutyPosition.txt file and reduce it to only WING Staff
$dutyPositions_all = Import-Csv .\DutyPosition.txt -ErrorAction Stop
$dutyPositions = $dutyPositions_all | Where-Object { $_.Lvl -eq "WING" }
$dutyPositions = $dutyPositions | Sort-Object CAPID -Unique

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome

# This function takes 2 arrays and compares them, return 3 Arrays
function Compare-UserIds {
    param (
        [array]$Array1,
        [array]$Array2
    )

    # Find user IDs that are in both arrays
    $inBoth = $Array1 | Where-Object { $Array2 -contains $_ }

    # Find user IDs that are only in Array1
    $AddtoTeams = $Array1 | Where-Object { $Array2 -notcontains $_ }

    # Find user IDs that are only in Array2
    $RemovefromTeams = $Array2 | Where-Object { $Array1 -notcontains $_ }

    # Output the results
    [PSCustomObject]@{
        InBoth       = $inBoth
        AddtoTeams = $AddtoTeams
        RemovefromTeams = $RemovefromTeams
    }
}
  
# Takes in an Array of userId's and prints the display names of each of them
function Get-DisplayNames {
    param (
        [array]$UserIds
    )

    foreach ($userId in $UserIds) {
        try {
            # Define the API endpoint
            $uri = "https://graph.microsoft.com/v1.0/users/$userId"
    
            # Make the API request to get the user details
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ContentType "application/json"
    
            # Extract and print the displayName
            $displayName = $response.displayName
            Write-Output "User ID: $userId, Display Name: $displayName"
        } catch {
            Write-Error "Failed to retrieve display name for User ID: $userId. Error: $_"
        }
    }
}

function GetAllUsers {
    # Fetch all users from Microsoft Entra ID (Azure AD)
    $allUsers = @()
    $uri = "https://graph.microsoft.com/beta/users?$select=id,displayName,officeLocation,companyName"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allUsers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    $allUsers
}

function GetAllGroups {
    # Get a list of groups that contain teams
    $allGroups = @()
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allGroups += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    $allGroups
}

function GetUnits {
    # Create a list of all Unit charter numbers and names in the Wing
    $organization_all = Import-Csv -Path $OrganizationFile
    $co_org = $organization_all | Where-Object { $_.Wing -eq "CO" } | Sort-Object Unit -Unique
    $co_org = $co_org | Select-Object Unit, Name
    # unitList will be a list of all the Teams required
    $unitList = @()
    foreach ($unit in $co_org) {
        if ($unit.Unit -ne "000" -and $unit.Unit -ne "999" -and $unit.Unit -ne "001") {
            $unitList += "CO-$($unit.Unit) $($unit.Name)"
        }
    }
    $unitList
}

function GetCommander {
    param (
        [Array]$allUsers,
        [string]$unit
    )

    $cn = $allUsers.Count
    Write-Host "Users are: $cn"

    $commanders_all = @()
    $commander = @()
    $commander_user = @()
    $commanders_all = Import-Csv -Path $CommandersFile
    $commander = $commanders_all | Where-Object { $_.Unit -eq $unit -and $_.Wing -eq 'CO' } 
    $CAPID = $commander.CAPID
    Write-Host "Commander CAPID is: $CAPID"
    $commander_user = $allUsers | Where-Object { $_.officeLocation -eq '138687' }
    $cu = $commander_user
    Write-Host "Commanderuser is: $cu"

    $commander_user
}

# Define the function to check if a team exists
function CheckTeamExists {
    param (
        [string]$teamName
    )

    # Define the endpoint to get the team by display name
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$teamName' and resourceProvisioningOptions/Any(x:x eq 'Team')"

    # Make the API request
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

    if ($response.value.Count -gt 0) {
        return $true
    } else {
        return $false
    }
}

function CheckTeams {
    param (
        [Array]$unitList,
        [Array]$allGroups,
        [Array]$allUsers
    )

    foreach ($unitName in $unitList) {
        $teamExists = $false
        $teamExists = CheckTeamExists -teamName $unitName
        if ($teamExists) {
            Write-Output "$unitName already exists"
        } else {
            Write-Output "$unitName needs to be created"
            $unitNumber = $unitName.Substring(3, 3)  # gets first 6 characters of string (co-XXX)
            Write-Host $unitNumber        
            $unitCommander = GetCommander -allUsers $allUsers -unit $unitNumber # gets first instance of commander
            Write-Host $unitCommander.displayName
            $mailName = "CO" + $unitNumber # makes mail for Team 'coxxx@cowg.cap.gov'
            Write-Host "mailname: $mailName"
            $commanderId = $unitCommander.mail # commander's email to make owner of team
            Write-Output $commanderId
            $uri = "https://graph.microsoft.com/v1.0/users('$commanderId')"
            if ($commanderId) {
                try {
                    # Create the team using Microsoft Graph API
                    $teamPayload = @{
                        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                        displayName = $unitName
                        mailNickname = $mailName
                        description = "Team for $unitName"
                        visibility = "Private"
                        members = @(
                            @{
                                "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                                roles = @(
                                "owner","member"
                            )
                            "user@odata.bind" = $uri
                        }
                    )                    
                }
                $response = New-MgTeam -BodyParameter $teamPayload
                    # Check if the response is not null
                if ($null -ne $response) {
                    Write-Output "Team created successfully. Team ID: $($response.id)"
                } else {
                    Write-Output "Team creation failed. No response received."
                }
                    # Verify the team creation by querying the team
                } catch {
                    Write-Error "An error occurred: $_"
                    Write-Error "Error Details: $($_.Exception.Message)"                
                }
            } else {
                Write-Output "No Commander found for $unitNumber"
            }
        }
    }
}
function PopulateTeams {
    param (
        [array]$allUsers,
        [array]$unitList 
    )
    
    foreach ($unitName in $unitList) {
        Write-Output "unitName is: $unitName"
        # Define the endpoint to get the team by display name
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$unitName') and resourceProvisioningOptions/Any(x:x eq 'Team')"
        $team = Invoke-MgGraphRequest -Method GET -Uri $uri
        $teamId = $team.value[0].id
        $uri = "https://graph.microsoft.com/v1.0/groups/$teamId"

        if (-not [string]::IsNullOrEmpty($teamId)) {

            Write-Output "Team Found: $teamId"

            $unitNumber = $unitName.Substring(0, 6) # get the "CO-XXX"
            $unitMembers = @()
            $unitMembers = $allUsers | Where-Object { $_.companyName -eq $unitNumber } # gets everyone in EntraID in the unit
            # Write-Output $unitMembers
            $unitMemberIds = $unitMembers | Select-Object -ExpandProperty ID    
            # Get all members on the Team
            # Define the API endpoint
            $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members"
            $teamMembers = @()
            do {
                $response = Invoke-MgGraphRequest -Method GET -Uri $apiEndpoint
                $teamMembers += $response.value
                $apiEndpoint = $response.'@odata.nextLink'
            } while ($apiEndpoint)
            
            # Check if the response is not null
            if ($null -ne $response) {
                # Extract user IDs from the response
                $teamUserIds = @()
                $teamUserIds = $teamMembers | ForEach-Object { $_.userId }
            }

            # Compare $groupMemberIds and $unitMemberIds
            Write-Host "Counts: $($unitMemberIds.Count) $($teamUserIds.Count)"
            $result = Compare-UserIds -Array1 $unitMemberIds -Array2 $teamUserIDs
            #  Display the results
#            Write-Output "User IDs in both arrays:"
#            Get-DisplayNames -UserIds $result.InBoth

            Write-Output "User IDs that should be added to the Group:"
            Get-DisplayNames -UserIds $result.AddtoTeams

            Write-Output "User IDs in Group that should be removed:"
            Get-DisplayNames -UserIds $result.RemovefromTeams

            # Add each user in AddtoTeamsto the team
            foreach ($userId in $result.AddtoTeams) {
                $uri = "https://graph.microsoft.com/v1.0/users('$userId')"
                Write-Output "URI is: $uri"
                New-MgGroupMember -GroupId $teamId -DirectoryObjectId $userId
                Write-Output "User $userId added successfully."
            }

            # Delete users in RemovefromTeams
            foreach ($userId in $result.RemovefromTeams) {
                try {
                    # Get the membership ID for the user in the team
                    $membership = Get-MgTeamMember -TeamId $teamId | Where-Object { $_.userId -eq $userId }
                
                    if ($null -ne $membership) {
                        # Remove the member from the team
                        Remove-MgTeamMember -TeamId $teamId -MembershipId $membership.id
                        Write-Output "User $userId removed successfully."
                    } else {
                        Write-Output "User $userId is not a member of the team."
                    }
                } catch {
                    Write-Error "Failed to remove user $userId"
                }
            }
        }
    }
}

function UpdateAnnouncements {
    param (
        [array]$unitList 
    )
    foreach ($unitName in $unitList) {
        Write-Output "unitName is: $unitName"
        # Define the endpoint to get the team by display name
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$unitName') and resourceProvisioningOptions/Any(x:x eq 'Team')"
        $team = Invoke-MgGraphRequest -Method GET -Uri $uri
        $teamValue = $team.value[0]
        $teamMail = $teamValue.mail
        $uri = "https://graph.microsoft.com/v1.0/groups/$teamMail"

        $distributionList = "announcements@cowg.cap.gov"
        Add-DistributionGroupMember -Identity $distributionList -Member $teamMail
        Write-Output "Team: $teamMail added to announcements"
    }
}

# Main Program starts here...
# ---------------------------

Write-Output "We are in main"
$unitList = GetUnits  # gets all units
$allGroups = GetAllGroups
$allUsers = GetAllUsers
CheckTeams -unitList $unitList -allGroups $allGroups -allUsers $allUsers # gets all teams and creates teams if needed
PopulateTeams -allUsers $allUsers -unitList $unitList
# UpdateAnnouncements -unitList $unitList