# Input bindings are passed in via param block.
param($Timer)

#region ImportClasses
<# 
    Using strongly typed classes allows us to both transform data
    (like dates) from [STRING] to their appropriate object types (like [DATETIME])
#>
#. "$PSScriptRoot\classes\Member.ps1"
#endregion

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
# Connect-MgGraph -Scopes "User.Read.All","TeamMember.ReadWrite.All","Team.ReadBasic.All" -DeviceCode
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Import-Module Microsoft.Graph

# Decode the JWT token
$tokenParts = $MSGraphAccessToken -split '\.'
$tokenPayload = $tokenParts[1]
$tokenPayload += '=' * ((4 - $tokenPayload.Length % 4) % 4)
$decodedPayload = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($tokenPayload))

# Validate and convert the JSON payload
try {
    $tokenObject = $decodedPayload | ConvertFrom-Json
    # Check the permissions
    $tokenObject.scp
} catch {
    Write-Warning "Failed to convert JSON: $_"
}
# Convert the JSON payload to a PowerShell object

# Check the permissions
$tokenObject.scp

Get-AzAccessToken -AsSecureString -ResourceUrl "https://graph.microsoft.com/"
Write-Output $MSGraphAccessToken
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
            # Get the user details
            $user = Get-MgUser -UserId $userId

            # Extract and print the displayName
            $displayName = $user.displayName
            Write-Output "User ID: $userId, Display Name: $displayName"
        } catch {
            Write-Output "Failed to retrieve display name for User ID: $userId"
        }
    }
}

# Fetch all users from Microsoft Entra ID (Azure AD)
$allUsers = @()
$uri = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,officeLocation"
do {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $allUsers += $response.value
    $uri = $response.'@odata.nextLink'
} while ($uri)

# Initialize the WINGSTAFF array
$wingStaffUsers = @()

# Create a hash table to store users by CAPID
$userHashTable = @{}
foreach ($user in $allUsers) {
    if ($user.officeLocation) {
        $userHashTable[$user.officeLocation] = $user.id
    }
}

# Process each entry in the DutyPosition file
foreach ($position in $dutyPositions) {
    $CAPID = $position.CAPID
    if ($userHashTable.ContainsKey($CAPID)) {
#        $EmployeeID = $userHashTable[$CAPID]
        if ($position.Lvl -eq "WING") {
#            $DEPARTMENT = "$($position.Lvl) $($position.Duty)"
            # Add the CAPID to the WINGSTAFF array
            $wingStaffUsers += $allUsers | Where-Object { $_.officeLocation -eq $CAPID }
        }
    }
}



# Define the Microsoft Team ID (Wing Staff)
$teamId = "4ea99e49-088a-4996-8dd5-966b6ca408c7"

# Get the current members of the team
$teamMembers = @()
$uri = "https://graph.microsoft.com/v1.0/teams/$teamId/members"
do {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $teamMembers += $response.value
    $uri = $response.'@odata.nextLink'
} while ($uri)


# Create an array to store current team members by userId
$currentTeamMemberIds = @()
foreach ($member in $teamMembers) {
    # $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($member.userId)?$select=officeLocation"
    # if ($user.officeLocation) {
        $currentTeamMemberIds += $member.userId
    # }
}

# Extract the user IDs
$wingUserIds = $wingStaffUsers | Select-Object -ExpandProperty id

$result = Compare-UserIds -Array1 $wingUserIds -Array2 $currentTeamMemberIds
# Display the results
Write-Output "User IDs in both arrays:"
Get-DisplayNames -UserIds $result.InBoth

Write-Output "User IDs that should be added to the Team:"
Get-DisplayNames -UserIds $result.AddtoTeams

Write-Output "User IDs in Teams that should be removed:"
Get-DisplayNames -UserIds $result.RemovefromTeams

# Add all new Staff members to the team
foreach ($userId in $result.AddtoTeams) {
    $user = Get-MgUser -Filter "id eq '$userId'"
    $body = @{
        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
        "roles" = @()
        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$($userId)"
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Body ($body | ConvertTo-Json -Depth 10)
    Write-Output "Added user: $userId to the team"
} 

# Remove any members from the team who are not in WINGSTAFF
foreach ($userId in $result.RemovefromTeams) {
    $membershipId = $teamMembers | Where-Object { $_.userId -eq $userId } | Select-Object -ExpandProperty id
    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members/$membershipId"
    Write-Output "Removed CAPID: $userId from the team"
}
