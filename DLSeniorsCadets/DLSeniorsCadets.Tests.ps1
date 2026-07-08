$ErrorActionPreference = "Stop"

function Assert-Equal {
    param (
        [object]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function Assert-ArrayEqual {
    param (
        [array]$Actual,
        [array]$Expected,
        [string]$Message
    )

    $actualText = @($Actual) -join ","
    $expectedText = @($Expected) -join ","
    if ($actualText -ne $expectedText) {
        throw "$Message Expected '$expectedText', got '$actualText'."
    }
}

$helpersPath = Join-Path $PSScriptRoot "DLSeniorsCadets.Helpers.ps1"
if (Test-Path $helpersPath) {
    . $helpersPath
}

if (-not (Get-Command Get-RegionalDistributionGroups -ErrorAction SilentlyContinue)) {
    throw "Get-RegionalDistributionGroups should be defined."
}

if (-not (Get-Command Get-RegionalDistributionGroupMembers -ErrorAction SilentlyContinue)) {
    throw "Get-RegionalDistributionGroupMembers should be defined."
}

$env:WING_DESIGNATOR = "CO"

$regionalGroups = @(Get-RegionalDistributionGroups)
Assert-Equal $regionalGroups.Count 4 "Four regional distribution groups should be configured."
Assert-Equal $regionalGroups[0].Name "Group 1 - Northern Colorado" "Group 1 display name should identify the region."
Assert-Equal $regionalGroups[0].EmailAddress "group1@cowg.cap.gov" "Group 1 address should match the requested address format."
Assert-ArrayEqual $regionalGroups[0].Units @("072", "191", "099", "136", "147", "022", "068") "Group 1 should include the requested northern Colorado units."
Assert-Equal $regionalGroups[3].Name "Group 4 - Central Colorado" "Group 4 display name should identify the region."
Assert-Equal $regionalGroups[3].EmailAddress "group4@cowg.cap.gov" "Group 4 address should match the requested address format."
Assert-ArrayEqual $regionalGroups[3].Units @("143", "157", "148", "162", "183", "031", "163", "186", "173") "Group 4 should include the requested central Colorado units."

$allUsers = @(
    [PSCustomObject]@{ companyName = "CO-072"; mail = "boulder.member@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-072"; mail = "boulder.member@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-191"; mail = "platte.member@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-999"; mail = "wrong.unit@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-068"; mail = "" },
    [PSCustomObject]@{ companyName = "CO-068"; mail = $null }
)

$group1Members = @(Get-RegionalDistributionGroupMembers -Group $regionalGroups[0] -AllUsers $allUsers)
Assert-ArrayEqual $group1Members @("boulder.member@cowg.cap.gov", "platte.member@cowg.cap.gov") "Regional group members should include only requested units, remove blanks, and deduplicate."

Write-Host "DLSeniorsCadets.Tests.ps1 passed"
