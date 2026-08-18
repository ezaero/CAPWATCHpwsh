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

function Assert-StringContains {
    param (
        [string]$Actual,
        [string]$Expected,
        [string]$Message
    )

    if ($Actual -notlike "*$Expected*") {
        throw "$Message Expected '$Actual' to contain '$Expected'."
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

if (Get-Command Get-RegionalDistributionGroupUpdateIdentity -ErrorAction SilentlyContinue) {
    throw "Get-RegionalDistributionGroupUpdateIdentity should not be defined; regional groups should be classic distribution groups addressed by SMTP."
}

if (Get-Command Get-RegionalDistributionGroupMemberSyncMethod -ErrorAction SilentlyContinue) {
    throw "Get-RegionalDistributionGroupMemberSyncMethod should not be defined; GroupMailbox recipients should not be adapted for regional distribution groups."
}

$env:WING_DESIGNATOR = "CO"

$organizationRows = @(
    [PSCustomObject]@{ ORGID = "1056"; Wing = "CO"; Unit = "001"; NextLevel = "226"; Name = "COLORADO WING HQ"; Type = "WING"; Scope = "WING"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1292"; Wing = "CO"; Unit = "167"; NextLevel = "1056"; Name = "GROUP 1 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1667"; Wing = "CO"; Unit = "169"; NextLevel = "1056"; Name = "GROUP 2 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "2654"; Wing = "CO"; Unit = "165"; NextLevel = "1056"; Name = "GROUP 3 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "2751"; Wing = "CO"; Unit = "164"; NextLevel = "1056"; Name = "GROUP 4 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "3000"; Wing = "CO"; Unit = "170"; NextLevel = "1056"; Name = "GROUP 5 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1059"; Wing = "CO"; Unit = "072"; NextLevel = "1292"; Name = "BOULDER COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "2876"; Wing = "CO"; Unit = "191"; NextLevel = "1292"; Name = "PLATTE VALLEY CADET SQUADRON"; Type = "CADET"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "423"; Wing = "CO"; Unit = "099"; NextLevel = "1292"; Name = "BROOMFIELD COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1253"; Wing = "CO"; Unit = "136"; NextLevel = "1292"; Name = "JEFFERSON COUNTY SENIOR SQDN"; Type = "SENIOR"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "872"; Wing = "CO"; Unit = "147"; NextLevel = "1292"; Name = "THOMPSON VALLEY COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1476"; Wing = "CO"; Unit = "022"; NextLevel = "1292"; Name = "VANCE BRAND CADET SQDN"; Type = "CADET"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1526"; Wing = "CO"; Unit = "068"; NextLevel = "1292"; Name = "NORTH VALLEY COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "15"; Wing = "CO"; Unit = "015"; NextLevel = "1667"; Name = "THUNDER MOUNTAIN COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1507"; Wing = "CO"; Unit = "030"; NextLevel = "2654"; Name = "COLORADO SPRINGS CADET SQDN"; Type = "CADET"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "873"; Wing = "CO"; Unit = "173"; NextLevel = "2751"; Name = "PARKER COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "4000"; Wing = "CO"; Unit = "201"; NextLevel = "3000"; Name = "FUTURE GROUP SQUADRON"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "5000"; Wing = "CO"; Unit = "202"; NextLevel = "3000"; Name = "INACTIVE GROUP SQUADRON"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "INACTIVE" },
    [PSCustomObject]@{ ORGID = "6000"; Wing = "WY"; Unit = "001"; NextLevel = "226"; Name = "WYOMING WING HQ"; Type = "WING"; Scope = "WING"; Status = "ACTIVE" }
)

$regionalGroups = @(Get-RegionalDistributionGroups -OrganizationRows $organizationRows)
Assert-Equal $regionalGroups.Count 5 "Regional distribution groups should be discovered from active CAPWATCH group organizations."
Assert-Equal $regionalGroups[0].Name "CO Group 1" "Group 1 display name should be generated from CAPWATCH group hierarchy."
Assert-Equal $regionalGroups[0].EmailAddress "group1@cowg.cap.gov" "Group 1 address should match the requested address format."
Assert-ArrayEqual $regionalGroups[0].Units @("022", "068", "072", "099", "136", "147", "191") "Group 1 should include active child units from Organization.txt NextLevel."
Assert-Equal $regionalGroups[3].Name "CO Group 4" "Group 4 display name should be generated from CAPWATCH group hierarchy."
Assert-Equal $regionalGroups[3].EmailAddress "group4@cowg.cap.gov" "Group 4 address should match the requested address format."
Assert-ArrayEqual $regionalGroups[3].Units @("173") "Group 4 should include active child units from Organization.txt NextLevel."
Assert-Equal $regionalGroups[4].EmailAddress "group5@cowg.cap.gov" "New CAPWATCH groups should create matching groupN addresses without code changes."
Assert-ArrayEqual $regionalGroups[4].Units @("201") "New CAPWATCH groups should include only active child units."

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

$unitList = @(
    [PSCustomObject]@{ Unit = "015"; Name = "Senior Squadron" },
    [PSCustomObject]@{ Unit = "147"; Name = "Composite Squadron" }
)

$plannerUsers = @(
    [PSCustomObject]@{ companyName = "CO-015"; employeeType = "SENIOR"; department = ""; mail = "senior.one@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-015"; employeeType = "SENIOR"; department = ""; mail = "senior.one@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-015"; employeeType = "CADET"; department = ""; mail = "cadet.one@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-015"; employeeType = "PARENT"; department = ""; mail = "parent.one@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-015"; employeeType = "SENIOR"; department = "CP"; mail = "cp.staff@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-147"; employeeType = "CADET"; department = ""; mail = "cadet.two@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-147"; employeeType = "SENIOR"; department = "EX"; mail = "ex.staff@cowg.cap.gov" },
    [PSCustomObject]@{ companyName = "CO-147"; employeeType = "SENIOR"; department = ""; mail = "" }
)

$squadronCadetPlans = @(Get-DLSquadronGroupPlans -MemberType "CADET" -UnitList $unitList -AllUsers $plannerUsers)
Assert-Equal $squadronCadetPlans.Count 2 "Cadet squadron planning should create one plan per unit."
Assert-Equal $squadronCadetPlans[0].Identity "CO-015 Cadets" "Cadet squadron plan should target the existing cadet group naming convention."
Assert-ArrayEqual $squadronCadetPlans[0].Members @("cadet.one@cowg.cap.gov", "cp.staff@cowg.cap.gov", "parent.one@cowg.cap.gov") "Cadet squadron members should include cadets, parents, and CP/EX staff with deduplication."

$wingCadetPlan = Get-DLWingGroupPlan -MemberType "CADET" -AllUsers $plannerUsers
Assert-Equal $wingCadetPlan.Identity "CO Wing CADETs" "Wing cadet plan should preserve the existing wing group identity."
Assert-ArrayEqual $wingCadetPlan.Members @("cadet.one@cowg.cap.gov", "cadet.two@cowg.cap.gov", "cp.staff@cowg.cap.gov", "ex.staff@cowg.cap.gov", "parent.one@cowg.cap.gov") "Wing cadet plan should include cadets, parents, and CP/EX staff across units."

$allPlans = @(New-DLGroupUpdatePlans -UnitList $unitList -AllUsers $plannerUsers -OrganizationRows $organizationRows)
Assert-Equal $allPlans.Count 13 "Planner should produce squadron senior/cadet/all plans, wing senior/cadet plans, and dynamically discovered regional plans."

$jobRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("dl-job-test-" + [guid]::NewGuid().ToString("N"))
$queueMessage = Save-DLGroupUpdateJob -Plan $wingCadetPlan -RunId "run-001" -RootPath $jobRoot
Assert-Equal $queueMessage.RunId "run-001" "Queue message should include the planner run id."
Assert-Equal $queueMessage.Identity "CO Wing CADETs" "Queue message should include group identity for logs."
Assert-Equal (Test-Path $queueMessage.JobPath) $true "Planner should persist the full job payload to disk."
Assert-Equal ($queueMessage.PSObject.Properties.Name -contains "Members") $false "Queue message should not carry full group membership."

$savedJob = Get-Content -Path $queueMessage.JobPath -Raw | ConvertFrom-Json
Assert-Equal $savedJob.Identity "CO Wing CADETs" "Saved job should contain the group identity."
Assert-Equal $savedJob.Members.Count 5 "Saved job should contain the full deduped member list."

$script:LoggedMessages = @()
function Write-Log {
    param ([string]$Message)
    $script:LoggedMessages += $Message
}

function Get-DistributionGroupMember {
    param (
        [string]$Identity,
        [string]$ResultSize,
        [string]$ErrorAction
    )

    throw "The operation couldn't be performed because object '$Identity' couldn't be found."
}

$missingGroupPlan = New-DLGroupUpdatePlan `
    -Identity "CO-189 Cadets" `
    -Members @("cadet@example.test") `
    -Category "SquadronCADET"
$missingGroupMessage = Save-DLGroupUpdateJob -Plan $missingGroupPlan -RunId "run-missing-group" -RootPath $jobRoot

Invoke-DLGroupUpdateJob -JobPath $missingGroupMessage.JobPath

$loggedText = $script:LoggedMessages -join "`n"
Assert-StringContains $loggedText "does not exist or could not be resolved" "Missing non-regional distribution groups should be logged and skipped instead of throwing."
Assert-StringContains $loggedText "CO-189 Cadets" "Missing group skip log should include the group identity."

Write-Host "DLSeniorsCadets.Tests.ps1 passed"
