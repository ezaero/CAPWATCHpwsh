$ErrorActionPreference = "Stop"

. "$PSScriptRoot\CAPGroups.Helpers.ps1"

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

function Assert-Contains {
    param (
        [array]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -notcontains $Expected) {
        throw "$Message Expected '$Expected' in '$($Actual -join ',')'."
    }
}

$env:WING_DESIGNATOR = "CO"

Assert-Equal (ConvertTo-CAPGroupNamePiece -Value "Director of IT") "Director-of-IT" "Name pieces should preserve readable words."
Assert-Equal (ConvertTo-CAPGroupMailNickname -DisplayName "CAP-Duty-WING-CO-IT-Director-of-IT") "CAPDutyWINGCOITDirectorofIT" "Mail nickname should strip punctuation."

$organizationRows = @(
    [PSCustomObject]@{ ORGID = "1056"; Wing = "CO"; Unit = "001"; Name = "COLORADO WING HQ"; Scope = "WING"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "4000"; Wing = "CO"; Unit = "004"; Name = "COLORADO GROUP 4"; Scope = "GROUP"; Type = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "2000"; Wing = "CO"; Unit = "162"; Name = "BLACK SHEEP SENIOR SQDN"; Scope = "UNIT"; Status = "ACTIVE"; NextLevel = "4000" }
)
$memberRows = @(
    [PSCustomObject]@{ CAPID = "100001"; Wing = "CO"; Region = "RMR"; Unit = "001"; Type = "SENIOR"; MbrStatus = "ACTIVE" },
    [PSCustomObject]@{ CAPID = "100002"; Wing = "CO"; Region = "RMR"; Unit = "162"; Type = "CADET"; MbrStatus = "ACTIVE" },
    [PSCustomObject]@{ CAPID = "100003"; Wing = "CO"; Region = "RMR"; Unit = "162"; Type = "SENIOR"; MbrStatus = "INACTIVE" },
    [PSCustomObject]@{ CAPID = "999999"; Wing = "WY"; Region = "RMR"; Unit = "001"; Type = "SENIOR"; MbrStatus = "ACTIVE" }
)
$dutyRows = @(
    [PSCustomObject]@{ CAPID = "100001"; Duty = "Director of IT"; FunctArea = "IT"; Lvl = "WING"; Asst = "0"; ORGID = "1056" },
    [PSCustomObject]@{ CAPID = "100002"; Duty = "Testing Officer"; FunctArea = "PD"; Lvl = "UNIT"; Asst = "1"; ORGID = "2000" },
    [PSCustomObject]@{ CAPID = "999999"; Duty = "Wyoming Officer"; FunctArea = "EX"; Lvl = "WING"; Asst = "0"; ORGID = "" }
)
$achievementRows = @(
    [PSCustomObject]@{ CAPID = "100001"; AchvID = "53"; Status = "ACTIVE" },
    [PSCustomObject]@{ CAPID = "100001"; AchvID = "57"; Status = "TRAINING" },
    [PSCustomObject]@{ CAPID = "100001"; AchvID = "103"; Status = "ACTIVE" },
    [PSCustomObject]@{ CAPID = "100002"; AchvID = "57"; Status = "TRAINING" },
    [PSCustomObject]@{ CAPID = "100002"; AchvID = "217"; Status = "TRAINING" },
    [PSCustomObject]@{ CAPID = "100003"; AchvID = "57"; Status = "TRAINING" }
)
$achievementLookupRows = @(
    [PSCustomObject]@{ AchvID = "53"; Achv = "GES - General Emergency Services"; FunctionalArea = "OPS-Emergency_Services" },
    [PSCustomObject]@{ AchvID = "57"; Achv = "MP - SAR/DR Mission Pilot"; FunctionalArea = "OPS-Emergency_Services" },
    [PSCustomObject]@{ AchvID = "103"; Achv = "MFC - Mountain Flying Certification"; FunctionalArea = "OPS-Emergency_Services" },
    [PSCustomObject]@{ AchvID = "217"; Achv = "ICUT - Introductory Communications User Training"; FunctionalArea = "Communications" }
)
$taskRows = @(
    [PSCustomObject]@{ CAPID = "100001"; TaskID = "69"; Status = "ACTIVE" },
    [PSCustomObject]@{ CAPID = "100002"; TaskID = "69"; Status = "EXPIRED" },
    [PSCustomObject]@{ CAPID = "100003"; TaskID = "69"; Status = "ACTIVE" }
)
$allUsers = @(
    [PSCustomObject]@{ id = "user-1"; employeeId = "100001"; companyName = "CO-162"; employeeType = "SENIOR"; department = "" },
    [PSCustomObject]@{ id = "user-2"; employeeId = "100002"; companyName = "CO-162"; employeeType = "CADET"; department = "" },
    [PSCustomObject]@{ id = "user-3"; employeeId = "100003"; companyName = "CO-162"; employeeType = "SENIOR"; department = "" },
    [PSCustomObject]@{ id = "user-parent"; employeeId = ""; companyName = "CO-162"; employeeType = "PARENT"; department = "" },
    [PSCustomObject]@{ id = "user-cp"; employeeId = ""; companyName = "CO-162"; employeeType = "SENIOR"; department = "CP" },
    [PSCustomObject]@{ id = "user-wy"; employeeId = "999999"; companyName = "WY-001"; employeeType = "SENIOR"; department = "" }
)

$plans = @(New-CAPSecurityGroupPlans -MemberRows $memberRows -DutyPositionRows $dutyRows -OrganizationRows $organizationRows -AchievementRows $achievementRows -AchievementLookupRows $achievementLookupRows -OpsQualAchievementIds @("53", "57", "103", "217") -TaskRows $taskRows -AllUsers $allUsers -WingDesignator "CO")
$planNames = @($plans | Select-Object -ExpandProperty DisplayName)

Assert-Contains $planNames "CAP-Wing-CO" "Plans should include a wing group."
Assert-Contains $planNames "CAP-Region-RMR" "Plans should include a region group."
Assert-Contains $planNames "CAP-Unit-CO-162" "Plans should include a unit group."
Assert-Contains $planNames "CAP-MemberType-SENIOR" "Plans should include member type groups."
Assert-Contains $planNames "CAP-MemberStatus-INACTIVE" "Plans should include member status groups."
Assert-Contains $planNames "CAP-Duty-WING-CO-IT-Director-of-IT" "Plans should include wing duty groups."
Assert-Contains $planNames "CAP-Duty-UNIT-CO-PD-Assistant-Testing-Officer" "Plans should include assistant duty groups."
Assert-Contains $planNames "CAP-WingAudience-CO-Cadets" "Plans should include wing cadet audience groups."
Assert-Contains $planNames "CAP-GroupAudience-CO-4-All" "Plans should include group audience groups."
Assert-Contains $planNames "CAP-UnitAudience-CO-162-All" "Plans should include unit audience groups."
Assert-Contains $planNames "CAP-UnitAudience-CO-162-Seniors" "Plans should include unit senior audience groups."
Assert-Contains $planNames "CAP-UnitAudience-CO-162-Cadets" "Plans should include unit cadet audience groups."
Assert-Contains $planNames "CAP-OpsQual-Pilot" "Plans should include ops-qual pilot groups."
Assert-Contains $planNames "CAP-OpsQual-GES-General-Emergency-Services" "Plans should include exact ops-qual achievement groups without numeric AchvIDs in the display name."
Assert-Contains $planNames "CAP-OpsQual-MFC-Mountain-Flying-Certification" "Plans should include exact ops-qual achievement groups from Achievements.txt."
Assert-Contains $planNames "CAP-OpsQual-Mission-Pilot" "Plans should include ops-qual mission pilot groups."
Assert-Contains $planNames "CAP-OpsQual-ES-Group-4" "Plans should include group-scoped ops-qual groups."
Assert-Contains $planNames "CAP-OpsQual-Communicator" "Plans should include training-status ops-qual groups."
if ($planNames -contains "CAP-OpsQual-53-GES-General-Emergency-Services" -or $planNames -contains "CAP-OpsQual-Achv-53-GES-General-Emergency-Services") {
    throw "Exact ops-qual group names should omit numeric AchvIDs."
}

$wingPlan = $plans | Where-Object DisplayName -eq "CAP-Wing-CO" | Select-Object -First 1
Assert-Equal $wingPlan.MemberIds.Count 2 "Wing group should include active CO members only."
Assert-Contains $wingPlan.MemberIds "user-1" "Wing group should include active senior member."
Assert-Contains $wingPlan.MemberIds "user-2" "Wing group should include active cadet member."

$inactivePlan = $plans | Where-Object DisplayName -eq "CAP-MemberStatus-INACTIVE" | Select-Object -First 1
Assert-Equal $inactivePlan.MemberIds.Count 1 "Status groups should include matching status members."
Assert-Contains $inactivePlan.MemberIds "user-3" "Inactive status group should include inactive user."

$unitCadetsPlan = $plans | Where-Object DisplayName -eq "CAP-UnitAudience-CO-162-Cadets" | Select-Object -First 1
Assert-Equal $unitCadetsPlan.MemberIds.Count 3 "Unit cadet audience should include cadets, parents, and CP/EX staff."
Assert-Contains $unitCadetsPlan.MemberIds "user-2" "Unit cadet audience should include cadets."
Assert-Contains $unitCadetsPlan.MemberIds "user-parent" "Unit cadet audience should include parents."
Assert-Contains $unitCadetsPlan.MemberIds "user-cp" "Unit cadet audience should include CP staff."

$pilotPlan = $plans | Where-Object DisplayName -eq "CAP-OpsQual-Pilot" | Select-Object -First 1
Assert-Equal $pilotPlan.MemberIds.Count 2 "Pilot ops-qual group should include active TaskID 69 users."
Assert-Contains $pilotPlan.MemberIds "user-1" "Pilot ops-qual group should include active pilot user."
Assert-Contains $pilotPlan.MemberIds "user-3" "Pilot ops-qual group should include active pilot task regardless member status."

$missionPilotPlan = $plans | Where-Object DisplayName -eq "CAP-OpsQual-Mission-Pilot" | Select-Object -First 1
Assert-Equal $missionPilotPlan.MemberIds.Count 2 "Mission Pilot source group should require active TaskID 69 plus mission pilot achievement."
Assert-Contains $missionPilotPlan.MemberIds "user-1" "Mission Pilot source group should include qualified active pilots."

$group4EsPlan = $plans | Where-Object DisplayName -eq "CAP-OpsQual-ES-Group-4" | Select-Object -First 1
Assert-Equal $group4EsPlan.MemberIds.Count 1 "Group 4 ES source group should include ES users in active Group 4 units."
Assert-Contains $group4EsPlan.MemberIds "user-1" "Group 4 ES source group should include matching Group 4 users."

$gesPlan = $plans | Where-Object DisplayName -eq "CAP-OpsQual-GES-General-Emergency-Services" | Select-Object -First 1
Assert-Equal $gesPlan.MemberIds.Count 1 "Exact ops-qual achievement groups should include active achievement users."
Assert-Contains $gesPlan.MemberIds "user-1" "Exact GES group should include active GES users."

$missionPilotExactPlan = $plans | Where-Object DisplayName -eq "CAP-OpsQual-MP-SAR-DR-Mission-Pilot" | Select-Object -First 1
Assert-Equal $missionPilotExactPlan.MemberIds.Count 3 "Exact ops-qual achievement groups should include active and training achievement users."
Assert-Contains $missionPilotExactPlan.MemberIds "user-1" "Exact mission pilot group should include training mission pilot users."
Assert-Contains $missionPilotExactPlan.MemberIds "user-2" "Exact mission pilot group should include training members even without active pilot task."
Assert-Contains $missionPilotExactPlan.MemberIds "user-3" "Exact mission pilot group should include training users regardless member status."

Write-Host "CAPGroups.Tests.ps1 passed"

