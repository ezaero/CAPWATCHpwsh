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
    [PSCustomObject]@{ ORGID = "2000"; Wing = "CO"; Unit = "162"; Name = "BLACK SHEEP SENIOR SQDN"; Scope = "UNIT"; Status = "ACTIVE" }
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
$allUsers = @(
    [PSCustomObject]@{ id = "user-1"; employeeId = "100001" },
    [PSCustomObject]@{ id = "user-2"; employeeId = "100002" },
    [PSCustomObject]@{ id = "user-3"; employeeId = "100003" },
    [PSCustomObject]@{ id = "user-wy"; employeeId = "999999" }
)

$plans = @(New-CAPSecurityGroupPlans -MemberRows $memberRows -DutyPositionRows $dutyRows -OrganizationRows $organizationRows -AllUsers $allUsers -WingDesignator "CO")
$planNames = @($plans | Select-Object -ExpandProperty DisplayName)

Assert-Contains $planNames "CAP-Wing-CO" "Plans should include a wing group."
Assert-Contains $planNames "CAP-Region-RMR" "Plans should include a region group."
Assert-Contains $planNames "CAP-Unit-CO-162" "Plans should include a unit group."
Assert-Contains $planNames "CAP-MemberType-SENIOR" "Plans should include member type groups."
Assert-Contains $planNames "CAP-MemberStatus-INACTIVE" "Plans should include member status groups."
Assert-Contains $planNames "CAP-Duty-WING-CO-IT-Director-of-IT" "Plans should include wing duty groups."
Assert-Contains $planNames "CAP-Duty-UNIT-CO-PD-Assistant-Testing-Officer" "Plans should include assistant duty groups."

$wingPlan = $plans | Where-Object DisplayName -eq "CAP-Wing-CO" | Select-Object -First 1
Assert-Equal $wingPlan.MemberIds.Count 2 "Wing group should include active CO members only."
Assert-Contains $wingPlan.MemberIds "user-1" "Wing group should include active senior member."
Assert-Contains $wingPlan.MemberIds "user-2" "Wing group should include active cadet member."

$inactivePlan = $plans | Where-Object DisplayName -eq "CAP-MemberStatus-INACTIVE" | Select-Object -First 1
Assert-Equal $inactivePlan.MemberIds.Count 1 "Status groups should include matching status members."
Assert-Contains $inactivePlan.MemberIds "user-3" "Inactive status group should include inactive user."

Write-Host "CAPGroups.Tests.ps1 passed"

