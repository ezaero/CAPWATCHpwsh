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

. (Join-Path $PSScriptRoot "Organization.Helpers.ps1")

$env:WING_DESIGNATOR = "CO"

$organizationRows = @(
    [PSCustomObject]@{ ORGID = "1056"; Wing = "CO"; Unit = "001"; NextLevel = "226"; Name = "COLORADO WING HQ"; Type = "WING"; Scope = "WING"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "1667"; Wing = "CO"; Unit = "169"; NextLevel = "1056"; Name = "GROUP 2 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "2751"; Wing = "CO"; Unit = "164"; NextLevel = "1056"; Name = "GROUP 4 HEADQUARTERS"; Type = "GROUP"; Scope = "GROUP"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "15"; Wing = "CO"; Unit = "015"; NextLevel = "1667"; Name = "THUNDER MOUNTAIN COMPOSITE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "674"; Wing = "CO"; Unit = "141"; NextLevel = "1667"; Name = "MONTROSE FLIGHT -- THUNDER MOUNTAIN COMPOSITE SQUADRON"; Type = "FLIGHT"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "3004"; Wing = "CO"; Unit = "200"; NextLevel = "2751"; Name = "THE FIREBORN CADET SQUADRON"; Type = "CADET"; Scope = "UNIT"; Status = "ACTIVE" },
    [PSCustomObject]@{ ORGID = "5000"; Wing = "CO"; Unit = "202"; NextLevel = "2751"; Name = "INACTIVE SQDN"; Type = "COMPOSITE"; Scope = "UNIT"; Status = "INACTIVE" },
    [PSCustomObject]@{ ORGID = "6000"; Wing = "WY"; Unit = "200"; NextLevel = "9999"; Name = "WRONG WING SQDN"; Type = "CADET"; Scope = "UNIT"; Status = "ACTIVE" }
)

$activeUnits = @(Get-CapwatchActiveWingUnits -OrganizationRows $organizationRows)
Assert-ArrayEqual ($activeUnits | Select-Object -ExpandProperty Unit) @("015", "141", "200") "Active wing units should include new squadrons and active flights, excluding inactive/admin/other-wing rows."

$group4Units = @(Get-CapwatchActiveUnitCodesForGroupNumber -OrganizationRows $organizationRows -GroupNumber "4")
Assert-ArrayEqual $group4Units @("200") "Group child lookup should discover new active Group 4 units from NextLevel."

$group2Units = @(Get-CapwatchActiveUnitCodesForGroupNumber -OrganizationRows $organizationRows -GroupNumber "2")
Assert-ArrayEqual $group2Units @("015", "141") "Group child lookup should keep active flights as standalone unit children."

Assert-Equal (ConvertTo-CapwatchCompanyName -Unit "200") "CO-200" "Company names should use the configured wing designator and unit code."

Write-Host "Organization.Helpers.Tests.ps1 passed"
