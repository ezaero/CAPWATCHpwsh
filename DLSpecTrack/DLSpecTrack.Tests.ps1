$ErrorActionPreference = "Stop"

. "$PSScriptRoot\DLSpecTrack.Helpers.ps1"

function Assert-Equal {
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function Assert-Contains {
    param(
        [array]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -notcontains $Expected) {
        throw "$Message Expected collection to contain '$Expected'. Actual: '$($Actual -join ',')'."
    }
}

function Assert-NotNull {
    param(
        [object]$Actual,
        [string]$Message
    )

    if ($null -eq $Actual) {
        throw $Message
    }
}

$address = Get-RecruitingGroupAddress -UnitCode "022"
Assert-Equal $address "co-022-recruiting@cowg.cap.gov" "Recruiting group address should match the deployed DL format."

$legacyAddress = "co-164recruiting@cowg.cap.gov"
$desiredAddress = Get-RecruitingGroupAddress -UnitCode "164"
Assert-Equal (Test-RecruitingGroupAddressNeedsRename -CurrentAddress $legacyAddress -DesiredAddress $desiredAddress) $true "Legacy no-hyphen recruiting addresses should be renamed."
Assert-Equal (Test-RecruitingGroupAddressNeedsRename -CurrentAddress $desiredAddress -DesiredAddress $desiredAddress) $false "Standard recruiting addresses should not be renamed."

$internalOnlyRecruitingGroup = [PSCustomObject]@{
    RequireSenderAuthenticationEnabled = $true
}
$externallyAllowedRecruitingGroup = [PSCustomObject]@{
    RequireSenderAuthenticationEnabled = $false
}

Assert-Equal (Test-RecruitingGroupRequiresExternalSenderUpdate -GroupName "CO-183 Recruiting" -Group $internalOnlyRecruitingGroup) $true "Recruiting groups that require authenticated senders should be updated to allow external senders."
Assert-Equal (Test-RecruitingGroupRequiresExternalSenderUpdate -GroupName "CO-183 Recruiting" -Group $externallyAllowedRecruitingGroup) $false "Recruiting groups already allowing external senders should not be updated."
Assert-Equal (Test-RecruitingGroupRequiresExternalSenderUpdate -GroupName "Emergency Services" -Group $internalOnlyRecruitingGroup) $false "Non-recruiting specialty track groups should not be changed by the recruiting external sender setting."

$plainRecruitingDistributionGroup = [PSCustomObject]@{
    RecipientTypeDetails = "MailUniversalDistributionGroup"
}
$securityRecruitingDistributionGroup = [PSCustomObject]@{
    RecipientTypeDetails = "MailUniversalSecurityGroup"
}

Assert-Equal (Get-DistributionGroupTypeForName -GroupName "CO-183 Recruiting") "Security" "Recruiting groups should be created as mail-enabled security groups."
Assert-Equal (Get-DistributionGroupTypeForName -GroupName "Emergency Services") "Distribution" "Non-recruiting specialty track groups should remain distribution groups."
Assert-Equal (Test-RecruitingGroupRequiresSecurityMigration -GroupName "CO-183 Recruiting" -Group $plainRecruitingDistributionGroup) $true "Existing recruiting distribution groups should be recreated as mail-enabled security groups."
Assert-Equal (Test-RecruitingGroupRequiresSecurityMigration -GroupName "CO-183 Recruiting" -Group $securityRecruitingDistributionGroup) $false "Existing recruiting mail-enabled security groups should not be recreated."
Assert-Equal (Test-RecruitingGroupRequiresSecurityMigration -GroupName "Emergency Services" -Group $plainRecruitingDistributionGroup) $false "Non-recruiting specialty track groups should not be recreated."

$recruitingMemberUpdateOptions = Get-DistributionGroupMemberUpdateOptions -GroupName "CO-183 Recruiting"
$specialtyTrackMemberUpdateOptions = Get-DistributionGroupMemberUpdateOptions -GroupName "Emergency Services"
$recruitingRemovalOptions = Get-DistributionGroupRemovalOptions -GroupName "CO-183 Recruiting"

Assert-Equal $recruitingMemberUpdateOptions.BypassSecurityGroupManagerCheck $true "Recruiting group membership updates should bypass the group-manager check."
Assert-Equal $recruitingMemberUpdateOptions.ErrorAction "Stop" "Recruiting group membership update failures should be terminating."
Assert-Equal $specialtyTrackMemberUpdateOptions.ContainsKey("BypassSecurityGroupManagerCheck") $false "Non-recruiting membership updates should not add the manager-check bypass."
Assert-Equal $recruitingRemovalOptions.BypassSecurityGroupManagerCheck $true "Recruiting group removal should bypass the group-manager check during destructive migration."
Assert-Equal $recruitingRemovalOptions.Confirm $false "Recruiting group removal should remain non-interactive during destructive migration."

$graphMembersUri = Get-GraphGroupMembersUri -GroupId "group-123"
Assert-Equal $graphMembersUri 'https://graph.microsoft.com/v1.0/groups/group-123/members?$select=id' "Graph member lookup URI should preserve the `$select query parameter."
Assert-Equal (Get-GraphGroupLookupMaxAttempts -GroupName "CO-183 Recruiting") 6 "Recruiting groups should retry Graph lookup after Exchange creation."
Assert-Equal (Get-GraphGroupLookupMaxAttempts -GroupName "Emergency Services") 1 "Existing specialty-track groups should keep one Graph lookup attempt."
Assert-Equal (Get-GraphGroupLookupDelaySeconds) 5 "Graph lookup retry delay should give Exchange-created groups time to replicate."

$specTracks = @(
    [PSCustomObject]@{ CAPID = "100001"; Track = "RECRUITING AND RETENTION OFFICER" },
    [PSCustomObject]@{ CAPID = "100002"; Track = "Emergency Services" }
)

$dutyPositions = @(
    [PSCustomObject]@{ CAPID = "200001"; Duty = "Recruiting Officer"; FunctArea = "" },
    [PSCustomObject]@{ CAPID = "300001"; Duty = "Squadron Commander"; FunctArea = "Command" },
    [PSCustomObject]@{ CAPID = "300002"; Duty = "Chief of Staff"; FunctArea = "Command" },
    [PSCustomObject]@{ CAPID = "400001"; Duty = "Safety Officer"; FunctArea = "Safety" }
)

$commanders = @(
    [PSCustomObject]@{ CAPID = "500001"; Wing = "CO"; Unit = "164"; NameLast = "Commander"; NameFirst = "Colorado" },
    [PSCustomObject]@{ CAPID = "500002"; Wing = "AL"; Unit = "024"; NameLast = "Commander"; NameFirst = "Alabama" }
)

$recruitingCapIds = @(Get-RecruitingCapIds -specTracks $specTracks -dutyPositions $dutyPositions -commanders $commanders -wingDesignator "CO")

Assert-Contains $recruitingCapIds "100001" "Recruiting specialty track CAPID should be included."
Assert-Contains $recruitingCapIds "200001" "Recruiting duty position CAPID should be included."
Assert-Contains $recruitingCapIds "300001" "Commander CAPID should be included."
Assert-Contains $recruitingCapIds "300002" "Chief of Staff CAPID should be included."
Assert-Contains $recruitingCapIds "500001" "CO commander CAPID from Commanders.txt should be included."
Assert-Equal ($recruitingCapIds -contains "400001") $false "Unrelated duty CAPID should not be included."
Assert-Equal ($recruitingCapIds -contains "500002") $false "Commanders from other wings should not be included."

$userWithoutEmployeeId = [PSCustomObject]@{
    employeeId = $null
}

Assert-Equal (Get-UserCapId -User $userWithoutEmployeeId) $null "CAPID resolution should return null when employeeId is missing."

$comparison = Compare-Arrays `
    -Array1 @(
        [PSCustomObject]@{ id = "user-1"; displayName = "User One"; mail = "one@example.test" }
    ) `
    -Array2 @($null, "", "   ", "stale-user")

Assert-NotNull $comparison "Compare-Arrays should return a result even when existing member ids contain null or blank values."
Assert-Equal $comparison.Add.Count 1 "Desired user should be marked for add."
Assert-Equal $comparison.Remove.Count 1 "Only nonblank stale existing ids should be marked for removal."
Assert-Equal $comparison.Remove[0] "stale-user" "Null and blank existing ids should not be passed into ContainsKey."

$co186Users = @(
    [PSCustomObject]@{
        id = "commander-user"
        employeeId = "211568"
        companyName = "CO-186"
        department = ""
        mail = "charles.sellers@cowg.cap.gov"
    },
    [PSCustomObject]@{
        id = "wrong-unit-user"
        employeeId = "211568"
        companyName = "CO-164"
        department = ""
        mail = "wrong.unit@cowg.cap.gov"
    }
)

$co186RecruitingMembers = @(Get-RecruitingMembersForUnit -allUsers $co186Users -unit "186" -recruitingCAPIDs @("211568"))

Assert-Equal $co186RecruitingMembers.Count 1 "CO-186 recruiting members should include the commander matched by employeeId."
Assert-Equal $co186RecruitingMembers[0].id "commander-user" "CO-186 recruiting selection should use employeeId CAPID and companyName unit."

$co186Commanders = @(
    [PSCustomObject]@{
        CAPID = "211568"
        Wing = "CO"
        Unit = "186"
        NameLast = "Sellers"
        NameFirst = "Charles"
    }
)

$co186UsersWithCurrentCompanyName = @(
    [PSCustomObject]@{
        id = "commander-user"
        employeeId = "211568"
        companyName = "CO-186"
        department = ""
        mail = "charles.sellers@cowg.cap.gov"
    }
)

$co186MembersFromCommanderFileOnly = @(Get-RecruitingMembersForUnit `
    -allUsers $co186UsersWithCurrentCompanyName `
    -unit "186" `
    -recruitingCAPIDs @() `
    -commanders $co186Commanders `
    -wingDesignator "CO")

Assert-Equal $co186MembersFromCommanderFileOnly.Count 1 "CO-186 commander should be included from Commanders.txt even when not present in the global recruiting CAPID list."
Assert-Equal $co186MembersFromCommanderFileOnly[0].employeeId "211568" "Commanders.txt unit match should resolve against employeeId."

$co186UsersWithStaleCompanyName = @(
    [PSCustomObject]@{
        id = "charles-sellers-user"
        displayName = "Charles Sellers, Lt Col"
        employeeId = "211568"
        companyName = "CO-164"
        department = ""
        mail = "Charles.Sellers@cowg.cap.gov"
    }
)

$co186MembersWithStaleCompanyName = @(Get-RecruitingMembersForUnit `
    -allUsers $co186UsersWithStaleCompanyName `
    -unit "186" `
    -recruitingCAPIDs @() `
    -commanders $co186Commanders `
    -wingDesignator "CO")

Assert-Equal $co186MembersWithStaleCompanyName.Count 1 "CO-186 commander from Commanders.txt should be included even when Entra companyName is stale."
Assert-Equal $co186MembersWithStaleCompanyName[0].id "charles-sellers-user" "Commanders.txt should override stale companyName for commander membership."

Write-Host "DLSpecTrack.Tests.ps1 passed"
