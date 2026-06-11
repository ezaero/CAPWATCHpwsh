$ErrorActionPreference = "Stop"

. "$PSScriptRoot\CrewLinkQualificationTransform.ps1"

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

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

$achievements = @(
    [PSCustomObject]@{ AchvID = "55"; Achv = "MS - Mission Scanner"; FunctionalArea = "Emergency Services" },
    [PSCustomObject]@{ AchvID = "57"; Achv = "MP - SAR/DR Mission Pilot"; FunctionalArea = "Emergency Services" },
    [PSCustomObject]@{ AchvID = "81"; Achv = "MO - Mission Observer"; FunctionalArea = "Emergency Services" },
    [PSCustomObject]@{ AchvID = "124"; Achv = "SET - Skills Evaluator Training"; FunctionalArea = "Emergency Services" }
)

$memberAchievements = @(
    [PSCustomObject]@{ CAPID = "100001"; AchvID = "55"; Status = "ACTIVE"; OriginallyAccomplished = "2025-01-01"; Completed = "2025-01-02"; Expiration = "2028-01-02"; Source = "Ops Quals"; RecID = "1"; ORGID = "CO-001" },
    [PSCustomObject]@{ CAPID = "100001"; AchvID = "81"; Status = "TRAINING"; OriginallyAccomplished = ""; Completed = ""; Expiration = ""; Source = "Ops Quals"; RecID = "2"; ORGID = "CO-001" },
    [PSCustomObject]@{ CAPID = "100002"; AchvID = "57"; Status = "SUSPENDED"; OriginallyAccomplished = "2024-01-01"; Completed = "2024-01-02"; Expiration = "2027-01-02"; Source = "Ops Quals"; RecID = "3"; ORGID = "CO-002" },
    [PSCustomObject]@{ CAPID = "100003"; AchvID = "999"; Status = "ACTIVE"; OriginallyAccomplished = "2026-01-01"; Completed = "2026-01-02"; Expiration = ""; Source = "Ops Quals"; RecID = "4"; ORGID = "CO-003" }
)

$snapshot = @(ConvertTo-CrewLinkQualificationSnapshot `
    -MemberAchievements $memberAchievements `
    -Achievements $achievements `
    -AchievementIds @("55", "57", "81", "999") `
    -SyncTime "2026-06-08T14:00:00.0000000Z")

Assert-Equal $snapshot.Count 4 "Snapshot should include every configured achievement record."

$scanner = $snapshot | Where-Object { $_.CAPID -eq "100001" -and $_.AchvID -eq "55" } | Select-Object -First 1
Assert-Equal $scanner.id "100001-55" "Document id should be stable by CAPID and AchvID."
Assert-Equal $scanner.AchievementName "MS - Mission Scanner" "Achievement lookup should populate the display name."
Assert-Equal $scanner.EligibilityStatus "FULLY_QUALIFIED" "ACTIVE records should be fully qualified."
Assert-True $scanner.IsFullyQualified "ACTIVE records should set IsFullyQualified."
Assert-True $scanner.IsEligibleForCrewMatching "ACTIVE records should be eligible for matching."
Assert-Equal $scanner.SyncSource "CrewLinkQualifications" "Snapshot should identify its sync source."
Assert-Equal $scanner.LastUpdated "2026-06-08T14:00:00.0000000Z" "Snapshot should use the provided sync time."

$observer = $snapshot | Where-Object { $_.CAPID -eq "100001" -and $_.AchvID -eq "81" } | Select-Object -First 1
Assert-Equal $observer.EligibilityStatus "TRAINING" "TRAINING records should preserve trainee eligibility."
Assert-True $observer.IsInTraining "TRAINING records should set IsInTraining."
Assert-True $observer.IsEligibleForCrewMatching "TRAINING records should be visible for training missions."

$suspendedPilot = $snapshot | Where-Object { $_.CAPID -eq "100002" -and $_.AchvID -eq "57" } | Select-Object -First 1
Assert-Equal $suspendedPilot.EligibilityStatus "NOT_ELIGIBLE" "SUSPENDED records should not satisfy matching."
Assert-Equal $suspendedPilot.IsEligibleForCrewMatching $false "SUSPENDED records should not be eligible for matching."

$missingLookup = $snapshot | Where-Object { $_.CAPID -eq "100003" -and $_.AchvID -eq "999" } | Select-Object -First 1
Assert-True $missingLookup.LookupMissing "Unknown AchvIDs should be retained and marked for review."
Assert-Equal $missingLookup.AchievementName $null "Unknown AchvIDs should not invent an achievement name."

$duplicateMemberAchievements = @(
    [PSCustomObject]@{ CAPID = "200001"; AchvID = "55"; Status = "TRAINING"; OriginallyAccomplished = ""; Completed = ""; Expiration = ""; Source = "Ops Quals"; RecID = "5"; ORGID = "CO-004"; DateMod = "2026-01-01" },
    [PSCustomObject]@{ CAPID = "200001"; AchvID = "55"; Status = "ACTIVE"; OriginallyAccomplished = "2026-02-01"; Completed = "2026-02-02"; Expiration = ""; Source = "Ops Quals"; RecID = "6"; ORGID = "CO-004"; DateMod = "2026-02-03" },
    [PSCustomObject]@{ CAPID = "200001"; AchvID = "55"; Status = "EXPIRED"; OriginallyAccomplished = "2025-02-01"; Completed = "2025-02-02"; Expiration = ""; Source = "Ops Quals"; RecID = "7"; ORGID = "CO-004"; DateMod = "2026-03-01" }
)

$dedupedSnapshot = @(ConvertTo-CrewLinkQualificationSnapshot `
    -MemberAchievements $duplicateMemberAchievements `
    -Achievements $achievements `
    -AchievementIds @("55") `
    -SyncTime "2026-06-08T14:00:00.0000000Z")

Assert-Equal $dedupedSnapshot.Count 1 "Duplicate CAPID-AchvID rows should produce one snapshot."
Assert-Equal $dedupedSnapshot[0].Status "ACTIVE" "Duplicate rows should prefer an active qualification over training or expired rows."
Assert-Equal $dedupedSnapshot[0].RecID "6" "The selected duplicate source row should be retained in the snapshot."

$desiredForComparison = @(
    [PSCustomObject]@{
        id = "300001-55"; CAPID = "300001"; AchvID = "55"; AchievementName = "MS - Mission Scanner"; FunctionalArea = "Emergency Services"
        Status = "ACTIVE"; EligibilityStatus = "FULLY_QUALIFIED"; IsFullyQualified = $true; IsInTraining = $false; IsEligibleForCrewMatching = $true
        LookupMissing = $false; OriginallyAccomplished = "2025-01-01"; Completed = "2025-01-02"; Expiration = "2028-01-02"; Source = "Ops Quals"; RecID = "8"; ORGID = "CO-005"
        LastUpdated = "2026-06-08T14:00:00.0000000Z"; SyncSource = "CrewLinkQualifications"
    },
    [PSCustomObject]@{
        id = "300002-81"; CAPID = "300002"; AchvID = "81"; AchievementName = "MO - Mission Observer"; FunctionalArea = "Emergency Services"
        Status = "TRAINING"; EligibilityStatus = "TRAINING"; IsFullyQualified = $false; IsInTraining = $true; IsEligibleForCrewMatching = $true
        LookupMissing = $false; OriginallyAccomplished = ""; Completed = ""; Expiration = ""; Source = "Ops Quals"; RecID = "9"; ORGID = "CO-006"
        LastUpdated = "2026-06-08T14:00:00.0000000Z"; SyncSource = "CrewLinkQualifications"
    },
    [PSCustomObject]@{
        id = "300003-57"; CAPID = "300003"; AchvID = "57"; AchievementName = "MP - SAR/DR Mission Pilot"; FunctionalArea = "Emergency Services"
        Status = "ACTIVE"; EligibilityStatus = "FULLY_QUALIFIED"; IsFullyQualified = $true; IsInTraining = $false; IsEligibleForCrewMatching = $true
        LookupMissing = $false; OriginallyAccomplished = "2025-03-01"; Completed = "2025-03-02"; Expiration = ""; Source = "Ops Quals"; RecID = "10"; ORGID = "CO-007"
        LastUpdated = "2026-06-08T14:00:00.0000000Z"; SyncSource = "CrewLinkQualifications"
    }
)

$existingForComparison = @(
    [PSCustomObject]@{
        id = "300001-55"; CAPID = "300001"; AchvID = "55"; AchievementName = "MS - Mission Scanner"; FunctionalArea = "Emergency Services"
        Status = "ACTIVE"; EligibilityStatus = "FULLY_QUALIFIED"; IsFullyQualified = $true; IsInTraining = $false; IsEligibleForCrewMatching = $true
        LookupMissing = $false; OriginallyAccomplished = "2025-01-01"; Completed = "2025-01-02"; Expiration = "2028-01-02"; Source = "Ops Quals"; RecID = "8"; ORGID = "CO-005"
        LastUpdated = "older-value-should-not-force-upsert"; timestamp = "older-timestamp-should-not-force-upsert"; SyncSource = "CrewLinkQualifications"
    },
    [PSCustomObject]@{
        id = "300002-81"; CAPID = "300002"; AchvID = "81"; AchievementName = "MO - Mission Observer"; FunctionalArea = "Emergency Services"
        Status = "ACTIVE"; EligibilityStatus = "FULLY_QUALIFIED"; IsFullyQualified = $true; IsInTraining = $false; IsEligibleForCrewMatching = $true
        LookupMissing = $false; OriginallyAccomplished = ""; Completed = ""; Expiration = ""; Source = "Ops Quals"; RecID = "9"; ORGID = "CO-006"
        LastUpdated = "older-value"; SyncSource = "CrewLinkQualifications"
    },
    [PSCustomObject]@{
        id = "300004-124"; CAPID = "300004"; AchvID = "124"; Status = "ACTIVE"; SyncSource = "CrewLinkQualifications"
    }
)

$comparison = Compare-CrewLinkQualificationSnapshots -DesiredSnapshots $desiredForComparison -ExistingSnapshots $existingForComparison

Assert-Equal $comparison.ToUpsert.Count 2 "Comparison should upsert only new and changed records."
Assert-Equal $comparison.ToDelete.Count 1 "Comparison should delete stale records no longer present in CAPWATCH."
Assert-Equal $comparison.UnchangedCount 1 "Comparison should ignore LastUpdated and timestamp when detecting changes."
Assert-True (($comparison.ToUpsert | Select-Object -ExpandProperty id) -contains "300002-81") "Changed record should be marked for upsert."
Assert-True (($comparison.ToUpsert | Select-Object -ExpandProperty id) -contains "300003-57") "New record should be marked for upsert."
Assert-Equal $comparison.ToDelete[0].id "300004-124" "Stale existing record should be marked for delete."

Write-Host "CrewLinkQualificationTransform.Tests.ps1 passed"
