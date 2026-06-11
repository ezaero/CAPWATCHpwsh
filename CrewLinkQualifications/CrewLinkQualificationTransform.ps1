$script:DefaultCrewLinkAchievementIds = @(
    "44",  # VFR Pilot
    "55",  # MS - Mission Scanner
    "56",  # TMP - Transport Mission Pilot
    "57",  # MP - SAR/DR Mission Pilot
    "59",  # Instructor Pilot - Airplane
    "60",  # Check Pilot - Airplane
    "81",  # MO - Mission Observer
    "90",  # Mission Check Pilot
    "91",  # Orientation Pilot - Airplane
    "124", # SET - Skills Evaluator Training
    "164", # Check Pilot Examiner - Airplane
    "166", # Mission Check Pilot Examiner
    "184", # Instrument Pilot
    "186", # FRO
    "193", # AP - Airborne Photographer
    "238", # CAP Aircrew rating
    "239", # CAP Aircrew rating
    "240", # CAP Aircrew rating
    "257", # sUAS Technician
    "258", # sUAS Mission Pilot
    "262", # sUAS Mission Technician
    "263"  # sUAS Instructor
)

function Get-CrewLinkEligibilityStatus {
    param(
        [AllowNull()]
        [string]$Status
    )

    switch (($Status ?? "").Trim().ToUpperInvariant()) {
        "ACTIVE" { return "FULLY_QUALIFIED" }
        "TRAINING" { return "TRAINING" }
        default { return "NOT_ELIGIBLE" }
    }
}

function Get-CrewLinkStatusRank {
    param(
        [AllowNull()]
        [string]$Status
    )

    switch (($Status ?? "").Trim().ToUpperInvariant()) {
        "ACTIVE" { return 1 }
        "TRAINING" { return 2 }
        default { return 3 }
    }
}

function Get-CrewLinkSourceDate {
    param(
        [object]$MemberAchievement
    )

    foreach ($propertyName in @("DateMod", "Completed", "AuthDate", "DateCreated")) {
        $rawValue = $MemberAchievement.$propertyName
        if ([string]::IsNullOrWhiteSpace([string]$rawValue)) {
            continue
        }

        try {
            return [datetime]$rawValue
        } catch {
            continue
        }
    }

    return [datetime]::MinValue
}

function Select-CrewLinkPreferredMemberAchievement {
    param(
        [Parameter(Mandatory = $true)]
        [array]$MemberAchievements
    )

    return $MemberAchievements |
        Sort-Object `
            @{ Expression = { Get-CrewLinkStatusRank -Status $_.Status }; Ascending = $true },
            @{ Expression = { Get-CrewLinkSourceDate -MemberAchievement $_ }; Ascending = $false } |
        Select-Object -First 1
}

function Get-CrewLinkComparableFieldNames {
    return @(
        "CAPID",
        "AchvID",
        "AchievementName",
        "FunctionalArea",
        "Status",
        "EligibilityStatus",
        "IsFullyQualified",
        "IsInTraining",
        "IsEligibleForCrewMatching",
        "LookupMissing",
        "OriginallyAccomplished",
        "Completed",
        "Expiration",
        "Source",
        "RecID",
        "ORGID",
        "SyncSource"
    )
}

function ConvertTo-CrewLinkComparableValue {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [bool]) {
        return $Value
    }

    return ([string]$Value).Trim()
}

function Test-CrewLinkQualificationChanged {
    param(
        [Parameter(Mandatory = $true)]
        [object]$DesiredSnapshot,

        [Parameter(Mandatory = $true)]
        [object]$ExistingSnapshot
    )

    foreach ($fieldName in (Get-CrewLinkComparableFieldNames)) {
        $desiredValue = ConvertTo-CrewLinkComparableValue -Value $DesiredSnapshot.$fieldName
        $existingValue = ConvertTo-CrewLinkComparableValue -Value $ExistingSnapshot.$fieldName

        if ($desiredValue -ne $existingValue) {
            return $true
        }
    }

    return $false
}

function Compare-CrewLinkQualificationSnapshots {
    param(
        [Parameter(Mandatory = $true)]
        [array]$DesiredSnapshots,

        [Parameter(Mandatory = $true)]
        [array]$ExistingSnapshots
    )

    $existingLookup = @{}
    foreach ($existingSnapshot in $ExistingSnapshots) {
        if ($existingSnapshot.id) {
            $existingLookup[[string]$existingSnapshot.id] = $existingSnapshot
        }
    }

    $desiredLookup = @{}
    $toUpsert = @()
    $unchangedCount = 0

    foreach ($desiredSnapshot in $DesiredSnapshots) {
        if (-not $desiredSnapshot.id) {
            continue
        }

        $snapshotId = [string]$desiredSnapshot.id
        $desiredLookup[$snapshotId] = $desiredSnapshot

        if (-not $existingLookup.ContainsKey($snapshotId)) {
            $toUpsert += $desiredSnapshot
            continue
        }

        if (Test-CrewLinkQualificationChanged -DesiredSnapshot $desiredSnapshot -ExistingSnapshot $existingLookup[$snapshotId]) {
            $toUpsert += $desiredSnapshot
        } else {
            $unchangedCount++
        }
    }

    $toDelete = @()
    foreach ($existingSnapshotId in $existingLookup.Keys) {
        if (-not $desiredLookup.ContainsKey($existingSnapshotId)) {
            $toDelete += $existingLookup[$existingSnapshotId]
        }
    }

    return [PSCustomObject]@{
        ToUpsert = @($toUpsert)
        ToDelete = @($toDelete)
        UnchangedCount = $unchangedCount
    }
}

function ConvertTo-CrewLinkQualificationSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [array]$MemberAchievements,

        [Parameter(Mandatory = $true)]
        [array]$Achievements,

        [string[]]$AchievementIds = $script:DefaultCrewLinkAchievementIds,

        [string]$SyncTime = (Get-Date -Format o)
    )

    $achievementLookup = @{}
    foreach ($achievement in $Achievements) {
        if ($achievement.AchvID) {
            $achievementLookup[[string]$achievement.AchvID] = $achievement
        }
    }

    $configuredAchievementIds = @{}
    foreach ($achievementId in $AchievementIds) {
        if (-not [string]::IsNullOrWhiteSpace($achievementId)) {
            $configuredAchievementIds[[string]$achievementId] = $true
        }
    }

    $memberAchievementLookup = @{}
    foreach ($memberAchievement in $MemberAchievements) {
        $capId = [string]$memberAchievement.CAPID
        $achievementId = [string]$memberAchievement.AchvID

        if ([string]::IsNullOrWhiteSpace($capId) -or [string]::IsNullOrWhiteSpace($achievementId)) {
            continue
        }

        if ($configuredAchievementIds.Count -gt 0 -and -not $configuredAchievementIds.ContainsKey($achievementId)) {
            continue
        }

        $snapshotKey = "$capId-$achievementId"
        if (-not $memberAchievementLookup.ContainsKey($snapshotKey)) {
            $memberAchievementLookup[$snapshotKey] = @()
        }

        $memberAchievementLookup[$snapshotKey] += $memberAchievement
    }

    foreach ($snapshotKey in ($memberAchievementLookup.Keys | Sort-Object)) {
        $memberAchievement = Select-CrewLinkPreferredMemberAchievement -MemberAchievements $memberAchievementLookup[$snapshotKey]
        $capId = [string]$memberAchievement.CAPID
        $achievementId = [string]$memberAchievement.AchvID
        $achievement = $achievementLookup[$achievementId]
        $eligibilityStatus = Get-CrewLinkEligibilityStatus -Status $memberAchievement.Status

        [PSCustomObject]@{
            id = "$capId-$achievementId"
            CAPID = $capId
            AchvID = $achievementId
            AchievementName = if ($achievement) { $achievement.Achv } else { $null }
            FunctionalArea = if ($achievement) { $achievement.FunctionalArea } else { $null }
            Status = if ($memberAchievement.Status) { ([string]$memberAchievement.Status).Trim().ToUpperInvariant() } else { $null }
            EligibilityStatus = $eligibilityStatus
            IsFullyQualified = ($eligibilityStatus -eq "FULLY_QUALIFIED")
            IsInTraining = ($eligibilityStatus -eq "TRAINING")
            IsEligibleForCrewMatching = ($eligibilityStatus -in @("FULLY_QUALIFIED", "TRAINING"))
            LookupMissing = ($null -eq $achievement)
            OriginallyAccomplished = $memberAchievement.OriginallyAccomplished
            Completed = $memberAchievement.Completed
            Expiration = $memberAchievement.Expiration
            Source = $memberAchievement.Source
            RecID = $memberAchievement.RecID
            ORGID = $memberAchievement.ORGID
            LastUpdated = $SyncTime
            SyncSource = "CrewLinkQualifications"
        }
    }
}
