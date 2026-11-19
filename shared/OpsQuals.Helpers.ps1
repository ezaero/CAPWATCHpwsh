function Get-CAPOpsQualCapIds {
    param (
        [array]$Rows,
        [string]$IdProperty,
        [array]$QualificationIds,
        [array]$Statuses,
        [array]$ManualCapIds = @()
    )

    $eligibleByQualification = @{}
    $suspendedByQualification = @{}
    $normalizedIds = @($QualificationIds | ForEach-Object { "$_".Trim() })
    $normalizedStatuses = @($Statuses | ForEach-Object { "$_".Trim() })

    foreach ($row in $Rows) {
        $qualificationId = "$($row.$IdProperty)".Trim()
        if ($qualificationId -notin $normalizedIds) {
            continue
        }

        $capId = "$($row.CAPID)".Trim()
        if (-not $capId) {
            continue
        }

        $key = "$qualificationId/$capId"
        $status = "$($row.Status)".Trim()
        if ($status -eq "SUSPENDED") {
            $suspendedByQualification[$key] = $capId
        } elseif ($status -in $normalizedStatuses) {
            $eligibleByQualification[$key] = $capId
        }
    }

    $eligibleCapIds = @(
        foreach ($key in $eligibleByQualification.Keys) {
            if (-not $suspendedByQualification.ContainsKey($key)) {
                $eligibleByQualification[$key]
            }
        }
    )

    # Manual inclusions still require the member to be free of an explicit suspension.
    $eligibleManualCapIds = @($ManualCapIds | Where-Object { $_ -notin $suspendedByQualification.Values })
    $eligibleCapIds + $eligibleManualCapIds | Sort-Object -Unique
}

function Get-CAPAchievementCapIds {
    param (
        [array]$AchievementRows,
        [array]$AchievementIds,
        [array]$Statuses
    )

    Get-CAPOpsQualCapIds -Rows $AchievementRows -IdProperty "AchvID" -QualificationIds $AchievementIds -Statuses $Statuses
}

function Get-CAPPilotCapIds {
    param (
        [array]$TaskRows,
        [array]$ManualCapIds = @()
    )

    Get-CAPOpsQualCapIds -Rows $TaskRows -IdProperty "TaskID" -QualificationIds @("69") -Statuses @("ACTIVE") -ManualCapIds $ManualCapIds
}
