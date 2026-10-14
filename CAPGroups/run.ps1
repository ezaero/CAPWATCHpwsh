param($CapwatchDownloadMessage)

$ErrorActionPreference = "Stop"

. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\CAPGroups.Helpers.ps1"

$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
$execute = $env:CAP_GROUPS_EXECUTE -match '^(?i:true|yes|1)$'
$queueInPreview = $env:CAP_GROUPS_QUEUE_DLDUTYPOS_IN_PREVIEW -match '^(?i:true|yes|1)$'
$batchSize = 35
if (-not [string]::IsNullOrWhiteSpace($env:CAP_GROUPS_PLAN_BATCH_SIZE)) {
    $batchSize = [Math]::Max(1, [int]$env:CAP_GROUPS_PLAN_BATCH_SIZE)
}

$triggerMessage = $CapwatchDownloadMessage
if ($CapwatchDownloadMessage -is [string] -and -not [string]::IsNullOrWhiteSpace($CapwatchDownloadMessage)) {
    try {
        $triggerMessage = $CapwatchDownloadMessage | ConvertFrom-Json -ErrorAction Stop
    } catch {
        $triggerMessage = [PSCustomObject]@{
            Source = "unknown"
            RawMessage = $CapwatchDownloadMessage
        }
    }
}

$startIndex = 0
if ($null -ne $triggerMessage -and $null -ne $triggerMessage.StartPlanIndex) {
    $startIndex = [Math]::Max(0, [int]$triggerMessage.StartPlanIndex)
}

Write-Log "CAPGroups security-group reconciler started. Execute=$execute StartPlanIndex=$startIndex BatchSize=$batchSize ------------------------------------------------"

Push-Location $CAPWATCHDATADIR
try {
    $msGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
    Connect-MgGraph -AccessToken $msGraphAccessToken -NoWelcome

	Write-Log "CAPGroups loading CAPWATCH source files."
    $memberRows = @(Import-Csv "$CAPWATCHDATADIR\Member.txt" -ErrorAction Stop)
    $dutyPositionRows = @(Import-Csv "$CAPWATCHDATADIR\DutyPosition.txt" -ErrorAction Stop)
    $organizationRows = @(Import-CapwatchOrganizationRows -CapwatchDataDir $CAPWATCHDATADIR)
	Write-Log "CAPGroups loaded $($memberRows.Count) members, $($dutyPositionRows.Count) duty positions, and $($organizationRows.Count) organizations."

	Write-Log "CAPGroups loading Entra users for CAPID matching."
    $allUsers = @(GetAllUsers -SelectFields "id,displayName,mail,userPrincipalName,employeeId")
	Write-Log "CAPGroups loaded $($allUsers.Count) Entra users."

    Write-Log "CAPGroups building security-group plans."
    $plans = @(New-CAPSecurityGroupPlans `
        -MemberRows $memberRows `
        -DutyPositionRows $dutyPositionRows `
        -OrganizationRows $organizationRows `
        -AllUsers $allUsers `
        -WingDesignator $env:WING_DESIGNATOR)

    Write-Log "CAPGroups planned $($plans.Count) security groups."

    $endIndex = [Math]::Min($plans.Count - 1, $startIndex + $batchSize - 1)
    $plansToProcess = @()
    if ($startIndex -lt $plans.Count) {
        $plansToProcess = @($plans[$startIndex..$endIndex])
    }

    Write-Log "CAPGroups processing $($plansToProcess.Count) plan(s), indexes $startIndex through $endIndex of $($plans.Count - 1)."

    foreach ($plan in $plansToProcess) {
        $result = Sync-CAPSecurityGroupPlan -Plan $plan -Execute:$execute
        Write-Log "CAPGroups '$($result.DisplayName)': desired=$($result.DesiredCount), add=$($result.AddCount), remove=$($result.RemoveCount)."
    }

    $nextIndex = $endIndex + 1
    if ($nextIndex -lt $plans.Count) {
        $continuationMessage = [PSCustomObject]@{
            Source = "CAPGroups"
            OriginalMessage = $triggerMessage
            StartPlanIndex = $nextIndex
            SecurityGroupPlanCount = $plans.Count
            Execute = $execute
            QueuedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        }
        Write-Log "CAPGroups pushing continuation message for plan index $nextIndex."
        Push-OutputBinding -Name CAPGroupsQueue -Value ($continuationMessage | ConvertTo-Json -Depth 10 -Compress)
        Write-Log "CAPGroups queued continuation starting at plan index $nextIndex."
    } elseif ($execute -or $queueInPreview) {
        $completedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        $queueMessage = [PSCustomObject]@{
            CompletedAtUtc = $completedAtUtc
            Source = "CAPGroups"
            CapwatchDownloadMessage = $triggerMessage
            SecurityGroupPlanCount = $plans.Count
            Execute = $execute
        }
        Write-Log "CAPGroups pushing DLDutyPos handoff message."
        Push-OutputBinding -Name DLDutyPosQueue -Value ($queueMessage | ConvertTo-Json -Depth 10 -Compress)
        Write-Log "CAPGroups queued DLDutyPos after successful security-group reconciliation."
    } else {
        Write-Log "CAPGroups did not queue DLDutyPos because CAP_GROUPS_EXECUTE is not true. Set CAP_GROUPS_QUEUE_DLDUTYPOS_IN_PREVIEW=true to test the downstream planner without writing security groups."
    }
} catch {
    Write-Log "CAPGroups failed. $($_ | Out-String)"
    throw
} finally {
    Pop-Location
}

Write-Log "CAPGroups security-group reconciler ended. ------------------------------------------------"

