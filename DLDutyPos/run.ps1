param($CAPGroupsMessage)

. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\DLDutyPos.Helpers.ps1"

Write-Log "DLDutyPos planner started. ------------------------------------------------"

$configPath = Get-DLDutyPosConfigPath
$definitions = @(Import-DLDutyPosConfig -Path $configPath)
if ($definitions.Count -eq 0) {
    Write-Log "DLDutyPos did not find distribution-list definitions at $configPath. No jobs were queued."
    return
}

$msGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Connect-MgGraph -AccessToken $msGraphAccessToken -NoWelcome

$runId = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
$allUsers = @(GetAllUsers -SelectFields "id,displayName,mail,userPrincipalName,employeeId")
$sourceGroupPatternPrefixes = @(Get-DLDutyPosSourceGroupPatternPrefixes -Definitions $definitions)
if ($sourceGroupPatternPrefixes.Count -gt 0) {
    $candidateSourceGroupNames = @(Get-DLDutyPosGraphGroupDisplayNamesByPrefix -Prefixes $sourceGroupPatternPrefixes)
    $definitions = @(Resolve-DLDutyPosDefinitionGroupPatterns -Definitions $definitions -CandidateGroupNames $candidateSourceGroupNames)
    Write-Log "DLDutyPos resolved $($candidateSourceGroupNames.Count) candidate source groups from $($sourceGroupPatternPrefixes.Count) configured prefixes."
}

$sourceGroupNames = @(Get-DLDutyPosSourceGroupNames -Definitions $definitions)
$sourceGroupResult = Get-DLDutyPosSourceGroupMembers -GroupNames $sourceGroupNames
$plans = @(New-DLDutyPosDistributionGroupPlans `
    -Definitions $definitions `
    -AllUsers $allUsers `
    -SourceMembersByGroup $sourceGroupResult.MembersByGroup `
    -MissingGroups $sourceGroupResult.MissingGroups)

Write-Log "DLDutyPos planned $($plans.Count) distribution group update jobs for run $runId."

foreach ($plan in $plans) {
    $queueMessage = Save-DLGroupUpdateJob -Plan $plan -RunId $runId
    Push-OutputBinding -Name GroupUpdateQueue -Value ($queueMessage | ConvertTo-Json -Compress)
    Write-Log "Queued DLDutyPos list '$($queueMessage.Identity)' with $($queueMessage.MemberCount) members using job file $($queueMessage.JobPath)."
}

Write-Log "DLDutyPos planner ended. ------------------------------------------------"

