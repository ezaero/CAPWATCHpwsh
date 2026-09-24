# Input bindings are passed in via param block.
param($Timer)

. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\DLSeniorsCadets.Helpers.ps1"

$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
$OrganizationFile = "$CAPWATCHDATADIR/Organization.txt"

function Get-DLUnits {
    param (
        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows
    )

    return @(
        Get-CapwatchActiveWingUnits -OrganizationRows $OrganizationRows |
            Select-Object Unit, Name
    )
}

Write-Log "Squadron Seniors/Cadets planner started. ------------------------------------------------"

Push-Location $CAPWATCHDATADIR
try {
    $msGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
    Connect-MgGraph -AccessToken $msGraphAccessToken -NoWelcome

    $runId = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
    $organizationRows = @(Import-Csv -Path $OrganizationFile)
    $unitList = Get-DLUnits -OrganizationRows $organizationRows
    $allUsers = GetAllUsers
    $plans = @(New-DLGroupUpdatePlans -UnitList $unitList -AllUsers $allUsers -OrganizationRows $organizationRows)

    Write-Log "Planner created $($plans.Count) distribution group update jobs for run $runId."

    foreach ($plan in $plans) {
        $queueMessage = Save-DLGroupUpdateJob -Plan $plan -RunId $runId
        Push-OutputBinding -Name GroupUpdateQueue -Value ($queueMessage | ConvertTo-Json -Compress)
        Write-Log "Queued '$($queueMessage.Identity)' ($($queueMessage.MemberCount) members) using job file $($queueMessage.JobPath)."
    }
} finally {
    Pop-Location
}

Write-Log "Squadron Seniors/Cadets planner ended. ------------------------------------------------"
