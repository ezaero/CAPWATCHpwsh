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

    $wingOrganization = $OrganizationRows |
        Where-Object {
            $_.Wing -eq $env:WING_DESIGNATOR -and
            $_.Status -eq "ACTIVE" -and
            $_.Scope -eq "UNIT"
        } |
        Sort-Object Unit -Unique |
        Select-Object Unit, Name

    return @(
        $wingOrganization |
            Where-Object { $_.Unit -ne "000" -and $_.Unit -ne "999" -and $_.Unit -ne "001" }
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
