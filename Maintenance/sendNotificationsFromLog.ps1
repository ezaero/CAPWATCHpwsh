param(
    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [string]$CapwatchDataDir = "$($env:HOME)\data\CAPWatch",

    [switch]$Execute
)

. "$PSScriptRoot\..\shared\shared.ps1"

if (-not (Test-Path -Path $LogPath)) {
    throw "LogPath not found: $LogPath"
}

$memberPath = Join-Path -Path $CapwatchDataDir -ChildPath "Member.txt"

$logLines = Get-Content -Path $LogPath
$memberRows = @()
if (Test-Path -Path $memberPath) {
    $memberRows = Import-Csv -Path $memberPath -ErrorAction Stop
} else {
    Write-Log "CAPWATCH Member.txt not found at $memberPath. Replay will use units present in the log and mark other units as Unknown."
}
$deletedMembers = ConvertFrom-MaintenanceDeletionLog -LogLines $logLines -MemberRows $memberRows

if ($deletedMembers.Count -eq 0) {
    Write-Log "No deleted account entries found in $LogPath."
    return
}

Write-Log "Reconstructed $($deletedMembers.Count) deleted account notification entries from $LogPath."
$deletedMembers |
    Group-Object -Property Unit |
    Sort-Object Name |
    ForEach-Object {
        Write-Log "Replay notification group: Unit $($_.Name) - $($_.Count) members"
    }

if ($Execute) {
    $env:EXECUTE = "true"
}

if ((Test-ExecutionMode) -and ($deletedMembers | Where-Object { $_.Unit -eq "Unknown" } | Select-Object -First 1)) {
    throw "Cannot execute replay while any deleted entries have Unknown unit. Provide a CapwatchDataDir containing Member.txt so units can be resolved."
}

if (-not (Test-ExecutionMode)) {
    Write-Log "[DRY-RUN] Would send replay notifications. Set EXECUTE=true or pass -Execute to send."
    return
}

$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome

$allUsers = GetAllUsers
Send-ExpiredMembersNotification -deletedMembers $deletedMembers -allUsers $allUsers
