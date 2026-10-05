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

$script:WrittenLogs = @()
function Write-Log {
    param ([string]$Message)
    $script:WrittenLogs += $Message
}

$sharedPath = Join-Path $PSScriptRoot ".." "shared" "shared.ps1"
$dlHelpersPath = Join-Path $PSScriptRoot ".." "DLSeniorsCadets" "DLSeniorsCadets.Helpers.ps1"
$workerPath = Join-Path $PSScriptRoot "run.ps1"
$script = Get-Content -Path $workerPath -Raw
$script = $script -replace [regex]::Escape('. "$PSScriptRoot\..\shared\shared.ps1"'), "# shared mocked"
$script = $script -replace [regex]::Escape('. "$PSScriptRoot\..\DLSeniorsCadets\DLSeniorsCadets.Helpers.ps1"'), "# helpers mocked"

$QueueItem = ""
. ([scriptblock]::Create($script))

Assert-Equal (($script:WrittenLogs -join "`n") -like "*received an empty queue message*") $true "Empty worker invocations should log and return without throwing."

$message = ConvertFrom-DLQueueMessage -Message '{"JobPath":"C:\\jobs\\one.json","Identity":"CO-200 Seniors","RunId":"run-1"}'
Assert-Equal $message.JobPath "C:\jobs\one.json" "Valid queue JSON should parse into an object."

$malformedFailed = $false
try {
    $null = ConvertFrom-DLQueueMessage -Message "not-json"
} catch {
    $malformedFailed = "$_" -like "*not valid JSON*"
}
Assert-Equal $malformedFailed $true "Malformed nonblank queue messages should produce a clear JSON error."

Write-Host "DLGroupUpdateWorker.Tests.ps1 passed"
