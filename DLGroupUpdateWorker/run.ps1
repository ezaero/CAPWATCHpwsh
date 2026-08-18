# Input bindings are passed in via param block.
param($QueueItem)

. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\..\DLSeniorsCadets\DLSeniorsCadets.Helpers.ps1"

function ConvertFrom-DLQueueMessage {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Message
    )

    if ($Message -is [string]) {
        return ($Message | ConvertFrom-Json)
    }

    return $Message
}

$message = ConvertFrom-DLQueueMessage -Message $QueueItem

if ([string]::IsNullOrWhiteSpace($message.JobPath)) {
    throw "Queue message did not include a JobPath. Message: $($QueueItem | ConvertTo-Json -Compress)"
}

Write-Log "DL group update worker started for '$($message.Identity)' from run $($message.RunId)."

Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

try {
    Invoke-DLGroupUpdateJob -JobPath $message.JobPath
} catch {
    Write-Log "DL group update worker failed for '$($message.Identity)' using job file '$($message.JobPath)'. Error: $_"
    throw
}

Write-Log "DL group update worker ended for '$($message.Identity)'."
