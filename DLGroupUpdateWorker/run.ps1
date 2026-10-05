# Input bindings are passed in via param block.
param($QueueItem)

. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\..\DLSeniorsCadets\DLSeniorsCadets.Helpers.ps1"

function ConvertFrom-DLQueueMessage {
    param (
        [object]$Message
    )

    if ($null -eq $Message) {
        return $null
    }

    if ($Message -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Message)) {
            return $null
        }

        try {
            return ($Message | ConvertFrom-Json -ErrorAction Stop)
        } catch {
            throw "Queue message was not valid JSON. Message: $($Message | ConvertTo-Json -Compress)"
        }
    }

    return $Message
}

$message = ConvertFrom-DLQueueMessage -Message $QueueItem

if ($null -eq $message) {
    Write-Log "DL group update worker received an empty queue message. This commonly happens when the function is invoked manually from the portal or host APIs; no job will be processed."
    return
}

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
