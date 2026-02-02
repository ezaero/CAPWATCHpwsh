<#
.SYNOPSIS
    Timer-triggered function to escalate pilot invitations for events with unfilled slots 24+ hours after creation

.DESCRIPTION
    This script:
    1. Runs every 6 hours
    2. Queries Cosmos DB for events created 24+ hours ago with unfilled pilot slots
    3. Expires all pending pilot invitations before escalation
    4. Sends urgent escalation emails to all qualified orientation pilots
    5. Updates event escalation status to prevent duplicate escalations
    6. Logs all actions for auditing

.NOTES
    Requires Cosmos DB and Microsoft Graph API configuration in environment variables
#>

param($Timer)

# Include shared functions
. "$PSScriptRoot\..\shared\shared.ps1"

$logPrefix = '🚨 [escalatePilotInvitations]'

Write-Log "$logPrefix Timer trigger function started at $(Get-Date -Format o)"

# Initialize statistics
$stats = @{
    EventsChecked = 0
    EventsEligible = 0
    InvitationsExpired = 0
    InvitationsSent = 0
    Errors = 0
}

# Helper Functions
function Query-CosmosDbContainer {
    param (
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [string]$Query
    )

    try {
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        $uri = "$endpoint/dbs/$Database/colls/$Container/docs"

        $verb = "post"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/$Container"
        $date = [DateTime]::UtcNow.ToString('r')

        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)

        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

        $headers = @{
            "Authorization" = $authToken
            "x-ms-date" = $date
            "x-ms-version" = "2020-07-15"
            "x-ms-documentdb-isquery" = "true"
            "x-ms-documentdb-query-enablecrosspartition" = "true"
        }

        $queryBody = @{
            query = $Query
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $queryBody -ContentType "application/query+json" -ErrorAction Stop

        return $response.Documents

    } catch {
        Write-Log "$logPrefix Failed to query Cosmos DB container $Container. Error: $($_.Exception.Message)"
        return @()
    }
}

function Update-CosmosDbDocument {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Document,
        [Parameter(Mandatory=$true)]
        [string]$Container,
        [Parameter(Mandatory=$true)]
        [array]$PartitionKeyValues,
        [string]$ConnectionString,
        [string]$Database
    )

    try {
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        $uri = "$endpoint/dbs/$Database/colls/$Container/docs/$($Document.id)"

        $verb = "put"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/$Container/docs/$($Document.id)"
        $date = [DateTime]::UtcNow.ToString('r')

        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)

        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

        # Build hierarchical partition key
        $partitionKeyJson = ($PartitionKeyValues | ForEach-Object { "`"$_`"" }) -join ","

        $headers = @{
            "Authorization" = $authToken
            "x-ms-date" = $date
            "x-ms-version" = "2020-07-15"
            "x-ms-documentdb-partitionkey" = "[$partitionKeyJson]"
        }

        $body = $Document | ConvertTo-Json -Depth 10

        $response = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop

        return $response

    } catch {
        Write-Log "$logPrefix Failed to update Cosmos DB document $($Document.id). Error: $($_.Exception.Message)"
        throw
    }
}

function Save-CosmosDbPilotInvitation {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Document,
        [string]$ConnectionString,
        [string]$Database,
        [string]$PartitionKeyValue
    )

    try {
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        $uri = "$endpoint/dbs/$Database/colls/pilotInvitations/docs"

        $verb = "post"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/pilotInvitations"
        $date = [DateTime]::UtcNow.ToString('r')

        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)

        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

        $headers = @{
            "Authorization" = $authToken
            "x-ms-date" = $date
            "x-ms-version" = "2020-07-15"
            "x-ms-documentdb-partitionkey" = "[`"$PartitionKeyValue`"]"
        }

        $body = $Document | ConvertTo-Json -Depth 10

        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop

        return $response

    } catch {
        Write-Log "$logPrefix Failed to save pilot invitation. Error: $($_.Exception.Message)"
        throw
    }
}

function New-SecureToken {
    $bytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($bytes)
    return [BitConverter]::ToString($bytes).Replace('-', '').ToLower()
}

function Send-PilotEscalationEmail {
    param(
        [string]$ToAddress,
        [string]$PilotName,
        [hashtable]$Event,
        [string]$InvitationToken,
        [int]$ExpiryHours = 48
    )

    try {
        # Test email override
        if ($env:TEST_EMAIL_OVERRIDE) {
            Write-Log "$logPrefix 🧪 TEST MODE: Redirecting email from $ToAddress to $($env:TEST_EMAIL_OVERRIDE)"
            $ToAddress = $env:TEST_EMAIL_OVERRIDE
        }

        $baseUrl = $env:FRONTEND_URL
        if (-not $baseUrl) {
            $baseUrl = "https://orientationflights.cowg.cap.gov"
        }

        $yesUrl = "$baseUrl/api/pilot-invitations/respond?token=$InvitationToken&response=yes"
        $noUrl = "$baseUrl/api/pilot-invitations/respond?token=$InvitationToken&response=no"

        $subject = "URGENT: Pilot Still Needed - O-Flight Event at $($Event.airport) on $($Event.date)"

        $aircraftInfo = if ($Event.aircraft -and $Event.aircraft.Count -gt 0) {
            ($Event.aircraft | ForEach-Object { "$($_.tailNumber) ($($_.model))" }) -join ", "
        } else {
            "TBD"
        }

        $timeInfo = if ($Event.time) { $Event.time } else { "TBD" }

        $coordinatorInfo = if ($Event.coordinatorName) {
            @"
        <tr>
          <td style="padding: 20px 0; border-top: 1px solid #e0e0e0;">
            <p style="margin: 0 0 10px 0; font-size: 14px; color: #666;">
              <strong>Event Coordinator:</strong><br/>
              $($Event.coordinatorName)$(if ($Event.coordinatorPhone) { " | $($Event.coordinatorPhone)" })$(if ($Event.coordinatorEmail) { " | <a href='mailto:$($Event.coordinatorEmail)' style='color: #2563eb;'>$($Event.coordinatorEmail)</a>" })
            </p>
          </td>
        </tr>
"@
        } else {
            ""
        }

        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      margin: 0;
      padding: 0;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 600px;
      margin: 20px auto;
      background-color: #ffffff;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    }
    .header {
      background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%);
      color: white;
      padding: 30px 20px;
      text-align: center;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .header p {
      margin: 10px 0 0 0;
      font-size: 16px;
      opacity: 0.9;
    }
    .urgent-banner {
      background-color: #fef3c7;
      border-left: 5px solid #f59e0b;
      padding: 20px;
      margin: 0;
    }
    .urgent-banner h2 {
      margin: 0 0 10px 0;
      font-size: 20px;
      color: #92400e;
    }
    .urgent-banner p {
      margin: 0;
      font-size: 15px;
      color: #78350f;
    }
    .content {
      padding: 30px 20px;
    }
    .greeting {
      font-size: 18px;
      margin-bottom: 20px;
      color: #1a1a1a;
    }
    .intro {
      font-size: 16px;
      margin-bottom: 30px;
      color: #555;
    }
    .event-details {
      background-color: #f8f9fa;
      border-left: 4px solid #dc2626;
      padding: 20px;
      margin: 20px 0;
      border-radius: 4px;
    }
    .event-details h2 {
      margin: 0 0 15px 0;
      font-size: 18px;
      color: #1a1a1a;
    }
    .detail-row {
      display: flex;
      margin: 10px 0;
      font-size: 15px;
    }
    .detail-label {
      font-weight: 600;
      min-width: 100px;
      color: #666;
    }
    .detail-value {
      color: #1a1a1a;
    }
    .expiry-notice {
      background-color: #fff3cd;
      border-left: 4px solid #ffc107;
      padding: 15px;
      margin: 20px 0;
      border-radius: 4px;
      font-size: 14px;
      color: #856404;
    }
    .expiry-notice strong {
      display: block;
      margin-bottom: 5px;
    }
    .button-container {
      text-align: center;
      margin: 30px 0;
      padding: 20px 0;
    }
    .button {
      display: inline-block;
      padding: 16px 40px;
      margin: 10px;
      font-size: 16px;
      font-weight: 600;
      text-decoration: none;
      border-radius: 6px;
      transition: all 0.3s ease;
    }
    .button-yes {
      background-color: #28a745;
      color: white !important;
    }
    .button-yes:hover {
      background-color: #218838;
    }
    .button-no {
      background-color: #6c757d;
      color: white !important;
    }
    .button-no:hover {
      background-color: #5a6268;
    }
    .footer {
      background-color: #f8f9fa;
      padding: 20px;
      text-align: center;
      font-size: 13px;
      color: #666;
      border-top: 1px solid #e0e0e0;
    }
    .footer p {
      margin: 5px 0;
    }
    .footer a {
      color: #2563eb;
      text-decoration: none;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>✈️ Urgent: Pilot Needed</h1>
      <p>Orientation Flight Event</p>
    </div>

    <div class="urgent-banner">
      <h2>⚠️ We Still Need Your Help!</h2>
      <p>This event was created over 24 hours ago and we still haven't secured enough pilots. We're reaching out to all qualified orientation pilots to help make this event happen for our cadets.</p>
    </div>

    <div class="content">
      <div class="greeting">
        Hello $PilotName,
      </div>

      <div class="intro">
        We urgently need qualified pilots for an Orientation Flight Event at <strong>$($Event.airport)</strong>. Despite our initial outreach, we still have unfilled pilot positions and need your support to give our cadets this valuable flying experience.
      </div>

      <div class="event-details">
        <h2>Event Details</h2>
        <div class="detail-row">
          <div class="detail-label">📍 Location:</div>
          <div class="detail-value">$($Event.airport)</div>
        </div>
        <div class="detail-row">
          <div class="detail-label">📅 Date:</div>
          <div class="detail-value">$($Event.date)</div>
        </div>
        <div class="detail-row">
          <div class="detail-label">🕐 Time:</div>
          <div class="detail-value">$timeInfo</div>
        </div>
        <div class="detail-row">
          <div class="detail-label">✈️ Aircraft:</div>
          <div class="detail-value">$aircraftInfo</div>
        </div>
        <div class="detail-row">
          <div class="detail-label">👥 Pilots Needed:</div>
          <div class="detail-value">$($Event.numberOfPilotsRequired)</div>
        </div>
      </div>

      <div class="expiry-notice">
        <strong>⏰ First Come, First Served</strong>
        This escalation invitation has been sent to all qualified orientation pilots. The first pilot(s) to accept will be assigned to this flight. Please respond as soon as possible if you can help. This invitation expires in $ExpiryHours hours.
      </div>

      <div class="button-container">
        <a href="$yesUrl" class="button button-yes">
          ✓ Yes, I'm Available
        </a>
        <br/>
        <a href="$noUrl" class="button button-no">
          ✗ Not Available
        </a>
      </div>

      <table width="100%" cellpadding="0" cellspacing="0">
        $coordinatorInfo
      </table>

      <p style="font-size: 14px; color: #666; margin-top: 30px;">
        This urgent invitation was sent to all qualified orientation pilots because we need additional support for this event. Thank you for your dedication to the cadet program!
      </p>
    </div>

    <div class="footer">
      <p><strong>COFLICS</strong></p>
      <p>Cadet Orientation Flight Information & Coordination System</p>
      <p><a href="$baseUrl">Visit COFLICS</a></p>
      <p style="margin-top: 10px; color: #999; font-size: 12px;">
        Colorado Wing Civil Air Patrol
      </p>
    </div>
  </div>
</body>
</html>
"@

        $fromAddress = $env:LOG_EMAIL_FROM_ADDRESS
        if (-not $fromAddress) {
            $fromAddress = "OFlights@cowg.cap.gov"
        }

        $emailPayload = @{
            message = @{
                subject = $subject
                body = @{
                    contentType = "HTML"
                    content = $emailBody
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $ToAddress
                        }
                    }
                )
            }
            saveToSentItems = $true
        } | ConvertTo-Json -Depth 10

        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$fromAddress/sendMail" -Body $emailPayload -ContentType "application/json"

        Write-Log "$logPrefix ✅ Email sent to $ToAddress"

    } catch {
        Write-Log "$logPrefix ❌ Failed to send email to ${ToAddress}: $($_.Exception.Message)"
        throw
    }
}

try {
    # Authenticate to Microsoft Graph using managed identity
    try {
        $MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
        Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
        Write-Log "$logPrefix Successfully authenticated to Microsoft Graph"
    } catch {
        Write-Log "$logPrefix Failed to authenticate to Microsoft Graph. Error: $($_.Exception.Message)"
        throw
    }

    # Get Cosmos DB configuration
    $cosmosConfig = Get-CosmosDbConnection
    $cosmosConnectionString = $cosmosConfig.ConnectionString
    $cosmosDatabase = $cosmosConfig.Database

    if (-not $cosmosConnectionString -or -not $cosmosDatabase) {
        Write-Log "$logPrefix Error: Cosmos DB configuration incomplete"
        throw "Cosmos DB configuration incomplete"
    }

    Write-Log "$logPrefix Cosmos DB Database: $cosmosDatabase"

    # Calculate 24 hours ago timestamp
    $escalationThreshold = (Get-Date).AddHours(-24).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    Write-Log "$logPrefix Escalation threshold: Events created before $escalationThreshold"

    # Query for events needing escalation
    Write-Log "$logPrefix Querying events created 24+ hours ago with unfilled pilot slots..."

    $eventsQuery = @"
SELECT * FROM c
WHERE c.createdAt < '$escalationThreshold'
  AND c.status = 'scheduled'
  AND (IS_NULL(c.escalationStatus) OR c.escalationStatus.initialInvitationsSent = false OR NOT IS_DEFINED(c.escalationStatus.initialInvitationsSent))
  AND c.numberOfPilotsRequired > 0
ORDER BY c.createdAt ASC
"@

    $events = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                     -Database $cosmosDatabase `
                                     -Container "events" `
                                     -Query $eventsQuery

    $stats.EventsChecked = $events.Count
    Write-Log "$logPrefix Found $($events.Count) event(s) to check"

    if ($events.Count -eq 0) {
        Write-Log "$logPrefix ✅ No events require escalation at this time"
        return
    }

    # Get all orientation pilots
    Write-Log "$logPrefix Fetching all Orientation Pilots..."
    $pilotsQuery = 'SELECT * FROM c WHERE ARRAY_CONTAINS(c.roles, "pilot")'
    $allPilots = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                        -Database $cosmosDatabase `
                                        -Container "users" `
                                        -Query $pilotsQuery

    Write-Log "$logPrefix Found $($allPilots.Count) pilot(s) in database"

    # Process each event
    foreach ($event in $events) {
        Write-Log "$logPrefix ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Write-Log "$logPrefix 🎯 Processing Event: $($event.id)"
        Write-Log "$logPrefix    Airport: $($event.airport) | Date: $($event.date) | Time: $($event.time)"
        Write-Log "$logPrefix    Created: $($event.createdAt)"

        # Check pilot slots
        $pilotSlots = $event.pilotSlots
        $totalSlotsRequired = $event.numberOfPilotsRequired
        $filledSlots = ($pilotSlots | Where-Object { $null -ne $_.pilotId -and $_.pilotId -ne "" }).Count
        $unfilledSlots = $totalSlotsRequired - $filledSlots

        Write-Log "$logPrefix    Pilot Slots: $filledSlots / $totalSlotsRequired filled ($unfilledSlots unfilled)"

        if ($unfilledSlots -le 0) {
            Write-Log "$logPrefix    ✅ All pilot slots filled - skipping"
            continue
        }

        $stats.EventsEligible++

        # Expire all pending invitations before sending escalation
        $pendingInvitationsQuery = "SELECT * FROM c WHERE c.eventId = '$($event.id)' AND c.status = 'pending'"
        $pendingInvitations = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                                      -Database $cosmosDatabase `
                                                      -Container "pilotInvitations" `
                                                      -Query $pendingInvitationsQuery

        if ($pendingInvitations.Count -gt 0) {
            Write-Log "$logPrefix    ⏰ Expiring $($pendingInvitations.Count) pending invitation(s)..."

            foreach ($invitation in $pendingInvitations) {
                try {
                    $invitation.status = "expired"
                    $invitation.respondedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                    $invitation.expiryReason = "Event escalated to all pilots after 24 hours"

                    Update-CosmosDbDocument -Document $invitation `
                                          -Container "pilotInvitations" `
                                          -PartitionKeyValues @($event.id) `
                                          -ConnectionString $cosmosConnectionString `
                                          -Database $cosmosDatabase

                    $stats.InvitationsExpired++
                    Write-Log "$logPrefix       ✅ Expired invitation for pilot: $($invitation.pilotName)"
                } catch {
                    Write-Log "$logPrefix       ⚠️ Failed to expire invitation for $($invitation.pilotName): $($_.Exception.Message)"
                }
            }

            Write-Log "$logPrefix    ✅ Pending invitations expired"
        }

        # Filter pilots: exclude those already assigned
        $assignedPilotIds = $pilotSlots | Where-Object { $null -ne $_.pilotId } | ForEach-Object { $_.pilotId }
        $eligiblePilots = $allPilots | Where-Object {
            $_.id -notin $assignedPilotIds -and
            $null -ne $_.email -and
            $_.email -ne ""
        }

        Write-Log "$logPrefix    Eligible for escalation: $($eligiblePilots.Count) pilot(s)"

        if ($eligiblePilots.Count -eq 0) {
            Write-Log "$logPrefix    ⚠️ No eligible pilots to invite"
            continue
        }

        # Send invitation to each eligible pilot
        foreach ($pilot in $eligiblePilots) {
            Write-Log "$logPrefix       📨 Sending escalation invitation to: $($pilot.name) ($($pilot.email))"

            try {
                # Generate secure token
                $token = New-SecureToken

                # Create invitation record
                $invitation = @{
                    id = (New-Guid).ToString()
                    eventId = $event.id
                    pilotId = $pilot.id
                    pilotEmail = $pilot.email
                    pilotName = $pilot.name
                    pilotCapid = $pilot.capid
                    pilotPhone = $pilot.phone
                    token = $token
                    status = "pending"
                    sentAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                    expiresAt = (Get-Date).AddHours(48).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                    respondedAt = $null
                    isEscalation = $true
                    escalationType = "24h-post-creation"
                    createdAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                }

                # Save invitation to Cosmos DB
                Save-CosmosDbPilotInvitation -Document $invitation `
                                            -ConnectionString $cosmosConnectionString `
                                            -Database $cosmosDatabase `
                                            -PartitionKeyValue $event.id

                # Send email
                Send-PilotEscalationEmail -ToAddress $pilot.email `
                                         -PilotName $pilot.name `
                                         -Event $event `
                                         -InvitationToken $token `
                                         -ExpiryHours 48

                $stats.InvitationsSent++
                Write-Log "$logPrefix          ✅ Invitation sent and recorded"
            } catch {
                $stats.Errors++
                Write-Log "$logPrefix          ❌ Error: $($_.Exception.Message)"
            }
        }

        # Update event escalation status
        Write-Log "$logPrefix    📝 Updating event escalation status..."

        if ($null -eq $event.escalationStatus) {
            $event.escalationStatus = @{
                initialInvitationsSent = $true
                escalation24hSent = $false
                escalation48hSent = $false
                cancellationNotified = $false
            }
        } else {
            $event.escalationStatus.initialInvitationsSent = $true
        }

        $event.updatedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

        Update-CosmosDbDocument -Document $event `
                              -Container "events" `
                              -PartitionKeyValues @($event.airport, $event.date) `
                              -ConnectionString $cosmosConnectionString `
                              -Database $cosmosDatabase

        Write-Log "$logPrefix    ✅ Event updated"
        Write-Log ""
    }

    # Log summary
    Write-Log "$logPrefix ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Write-Log "$logPrefix 📊 Escalation Summary"
    Write-Log "$logPrefix    Events Checked: $($stats.EventsChecked)"
    Write-Log "$logPrefix    Events Eligible for Escalation: $($stats.EventsEligible)"
    Write-Log "$logPrefix    Original Invitations Expired: $($stats.InvitationsExpired)"
    Write-Log "$logPrefix    Escalation Invitations Sent: $($stats.InvitationsSent)"
    Write-Log "$logPrefix    Errors: $($stats.Errors)"
    Write-Log "$logPrefix ✅ Pilot Invitation Escalation Process Complete!"

} catch {
    Write-Log "$logPrefix ❌ Error in timer trigger: $($_.Exception.Message)"
    Write-Log "$logPrefix Stack Trace: $($_.ScriptStackTrace)"
    throw
}
