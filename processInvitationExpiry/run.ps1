# PowerShell Timer Function: Process Invitation Expiry
# Runs every hour to process expired invitations and send additional opportunity notices when pending responses expire
# Trigger Schedule: "0 0 * * * *" (every hour at the top of the hour)

param($Timer)

# Include shared functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Initialize counters
$stats = @{
    checked   = 0
    expired   = 0
    cascaded  = 0
    errors    = 0
}

# Configuration constants
$containerConfig = @{
    invitations       = @{ name = "invitations"; partitionKey = "/eventId" }
    events            = @{ name = "events"; partitionKey = "/airport" }
    invitationQueues  = @{ name = "invitationQueues"; partitionKey = "/eventId" }
}

$retryConfig = @{
    timeoutSeconds = 30
}

function Query-CosmosDbContainer {
    param (
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [string]$Query,
        [switch]$ThrowOnFailure
    )
    
    try {
        # Parse connection string
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }
        
        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']
        
        # Build URI for query
        $uri = "$endpoint/dbs/$Database/colls/$Container/docs"
        
        # Generate auth header
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
            "x-ms-documentdb-query-enable-scan" = "true"
            "x-ms-max-item-count" = "-1"
        }
        
        $queryBody = @{
            query = $Query
        } | ConvertTo-Json
        
        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $queryBody -ContentType "application/query+json" -ErrorAction Stop
        
        return $response.Documents
        
    } catch {
        $errorMessage = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorMessage = "$errorMessage Response: $($_.ErrorDetails.Message)"
        }

        Write-Log "Failed to query Cosmos DB container $Container. Error: $errorMessage"

        if ($ThrowOnFailure) {
            throw
        }

        return @()
    }
}

function Save-CosmosDbItem {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Item,
        [Parameter(Mandatory=$true)]
        [string]$Container,
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase
    )
    
    try {
        # Ensure item has required id field
        if (-not $Item.id) {
            $Item | Add-Member -MemberType NoteProperty -Name "id" -Value ([guid]::NewGuid().ToString()) -ErrorAction SilentlyContinue
        }

        $partitionKeyPath = $containerConfig[$Container].partitionKey
        if (-not $partitionKeyPath) {
            throw "Partition key configuration not found for container '$Container'"
        }

        $partitionPropertyName = $partitionKeyPath.TrimStart('/')
        $partitionKeyValue = $Item.$partitionPropertyName
        if ($null -eq $partitionKeyValue -or [string]::IsNullOrWhiteSpace([string]$partitionKeyValue)) {
            throw "Item $($Item.id) is missing required partition key property '$partitionPropertyName' for container '$Container'"
        }
        
        # Parse connection string
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }
        
        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']
        
        # Build URI for upsert
        $uri = "$endpoint/dbs/$Database/colls/$Container/docs"
        
        # Generate auth header
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
            "Authorization"                  = $authToken
            "x-ms-date"                      = $date
            "x-ms-version"                   = "2020-07-15"
            "x-ms-documentdb-is-upsert"      = "true"
            "x-ms-documentdb-partitionkey"   = "[`"$partitionKeyValue`"]"
        }
        
        $body = $Item | ConvertTo-Json -Depth 10
        
        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
        
        Write-Log "Saved item to Cosmos DB: $($Item.id)"
        return $true
        
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Failed to save item to Cosmos DB: $($Item.id). Error: $errorMessage"
        return $false
    }
}

function Get-CosmosDbItemByQuery {
    param (
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [string]$Query
    )
    
    $results = Query-CosmosDbContainer -ConnectionString $ConnectionString `
                                       -Database $Database `
                                       -Container $Container `
                                       -Query $Query
    
    if ($results -and $results.Count -gt 0) {
        return $results[0]
    }
    return $null
}

function Validate-EnvironmentVariables {
    Write-Log "🔍 Validating environment variables..."

    $required = @("CosmosDbConnectionString", "CosmosDbDatabase")
    $missing = @()

    foreach ($var in $required) {
        $value = Get-Item -Path "env:$var" -ErrorAction SilentlyContinue
        if (-not $value) {
            $missing += $var
        }
    }

    if ($missing.Count -gt 0) {
        throw "Missing required environment variables: $($missing -join ', ')"
    }

    # Validate FRONTEND_URL format if present
    $frontendUrl = $env:FRONTEND_URL
    if ($frontendUrl -and -not $frontendUrl.StartsWith("https://")) {
        throw "Invalid FRONTEND_URL format - must start with https://"
    }

    if (-not $frontendUrl) {
        Write-Log "⚠️ FRONTEND_URL not configured - invitation emails will be skipped"
    }

    Write-Log "✅ All environment variables validated"
}

try {
    Write-Log "⏰ Starting invitation expiry processing..."
    
    # Validate environment early
    Validate-EnvironmentVariables
    
    # Get current time
    $now = Get-Date
    
    # Get Cosmos DB configuration
    $cosmosConfig = Get-CosmosDbConnection
    $cosmosConnectionString = $cosmosConfig.ConnectionString
    $cosmosDatabase = $cosmosConfig.Database
    
    if (-not $cosmosConnectionString -or -not $cosmosDatabase) {
        Write-Log "❌ Error: Cosmos DB configuration incomplete"
        throw "Cosmos DB configuration incomplete"
    }
    
    Write-Log "📡 Cosmos DB Database: $cosmosDatabase"
    
    # Get Graph API access token (needed for sending emails)
    Write-Log "🔐 Acquiring Graph API access token..."
    $graphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
    if (-not $graphAccessToken) {
        throw "Failed to acquire Graph API access token"
    }
    Write-Log "✅ Graph API access token acquired"
    
    # Helper function to send invitation email via Graph API
    function Send-InvitationEmail {
        param(
            [string]$CadetName,
            [string]$CadetEmail,
            [string]$EventName,
            [string]$EventDate,
            [string]$InvitationToken,
            [int]$ExpiryHours,
            [string]$GraphAccessToken
        )
        
        try {
            # Test email override - remove TEST_EMAIL_OVERRIDE when ready for production
            if ($env:TEST_EMAIL_OVERRIDE) {
                Write-Log "🧪 TEST MODE: Redirecting email from $CadetEmail to $($env:TEST_EMAIL_OVERRIDE)"
                $CadetEmail = $env:TEST_EMAIL_OVERRIDE
            }
            
            $frontendUrl = $env:FRONTEND_URL
            if (-not $frontendUrl) {
                Write-Log "⚠️ FRONTEND_URL not configured, skipping email"
                return $false
            }
            
            $invitationUrl = "$frontendUrl/invitation/$InvitationToken"
            $yesUrl = "$frontendUrl/api/invitations/respond?token=$InvitationToken&response=yes"
            $noUrl = "$frontendUrl/api/invitations/respond?token=$InvitationToken&response=no"
            
            # Format event date if it exists
            $eventDateFormatted = if ($EventDate) { [System.DateTime]::Parse($EventDate).ToString("MMMM d, yyyy") } else { "TBD" }
            
            $emailSubject = "🛩️ Orientation Flight Opportunity at $EventName"
            
            # ... email body HTML ... (unchanged)
            
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
      background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
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
      border-left: 4px solid #2563eb;
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
      min-width: 80px;
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
      font-size: 14px;
      color: #666;
      border-top: 1px solid #e0e0e0;
    }
    .footer p {
      margin: 5px 0;
    }
    @media only screen and (max-width: 600px) {
      .button {
        display: block;
        margin: 10px auto;
        max-width: 250px;
      }
      .detail-row {
        flex-direction: column;
      }
      .detail-label {
        margin-bottom: 5px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🛩️ Orientation Flight Opportunity</h1>
      <p>Limited slots are available for this event</p>
    </div>

    <div class="content">
      <div class="greeting">
        Hello $CadetName,
      </div>

      <div class="intro">
        Great news! Based on your priority status, you are receiving an opportunity notice for this Civil Air Patrol orientation flight event. Confirmed slots are assigned as responses come in, and additional "yes" responses are automatically added to the waitlist.
      </div>

      <div class="event-details">
        <h2>Event Details</h2>
        <div class="detail-row">
          <div class="detail-label">Event:</div>
          <div class="detail-value">$EventName</div>
        </div>
        <div class="detail-row">
          <div class="detail-label">Date:</div>
          <div class="detail-value">$eventDateFormatted</div>
        </div>
      </div>

      <div class="expiry-notice">
        <strong>⏰ Time-Sensitive Response Window</strong>
        This response link expires in $ExpiryHours hours. If slots are still open when you respond, you will be confirmed. If the event fills first, your "yes" response will be recorded on the waitlist.
      </div>

      <div class="button-container">
        <a href="$yesUrl" class="button button-yes">✅ YES, I CAN ATTEND</a>
        <a href="$noUrl" class="button button-no">❌ NO, I CANNOT ATTEND</a>
      </div>

      <p style="text-align: center; color: #666; font-size: 14px; margin-top: 20px;">
        Please respond as soon as possible so we can confirm slots and manage the waitlist.
      </p>
    </div>

    <div class="footer">
      <p><strong>Colorado Wing CAP</strong></p>
      <p>COFLICS - Cadet Orientation Flight Information & Coordination System</p>
      <p><a href="$frontendUrl" style="color: #2563eb;">Visit COFLICS</a></p>
    </div>
  </div>
</body>
</html>
"@
            
            $senderEmail = $env:SENDER_EMAIL -or "noreply@orientationflight.cowg.cap.gov"
            
            # Build Graph API email payload
            $emailPayload = @{
                message = @{
                    subject = $emailSubject
                    body    = @{
                        contentType = "HTML"
                        content     = $emailBody
                    }
                    toRecipients = @(
                        @{
                            emailAddress = @{
                                address = $CadetEmail
                                name    = $CadetName
                            }
                        }
                    )
                    from = @{
                        emailAddress = @{
                            address = $senderEmail
                            name    = "Colorado Wing CAP"
                        }
                    }
                }
                saveToSentItems = $true
            } | ConvertTo-Json -Depth 10
            
            Write-Log "📧 Sending invitation email to $CadetEmail via Graph API"
            
            $graphApiUrl = "https://graph.microsoft.com/v1.0/me/sendMail"
            $headers = @{
                "Authorization" = "Bearer $GraphAccessToken"
                "Content-Type"  = "application/json"
            }
            
            $response = Invoke-RestMethod -Uri $graphApiUrl -Method POST -Headers $headers -Body $emailPayload -ErrorAction Stop -TimeoutSec $retryConfig.timeoutSeconds
            
            Write-Log "✅ Email sent to $CadetEmail"
            return $true
            
        } catch {
            Write-Log "⚠️ Error sending email to $CadetEmail : $_"
            return $false
        }
    }

    
    # Query all pending invitations
    Write-Log "🔍 Querying pending invitations..."
    $query = "SELECT * FROM c WHERE c.status = 'pending'"
    $invitations = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                           -Database $cosmosDatabase `
                                           -Container $containerConfig.invitations.name `
                                           -Query $query `
                                           -ThrowOnFailure
    $stats.checked = $invitations.Count
    
    Write-Log "📊 Found $($invitations.Count) pending invitation(s) to check"
    
    if ($invitations.Count -eq 0) {
        $diagnosticInvitations = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                                         -Database $cosmosDatabase `
                                                         -Container $containerConfig.invitations.name `
                                                         -Query "SELECT TOP 10 c.id, c.status, c.eventId, c.expiresAt FROM c"

        if ($diagnosticInvitations.Count -gt 0) {
            $diagnosticSummary = $diagnosticInvitations | ForEach-Object {
                $status = if ($_.status) { $_.status } else { "<null>" }
                $eventId = if ($_.eventId) { $_.eventId } else { "<no-event>" }
                $expiresAt = if ($_.expiresAt) { $_.expiresAt } else { "<no-expiry>" }
                "{0} status={1} eventId={2} expiresAt={3}" -f $_.id, $status, $eventId, $expiresAt
            }
            Write-Log "ℹ️ Invitation diagnostics (top 10 docs): $($diagnosticSummary -join ' | ')"
        } else {
            Write-Log "ℹ️ Invitation diagnostics: invitations container returned no documents at all"
        }

        Write-Log "✅ No pending invitations to process"
        Write-Log @"
📊 Invitation Expiry Processing Summary:
   - Checked: $($stats.checked)
   - Expired: $($stats.expired)
   - Cascaded: $($stats.cascaded)
   - Errors: $($stats.errors)
"@
        return
    }
    
    # Process each pending invitation
    foreach ($invitation in $invitations) {
        try {
            $expiresAt = [System.DateTime]::Parse($invitation.expiresAt)
            
            # Check if expired
            if ($now -gt $expiresAt) {
                Write-Log "⏱️ Invitation $($invitation.id) has expired (cadet: $($invitation.cadetName))"
                
                # Mark as expired
                $invitation.status = "expired"
                $invitation.respondedAt = $now.ToUniversalTime().ToString("o")
                $invitationSaved = Save-CosmosDbItem -Item $invitation -Container $containerConfig.invitations.name -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
                if (-not $invitationSaved) {
                    $stats.errors++
                    Write-Log "❌ Failed to persist expired status for invitation $($invitation.id); skipping cascade"
                    continue
                }

                $stats.expired++
                Write-Log "✅ Marked invitation $($invitation.id) as expired"
                
                # Query event to check auto-fill status
                $eventQuery = "SELECT * FROM c WHERE c.id = '$($invitation.eventId)'"
                $event = Get-CosmosDbItemByQuery -ConnectionString $cosmosConnectionString `
                                                 -Database $cosmosDatabase `
                                                 -Container $containerConfig.events.name `
                                                 -Query $eventQuery
                
                if (-not $event) {
                    Write-Log "⚠️ Event $($invitation.eventId) not found for expired invitation"
                    continue
                }
                $openSlots = @($event.slots | Where-Object { -not $_.cadetId }).Count
                if ($openSlots -eq 0) {
                    Write-Log "ℹ️ Event $($event.id) is full, skipping cascade"
                    continue
                }
                
                # Check if auto-fill is enabled (default to true)
                if ($event.autoFillEnabled -eq $false) {
                    Write-Log "ℹ️ Auto-fill disabled for event $($event.id), skipping cascade"
                    continue
                }
                
                # Check if event is still scheduled
                if ($event.status -ne "scheduled") {
                    Write-Log "ℹ️ Event $($event.id) status is $($event.status), skipping cascade"
                    continue
                }
                
                # Trigger cascade to next cadet
                Write-Log "🔄 Triggering cascade for event $($event.id)..."
                
                # Query invitation queue for this event
                $queueQuery = "SELECT * FROM c WHERE c.eventId = '$($event.id)' AND c.status = 'active'"
                $queue = Get-CosmosDbItemByQuery -ConnectionString $cosmosConnectionString `
                                                 -Database $cosmosDatabase `
                                                 -Container $containerConfig.invitationQueues.name `
                                                 -Query $queueQuery
                
                if (-not $queue) {
                    Write-Log "⚠️ No active invitation queue found for event $($event.id)"
                    continue
                }
                $nextCadet = $queue.cadets | Where-Object { -not $_.invitationSent -and -not $_.removed } | Select-Object -First 1
                
                if (-not $nextCadet) {
                    Write-Log "ℹ️ No more cadets in queue for event $($event.id)"
                    $queue.status = "completed"
                    # Update queue status
                    $queueSaved = Save-CosmosDbItem -Item $queue -Container $containerConfig.invitationQueues.name -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
                    if (-not $queueSaved) {
                        $stats.errors++
                        Write-Log "❌ Failed to mark invitation queue $($queue.id) as completed"
                    }
                    continue
                }
                
                try {
                    # Create invitation for next cadet
                    $expiryHours = $event.invitationExpiryHours -or 24
                    $invitationToken = [guid]::NewGuid().ToString("N") + [guid]::NewGuid().ToString("N")
                    $position = ($queue.cadets.IndexOf($nextCadet)) + 1
                    
                    $nowCascade = Get-Date
                    $expiresAtCascade = $nowCascade.AddHours($expiryHours)
                    
                    $newInvitation = @{
                        id                = [guid]::NewGuid().ToString()
                        eventId           = $event.id
                        cadetId           = $nextCadet.cadetId
                        cadetEmail        = $nextCadet.email
                        cadetName         = $nextCadet.name
                        cadetCapid        = $nextCadet.capid
                        cadetSquadron     = $nextCadet.squadron
                        token             = $invitationToken
                        status            = "pending"
                        sentAt            = $nowCascade.ToUniversalTime().ToString("o")
                        expiresAt         = $expiresAtCascade.ToUniversalTime().ToString("o")
                        respondedAt       = $null
                        position          = $position
                        priorityScore     = $nextCadet.priorityScore
                        priorityTier      = $nextCadet.priorityTier
                        createdAt         = $nowCascade.ToUniversalTime().ToString("o")
                    }
                    
                    # Save new invitation
                    $newInvitationSaved = Save-CosmosDbItem -Item $newInvitation -Container $containerConfig.invitations.name -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
                    if (-not $newInvitationSaved) {
                        throw "Failed to save new invitation $($newInvitation.id)"
                    }
                    
                    # Mark cadet as invited in queue
                    $nextCadet.invitationSent = $true
                    $nextCadet.invitationId = $newInvitation.id
                    $queue.updatedAt = $nowCascade.ToUniversalTime().ToString("o")
                    
                    # Update queue
                    $queueUpdated = Save-CosmosDbItem -Item $queue -Container $containerConfig.invitationQueues.name -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
                    if (-not $queueUpdated) {
                        throw "Failed to update invitation queue $($queue.id) after creating invitation $($newInvitation.id)"
                    }
                    
                    # Send invitation email
                    $emailSent = Send-InvitationEmail -CadetName $nextCadet.name -CadetEmail $nextCadet.email -EventName $event.name -EventDate $event.eventDate -InvitationToken $invitationToken -ExpiryHours $expiryHours -GraphAccessToken $graphAccessToken
                    
                    if ($emailSent) {
                        $stats.cascaded++
                        Write-Log "✅ Cascade successful: Invited $($nextCadet.name) ($($nextCadet.email))"
                    } else {
                        Write-Log "⚠️ Invitation created but email sending may have failed for $($nextCadet.name)"
                    }
                    
                } catch {
                    Write-Log "❌ Error cascading to next cadet: $_"
                }
                
            }
        } catch {
            $stats.errors++
            Write-Log "❌ Error processing invitation $($invitation.id): $_"
        }
    }
    
    # Log summary
    Write-Log @"
📊 Invitation Expiry Processing Summary:
   - Checked: $($stats.checked)
   - Expired: $($stats.expired)
   - Cascaded: $($stats.cascaded)
   - Errors: $($stats.errors)
"@

} catch {
    Write-Log "❌ FATAL ERROR in invitation expiry processing: $_"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)"
    throw
}
