<#
.SYNOPSIS
    Timer-triggered function to send orientation flight reminders

.DESCRIPTION
    This script:
    1. Runs daily at 6 PM MST (01:00 UTC next day)
    2. Queries Cosmos DB for flights and events scheduled for tomorrow
    3. Sends reminder emails to cadets assigned to flights/events
    4. Records notifications to prevent duplicate reminders
    5. Logs all actions for auditing

.NOTES
    Requires Cosmos DB and Microsoft Graph API configuration in environment variables
#>

param($myTimer)

# Include shared functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Define helper functions
function Query-CosmosDbContainer {
    param (
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [string]$Query
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
        }
        
        $queryBody = @{
            query = $Query
        } | ConvertTo-Json
        
        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $queryBody -ContentType "application/query+json" -ErrorAction Stop
        
        return $response.Documents
        
    } catch {
        Write-Log "Failed to query Cosmos DB container $Container. Error: $($_.Exception.Message)"
        return @()
    }
}

function Save-NotificationItem {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Item,
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase
    )
    
    try {
        # Ensure item has required id field
        if (-not $Item.id) {
            $Item | Add-Member -MemberType NoteProperty -Name "id" -Value ([guid]::NewGuid().ToString()) -ErrorAction SilentlyContinue
        }
        
        # Add timestamp if not present
        if (-not $Item.timestamp) {
            $Item | Add-Member -MemberType NoteProperty -Name "timestamp" -Value (Get-Date -Format o) -ErrorAction SilentlyContinue
        }
        
        # Parse connection string
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }
        
        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']
        
        # Build URI for upsert (notifications container)
        $uri = "$endpoint/dbs/$Database/colls/notifications/docs"
        
        # Generate auth header
        $verb = "post"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/notifications"
        $date = [DateTime]::UtcNow.ToString('r')
        
        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"
        
        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)
        
        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)
        
        # Use userId as partition key (notifications container uses /userId)
        $partitionKeyValue = $Item.userId
        
        $headers = @{
            "Authorization"                  = $authToken
            "x-ms-date"                      = $date
            "x-ms-version"                   = "2020-07-15"
            "x-ms-documentdb-is-upsert"      = "true"
            "x-ms-documentdb-partitionkey"   = "[`"$partitionKeyValue`"]"
        }
        
        $body = $Item | ConvertTo-Json -Depth 10
        
        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
        
        Write-Log "Successfully saved notification to Cosmos DB: $($Item.id)"
        return $true
        
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Failed to save notification to Cosmos DB: $($Item.id). Error: $errorMessage"
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

function Send-OrientationReminderEmail {
    param (
        [Parameter(Mandatory=$true)]
        [string]$CadetEmail,
        [Parameter(Mandatory=$true)]
        [string]$CadetName,
        [Parameter(Mandatory=$true)]
        [string]$CadetLastName,
        [Parameter(Mandatory=$true)]
        [string]$AirportCode,
        [string]$CoordinatorName,
        [string]$CoordinatorPhone,
        [string]$CoordinatorEmail
    )
    
    $subject = "Reminder: Your CAP Orientation Flight Tomorrow at $AirportCode"
    
    # Build coordinator section if coordinator info is provided
    $coordinatorSection = ""
    if ($CoordinatorName) {
        $coordinatorSection = @"
    <p>If you have any questions before the activity, please contact the event coordinator:</p>
    <p style="margin-left: 20px;">
      <strong>$CoordinatorName</strong><br/>
"@
        if ($CoordinatorPhone) {
            $coordinatorSection += "      Phone: $CoordinatorPhone<br/>`n"
        }
        if ($CoordinatorEmail) {
            $coordinatorSection += "      Email: <a href=`"mailto:$CoordinatorEmail`">$CoordinatorEmail</a><br/>`n"
        }
        $coordinatorSection += "    </p>`n"
    }
    
    $body = @"
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    h2 { color: #1a365d; }
    ul { margin: 10px 0; padding-left: 20px; }
    .section { margin: 20px 0; }
    a { color: #2563eb; text-decoration: none; }
  </style>
</head>
<body>
  <p>Cadet $CadetLastName,</p>

  <p>This is a reminder that you are scheduled for your Civil Air Patrol Orientation Flight tomorrow at <strong>$AirportCode</strong>.</p>

  $coordinatorSection

  <div class="section">
    <p>Please plan for the activity to take approximately 3–4 hours. The schedule will include:</p>
    <ul>
      <li>Pre-flight briefing</li>
      <li>Pre-flight inspection</li>
      <li>Orientation flight</li>
      <li>Post-flight briefing (Q&A)</li>
    </ul>
  </div>

  <div class="section">
    <h2>What to Bring</h2>
    <ul>
      <li><strong>CAP Identification Card</strong></li>
      <li><strong>Completed CAPF 161 – Emergency Contact Info</strong>. You must have this on your person at all times.</li>
      <li><strong>Completed CAPF 60-80 – Activity Permission Slip</strong><br/>
        Link: <a href="https://www.gocivilairpatrol.com/media/cms/CAPF_6080_Permission_Slip_6DEA6DC6D05D1.pdf">https://www.gocivilairpatrol.com/media/cms/CAPF_6080_Permission_Slip_6DEA6DC6D05D1.pdf</a>
      </li>
      <li><strong>Approved CAP Uniform</strong> (Blues or ABUs). Bring a jacket, as it can be cold on the ramp.</li>
      <li>Sunglasses</li>
      <li>Chewing gum</li>
      <li>Snacks and water</li>
      <li>Camera or video camera (optional)</li>
    </ul>
  </div>

  <p>We look forward to seeing you tomorrow and hope you enjoy your flight experience!</p>

  <p>
    Respectfully,<br/>
    <strong>Civil Air Patrol Colorado Wing</strong>
  </p>
</body>
</html>
"@
    
    # Send email via Microsoft Graph
    try {
        $fromAddress = $env:LOG_EMAIL_FROM_ADDRESS
        if (-not $fromAddress) {
            $fromAddress = "noreply@cowg.cap.gov"
        }
        
        # Build toRecipients array
        $toRecipients = @(
            @{
                emailAddress = @{
                    address = $CadetEmail
                }
            }
        )
        
        # Build ccRecipients array (only add coordinator if email is provided)
        $ccRecipients = @()
        if ($CoordinatorEmail) {
            $ccRecipients += @{
                emailAddress = @{
                    address = $CoordinatorEmail
                }
            }
        }
        
        $emailBody = @{
            message = @{
                subject = $subject
                body = @{
                    contentType = "HTML"
                    content = $body
                }
                toRecipients = $toRecipients
            }
            saveToSentItems = $true
        }
        
        # Add CC recipients if any exist
        if ($ccRecipients.Count -gt 0) {
            $emailBody.message["ccRecipients"] = $ccRecipients
        }
        
        # Send email from the configured service account
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$fromAddress/sendMail" -Body ($emailBody | ConvertTo-Json -Depth 10) -ContentType "application/json"
        
        Write-Log "$logPrefix ✅ Email sent successfully to $CadetEmail"
        
    } catch {
        Write-Log "$logPrefix ❌ Failed to send email to ${CadetEmail}: $($_.Exception.Message)"
        throw
    }
}

$logPrefix = '⏰ [sendReminders]'

Write-Log "$logPrefix Timer trigger function started at $(Get-Date -Format o)"

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

    # Calculate tomorrow's date in YYYY-MM-DD format (Mountain time zone)
    $mountainZone = [System.TimeZoneInfo]::FindSystemTimeZoneById('Mountain Standard Time')
    $now = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $mountainZone)
    $tomorrow = $now.AddDays(1)
    $tomorrowStr = $tomorrow.ToString('yyyy-MM-dd')
    
    Write-Log "$logPrefix Looking for flights and events on $tomorrowStr"

    # Get Cosmos DB configuration
    $cosmosConfig = Get-CosmosDbConnection
    $cosmosConnectionString = $cosmosConfig.ConnectionString
    $cosmosDatabase = $cosmosConfig.Database
    
    if (-not $cosmosConnectionString -or -not $cosmosDatabase) {
        Write-Log "$logPrefix Error: Cosmos DB configuration incomplete"
        throw "Cosmos DB configuration incomplete"
    }

    # Initialize counters
    $emailsSent = 0
    $emailsFailed = 0
    $emailsSkipped = 0

    # Query for flights tomorrow
    Write-Log "$logPrefix Querying flights container for date $tomorrowStr"
    $flights = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                       -Database $cosmosDatabase `
                                       -Container "flights" `
                                       -Query "SELECT * FROM c WHERE c.date = '$tomorrowStr' AND c.status = 'scheduled'"
    
    Write-Log "$logPrefix Found $($flights.Count) scheduled flight(s) for tomorrow"

    # Query for events tomorrow
    Write-Log "$logPrefix Querying events container for date $tomorrowStr"
    $events = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                      -Database $cosmosDatabase `
                                      -Container "events" `
                                      -Query "SELECT * FROM c WHERE c.date = '$tomorrowStr' AND c.status = 'scheduled'"
    
    Write-Log "$logPrefix Found $($events.Count) scheduled event(s) for tomorrow"

    # Process flights
    foreach ($flight in $flights) {
        Write-Log "$logPrefix Processing flight $($flight.id) at $($flight.airport)"
        
        if ($flight.seats) {
            foreach ($seat in $flight.seats) {
                if ($seat.cadetId) {
                    try {
                        # Check if reminder already sent (duplicate prevention)
                        $notificationId = "reminder-$tomorrowStr-flight-$($flight.id)-$($seat.cadetId)"
                        
                        $existingNotification = Get-CosmosDbItemByQuery -ConnectionString $cosmosConnectionString `
                                                                        -Database $cosmosDatabase `
                                                                        -Container "notifications" `
                                                                        -Query "SELECT * FROM c WHERE c.id = '$notificationId'"
                        
                        if ($existingNotification) {
                            Write-Log "$logPrefix ⏭️  Skipping cadet $($seat.cadetId) - reminder already sent"
                            $emailsSkipped++
                            continue
                        }

                        Write-Log "$logPrefix Sending reminder to cadet $($seat.cadetId) ($($seat.cadetName))"
                        
                        # Get detailed cadet info from Graph API
                        $cadetInfo = Get-MgUser -UserId $seat.cadetId -Property "mail,displayName,employeeId" -ErrorAction Stop
                        
                        if (-not $cadetInfo -or -not $cadetInfo.Mail) {
                            Write-Log "$logPrefix ❌ Could not get email for cadet $($seat.cadetId)"
                            $emailsFailed++
                            continue
                        }

                        # Extract last name from display name (assumes format "FirstName LastName, Rank")
                        $displayNameParts = $cadetInfo.DisplayName -split ','
                        $namePart = $displayNameParts[0].Trim()
                        $nameWords = $namePart -split '\s+'
                        $lastName = if ($nameWords.Count -gt 0) { $nameWords[-1] } else { "Cadet" }

                        # Send orientation reminder email (flights don't have coordinators)
                        Send-OrientationReminderEmail -CadetEmail $cadetInfo.Mail `
                                                     -CadetName $cadetInfo.DisplayName `
                                                     -CadetLastName $lastName `
                                                     -AirportCode $flight.airport `
                                                     -CoordinatorName $null `
                                                     -CoordinatorPhone $null `
                                                     -CoordinatorEmail $null

                        # Record that reminder was sent
                        $notificationItem = @{
                            id = $notificationId
                            userId = $seat.cadetId
                            flightId = $flight.id
                            cadetId = $seat.cadetId
                            cadetEmail = $cadetInfo.Mail
                            type = "reminder"
                            activityType = "flight"
                            activityDate = $tomorrowStr
                            sentAt = (Get-Date -Format o)
                        }
                        
                        $saved = Save-NotificationItem -Item $notificationItem `
                                                      -ConnectionString $cosmosConnectionString `
                                                      -Database $cosmosDatabase
                        
                        if ($saved) {
                            $emailsSent++
                            Write-Log "$logPrefix ✅ Sent reminder to $($cadetInfo.Mail)"
                        } else {
                            Write-Log "$logPrefix ⚠️  Email sent but failed to record notification (non-critical)"
                            $emailsSent++
                        }
                        
                    } catch {
                        Write-Log "$logPrefix ❌ Error sending reminder to cadet $($seat.cadetId): $($_.Exception.Message)"
                        $emailsFailed++
                    }
                }
            }
        }
    }

    # Process events
    foreach ($event in $events) {
        Write-Log "$logPrefix Processing event $($event.id) at $($event.airport)"
        
        if ($event.slots) {
            foreach ($slot in $event.slots) {
                if ($slot.cadetId) {
                    try {
                        # Check if reminder already sent (duplicate prevention)
                        $notificationId = "reminder-$tomorrowStr-event-$($event.id)-$($slot.cadetId)"
                        
                        $existingNotification = Get-CosmosDbItemByQuery -ConnectionString $cosmosConnectionString `
                                                                        -Database $cosmosDatabase `
                                                                        -Container "notifications" `
                                                                        -Query "SELECT * FROM c WHERE c.id = '$notificationId'"
                        
                        if ($existingNotification) {
                            Write-Log "$logPrefix ⏭️  Skipping cadet $($slot.cadetId) - reminder already sent"
                            $emailsSkipped++
                            continue
                        }

                        # Use cadetName or displayName field from slot
                        $slotName = if ($slot.cadetName) { $slot.cadetName } else { $slot.displayName }
                        Write-Log "$logPrefix Sending reminder to cadet $($slot.cadetId) ($slotName)"
                        
                        # Get detailed cadet info from Graph API
                        $cadetInfo = Get-MgUser -UserId $slot.cadetId -Property "mail,displayName,employeeId" -ErrorAction Stop
                        
                        if (-not $cadetInfo -or -not $cadetInfo.Mail) {
                            Write-Log "$logPrefix ❌ Could not get email for cadet $($slot.cadetId)"
                            $emailsFailed++
                            continue
                        }

                        # Extract last name from display name
                        $displayNameParts = $cadetInfo.DisplayName -split ','
                        $namePart = $displayNameParts[0].Trim()
                        $nameWords = $namePart -split '\s+'
                        $lastName = if ($nameWords.Count -gt 0) { $nameWords[-1] } else { "Cadet" }

                        # Send orientation reminder email with coordinator info if available
                        Send-OrientationReminderEmail -CadetEmail $cadetInfo.Mail `
                                                     -CadetName $cadetInfo.DisplayName `
                                                     -CadetLastName $lastName `
                                                     -AirportCode $event.airport `
                                                     -CoordinatorName $event.coordinatorName `
                                                     -CoordinatorPhone $event.coordinatorPhone `
                                                     -CoordinatorEmail $event.coordinatorEmail

                        # Record that reminder was sent
                        $notificationItem = @{
                            id = $notificationId
                            userId = $slot.cadetId
                            flightId = $event.id
                            cadetId = $slot.cadetId
                            cadetEmail = $cadetInfo.Mail
                            type = "reminder"
                            activityType = "event"
                            activityDate = $tomorrowStr
                            sentAt = (Get-Date -Format o)
                        }
                        
                        $saved = Save-NotificationItem -Item $notificationItem `
                                                      -ConnectionString $cosmosConnectionString `
                                                      -Database $cosmosDatabase
                        
                        if ($saved) {
                            $emailsSent++
                            Write-Log "$logPrefix ✅ Sent reminder to $($cadetInfo.Mail)"
                        } else {
                            Write-Log "$logPrefix ⚠️  Email sent but failed to record notification (non-critical)"
                            $emailsSent++
                        }
                        
                    } catch {
                        Write-Log "$logPrefix ❌ Error sending reminder to cadet $($slot.cadetId): $($_.Exception.Message)"
                        $emailsFailed++
                    }
                }
            }
        }
    }

    Write-Log "$logPrefix ✅ Reminder job completed: $emailsSent sent, $emailsFailed failed, $emailsSkipped skipped (already sent)"

} catch {
    Write-Log "$logPrefix ❌ Error in timer trigger: $($_.Exception.Message)"
    throw
}
