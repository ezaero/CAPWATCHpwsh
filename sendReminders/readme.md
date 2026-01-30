# sendReminders Function

## Overview
Timer-triggered Azure Function that sends orientation event reminder emails to cadets scheduled for flying activities the following day.

**Schedule:** Runs daily at 6 PM MST (01:00 UTC the next day)

## What It Does
1. Calculates tomorrow's date in Mountain Time zone
2. Queries Cosmos DB for:
   - Events scheduled for tomorrow with status "scheduled"
3. For each cadet assigned to an event:
   - Fetches detailed cadet info from Microsoft Graph API
   - Sends an HTML-formatted reminder email
   - CCs the event coordinator (if available)
   - Records the notification to Cosmos DB to prevent duplicate sends
4. Logs a summary: emails sent, failed, and skipped (already sent)

## Requirements

### Environment Variables
```
CosmosDbConnectionString  - Cosmos DB account connection string
CosmosDbDatabase         - Cosmos DB database name (e.g., "orientation-flights")
LOG_EMAIL_FROM_ADDRESS   - Email address to send reminders from (fallback: noreply@cowg.cap.gov)
```

### Azure Services
- **Cosmos DB:** Containers must exist:
  - `events` (partition key: `/CAPID`)
  - `notifications` (partition key: `/userId`)
- **Microsoft Graph API:** Requires authentication via managed identity
- **Azure Key Vault:** (optional) For storing secrets

### Data Schema

#### Flights Container
```json
{
  "id": "flight-123",
  "CAPID": "12345",
  "date": "2026-01-26",
  "status": "scheduled",
  "airport": "KAPA",
  "aircraft": "N1234",
  "pilotName": "John Doe",
  "seats": [
    {
      "cadetId": "uuid-here",
      "cadetName": "Jane Smith"
    }
  ]
}
```

#### Events Container
```json
{
  "id": "event-456",
  "CAPID": "12345",
  "date": "2026-01-26",
  "status": "scheduled",
  "airport": "KAPA",
  "coordinatorName": "John Coordinator",
  "coordinatorPhone": "555-1234",
  "coordinatorEmail": "coordinator@example.com",
  "slots": [
    {
      "cadetId": "uuid-here",
      "cadetName": "Jane Smith"
    }
  ]
}
```

#### Notifications Container (Created by function)
```json
{
  "id": "reminder-2026-01-26-flight-flight-123-uuid",
  "userId": "uuid-here",
  "flightId": "flight-123",
  "cadetId": "uuid-here",
  "cadetEmail": "jane@example.com",
  "type": "reminder",
  "activityType": "flight",
  "activityDate": "2026-01-26",
  "sentAt": "2026-01-25T02:57:06Z"
}
```

## Email Template
Sends an HTML-formatted email with:
- Greeting addressing cadet by last name
- Activity details (airport, date, time estimate)
- Coordinator contact info (if available)
- What to bring checklist (CAP ID, forms, uniform, etc.)
- Links to required forms (CAPF 60-80)
- Professional CAP Colorado Wing signature

**CC Recipients:**
- Event coordinator email (if provided in event/flight record)

## Local Testing

### Prerequisites
```bash
# Install Azure Functions Core Tools
# macOS with Homebrew:
brew tap azure/formulae
brew install azure-functions-core-tools@4

# Install PowerShell modules (if needed)
pip install azure-cli
```

### Run Locally
```bash
cd /path/to/CAPWATCHSyncPWSH
func start
```

### Test via Code+Test
1. Open VS Code
2. Right-click on `sendReminders` function folder
3. Select **Code + Test**
4. Click **Run** or **Test/Run** button
5. Check function output and logs

### Manual Test Command
```bash
func azure functionapp fetch-app-settings <function-app-name>
func start --verbose
```

## Logs
- Location: `$HOME/logs/script_log_YYYY-MM-DD.txt`
- Format: `timestamp - [prefix] message`
- Prefix: `⏰ [sendReminders]`

### Example Log Output
```
2026-01-25 03:00:07 - ⏰ [sendReminders] Timer trigger function started at 2026-01-25T03:00:07.7194759+00:00
2026-01-25 03:00:11 - ⏰ [sendReminders] Successfully authenticated to Microsoft Graph
2026-01-25 03:00:12 - ⏰ [sendReminders] Found 1 scheduled event(s) for tomorrow
2026-01-25 03:00:20 - ⏰ [sendReminders] ✅ Email sent successfully to cadet@example.com
2026-01-25 03:00:20 - ⏰ [sendReminders] ✅ Reminder job completed: 1 sent, 0 failed, 0 skipped (already sent)
```

## Error Handling

### Common Issues

**400 Bad Request on notification save:**
- Usually a partition key mismatch
- Ensure `notifications` container has `/userId` as partition key
- Non-critical: email is still sent

**404 Not Found on cadet lookup:**
- Cadet ID doesn't exist in Microsoft Graph
- Function logs error and skips that cadet
- Returns 0 failed emails (email send took precedence)

**Timeout:**
- Default timeout: 10 minutes
- If exceeding, check Graph API or Cosmos DB performance
- Consider batching or splitting queries

## Conversion Notes

### From Node.js
This function was converted from TypeScript/Node.js (`index.js`). Key differences:
- All logic moved to `run.ps1` (PowerShell)
- Dependencies (Microsoft Graph SDK) replaced with native PowerShell cmdlets
- No external npm packages required
- Runs with PowerShell worker runtime

### Removed Files
Not needed after conversion:
- `index.js` (main entry point)
- `config.js`
- `middleware/` folder
- `services/` folder
- `shared/` folder (Node.js utilities)

## Related Functions
- **OFlights:** Processes O-Flight data from CAPWATCH
- **emailLogFile:** Sends log summaries

## Future Enhancements
- [ ] Retry logic for failed email sends
- [ ] Support for SMS reminders
- [ ] Customizable reminder timing (1 day, 3 days, 1 week before)
- [ ] Template customization per wing
- [ ] Analytics/reporting on reminder delivery rates

## Support
For issues or questions, contact the CAP Colorado Wing IT team.
