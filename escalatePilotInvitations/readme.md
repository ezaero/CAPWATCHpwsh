# Pilot Invitation Escalation Function

## Overview

This Azure Function automatically escalates pilot invitations to all Orientation Pilots when pilot slots remain unfilled within 7 days of the scheduled event date.

## Schedule

- **Trigger**: Timer (CRON expression: `0 0 6,18 * * *`)
- **Frequency**: Twice daily at 6:00 AM and 6:00 PM Mountain Time
- **Run on Startup**: No

## Functionality

1. **Query Events**: Finds scheduled events happening within the next 7 days that:
   - Are in `scheduled` status
   - Have unfilled pilot slots (`numberOfPilotsRequired > 0`)
   - Haven't already had escalation invitations sent (`escalationStatus.initialInvitationsSent = false`)

2. **Get Pilots**: Retrieves all users with `pilot` role from the `users` container

3. **Expire Pending Invitations**: Before escalating:
   - Finds all pending pilot invitations for the event
   - Marks them as `expired` with reason: "Event escalated to all pilots within 7 days of scheduled date"
   - Prevents confusion between original invitations and escalation invitations

4. **Filter Eligible Pilots**: Excludes pilots who:
   - Are already assigned to this event
   - Don't have a valid email address
   - Note: Previously invited pilots ARE included (their invitations are now expired)

5. **Send Invitations**: For each eligible pilot:
   - Generates a secure random token
   - Creates an invitation record in `pilotInvitations` container
   - Sends escalation email via Microsoft Graph API

6. **Update Event**: Sets `escalationStatus.initialInvitationsSent = true` to prevent duplicate escalations

## Cosmos DB Containers Used

- **events**: Read/Update - Query for events needing escalation, update escalation status
- **users**: Read - Get all pilots
- **pilotInvitations**: Read/Write/Update - Expire old invitations, create new escalation invitations

## Environment Variables Required

- `CosmosDbConnectionString`: Connection string for Cosmos DB
- `CosmosDbDatabase`: Database name (typically "OFlightCoordinator")
- `LOG_EMAIL_FROM_ADDRESS`: Email address to send from (default: "OFlights@cowg.cap.gov")
- `FRONTEND_URL`: Base URL for invitation links (default: "https://orientationflights.cowg.cap.gov")
- `TEST_EMAIL_OVERRIDE`: (Optional) Override all email recipients for testing

## Email Template

The escalation email includes:
- ⚠️ Urgent banner explaining why this is an escalation
- 📋 Event details (location, date, time, aircraft, pilots needed)
- ⏰ Expiration notice (48 hours by default)
- ✅ Accept/Decline buttons with secure token links
- 👤 Coordinator info (if available)
- 📧 Footer with COFLICS branding

Subject: `URGENT: Pilot Still Needed - O-Flight Event at {AIRPORT} on {DATE}`

## Partition Key Structure

- **events**: `[airport, date]` (hierarchical)
- **pilotInvitations**: `[eventId]` (single)
- **users**: Cross-partition queries enabled

## Logging

All actions are logged with prefix `🚨 [escalatePilotInvitations]`:
- Events checked and eligible
- Invitations expired and sent
- Errors encountered
- Summary statistics

## Integration with Existing System

This function complements:
- `sendReminders`: Notifies cadets 48 hours before events
- `processInvitationExpiry`: Cascades cadet invitations when they expire
- **`escalatePilotInvitations`**: Escalates to all pilots when slots remain unfilled within 7 days of the event

## Testing

Set `TEST_EMAIL_OVERRIDE` environment variable to redirect all emails to a test address:

```powershell
$env:TEST_EMAIL_OVERRIDE = "test@example.com"
```

## Statistics Tracked

- `EventsChecked`: Total events queried
- `EventsEligible`: Events meeting escalation criteria
- `InvitationsExpired`: Original invitations marked as expired
- `InvitationsSent`: New escalation invitations created and sent
- `Errors`: Count of failures during processing

## Error Handling

- Continues processing remaining events/pilots if individual operations fail
- Logs all errors with details
- Throws fatal errors to trigger Azure Function retry logic
