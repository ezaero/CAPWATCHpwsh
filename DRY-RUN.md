# Dry-Run Mode Guide

## Overview

Dry-run mode is a **safety feature** that allows you to preview what changes a function would make **without actually applying them**. This is essential for:

- ✅ Testing before production deployment
- ✅ Validating configuration before making changes  
- ✅ Multi-wing deployments (test on one wing before rolling out to others)
- ✅ Understanding impact before execution
- ✅ Auditing changes in logs

## Default Behavior

**By default, all functions run in DRY-RUN mode** (safe/read-only). This means:

- Functions will **query and analyze** data
- Functions will **log what would happen**
- Functions will **NOT make any actual changes**
- No Teams created, no users modified, no emails sent

## Enabling Dry-Run (Explicit)

### Option 1: Azure Portal

1. Go to Function App → **Configuration**
2. Click **+ New application setting**
3. Add setting: `EXECUTE` = `false`
4. Click **Save**

Then test functions - they will run in dry-run mode with "🔍 [DRY-RUN]" prefixes in logs.

### Option 2: Azure CLI

```bash
az functionapp config appsettings set \
  --resource-group capwatch-MT-rg \
  --name capwatch-MT-func \
  --settings "EXECUTE=false"
```

### Option 3: Local Development

In `local.settings.json`:

```json
{
  "Values": {
    "EXECUTE": "false"
  }
}
```

## Enabling Execution Mode

Once you've reviewed dry-run logs and are confident in the changes:

### Option 1: Azure Portal

1. Go to Function App → **Configuration**
2. Click on the `EXECUTE` setting
3. Change value to `true`
4. Click **Save**

### Option 2: Azure CLI

```bash
az functionapp config appsettings set \
  --resource-group capwatch-MT-rg \
  --name capwatch-MT-func \
  --settings "EXECUTE=true"
```

## What Gets Logged in Dry-Run Mode?

When `EXECUTE=false`, the logs will show:

```
2026-02-26 14:30:45 - 🔍 [DRY-RUN] Creating Team - Team: CO-001 Boulder Composite Squad
2026-02-26 14:30:46 - 🔍 [DRY-RUN] Adding member to Team - User: John Smith (john.smith@cowg.cap.gov)
2026-02-26 14:30:47 - 🔍 [DRY-RUN] Would send email to - admin@cowg.cap.gov
```

When `EXECUTE=true`, the logs will show:

```
2026-02-26 14:31:45 - ⚡ [EXECUTING] Creating Team - Team: CO-001 Boulder Composite Squad
2026-02-26 14:31:46 - ⚡ [EXECUTING] Adding member to Team - User: John Smith (john.smith@cowg.cap.gov)
2026-02-26 14:31:47 - ⚡ [EXECUTING] Send email to - admin@cowg.cap.gov
```

## Multi-Wing Deployment Workflow

This is the recommended workflow when deploying to multiple wings:

### Phase 1: Dry-Run Testing

```bash
# 1. Deploy code to new wing
func azure functionapp publish capwatch-MT-func --build remote

# 2. Ensure EXECUTE is set to false (or leave unset - default is dry-run)
az functionapp config appsettings set \
  --resource-group capwatch-MT-rg \
  --name capwatch-MT-func \
  --settings "EXECUTE=false"

# 3. Test each function in portal or via timer triggers
# 4. Review logs in Application Insights
# 5. Verify the proposed changes look correct
```

### Phase 2: Enable Execution

```bash
# When confident in results, enable execution
az functionapp config appsettings set \
  --resource-group capwatch-MT-rg \
  --name capwatch-MT-func \
  --settings "EXECUTE=true"

# 6. Run functions again to apply changes
# 7. Verify in Teams, Exchange, Azure AD that changes were applied
# 8. Monitor logs for any errors
```

### Phase 3: Production Confidence

```bash
# Once successful in one wing, repeat for other wings
# Each wing can be rolled out independently with safety
```

## Which Functions Support Dry-Run?

### Full Support (Safe to use)

These functions fully respect the `EXECUTE` flag:

- ✅ `updateTeams` - Create/update Teams without modifying
- ✅ `DLAnnouncements` - Preview distribution list changes
- ✅ `DLOpsQuals` - Preview operational qualification lists
- ✅ `DLSeniorsCadets` - Preview senior/cadet lists
- ✅ `DLSpecTrack` - Preview specialty track lists
- ✅ `checkAccounts` - Preview account creation/updates
- ✅ `update-user-names-ranks` - Preview user attribute updates
- ⚠️ `download-extract-capwatch` - Downloads are read-only (safe)
- ⚠️ `Maintenance` - Should be reviewed before running

### Functions with Limited State Changes

- 🟡 `sendReminders` - Can be tested in dry-run
- 🟡 `emailLogFile` - Can be tested in dry-run
- 🟡 `escalatePilotInvitations` - Can be tested in dry-run

### State-Modifying Functions (Use Caution)

- 🔴 `processInvitationExpiry` - Updates Cosmos DB state; test carefully
- 🔴 `OFlights` - May update Cosmos DB; test carefully
- 🔴 `OFlightMetrics` - May update Cosmos DB; test carefully

## Example: Testing updateTeams in Dry-Run

### Step 1: Verify Dry-Run is Enabled

```bash
az functionapp config appsettings list \
  --resource-group capwatch-MT-rg \
  --name capwatch-MT-func \
  --query "[?name=='EXECUTE'].value" --output tsv
```

Expected output: *(empty)* or `false`

### Step 2: Run Function from Portal

1. Go to Function App → **Functions** → **updateTeams**
2. Click **Code + Test** tab
3. Click **Run** button
4. Check **Output** - should see:
   ```
   🔍 [DRY-RUN] Creating Team - Team: MT-001 ...
   🔍 [DRY-RUN] Adding 5 members to Team...
   ```

### Step 3: Review Application Insights

1. Go to Function App → **Monitor**
2. Check logs for `[DRY-RUN]` entries
3. Verify the proposed changes

### Step 4: Enable Execution

Once confident:

```bash
az functionapp config appsettings set \
  --resource-group capwatch-MT-rg \
  --name capwatch-MT-func \
  --settings "EXECUTE=true"
```

### Step 5: Run Again

1. Go to Function App → **Functions** → **updateTeams**
2. Click **Run** button again
3. Check **Output** - should see:
   ```
   ⚡ [EXECUTING] Creating Team - Team: MT-001 ...
   ⚡ [EXECUTING] Adding 5 members to Team...
   ```
4. Teams should now be created and members added

## Troubleshooting

### Q: I set EXECUTE=true but still see [DRY-RUN] messages

**A:** The setting may not have propagated. Try:
1. Restart the Function App
2. Re-read the setting to verify it changed
3. Check for a different `EXECUTE` environment variable (may be lowercase `execute`)

### Q: How do I know if my changes were actually applied?

**A:** Check for:
- `⚡ [EXECUTING]` prefix in logs (not `🔍 [DRY-RUN]`)
- Changes in Teams (new teams created, members added)
- Changes in Exchange (distribution lists updated)
- Changes in Azure AD (user attributes updated)

### Q: Can I switch between dry-run and execute multiple times?

**A:** Yes! You can:
1. Run in dry-run mode to see what would happen
2. Switch to execute and run to make changes
3. Switch back to dry-run to verify next changes
4. Repeat as needed

This is safe and recommended for testing.

## Best Practices

✅ **DO:**
- Always start with dry-run mode (`EXECUTE=false`)
- Review logs thoroughly before switching to execute
- Test on a non-production wing first
- Document what changes you expect before running
- Run all related functions in dry-run before executing any

❌ **DON'T:**
- Run functions in execute mode without reviewing dry-run logs first
- Skip dry-run testing when deploying to new wings
- Change `EXECUTE` setting and immediately run functions without checking
- Run all functions at once - test one at a time for clarity

## Reference: Dry-Run Helper Functions

These helper functions are available in `shared/shared.ps1`:

### Get-DryRunMode

Returns `$true` if in dry-run mode, `$false` if in execute mode.

```powershell
$isDryRun = Get-DryRunMode
if ($isDryRun) {
    Write-Log "Running in DRY-RUN mode"
}
```

### Write-OperationLog

Logs an operation with automatic `[DRY-RUN]` or `[EXECUTING]` prefix.

```powershell
Write-OperationLog -Operation "Creating Team" -Details "Team: MyTeam"
# Output: 🔍 [DRY-RUN] Creating Team - Team: MyTeam
```

### Should-ExecuteOperation

Convenience function to check if operation should execute.

```powershell
if (Should-ExecuteOperation) {
    # Make actual changes
    Create-MgTeam @params
} else {
    Write-OperationLog "Would create Team" "Team: MyTeam"
}
```

## Summary

- 🔍 **Dry-Run Mode**: Preview changes, nothing is modified
- ⚡ **Execute Mode**: Actually apply changes
- 📋 **Default**: Dry-run (safe)
- ✅ **Best Practice**: Always test in dry-run first
- 🚀 **Multi-Wing**: Test on one wing, roll out to others with confidence
