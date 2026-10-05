# Teams Onboarding Procedure for Manual Guests

Prepared: 2026-10-02

## Purpose

This procedure gives COWG administrators a repeatable way to onboard users who are not in CAPWATCH but need access to a Microsoft Team, such as Winter Encampment staff from outside the wing.

The goal is to avoid solving access issues one member at a time by using a tracked CSV, a dry-run-first process, and clear troubleshooting steps.

## Roles

- **Sponsor**: COWG member who approves the guest's need for access and confirms the correct Team.
- **IT administrator**: Maintains `ManualGuests.csv`, runs the automation, and resolves account or Teams issues.
- **Guest user**: Accepts the invitation, completes MFA setup, and signs into Teams with the invited email address.

## Source of Truth

Manual guests are tracked in:

```text
CAPWATCHpwsh/ManualGuests.csv
```

Required or recommended columns:

```csv
Email,DisplayName,NameFirst,NameLast,Grade,Type,Unit,Teams,Sponsor,EmployeeId,Enabled
```

Example:

```csv
Email,DisplayName,NameFirst,NameLast,Grade,Type,Unit,Teams,Sponsor,EmployeeId,Enabled
kaitlin.reinhart@utwg.cap.gov,Kaitlin Reinhart,Kaitlin,Reinhart,Winter Encampment Staff,MANUAL,WINTER-ENCAMPMENT,"COWG Winter Encampment",Stephanie.Moffitt@cowg.cap.gov,,TRUE
```

## Pre-Onboarding Checklist

Before running automation:

1. Confirm the sponsor approved access.
2. Confirm the guest's exact email address.
3. Confirm the exact Team display name.
4. Confirm whether the user needs access to one Team or multiple Teams.
5. Add or update the row in `ManualGuests.csv`.
6. Keep `EmployeeId` blank unless there is a specific reason to supply one.
7. Set `Enabled=TRUE`.

## Step 1: Dry-Run Account Creation

Run `checkAccounts` with `EXECUTE` unset or `EXECUTE=false`.

Expected dry-run signals:

```text
Loaded 1 manual guest record(s) from .../ManualGuests.csv
[DRY-RUN] Guest user creation skipped (would create and send invitation)
```

If the account already exists, confirm the dry-run indicates it will be adopted or updated rather than duplicated.

## Step 2: Execute Account Creation

Set:

```text
EXECUTE=true
```

Run `checkAccounts`.

Expected result:

- B2B guest is created or existing guest is updated.
- Invitation is sent.
- Account receives a stable `MANUAL-<hash>` employee ID if `EmployeeId` was blank.
- Account metadata is stamped for unit, title, type, and mail.
- Sponsor appears in notification output when present.

## Step 3: User Accepts Invitation and Sets Up MFA

Send the member-facing MFA setup guide to the user and sponsor.

The user must:

1. Accept the invitation using the same email address listed in `ManualGuests.csv`.
2. Complete MFA setup.
3. Sign into Teams.
4. Switch to the COWG tenant/organization if prompted.

## Step 4: Dry-Run Team Assignment

Run `updateTeams` with `EXECUTE` unset or `EXECUTE=false`.

Expected dry-run signal:

```text
Would add manual guest to Team - kaitlin.reinhart@utwg.cap.gov -> COWG Winter Encampment (Sponsor: Stephanie.Moffitt@cowg.cap.gov)
```

If the guest cannot be found, run `checkAccounts` first or wait for the B2B invitation/account creation to complete.

If the Team cannot be found, verify the `Teams` column matches the Team display name exactly.

## Step 5: Execute Team Assignment

Set:

```text
EXECUTE=true
```

Run `updateTeams`.

Expected result:

- User is added to each Team listed in `Teams`.
- Existing Team membership is detected and skipped.
- Sponsor appears in assignment logs for audit context.

## Bulk Onboarding for Events

For events such as Winter Encampment:

1. Collect the roster in spreadsheet form.
2. Normalize names, emails, sponsors, and Teams.
3. Paste/export into `ManualGuests.csv`.
4. Run `checkAccounts` in dry-run.
5. Fix invalid emails or duplicate rows.
6. Run `checkAccounts` with `EXECUTE=true`.
7. Send the MFA setup guide to guests and sponsors.
8. Run `updateTeams` in dry-run.
9. Fix missing accounts or Team name mismatches.
10. Run `updateTeams` with `EXECUTE=true`.

## Troubleshooting

### Guest Account Was Not Created

Check:

- `Email` is populated and valid.
- `Enabled` is not `FALSE`, `NO`, `0`, or `DISABLED`.
- Existing account does not already use the same mail address.
- Domain is not blocked by cross-tenant access policy.
- Exchange recipient conflict did not block invitation.

### Invitation Was Sent but User Cannot Sign In

Check:

- User is using the invited email address.
- User completed MFA setup.
- User is selecting the COWG tenant in Teams.
- Entra sign-in logs for MFA, Conditional Access, or cross-tenant errors.

### User Is Not in the Team

Check:

- `updateTeams` ran after `checkAccounts`.
- `Teams` column matches exact Team display name.
- Guest account exists in Microsoft Graph.
- The user accepted the invitation if Teams access requires redemption.

### Sponsor or Approval Is Unclear

Do not create or maintain access without a sponsor. Update `Sponsor` before executing account creation or Team assignment.

## Offboarding

When access is no longer approved:

1. Set `Enabled=FALSE` or remove the row from `ManualGuests.csv`.
2. Remove the user from event Teams if immediate removal is required.
3. Let maintenance identify accounts no longer represented in CAPWATCH or enabled `ManualGuests.csv`, or remove the account manually if urgent.
4. Keep enough record of the sponsor and event purpose for audit review.

## Operating Notes

- Dry-run is the default because `EXECUTE` must be explicitly set to `true`.
- `checkAccounts` creates or updates accounts; `updateTeams` manages Team membership.
- The sponsor is recorded in source data and logs, not in Microsoft Entra ID.
- A user may need several minutes after invitation redemption before Teams membership works cleanly across all clients.

