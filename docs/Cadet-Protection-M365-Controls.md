# Civil Air Patrol Cadet Protection in Microsoft 365

Working group discussion brief

Prepared: 2026-07-09

## Purpose

Civil Air Patrol's Cadet Protection expectations require adult involvement in cadet interactions. In physical settings, this is often framed as either two adults present with a cadet or two cadets present with an adult. The same risk pattern exists in digital communication: email, Teams chats, Teams channels, meetings, calls, and file collaboration can all create one-to-one adult/cadet communication paths.

This document outlines Microsoft 365 and Exchange Online controls that could reduce unsupervised digital communication, improve supervision, and provide an audit trail for review. It is intended to support working group discussion, not to serve as legal or policy approval.

## Executive Summary

There does not appear to be one Microsoft 365 switch that enforces the Cadet Protection rule across every communication surface. The best fit is a layered model:

1. Prevent unsupervised communication where native controls exist.
2. Route or copy an approved adult where Exchange controls make that practical.
3. Monitor and investigate communications that cannot be perfectly prevented.
4. Use CAPWATCH/Entra identity data to keep parent and senior member relationships current.

The strongest native control is Microsoft Teams supervised chat. The weakest area is inbound email, because Exchange mail-flow rules can add fixed recipients or the sender's manager, but they do not natively calculate "recipient.employeeId + 'P'" and Bcc that account.

## Guiding Principles

- Cadets should not be able to initiate private unsupervised digital communications with adults.
- Adults should not be able to create persistent private communication channels with a cadet unless another approved adult or parent is included.
- Routine cadet communication should happen in supervised team/channel contexts where practical.
- Parents and senior members should be included through clear policy choices, not hidden or inconsistent routing.
- All controls should be automatable at 30,000-user scale.
- The solution should avoid creating 30,000 shared mailboxes unless there is a strong operational reason.
- Detection and audit should supplement prevention; they should not be the only control.

## Identity Model

The existing CAPWATCH/Entra pattern is useful:

- Cadet account: `employeeId = <CAPID>`
- Parent account: `employeeId = <CAPID>P`

That relationship can be used by automation to resolve a cadet's parent account. It is already present in this repository's shared PowerShell helpers, where `Get-ParentEmailByCapid` appends `P` to the cadet CAPID and looks up the parent through Microsoft Graph.

Recommended additions for discussion:

- Maintain groups for `Cadets`, `SeniorMembers`, `Parents`, and `CadetProtectionReviewers`.
- Consider setting each cadet account's `manager` attribute to a designated parent or senior-supervisor account if Exchange mail-flow rules will use "sender's manager" actions.
- Maintain a secondary mapping for cadets who do not have a usable parent account.
- Track parent consent and parent account readiness separately from the mere existence of a parent account.

## Recommended Control Stack

### 1. Microsoft Teams Supervised Chat

Microsoft Teams supervised chat is the closest native match for the cadet protection model. It allows designated users with full permissions to supervise chats, while restricted users can only initiate chats with users who have full permissions.

Potential role model:

| User type | Teams supervised chat role | Rationale |
| --- | --- | --- |
| Cadets | Restricted | Cadets need supervised communication. |
| Senior members approved for cadet supervision | Full | They can initiate and supervise conversations. |
| Other staff or helpers | Limited | They can communicate in supervised contexts but cannot initiate unsupervised cadet chats. |
| Parents | Depends on tenant model | Guest users cannot be assigned supervised chat roles; parent accounts may need to be member accounts if they are part of this model. |

Important constraints:

- Supervised chat applies to new private chats created after it is enabled.
- It does not cover existing chats, meeting chats, or channel conversations.
- It must be enabled tenant-wide, not only for a subset of users.
- Guest users cannot be assigned supervised chat roles.
- If a supervising adult leaves, the organization must ensure another full-permission user is added so the chat remains supervised.

Working group decision:

- Should CAP enable Teams supervised chat tenant-wide?
- Are parent accounts guest accounts or member accounts?
- Who is allowed to hold the `Full` role?

Microsoft reference:
https://learn.microsoft.com/en-us/microsoftteams/supervise-chats-edu

### 2. Teams Messaging Policies for Cadets

Create a cadet-specific Teams messaging policy that reduces risky behavior and preserves reviewability.

Recommended cadet settings for consideration:

- Disable user deletion of sent messages.
- Disable editing of sent messages.
- Disable audio messages because Teams audio messages are not captured in eDiscovery reporting.
- Enable "Report inappropriate content" so messages can be routed to communication compliance reviewers.
- Restrict chat behavior in alignment with supervised chat.
- Consider disabling video messages unless there is a strong training need.
- Consider stricter Giphy/meme/sticker settings for cadet contexts.

Working group decision:

- Should cadets be prevented from deleting or editing Teams messages?
- Should audio and video messages be disabled for cadets?
- Should "Report inappropriate content" be mandatory?

Microsoft reference:
https://learn.microsoft.com/en-us/microsoftteams/messaging-policies-in-teams

### 3. Teams and Channel Structure

Prefer Teams channels for routine communication rather than private chat.

Recommended structure:

- Squadron or wing Teams with at least two approved senior owners.
- Channels for operational topics such as announcements, orientation flights, training, logistics, and parent coordination.
- Cadets participate in channels where senior members are present.
- Private and shared channel creation should be disabled for cadets.
- Team ownership should be periodically reviewed to ensure at least two active senior owners remain.

Working group decision:

- Should cadets be allowed to create teams, private channels, or shared channels?
- Which communications should be required to occur in channels rather than private chats?

Microsoft reference:
https://learn.microsoft.com/en-us/microsoftteams/teams-policies

### 4. Teams Meetings and Calls

Meetings and calls need both policy and procedure. Microsoft Teams meeting policies can control organizer and participant capabilities, but they do not directly enforce the full Cadet Protection rule in every live scenario.

Recommended approach:

- Cadets should not be allowed to organize meetings for cadet-program purposes.
- Meetings should be scheduled from unit/team calendars or by approved senior members.
- Meeting chat should be restricted or tied to supervised channels where possible.
- Recording, transcription, lobby, and presenter settings should be reviewed for cadet safety.
- Meeting organizers should be trained that digital meetings follow Cadet Protection expectations.

Working group decision:

- Can cadets organize meetings at all?
- Should meeting chat be disabled or limited for cadets?
- Are recordings required, optional, or prohibited for cadet meetings?

Microsoft reference:
https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-overview

### 5. Exchange Online Mail Flow Controls

Exchange Online mail-flow rules can add recipients to To/Cc/Bcc, redirect messages, reject messages, apply disclaimers, and add the sender's manager as a recipient. This creates several possible email controls, but each has tradeoffs.

#### Option A: Bcc the sender's manager on cadet outbound mail

If each cadet's `manager` attribute is set to a parent or designated senior supervisor, a mail-flow rule can add the sender's manager as a Bcc recipient for outbound cadet mail.

Pros:

- Uses native Exchange behavior.
- Avoids 30,000 per-cadet transport rules.
- Can be applied by group membership.

Cons:

- Only supports one manager target.
- Requires accurate manager assignment.
- "Manager" may not semantically match parent/supervisor in other Microsoft 365 experiences.

#### Option B: Require an approved adult recipient on cadet outbound mail

Mail-flow rules could reject or warn when a cadet sends mail without a parent, senior member, or cadet-protection mailbox copied.

Pros:

- Encourages visible compliance.
- Reduces hidden copying concerns.

Cons:

- Harder to detect every acceptable adult relationship dynamically.
- May block legitimate operational messages unless exceptions are carefully designed.

#### Option C: Copy a central Cadet Protection mailbox

Cadet outbound or inbound messages could be copied to a central mailbox for review.

Pros:

- Simple to implement.
- Does not require one parent mapping per message.

Cons:

- Creates a high-volume review mailbox.
- Does not actually place a parent or second adult into the conversation.
- Needs retention, privacy, and reviewer workflow controls.

#### Option D: Automation-managed parent copy

CAPWATCH automation can use `employeeId + 'P'` to resolve parent accounts and maintain per-cadet forwarding, inbox, or transport-rule structures.

Pros:

- Uses the existing parent relationship.
- Can handle parent-specific routing.

Cons:

- More custom automation.
- Needs careful throttling, idempotent updates, logging, and error handling.
- Exchange transport rules are not well suited to 30,000 one-off dynamic mappings.

Working group decision:

- Is the goal visible adult participation, hidden Bcc supervision, compliance archive, or all three?
- Should parent copies apply to outbound only, inbound only, or both?
- Should mail without adult involvement be blocked, warned, copied, or allowed but audited?

Microsoft reference:
https://learn.microsoft.com/en-us/exchange/security-and-compliance/mail-flow-rules/mail-flow-rule-actions

### 6. Microsoft Purview Communication Compliance

Purview Communication Compliance should be considered the oversight layer. It can detect, capture, and route policy matches from Exchange email, Teams chats and channels, Microsoft 365 Copilot interactions, Viva Engage, and supported third-party imports.

Recommended policies for consideration:

- Adult-to-cadet one-to-one communication indicators.
- Cadet-to-adult messages without another cadet, parent, or senior member present.
- Harassment, threats, profanity, grooming language, adult content, and inappropriate images.
- Requests to move conversations to non-CAP channels such as personal phone, social media, or consumer messaging apps.
- Repeated policy matches by the same user or involving the same cadet.

Recommended reviewer model:

- Create a `CadetProtectionReviewers` group.
- Separate IT administration from message review where possible.
- Use documented escalation paths for serious findings.
- Export audit logs and review metrics periodically.

Working group decision:

- Who is authorized to review captured communications?
- What messages should trigger immediate escalation?
- What is the retention period for reviewed incidents?

Microsoft reference:
https://learn.microsoft.com/en-us/purview/communication-compliance

### 7. eDiscovery, Audit, and Retention

Microsoft Purview eDiscovery can search cloud content including Exchange Online mailboxes, Teams, SharePoint, OneDrive, Microsoft 365 Groups, and Viva Engage. This is the investigation layer for complaints, incident response, and policy validation.

Recommended approach:

- Define retention requirements for cadet communications.
- Ensure Teams and Exchange content remains discoverable.
- Limit eDiscovery permissions to trained staff.
- Document search, export, and evidence-handling procedures.

Working group decision:

- What retention period is required for cadet communications?
- Who can authorize an eDiscovery search?
- What evidence-handling process should apply to Cadet Protection incidents?

Microsoft reference:
https://learn.microsoft.com/en-us/purview/ediscovery-content-search

## Shared Mailbox Alternative

One idea is to give each cadet a shared mailbox and grant parent access. At 30,000 cadets, this is not recommended as the default architecture.

Concerns:

- A shared mailbox is intended for delegated group access, not as a one-mailbox-per-cadet supervision model.
- Parent access to shared mailboxes requires the parent to have a licensed Exchange Online mailbox in the organization.
- Shared mailbox accounts should not be used for direct sign-in.
- 30,000 shared mailboxes create substantial lifecycle, permission, support, and audit overhead.
- Shared mailbox access does not itself enforce the "two adults or two cadets" communication pattern.

Shared mailboxes might still be useful for limited program functions such as `cadetprotection@`, squadron inboxes, or monitored support addresses.

Microsoft reference:
https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes

## Proposed Target Architecture

Recommended starting architecture:

1. Use Entra groups to classify users: cadets, senior members, parents, and reviewers.
2. Enable Teams supervised chat after role assignments are validated.
3. Assign cadets a restrictive Teams messaging policy.
4. Require routine communication to occur in supervised Teams channels.
5. Use Exchange mail-flow rules to add an adult recipient or block non-compliant patterns where practical.
6. Use CAPWATCH automation to resolve parent accounts through `employeeId + 'P'`.
7. Use Purview Communication Compliance for detection, review, escalation, and reporting.
8. Use eDiscovery and retention policies for investigation readiness.

## Suggested Pilot

Pilot with one wing or a representative subset before national rollout.

Pilot scope:

- 500-1,000 cadets
- Parent accounts with valid `employeeId + 'P'`
- Senior members approved for cadet supervision
- One or more cadet-protection reviewer groups
- Teams supervised chat enabled in a test or controlled tenant if possible
- Exchange rules tested in audit/warn mode before blocking

Pilot metrics:

- Number of blocked or prevented unsupervised chats
- Number of adult-copy email actions
- Number of communication compliance alerts
- False positives and false negatives
- Parent/senior member support tickets
- Impact on normal cadet program communication

## Open Questions for the Working Group

1. Should parents be considered approved adult participants for all digital channels, or only for email?
2. Are parent accounts member accounts, guest accounts, or external addresses?
3. Should adult involvement be visible to all recipients, or is Bcc acceptable?
4. Should non-compliant email be blocked, copied, warned, or only audited?
5. Should cadet-to-cadet private chat be allowed without an adult?
6. Should cadets be allowed to initiate Teams meetings or calls?
7. What is the minimum required retention period for cadet communications?
8. Who is allowed to review message content under Communication Compliance?
9. What incident categories require immediate escalation?
10. Is the organization willing to enable Teams supervised chat tenant-wide?

## Implementation Considerations

- Use automation rather than manual per-user configuration.
- Treat CAPWATCH/Entra identity data quality as a dependency.
- Keep all group and policy assignment logic idempotent.
- Log every automated parent/senior mapping decision.
- Build reporting for missing parent accounts, invalid email addresses, missing senior supervisors, and policy assignment failures.
- Start in audit or warning mode where possible before enforcing hard blocks.
- Train senior members and parents on what the controls do and do not guarantee.

## Recommended Next Step

The working group should decide which outcome is required for each channel:

| Channel | Prevent | Copy/add adult | Monitor | Investigate |
| --- | --- | --- | --- | --- |
| Teams private chat | Strong native option through supervised chat | Limited | Yes | Yes |
| Teams channels | Team/channel ownership and policies | Channel membership | Yes | Yes |
| Teams meetings/calls | Partial through policies and procedure | Organizer/process driven | Partial | Yes |
| Exchange outbound email | Partial through mail-flow rules | Stronger if using manager or automation | Yes | Yes |
| Exchange inbound email | Limited native dynamic support | Possible through rules or automation | Yes | Yes |
| Files/OneDrive/SharePoint | Permissions and sharing policies | Not a communication copy model | Yes | Yes |

Recommended first decisions:

1. Confirm whether Teams supervised chat is acceptable tenant-wide.
2. Decide whether Exchange should use parent, senior member, central mailbox, or manager-based supervision.
3. Decide whether the first rollout should be audit-only, warn, or block.
4. Choose a pilot population and success metrics.
