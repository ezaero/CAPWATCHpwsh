# COWG Microsoft 365 MFA Standard

Prepared: 2026-10-02

## Purpose

This standard defines the expected multi-factor authentication (MFA) posture for Colorado Wing Microsoft 365 accounts, including normal member accounts, guest accounts, and privileged administrator accounts.

The goal is to reduce account compromise risk while keeping onboarding repeatable for members, encampment staff, and administrators.

## Scope

This standard applies to:

- COWG Microsoft 365 member accounts.
- COWG Microsoft 365 guest accounts created for approved mission support.
- Privileged accounts used to administer Microsoft 365, Microsoft Entra ID, Exchange Online, Teams, Azure, automation, or security controls.

## Definitions

- **Normal account**: A standard user account used for email, Teams, SharePoint, OneDrive, and member collaboration.
- **Guest account**: A Microsoft Entra B2B account invited into the COWG tenant for a user outside the organization.
- **Privileged account**: Any account with administrative roles, elevated Microsoft Graph/API permissions, ownership of automation credentials, or standing access that can materially change tenant security or user access.
- **Phishing-resistant MFA**: Authentication that resists credential replay and remote phishing, such as passkeys/FIDO2, Windows Hello for Business, platform credentials, or certificate-based authentication when configured.
- **Temporary Access Pass (TAP)**: A time-limited Microsoft Entra passcode used by administrators to bootstrap or recover authentication methods.

## MFA Requirement

All COWG Microsoft 365 accounts must use MFA when accessing COWG resources unless an explicit written exception is approved by the COWG IT authority.

MFA exceptions must be:

- Time limited.
- Documented with a business reason.
- Reviewed at least monthly while active.
- Removed as soon as the blocking condition is resolved.

## Observed Current Tenant Configuration

Observed with Azure CLI and Microsoft Graph on 2026-10-02 against tenant `Civil Air Patrol Colorado` (`cowg.cap.gov`).

### Conditional Access Policies Observed

| Policy | State | Scope | Grant/session behavior |
| --- | --- | --- | --- |
| `Require MFA for admins` | Enabled | Selected privileged directory roles, with one excluded user | Requires MFA. Persistent browser always enabled. Sign-in frequency is 180 days for primary and secondary authentication. |
| `Require MFA for external and guest users` | Enabled | `GuestsOrExternalUsers`, with one excluded user | Requires MFA. |
| `Require MFA for internal users (admins not included) - Basic` | Enabled | All users except guest/external users, one excluded user, and selected admin roles | Requires MFA. |
| `Block all legacy sign-ins that don't support MFA` | Enabled | All users, with one excluded user | Blocks Exchange ActiveSync and other legacy client app types. |
| `Require Microsoft Authenticator if login outside of United States` | Enabled | All users, all apps, all locations except named location `United States` | Requires custom authentication strength `Microsoft Authenticator`; allowed combinations observed were `fido2`, `deviceBasedPush`, and `password,microsoftAuthenticatorPush`. Sign-in frequency is every time. |
| `Require MFA for All Users` | Report-only | All users, all apps | Built-in `Multifactor authentication` strength. Not enforced. |
| `Microsoft-managed: Require phishing-resistant multifactor authentication for admins` | Report-only | Selected privileged admin roles, with one excluded user | Built-in `Phishing-resistant MFA` strength. Not enforced. |
| `Block Phishing Hacker` | Report-only | One included user; no cloud apps selected | Block control present but not enforced. |

### Named Locations Observed

| Named location | Type | Values |
| --- | --- | --- |
| `United States` | Country/region | `US` |
| `Afganistan` | Country/region | `AF` |
| `Phishing Hacker` | IP ranges | `104.168.115.0/24`, `185.92.209.0/24`, `69.167.30.158/32` |

### Authentication Methods Policy Observed

Observed from the Microsoft Entra admin center on 2026-10-02.

| Method | Target | Enabled | Notes |
| --- | --- | --- | --- |
| Passkey (FIDO2) | All users | Yes | Supports the privileged-account target standard and can be used by capable members. |
| Microsoft Authenticator | All users | Yes | Registration optional. Authentication mode observed as `Any`. This is the preferred day-to-day method. |
| SMS | All users | Yes | Registration optional. First-factor sign-in is enabled. Microsoft-provided SMS authentication retires February 1, 2027. |
| Temporary Access Pass | All users | Yes | Registration optional. TAP is still administrator-issued and should be treated as bootstrap/recovery, not routine sign-in. |
| Hardware OATH tokens (Preview) | All users | Yes | Registration optional. Use only when there is an administrative process to issue and track tokens. |
| Software OATH tokens | All users | Yes | Registration optional. Useful fallback, but not the preferred standard. |
| Voice call | All users | Yes | Registration optional. Microsoft-provided voice authentication retires February 1, 2027. |
| Email OTP | All users, excluding one group | Yes | Fallback/guest-adjacent method; not the preferred standard for COWG-managed accounts. |
| Certificate-based authentication | Not shown as targeted | No | Not part of the current COWG standard. |
| Verified ID | Not shown as targeted | No | Not part of the current COWG sign-in standard. |
| QR code | Not shown as targeted | No | Not part of the current COWG standard. |

### Observation Limits

The account/token used for observation could read Conditional Access policies and named locations. Microsoft Graph denied direct reads of:

- `policies/authenticationMethodsPolicy`
- `policies/identitySecurityDefaultsEnforcementPolicy`
- `identity/conditionalAccess/authenticationStrength/policies`

Because of that limitation, the authentication methods policy above was captured from the Microsoft Entra admin center rather than directly from the Graph endpoint.

## Current-State Interpretation

The current enforced configuration already requires MFA for:

- Internal non-admin users.
- External and guest users.
- Privileged/admin roles.

The tenant also blocks legacy sign-ins that cannot satisfy MFA.

The current enforced admin policy requires MFA, but it does not enforce phishing-resistant MFA for administrators. A Microsoft-managed phishing-resistant MFA policy for admins exists in report-only mode. This means the recommended privileged-account standard below is stronger than today's enforced state and should be treated as a target-state improvement rather than a description of current enforcement.

The current outside-United-States policy is stricter than generic MFA because it requires the custom `Microsoft Authenticator` authentication strength and re-prompts every time.

The current authentication methods policy is broad. Microsoft Authenticator, passkey/FIDO2, TAP, SMS, voice, and OATH methods are enabled for all users. This means COWG's present enforcement model is best described as **MFA required, with multiple methods allowed**. It is not currently "Microsoft Authenticator required" except for sign-ins covered by the outside-United-States Conditional Access policy.

For Winter Encampment and other bulk guest onboarding, this broad method availability may reduce onboarding friction. After the event, COWG should revisit legacy methods, especially SMS first-factor sign-in, SMS MFA, and voice call MFA, because Microsoft-provided SMS and voice authentication retire on February 1, 2027.

## Approved Authentication Methods

### Preferred Methods

Use these methods whenever practical:

- Passkey in Microsoft Authenticator.
- Passkey/FIDO2 security key.
- Windows Hello for Business.
- Platform credential for macOS, where supported.
- Microsoft Authenticator push notification with number matching.

### Bootstrap and Recovery Methods

These methods are for setup, recovery, and temporary support:

- Temporary Access Pass (TAP).
- Microsoft Authenticator one-time code.
- Software OATH token, when Authenticator push is not available.

### Restricted or Legacy Methods

The following methods are enabled today but should not be the normal standard:

- SMS text message.
- Voice call.
- Email one-time passcode.

Use these only when no stronger method is workable and the user has a documented support need. Privileged accounts must not rely on SMS, voice, or email OTP as their primary MFA method.

SMS first-factor sign-in is currently enabled for all users. COWG should avoid documenting SMS first-factor as a supported onboarding pattern and should review whether to disable or scope it after Winter Encampment onboarding is complete.

## Normal Account Standard

Normal member and guest accounts must have at least one approved MFA method registered before they are considered fully onboarded.

Recommended baseline:

1. Microsoft Authenticator app with number matching or passkey.
2. A backup method where appropriate, such as an Authenticator one-time code or second registered device.
3. Security info kept current by the user.

For guests, the sponsor or requesting staff member is responsible for making sure the guest receives setup instructions and knows where to request help.

During high-volume event onboarding, SMS and voice may remain available as fallback methods to prevent blocking users who cannot immediately configure Authenticator. Support instructions should still direct users to Microsoft Authenticator or passkey/FIDO2 first.

## Privileged Account Standard

Privileged accounts require a stronger standard because compromise creates tenant-wide risk.

Privileged accounts must:

- Use phishing-resistant MFA wherever technically possible.
- Have at least two registered strong authentication methods.
- Avoid SMS, voice, and email OTP except as a temporary break-glass recovery measure.
- Use separate admin accounts when practical instead of performing administrative work from day-to-day member accounts.
- Be reviewed periodically for role assignment, sign-in activity, and authentication method readiness.

Privileged accounts should use one or more of:

- FIDO2/passkey security key.
- Passkey in Microsoft Authenticator.
- Windows Hello for Business on a managed and protected device.
- Certificate-based authentication if COWG adopts and manages the certificate lifecycle.

## Temporary Access Pass Use

Administrators may issue a TAP when:

- A new user cannot complete MFA registration.
- A user lost access to all registered MFA methods.
- A privileged user needs to register a phishing-resistant method.
- A guest account invitation was redeemed but the user cannot complete security setup.

TAP handling rules:

- Use the shortest practical lifetime.
- Prefer single-use TAP when possible.
- Verify the user's identity before issuing the TAP.
- Communicate TAP securely and directly to the user.
- Document the reason in the support ticket or access log.
- Confirm the user registered an approved method after TAP use.

## Guest and Encampment Staff Accounts

One-off external users must be tracked in `ManualGuests.csv` at the project root.

Each row must include:

- External email address.
- Name.
- Purpose or role in `Grade`.
- Team assignment in `Teams`.
- Sponsor/approver in `Sponsor`.
- `Enabled=TRUE` while access is approved.

The sponsor is accountable for the business reason for access. IT is accountable for technical account creation, MFA readiness guidance, and Teams assignment.

## Troubleshooting Responsibilities

Members and guests should first use the member-facing MFA setup guide and troubleshooting checklist.

Administrators should:

1. Confirm the user is signing in with the correct account.
2. Confirm the account exists and is enabled.
3. Confirm the user has a registered authentication method.
4. Use TAP only when normal self-service recovery is not sufficient.
5. Escalate repeated or unexplained failures for sign-in log review.

## Review Cycle

This standard should be reviewed:

- Before major events requiring bulk guest onboarding.
- After any account compromise or significant MFA support incident.
- At least annually.

## References

- Microsoft Entra authentication methods overview.
- Microsoft Entra multifactor authentication overview.
- Microsoft Entra Temporary Access Pass guidance.
- Microsoft Entra passkeys/FIDO2 guidance.
- Microsoft guidance on SMS and voice retirement.
