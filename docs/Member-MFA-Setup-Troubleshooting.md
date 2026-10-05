# Member MFA Setup and Troubleshooting Guide

Prepared: 2026-10-02

## Purpose

This guide helps COWG members and guests set up Microsoft 365 MFA and resolve common sign-in problems without needing one-off administrator help.

## Before You Start

You need:

- Your COWG Microsoft 365 account or guest invitation email.
- A phone or device you control.
- The Microsoft Authenticator app, unless you were instructed to use a passkey/security key.
- Access to a web browser.

If you are an external guest, accept the invitation using the same email address where you received the invitation.

## Recommended Setup

### Option 1: Microsoft Authenticator

Use this for most members and guests.

1. Install Microsoft Authenticator from your device's app store.
2. Sign in to Microsoft 365 or open the guest invitation link.
3. When prompted for more information, choose Microsoft Authenticator.
4. Follow the QR code or app-based setup prompts.
5. Approve the test notification.
6. Keep the app installed while you need COWG access.

### Option 2: Passkey or Security Key

Use this when directed by IT or when you already use passkeys.

1. Go to `https://mysignins.microsoft.com/security-info` and sign in.
2. Choose to add a passkey/security key.
3. Follow the prompts for your device, browser, or hardware key.
4. Confirm the method appears in your security info.

### Backup Method

If available, add a second method so you do not lose access if you replace your phone.

Acceptable backup methods include:

- Another Microsoft Authenticator installation.
- Authenticator one-time code.
- A passkey or security key.

Avoid SMS or voice unless no stronger option works for your situation.

## Add Another MFA Method to Your Account

Use this when your account already works and you want to add another way to verify your identity, such as a new phone, Authenticator app, passkey, security key, or backup method.

### Standard Self-Service Steps

1. Go to:

   ```text
   https://mysignins.microsoft.com/security-info
   ```

2. Sign in with your COWG Microsoft 365 account or the external email address used for your guest invitation.
3. Complete MFA using one of your existing methods.
4. Select **Add sign-in method**.
5. Choose the method you want to add.
6. Follow the on-screen prompts.
7. After setup completes, confirm the new method appears on the **Security info** page.
8. If this replaces an old phone or method, test the new method before deleting the old one.

### Add Microsoft Authenticator

1. Install Microsoft Authenticator on your phone.
2. On the **Security info** page, select **Add sign-in method**.
3. Choose **Authenticator app**.
4. Follow the prompts to scan the QR code or pair the app.
5. Approve the test notification.

Use Microsoft Authenticator as the preferred method for most members and guests.

### Add a Passkey or Security Key

1. On the **Security info** page, select **Add sign-in method**.
2. Choose **Passkey**, **Passkey in Microsoft Authenticator**, or **Security key**, depending on what is available.
3. Follow the browser/device prompts.
4. If using a hardware key, have the key physically available and know whether it is USB, NFC, or another supported type.
5. Name the key or passkey clearly so you can recognize it later.

Passkeys and security keys are preferred for privileged/admin accounts when practical.

### Add SMS or Phone as a Backup

SMS and voice are enabled today, but they are fallback methods rather than the preferred standard.

1. On the **Security info** page, select **Add sign-in method**.
2. Choose **Phone** if available.
3. Enter the phone number.
4. Choose text message or call, if prompted.
5. Complete the verification code or call approval.

Do not use SMS or voice as your only method if you can use Microsoft Authenticator or a passkey.

### If You Cannot Sign In to Add a Method

If you lost your phone, replaced your device, or cannot complete MFA:

1. Contact your event sponsor or COWG IT.
2. Provide your name, email address, event/Team needed, and a screenshot or exact error message.
3. An administrator may verify your identity and issue a Temporary Access Pass.
4. Use the Temporary Access Pass to sign in.
5. Immediately add Microsoft Authenticator, passkey, or another approved method.
6. Confirm the new method works before ending the support session.

## Common Problems

### I Do Not See the Invitation

Check:

- Spam or junk mail.
- The exact email address the sponsor gave IT.
- Whether your organization blocks external Microsoft invitations.

Ask your sponsor to confirm the email address in `ManualGuests.csv`.

### I Accepted the Invitation but Teams Still Does Not Work

Possible causes:

- The account was created, but Team membership has not synced yet.
- You signed into Teams with the wrong account.
- You need to switch organizations/tenants in Teams.

Try:

1. Sign out of Teams completely.
2. Sign in with the same email address used for the invitation.
3. In Teams, switch to the COWG tenant/organization if prompted.
4. Wait a few minutes and try again.
5. Contact your sponsor if the Team is still missing.

### I Got a New Phone

If your old phone is still available:

1. Sign in to your security info.
2. Add Microsoft Authenticator on the new phone.
3. Confirm the new method works.
4. Remove the old phone.

If your old phone is not available, contact COWG IT or your event sponsor. An administrator may need to issue a Temporary Access Pass after verifying your identity.

### Authenticator Prompts Are Going to the Wrong Device

Try:

1. Open your security info.
2. Check which devices are registered.
3. Add your current device.
4. Remove old or unknown devices.

If you cannot sign in to reach security info, contact support.

### I Am Stuck in a Sign-In Loop

Try:

1. Open a private/incognito browser window.
2. Go directly to Microsoft 365 or Teams.
3. Sign in with the invited email address.
4. Clear old saved Microsoft credentials if the browser keeps choosing the wrong account.
5. Try the Teams web app before troubleshooting the desktop app.

### I See "You Cannot Access This Right Now"

This usually means a policy, device, account, or tenant-selection issue is blocking access.

Send the following to your sponsor or IT:

- Screenshot of the error.
- Time and date of the failed sign-in.
- Email address used.
- Whether you are using browser, desktop Teams, or mobile Teams.
- Whether you can reach other Microsoft 365 resources.

## Escalation Information to Collect

Collect this before escalating:

- User name and email address.
- Sponsor.
- Team/resource needed.
- Screenshot or exact error text.
- Approximate sign-in time with time zone.
- Device type and browser/app.
- Whether this is first-time setup, phone replacement, or recurring failure.
