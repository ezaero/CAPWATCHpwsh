# List Distribution Lists

## Summary

This function performs monthly maintenance tasks for CAPWATCHSyncPWSH, including:

- **Expired Member Account Deletion:**
  - Automatically deletes Azure AD and Exchange Online accounts for members whose CAPWATCH membership has expired.
  - Removes parent guest accounts associated with expired members.
  - Cleans up O365 accounts whose CAPIDs are not present in the current CAPWATCH member list (with exceptions handled).
  - **Special Accounts:** Accounts with a CAPID of `999999` are considered exceptions and are not deleted by this function.
- **Log File Cleanup:**
  - Deletes log files older than 30 days from the logs directory to conserve space and maintain audit hygiene.

### Schedule
- The maintenance function is triggered automatically on the 3rd day of each month.

### Prerequisites
- Requires access to CAPWATCH data files and Microsoft Graph/Exchange Online permissions.

### Purpose
- Ensures that Office 365 and Azure AD remain accurate and free of stale accounts, reducing security risks and administrative overhead.
