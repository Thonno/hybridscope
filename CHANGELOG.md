# Changelog

## 1.0.0 - 2026-11-17
First public release.

- GUI (WPF) with RVTools-style tabs: overview, users, one tab per group/OU, on-prem mailboxes, Exchange Online mailboxes, OneDrive, MFA, tenant licenses, attention list, history
- Sources: Active Directory, Exchange Server (Remote PowerShell or local snap-in), Microsoft Graph
- Mailbox location detected from Exchange (`Get-Mailbox` / `Get-RemoteMailbox`) and Graph (`provisionedPlans`), not from mailbox size
- Real MFA status (registered methods, default method, per-user state)
- Windows or explicit domain credentials, works from non-domain-joined PCs
- Export to XLSX (formatted) and CSV, per tab or all tabs
- Headless mode `-Auto` with certificate-based sign-in for scheduled reports and history tracking
- English and Italian UI
