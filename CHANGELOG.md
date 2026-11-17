# Changelog

## 1.4.0 - 2026-11-17
- **Command line without the GUI**: `-Only MFA | Licenses | Mailbox365 | OneDrive | MailboxOnPrem` collects a single dataset and contacts only the source it needs; `-PassThru` returns objects instead of writing files (English values, pipeline-friendly); `-Mode` overrides the infrastructure mode for one run
- `-Auto` / `-Only` reuse an existing `Connect-MgGraph` session and leave it open
- Data collection split into one function per source (`Get-HSEntraUser`, `Get-HSMfaRegistration`, `Get-HSMailbox365Usage`, `Get-HSOneDriveUsage`, `Get-HSExchangeMailbox`, ...); `Get-HSData` now only merges
- Fix: the Excel OVERVIEW sheet lost its formatting (section rows, GB number format, bold TOTAL) and the "Generated on" / legend / warning lines overwrote metric names from row 4 down. Cause: `$M` (metric list) and `$m` (loop variable) are the same variable in PowerShell
- Shared state moved from `$global:HS_*` to `$script:HS_*`: nothing is left in the caller's session
- Single-letter variables renamed to descriptive names
- Full comment-based help (`Get-Help .\HybridScope.ps1 -Full`) with examples

## 1.3.0 - 2026-11-17
- **Step indicators** under the connection panels: each step (Connect 365, Connect Exchange, Verify AD, Load data) shows its number, then turns amber while running, green with a check mark when done, red on error
- "Load data" button restyled: stays readable while disabled
- Tab names keep their underscores (e.g. MAILBOX_ONPREM, TENANT_LICENSES)
- Screenshots added to the README (sensitive data blurred)

## 1.2.0 - 2026-11-17
- **Infrastructure modes** at the top of the window: Full Cloud, Full Hybrid, Hybrid (365 + AD), On-Prem (Exchange + AD), AD only
- Every source is now optional: Microsoft 365 is no longer required (works with Exchange 2016/2019 + AD only)
- Connection panels, columns, tabs, overview and attention list adapt to the selected mode
- New `Mode` setting in `config.json` (starting mode and mode used by `-Auto`)
- The ActiveDirectory module is only requested when the selected mode uses AD

## 1.1.0 - 2026-11-17
- **Change account** button: signs out of Microsoft 365 and clears the cached sign-in; the sign-in is no longer persisted between sessions
- **Stop** button to cancel a running operation (data load, connections, exports)
- **Group by** selector in the GUI: top-level OU, parent OU, no grouping, custom rules (`GroupBy = Rules`)
- `GroupRules` are now applied only when `GroupBy` is `Rules`

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
