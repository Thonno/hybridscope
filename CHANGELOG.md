# Changelog

## 1.8.0 - 2026-11-26
- **Double-click details on every tab**: a number in OVERVIEW opens the list of the users behind it (group column or TOTAL; e.g. *MFA per-user enforced*, a license, inactive users, Defender), a license in TENANT_LICENSES lists who has it, a fix in SECURITY_FIXES lists the users of the PCs it fixes, and any other row (HISTORY, DEVICES_UNASSIGNED, ...) opens in a window with all its fields. The user lists have search, column filters, *Copy* and double-click for the user card. **Saved reports** (also made by someone else, in either language): the lists are recomputed when the report is opened, from its USERS and SECURITY tabs with the same rules and the report's date; a number stays without list only if the report lacks the columns it needs (e.g. a preset export)
- **Column filters** in every grid: click the ▾ next to a column name to choose the values to show (search, select all, blanks), Excel-style; a blue ▼ marks the filtered columns, the filters add up with the search box and the row count follows them
- OVERVIEW columns can no longer be sorted: a click on a header used to mix sections and rows in alphabetical order
- **Update check at startup**: when the GUI opens, HybridScope asks GitHub for the latest release in the background and shows a message only if a newer version exists (offline / up to date: nothing). New setting `CheckUpdatesAtStartup` (default `true`); not used by `-Auto`

## 1.7.0 - 2026-11-24
- **Security (Microsoft Defender)**: new **Defender** box next to *Connect 365* (setting `Defender`). With the FULL content, Load data and `-Auto` also read Microsoft Defender for Business / Defender for Endpoint (read-only) for the users with a Defender license
- Users are filtered by the **service plans** of their licenses (`DefenderLicensePlans`, default `MDE_SMB`, `MDE_LITE`, `WINDEFATP`): users with other licenses (e.g. Teams Essentials, Business Standard) and the PCs used only by them do not appear
- PCs are linked to users through the **users logged on** in Defender (`DOMAIN\user` matched with AD / Entra): no Intune or Entra join needed. Main user = latest interactive logon; logons older than `DefenderDays` (default 30, from the PC's last contact) are ignored
- New tabs: **SECURITY** (one row per licensed user and PC: OS, last seen, sensor, risk, exposure, antivirus, open alerts, vulnerabilities per severity, exploitable, missing KBs, software to update, software out of support, weak configurations, notes; users without a device get one row), **VULNERABILITIES** (PC + CVE from `DefenderMinSeverity`, default High: CVSS, exploit, software, fix), **SECURITY_FIXES** (KBs, software updates, software to replace and weak configurations ordered by priority, with the PCs they fix), **DEVICES_UNASSIGNED** (only with `DefenderUnassignedDevices`: onboarded devices not linked to a licensed user, with the reason and the accounts)
- OVERVIEW section **SECURITY (DEFENDER)**, ATTENTION categories 8 (exploitable or critical vulnerabilities) and 9 (Defender license without a device), one Defender section per PC in the user card, exposure score and Secure Score for devices in the report header
- Defender is read through the **HybridScope app registration** (one per tenant, `ClientId`): in the GUI every user signs in in the browser with their own account (delegated, no certificate or secret; needs a Defender role such as Security Reader); `-Auto` with a certificate reads it app-only. MSAL from Microsoft.Graph.Authentication, no new modules
- `tools\New-HybridScopeApp.ps1`: new `-Defender` (public client with redirect `http://localhost`, delegated WindowsDefenderATP permissions + admin consent, `Defender = true`, no certificate) and `-Certificate`; reuses the app already in `config.json` instead of creating a new one; application Defender permissions for `-Auto` when Defender is set up. README: how the sign-in works, who does what, setup by script and by hand
- Optional parts (antivirus, vulnerability export, software, configurations, alerts, scores) are skipped with a warning when the plan or a permission does not allow them; without the vulnerability export the basic list is used. Defender API limits are respected (pauses, retry on 429)
- A preset / Custom load skips Defender (it fills its own tabs, not user columns)
- Logon matching: `DOMAIN\user` with AD, `AzureAD\name` (Entra-joined PCs) with the UPN prefix or the display name; new log line that explains the devices (used by licensed users / only by users without a license / accounts not found / no interactive logon) to spot PCs that are not onboarded or not matched
- Logons of type Unlock or without a type are counted; PCs joined to Entra ID without logons in Defender fall back to their Entra owner / registered user (Graph `Device.Read.All`, optional scope of the GUI, application permission added by `New-HybridScopeApp.ps1`); new log line with the logons read; DEVICES_UNASSIGNED shows the raw logons when none is usable
- Last resort `DefenderMatchByName` (default on): a PC with no logons and no Entra owner (e.g. a Mac) is linked to the user whose username or unique surname is in the PC name, with a note
- User card: for users with a PC in Defender the right side is split - groups and roles on top, **Defender: what to fix** below (resizable): one line per PC with its state, and one row per KB, software update, software out of support and weak configuration, with severity and exploit, exploitable first. Included in *Copy*
- Fixed (since 1.6.0): MFA status, MFA % and the "no 365 size" note were empty for every AD user (a local variable hid the load parts): the OVERVIEW showed 0% MFA and everyone "Not in Entra"
- PC-name matching prefers users with a Defender license, then enabled users, then the longest match
- Fixed: **copy with headers** (Ctrl+C) wrote `System.Windows.Controls.TextBlock` instead of the column names, in every tab
- Recommended software version no longer repeated when it is already in the recommendation name
- README (EN / IT) rewritten: shorter overview, Microsoft Defender in the contents with checklist (onboarding, app, role), setup by script and by hand, troubleshooting table; practical guides for the GUI and for PowerShell / `-Auto`; configuration grouped by topic. Website: new Security section

## 1.6.0 - 2026-11-22
- **Export content** box next to the connection panels: choose **before Load data** what Export XLSX / CSV will write: **FULL** (every tab and column, as before), a **preset** or **Custom** columns. **Columns...** opens the column chooser (columns of the selected infrastructure mode, or of the loaded data); the box shows what will be exported, and the export buttons write that choice directly
- Built-in presets for the user list: *Sign-ins and inactivity*, *Licenses*, *Mailbox migration*, *MFA security*, *Directory*
- **Custom**: one checkbox per column of the USERS tab (only the columns of the current infrastructure mode), All / None; changing a preset column switches to Custom. **Save as preset...** / **Delete preset** keep your own presets in `config.json` (`ExportPresets`)
- Options for presets and custom exports: one sheet per group, include OVERVIEW, enabled accounts only. A CSV of the user list alone is written as a single file
- **Dynamic load** - "Read only the data needed" (default on, setting `ExportFastLoad`): with a preset or Custom, Load data and `-Auto` read only what the columns need (e.g. *Sign-ins and inactivity* skips Exchange mailboxes, usage reports, MFA and the sync check). The top bar shows **PARTIAL LOAD**, tabs and overview sections without data are hidden, exporting columns that were not read asks first, and a partial `-Auto` run does not update `history.csv`
- The last choice is remembered (`ExportPreset`, `ExportColumns`, `ExportGroupSheets`, `ExportOverview`, `ExportEnabledOnly`)
- Command line: `-ExportPreset <name>` and `-Columns <list>` (e.g. `-Auto -Columns Username,AD_LastLogon,M365_LastSignIn`); both imply `-Auto` and also shape the `-PassThru` objects. New setting `AutoExportPreset` (default `Full`) for scheduled reports; a preset report is written as `HybridScope_<preset>_<timestamp>.xlsx` and still copied to `HybridScope_LATEST.xlsx`. `history.csv` is always computed on the full data
- **Icon**: HybridScope has its own icon (`tools\HybridScope.ico`, 16-256 px): on the exe, on the main window and taskbar (also when started with `Start-HybridScope.cmd`) and on the user card, export and About windows
- `tools\Build-Exe.ps1`: uses the icon and writes the script version (e.g. 1.6.0.0), company and copyright in the exe file properties. The exe contains the script: rebuild it after every update

## 1.5.0 - 2026-11-19
- **User card**: double-click a user in any tab (USERS, groups, MFA, mailboxes, ATTENTION, SYNC...) to open one window with identity, AD, Microsoft 365, mailbox (on-prem and Exchange Online details), OneDrive, MFA, sync data and notes. Values can be selected; **Copy** puts the whole card on the clipboard
- The card loads the user's **groups and roles** in the background: AD groups (direct, nested via `LDAP_MATCHING_RULE_IN_CHAIN`, primary group) and Entra groups (Microsoft 365 / security / mail-enabled security / distribution, dynamic, synced or cloud) and Entra roles, with direct vs inherited membership and a search box
- New delegated Graph permissions for the card: `GroupMember.Read.All`, `RoleManagement.Read.Directory` (read-only). Without admin consent **Connect 365** falls back to the previous permissions and the card shows AD groups only
- **SYNC tab** (modes with AD + Microsoft 365): AD <-> Entra ID check based on the source anchor (`ms-DS-ConsistencyGuid` / `objectGUID` vs `onPremisesImmutableId`): last sync cycle and stale sync, ImmutableId that differs from the AD anchor, synced accounts missing from AD, leftover ImmutableId on cloud accounts, AD accounts that will soft-match or conflict with a cloud account, Entra provisioning errors, UPNs that fell back to `onmicrosoft.com` (a UPN taken from `mail` / Alternate Login ID is not flagged), duplicate anchors in AD. Other domains / forests and accounts outside `SearchBase` are not flagged; custom source anchors are detected and the anchor checks skipped
- Sync problems also appear in the USERS / group tabs (`SyncIssues` column), in ATTENTION (category 7) and in the overview
- AD and Entra accounts are paired by sync anchor first, names as a fallback. An AD account paired by name with a cloud account that is not synced from AD now shows Origin **AD + cloud (not linked)** instead of "AD synced"
- `-Only Sync` collects the sync check alone (AD + Graph)
- New setting `SyncStaleHours` (default 3)
- **Open a saved report**: File > Open report (Ctrl+O) or drag & drop. Reads a full or single-tab XLSX, a `HybridScope_CSV_*` folder or a single CSV, in English or Italian (headers and tab names are mapped back), restores dates and numbers, and shows it like live data: tabs, search, overview, re-export, user cards (groups and roles are not read for an imported report)
- Fix: Excel tables were not created with ImportExcel 7.8 (`T1` reads as a cell address, so ImportExcel skipped the table: no style and no header filters). Tables are now named `HS_1`, `HS_2`...
- **Menu bar**: File (exports, reports folder, exit) and Help (documentation with F1, **Check for updates**, report a problem, sponsor links, **About**)
- **Check for updates** compares the running version with the latest GitHub release (public API, only when clicked; works behind authenticated proxies)
- **About** window: version, author (Thonno) and project links, license, support links, environment info. The status bar shows the version and a link to the author
- AD collection moved to `Get-HSADUser`; the AD read also returns `proxyAddresses` and `mS-DS-ConsistencyGuid` (falls back to `objectGUID` on very old schemas)

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
