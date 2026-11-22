<div align="center">

<img src="docs/icon.png" alt="HybridScope" width="96">

# HybridScope

**One view of your hybrid identities: Active Directory + Exchange on-premises + Microsoft 365.**

Who is on-prem and who is already in the cloud, how much space every mailbox uses, who has a license, who has MFA, who hasn't signed in for months: in one window, exportable to Excel.

[![Sponsor on GitHub](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=githubsponsors&logoColor=white)](https://github.com/sponsors/Thonno)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/aascione956)
[![Star on GitHub](https://img.shields.io/github/stars/Thonno/hybridscope?style=for-the-badge&logo=github&label=Star)](https://github.com/Thonno/hybridscope/stargazers)

![PowerShell 5.1](https://img.shields.io/badge/PowerShell-5.1-5391FE?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011%20%7C%20Server-0078D6?logo=windows&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Languages](https://img.shields.io/badge/UI-English%20%7C%20Italiano-blue)

**English** · [Italiano](README.it.md)

<img src="docs/07-overview.png" alt="HybridScope overview" width="900">

</div>

---

## Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Requirements](#requirements)
- [Quick start](#quick-start)
- [Using the GUI](#using-the-gui)
- [Tabs and columns](#tabs-and-columns)
- [Configuration (config.json)](#configuration-configjson)
- [Permissions](#permissions)
- [Command line (no GUI)](#command-line-no-gui)
- [Scheduled reports (-Auto)](#scheduled-reports--auto)
- [Troubleshooting](#troubleshooting)
- [FAQ](#faq)
- [Roadmap](#roadmap)
- [Support the project](#support-the-project)
- [Thanks](#thanks)
- [License](#license)

## Features

- **RVTools-style GUI**: every dataset in its own tab, a search box that filters the current tab, sortable columns, copy with headers.
- **Three sources in one report**:
  - **Active Directory**: users, OUs, enabled/disabled, last logon, last password change
  - **Exchange Server on-premises**: mailboxes, size, item count, database, quota, last access, and remote mailboxes that were already migrated
  - **Microsoft 365 (Microsoft Graph)**: licenses, last sign-in, Exchange Online mailboxes, OneDrive, MFA registration and methods, per-user MFA state, purchased and free licenses
- **Reliable mailbox location**: `ON-PREM` / `365` / `NO MAILBOX` comes from Exchange (`Get-Mailbox`, `Get-RemoteMailbox`) and Graph (`provisionedPlans`), never from the mailbox size.
- **Real MFA status**: whether at least one method is registered, which methods, the default one, and the legacy per-user state.
- **User card**: double-click a user in any tab to see AD, mailbox, OneDrive, MFA, licenses and sync data in one window, together with their **groups and roles**: AD groups (direct, nested and primary) and Entra groups (Microsoft 365, security, distribution, dynamic) and Entra roles.
- **Sync check (SYNC tab)**: compares AD and Entra ID the way Entra Connect / Cloud Sync link them, and lists what is broken or about to break: `ImmutableId` that no longer matches the AD anchor, synced accounts missing from AD, cloud accounts with a leftover `ImmutableId`, AD accounts that will soft-match or conflict with a cloud account at their first sync, provisioning errors, UPNs that fell back to `onmicrosoft.com`, and when the last sync ran.
- **Overview** per group (OU) with totals, plus an **Attention** list:
  - on-prem mailboxes to migrate, sorted by size
  - licensed users without MFA
  - licenses assigned to disabled accounts
  - inactive accounts
  - AD accounts that are not synced
  - sync problems
- **Export**: formatted XLSX or CSV. **FULL** (every tab), a **preset** (sign-ins, licenses, migration, MFA, directory) or **custom**: tick only the columns you want. Save your own presets. Or only the current tab with the filter applied.
- **Open a saved report** (File > Open, Ctrl+O, or drag it onto the window): an XLSX report, a CSV export folder or a single CSV, in either language. Browse, search, re-export and open user cards without connecting to anything, e.g. the daily `HybridScope_LATEST.xlsx` on another PC.
- **Works from a PC outside the domain**: either Windows credentials or a domain user and password entered manually, with Kerberos.
- **Infrastructure modes**: *Full Cloud*, *Full Hybrid*, *Hybrid*, *On-Prem* or *AD*, selectable with one click. The app shows and merges only the sources you have.
- **Grouping in tabs** chosen from the window: top-level OU, the OU that contains the user, a single list, or custom rules.
- **Step indicators**: Connect 365 › Connect Exchange › Verify AD › Load data turn green with a ✓ when done, amber while running and red on error.
- **Change account** button to sign out of Microsoft 365 and clear the saved sign-in, and a **Stop** button to cancel a running load.
- **Command line**: `-Auto` for the full report, `-Only MFA` (or `Licenses`, `Mailbox365`, `OneDrive`, `MailboxOnPrem`, `Sync`) for a single dataset, `-PassThru` to get objects in the pipeline. Certificate-based sign-in for scheduled reports. Keeps a `HybridScope_LATEST.xlsx` copy and a history of trends (on-prem users, space, % MFA).
- **File** and **Help** menus: open a report, exports, reports folder, documentation (F1), **Check for updates** (compares with the latest GitHub release), problem reports, **About** window.
- **English / Italian** interface.
- **Read-only**: HybridScope never changes anything in AD, Exchange or Microsoft 365.

## Screenshots

**Overview**: counts per OU / group, mailbox location, space, licenses, MFA and inactivity at a glance.

![Overview](docs/07-overview.png)

**Users**: one row per user, merging AD, Exchange and Microsoft 365. Search, sort, export.

![Users](docs/08-users.png)

**Infrastructure modes**: one click and the window shows only the connections you need.

| Full Hybrid | Full Cloud |
|---|---|
| ![Full Hybrid](docs/01-full-hybrid.png) | ![Full Cloud](docs/02-full-cloud.png) |
| **Hybrid** | **On-Prem** |
| ![Hybrid](docs/03-hybrid.png) | ![On-Prem](docs/04-on-prem.png) |
| **AD** | **Connected, loading** |
| ![AD](docs/05-ad.png) | ![Loading](docs/06-connected-loading.png) |

<sub>Sensitive data (users, domains, servers, OU names) is blurred.</sub>

## Requirements

| | |
|---|---|
| OS | Windows 10 / 11 or Windows Server 2016+ |
| PowerShell | **Windows PowerShell 5.1**, already included in Windows |
| Modules | `Microsoft.Graph.Authentication`, `ImportExcel`, `ActiveDirectory` (RSAT). HybridScope offers to install the missing ones at start-up. |
| Network | HTTPS to Microsoft 365; HTTP 80 (WinRM) to the Exchange server; 9389 (ADWS) to a domain controller |
| Exchange | Exchange Server 2013 / 2016 / 2019 / SE. Optional: without it you get AD + Microsoft 365 only. |
| Microsoft 365 | Last sign-in and the MFA registration report need **Entra ID P1** (included in Business Premium, E3, E5). Without P1, HybridScope switches to a slower method and tells you. |

## Quick start

1. **Download** the latest release (or `git clone`) to a local folder, e.g. `C:\HybridScope`. Avoid OneDrive and Desktop folders.
2. **Allow local scripts** once. From PowerShell:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   Get-ChildItem C:\HybridScope -Recurse | Unblock-File
   ```
3. **Start** it by double-clicking **`Start-HybridScope.cmd`**, or from PowerShell:
   ```powershell
   cd C:\HybridScope
   .\HybridScope.ps1
   ```
4. If prompted, let it install the missing modules. For `ActiveDirectory` a visible PowerShell window asks for administrator confirmation.

### Optional: build an .exe

```powershell
.\tools\Build-Exe.ps1
```
It creates `HybridScope.exe` in the main folder with [PS2EXE](https://github.com/MScholtes/PS2EXE) and installs PS2EXE if missing. The exe gets the HybridScope icon (`tools\HybridScope.ico`) and the version of the script in its file properties. It must stay in the same folder as `tools\` and `config.json`.

> The exe **contains** the script: it does not run the `HybridScope.ps1` next to it. After every update (new release), run `.\tools\Build-Exe.ps1` again. `Start-HybridScope.cmd` always runs the current `HybridScope.ps1` and never needs rebuilding.

> ⚠️ Antivirus and sandbox engines (Defender, Bitdefender...) often flag PS2EXE executables as false positives. Sign the exe with a code-signing certificate (the command is printed at the end of the build), or simply use `Start-HybridScope.cmd`.

`config.json` is created next to the script on first run. `config.sample.json` shows every option.

## Using the GUI

0. **Infrastructure**: click the button that matches your environment at the top of the window: *Full Cloud*, *Full Hybrid*, *Hybrid*, *On-Prem* or *AD*. Only the panels you need are shown.
1. **Microsoft 365**: optionally enter the tenant (`contoso.onmicrosoft.com`), then **Connect 365**. A browser window opens. The first time, an administrator must consent to the read-only permissions.
2. **Domain credentials**:
   - leave **Use Windows credentials** checked on a domain-joined PC;
   - otherwise uncheck it and type user and password: `john@contoso.local`, `CONTOSO\john` or just `john`, in which case `@domain` is added for you.

   These credentials are used for both Exchange and AD. The user name is remembered; the password never is.

   Signed in to Microsoft 365 with the wrong account? Click **Change account**: it signs out and clears the saved sign-in, then **Connect 365** asks for an account again. The sign-in is never kept between sessions.
3. **Exchange on-premises**: enter the **internal** FQDN of an Exchange server (e.g. `exch01.contoso.local`) and choose a mode:
   - `http` + `Kerberos`: recommended, no certificate needed
   - `https`: only with a valid certificate for that name
   - `local`: when HybridScope runs on the Exchange server itself

   Then click **Connect Exchange**.
4. **Active Directory**: check the domain and click **Verify AD**.
5. Choose **Group by**:
   - *Top-level OU*: the first OU under the domain, e.g. `Sales`
   - *Parent OU*: the OU that directly contains the user, e.g. `Employees` or `External`; usually the most useful
   - *No grouping*: a single list
   - *Custom rules*: `GroupRules` in `config.json`

   You never have to map OUs by hand: grouping is automatic, and the choice is remembered.
   In **Export content** choose what the export will contain: **FULL**, a preset or only the columns you tick (**Columns...**). See [Export](#export-full-preset-or-custom-columns).
6. Click **Load data**. Progress appears in the **Log** panel; a few hundred users take a few minutes. **Stop** cancels the load; it waits for the current request to finish.
7. Search, sort, and **export**. To look at a report saved earlier, use **File > Open report** (Ctrl+O) or drag the `.xlsx` / CSV folder onto the window. **Double-click a user** in any tab (USERS, a group, MFA, mailboxes, ATTENTION, SYNC...) to open their card.

The language can be switched from the top-right menu (restart required). **Help > Check for updates** tells you whether a newer release is out; **Help > About** shows the version and the project links.

### Infrastructure modes

Pick your infrastructure with the buttons at the top of the window. The connection panels, tabs, columns and overview adapt to the sources you have. Switching takes one click: the choice is not tied to your data.

| Button | Sources | What you get |
|---|---|---|
| **Full Cloud** | Microsoft 365 | Users, licenses, Exchange Online mailboxes, OneDrive, MFA, last sign-in |
| **Full Hybrid** | Microsoft 365 + Exchange on-prem + AD | Everything: on-prem vs 365 location, mailboxes on both sides, MFA, licenses, AD data |
| **Hybrid** | Microsoft 365 + AD | AD users merged with their cloud accounts: licenses, Exchange Online, OneDrive, MFA |
| **On-Prem** | Exchange on-prem + AD | AD users with their on-prem mailboxes: size, database, quota, last access, inactivity |
| **AD** | Active Directory | Users, OUs, status, last logon, password age, inactivity |

For scheduled reports (`-Auto`) the mode comes from `Mode` in `config.json`.

### Export: FULL, preset or custom columns

Choose **what to export before loading the data**, in the **Export content** box next to the connection panels (the choice is saved and can also be changed after loading). **Export XLSX...** / **Export CSV...** then write exactly that, without asking again.

In the box, the drop-down picks the content and **Columns...** opens the column chooser (only the columns of the selected infrastructure mode are listed):

- **FULL**: every tab with every column, as before.
- **A preset**, ready to use for the most common requests:

  | Preset | Columns of the user list |
  |---|---|
  | Sign-ins and inactivity | first and last name, username, status, last AD logon, last Microsoft 365 sign-in, last password change, licenses |
  | Licenses | first and last name, username, UPN, status, licenses, last Microsoft 365 sign-in, notes |
  | Mailbox migration | group, first and last name, username, status, Location, mailbox type, on-prem / 365 / OneDrive size |
  | MFA security | first and last name, username, UPN, status, licenses, MFA status, default method, methods, per-user state |
  | Directory | group, first and last name, username, UPN, OU, status, origin |

- **Custom**: tick only the columns you want, e.g. just *AD_Username*, *AD_LastLogon* and *M365_LastSignIn*. Changing a column of a preset switches to Custom; **Save as preset...** keeps the selection with a name of your choice (stored in `config.json`, shown with a ★).

Options for presets and custom exports: **one sheet per group**, **include OVERVIEW**, **enabled accounts only**. A CSV with only the user list is a single file; with more sheets you get a folder with one CSV per tab. Next to the drop-down you see what will be exported, e.g. *3 columns: AD_Username, AD_LastLogon, M365_LastSignIn*.

**Read only the data needed** (ticked by default, used with a preset or Custom): **Load data** reads only what the chosen columns need, so it is much faster. *Sign-ins and inactivity*, for example, reads AD and the Microsoft 365 user list, but not Exchange mailboxes, usage reports, MFA or the sync check. The top bar then shows **PARTIAL LOAD**: the overview, the tabs and the user cards are incomplete. If you later export columns that were not read, HybridScope warns you. Untick it, or choose FULL, to read everything.

## Tabs and columns

| Tab | Content |
|---|---|
| **OVERVIEW** | Counts per group and **TOTAL**: users, mailbox location, space (GB), licenses, MFA, inactivity |
| **USERS** | All users, one row each |
| *one tab per group* | Users of each OU / group (see `GroupBy` and `GroupRules`) |
| **OTHER** | AD users that match no rule |
| **CLOUD_ONLY** | Entra users without an AD account |
| **MAILBOX_ONPREM** | On-prem mailboxes: type, database, size, items, deleted items, last access, quota |
| **MAILBOX_365** | Exchange Online mailboxes from the Microsoft usage report |
| **ONEDRIVE** | OneDrive accounts: used space, quota, files, last activity |
| **MFA** | Every Entra member: registered, capable, default method, methods, per-user state, SSPR, admin |
| **TENANT_LICENSES** | Purchased / assigned / available / suspended per SKU |
| **SYNC** | AD ↔ Entra ID sync check (modes with both AD and Microsoft 365); see [Sync check](#sync-check) |
| **ATTENTION** | Action list: mailboxes to migrate, licensed without MFA, licensed but disabled, inactive accounts, AD accounts not synced, sync problems |
| **HISTORY** | One row per `-Auto` run, to follow the trends |

<details>
<summary><b>User columns</b> (USERS and group tabs)</summary>

| Column | Description |
|---|---|
| Group | Group / OU tab the user belongs to |
| FirstName, LastName, Username | From AD (or Entra for cloud-only users) |
| OU | Path of the user's OU, e.g. `CONTOSO\Sales\Employees` |
| Status | Enabled / Disabled |
| AD_LastLogon | `lastLogonTimestamp`, replicated, accurate to about 14 days |
| M365_LastSignIn | Latest interactive or non-interactive sign-in |
| Licenses | Friendly license names (see `LicenseNames`) |
| Mailbox365_MB / MailboxOnPrem_MB / OneDrive_MB | Used space in MB (2 decimals). Empty when Microsoft does not report the size |
| PasswordLastSet | Last password change (AD) |
| **Location** | `365`, `ON-PREM` or `NO MAILBOX`; see below |
| MailboxType | `UserMailbox`, `SharedMailbox`, `RemoteUserMailbox`, `Exchange Online`... |
| MailboxSource | Where the location information comes from |
| MFA_Status | Configured / Not configured / Not in Entra |
| MFA_DefaultMethod, MFA_Methods | Default MFA method and all registered methods |
| MFA_PerUser | Legacy per-user MFA: `disabled` / `enabled` / `enforced` |
| UPN | Microsoft 365 user principal name |
| Origin | AD synced / AD only / AD + cloud (not linked) / Cloud only / Guest. *Not linked*: an AD account and a cloud account with the same name, but Entra does not sync one from the other |
| SyncIssues | Sync problems of this user (details in the SYNC tab) |
| Notes | Anomalies, e.g. "Licensed without MFA" or "Password never expires" |
</details>

**How Location is decided:**
1. A mailbox exists on Exchange on-prem (`Get-Mailbox`): **ON-PREM**
2. A `RemoteMailbox` exists on Exchange on-prem, i.e. the mailbox was migrated: **365**
3. Exchange Online is provisioned for the user (Graph `provisionedPlans`) or the user appears in the usage report: **365**
4. None of the above: **NO MAILBOX**

### Sync check

The SYNC tab compares AD with Entra ID the way Entra Connect and Cloud Sync do: through the **source anchor**. In AD the anchor is `ms-DS-ConsistencyGuid` (or `objectGUID` when it is empty); in Entra it is stored, base64-encoded, in `onPremisesImmutableId`. Users are also paired by anchor in the other tabs, with names as a fallback.

| Category | What it means |
|---|---|
| 0 - Directory sync status | When the last sync ran; flagged when it is older than `SyncStaleHours`, or when sync is not enabled in the tenant |
| 1 - ImmutableId differs from the AD anchor | Same account name, but Entra is linked to another object. Usually the AD account was deleted and recreated |
| 2 - Synced account not found in AD | Synced from the domain / OU you read, but no AD account has its anchor |
| 3 - Cloud account with a leftover ImmutableId | Sync is off for the account but the `ImmutableId` is still set: an AD account with that anchor would take it over |
| 4 - Will soft-match a cloud account | An AD account not yet synced and a cloud account with the same UPN or primary SMTP: they will be merged at the first sync |
| 5 - Conflicts with a cloud account | Same UPN / SMTP, but the cloud account already has another `ImmutableId`: sync will fail with a duplicate attribute error |
| 6 - Provisioning error in Entra | Errors Entra reports on the object (duplicate `proxyAddresses`, UPN...) |
| 7 - Entra UPN fell back to onmicrosoft.com | The AD attribute the sync uses as UPN (`userPrincipalName`, or `mail` with Alternate Login ID) has a suffix that is not a verified domain, so the user got `@tenant.onmicrosoft.com`. A UPN that differs from AD but uses a verified domain (e.g. `.local` in AD, `.com` in Entra) is a normal setup and is not flagged |
| 8 - Same anchor on more than one AD account | Copied or restored objects with the same `ms-DS-ConsistencyGuid` |

Accounts synced from another domain or forest, or from outside `SearchBase`, are not reported as missing. If most synced accounts do not match by anchor (a custom source anchor such as `employeeID`), categories 1 and 2 are skipped and a warning explains why.

## Configuration (config.json)

| Key | Default | Description |
|---|---|---|
| `Language` | system | `en` or `it` |
| `Mode` | `FullHybrid` | `Cloud`, `FullHybrid`, `Hybrid`, `OnPrem` or `AD`. Starting mode of the GUI and mode used by `-Auto` |
| `TenantId` | *(empty)* | Tenant domain or ID. Empty means the tenant of the signed-in account |
| `AuthMode` | `Interactive` | `Interactive` (browser) or `Certificate` (for `-Auto`) |
| `ClientId`, `CertificateThumbprint` | | Set by `tools\New-HybridScopeApp.ps1` |
| `IncludePerUserMfaState` | `true` | Reads the legacy per-user MFA state (one Graph call per user) |
| `DomainFqdn` | `%USERDNSDOMAIN%` | AD domain, e.g. `contoso.local`. Empty skips AD |
| `SearchBase` | *(empty)* | Limit AD to an OU, e.g. `OU=Company,DC=contoso,DC=local` |
| `UseWindowsCredentials`, `DomainUser` | `true`, empty | Last values used in the GUI. The password is never saved |
| `ExchangeServer` | *(empty)* | Internal FQDN of an Exchange server. Empty skips Exchange |
| `ExchangeProtocol` | `http` | `http`, `https` or `local` |
| `ExchangeAuth` | `Kerberos` | `Kerberos`, `Negotiate` or `Basic` |
| `ExchangeSkipCertCheck` | `false` | Skip certificate checks with `https` |
| `OutputFolder` | `Documents\HybridScope` | Reports, `history.csv` and logs |
| `RetentionDays` | `90` | `-Auto` deletes older reports and logs (0 = never) |
| `InactiveDays` | `90` | Inactivity threshold |
| `PasswordAgeDays` | `365` | Old password threshold |
| `SyncStaleHours` | `3` | SYNC tab: flag the directory sync when the last cycle is older than this |
| `LicenseFilter` | `All` | `All` shows every SKU (unknown ones with the raw name); `MappedOnly` shows only SKUs listed in `LicenseNames` |
| `LicenseNames` | common SKUs | `"SKU_PART_NUMBER": "Friendly name"`. Add or rename freely |
| `GroupBy` | `TopLevelOU` | `TopLevelOU`, `ParentOU`, `None` (single group) or `Rules` (uses `GroupRules`). Also selectable in the GUI |
| `GroupRules` | `[]` | Used only with `GroupBy = Rules`. `[{ "Group": "EMPLOYEES", "Match": "OU=Employees," }]`; the first match on the DistinguishedName wins |
| `UnmappedGroup`, `CloudOnlyGroup` | `OTHER`, `CLOUD_ONLY` | Custom names for those groups |
| `IncludeGuests` | `false` | Include Entra guests |
| `ExportPreset`, `ExportColumns` | `Full`, `[]` | Choice in the **Export content** box; `ExportColumns` are the **Custom** columns (English names, e.g. `["Username", "AD_LastLogon", "M365_LastSignIn"]`) |
| `ExportGroupSheets`, `ExportOverview`, `ExportEnabledOnly` | `false` | Export options: one sheet per group, include OVERVIEW, enabled accounts only (also used by `-Auto` with a preset) |
| `ExportPresets` | `{}` | Your presets: `{ "Inactive": ["Username", "AD_LastLogon", "M365_LastSignIn"] }`. Saved by the export window |
| `ExportFastLoad` | `true` | With a preset / Custom, read only the data the columns need (GUI checkbox **Read only the data needed**, also used by `-Auto`). A partial `-Auto` run does not update `history.csv` |
| `AutoExportPreset` | `Full` | Content of the `-Auto` report: `Full`, `Access`, `Licenses`, `Migration`, `MFA`, `Directory`, `Custom` or the name of one of your presets |

Example `GroupRules`:
```json
"GroupRules": [
  { "Group": "DISABLED",  "Match": "OU=Disabled Users," },
  { "Group": "EMPLOYEES", "Match": "OU=Employees," },
  { "Group": "EXTERNAL",  "Match": "OU=External," }
]
```
> Keep the trailing comma in `Match`: it prevents partial matches such as `OU=EmployeesOld`.

To show **only some licenses** with names of your choice, set `"LicenseFilter": "MappedOnly"` and keep only those SKUs in `LicenseNames`:
```json
"LicenseFilter": "MappedOnly",
"LicenseNames": {
  "O365_BUSINESS_PREMIUM":    "M365 BUSINESS STANDARD",
  "O365_BUSINESS_ESSENTIALS": "M365 BUSINESS BASIC",
  "SPB":                      "M365 BUSINESS PREMIUM"
}
```

## Permissions

HybridScope only **reads**. Least privilege is enough; you do not need Domain Admin or Global Admin for everyday use.

| System | Minimum |
|---|---|
| Active Directory | Any domain user |
| Exchange on-premises | Role group **View-Only Organization Management** and `RemotePowerShellEnabled = $true` |
| Microsoft 365 | **Global Reader**, plus a one-time admin consent to the permissions below |

Microsoft Graph permissions (delegated in the GUI, application for `-Auto`):
`User.Read.All`, `AuditLog.Read.All`, `Reports.Read.All`, `UserAuthenticationMethod.Read.All`, `Organization.Read.All`, `Policy.Read.All`.

The GUI also asks for `GroupMember.Read.All` and `RoleManagement.Read.Directory`, read-only, used only by the user card to show group and role names. If an administrator has not consented to them yet, **Connect 365** still connects with the permissions above and the card shows AD groups only; consent once as an administrator to get Entra groups and roles too.

## Command line (no GUI)

Every run without the window uses `config.json`; full help with examples: `Get-Help .\HybridScope.ps1 -Full`.

| Parameter | What it does |
|---|---|
| `-Auto` | Full report for the configured `Mode`, written to `OutputFolder` (plus `history.csv` and `HybridScope_LATEST.xlsx`) |
| `-Only <dataset>` | Just one dataset: `MFA`, `Licenses`, `Mailbox365`, `OneDrive`, `MailboxOnPrem` or `Sync` (AD + Microsoft 365). Only the sources it needs are contacted |
| `-PassThru` | Returns objects instead of writing files: the `-Only` rows, or one merged row per user. Values are always in English |
| `-Mode <mode>` | Overrides `Mode` for this run (`Cloud`, `FullHybrid`, `Hybrid`, `OnPrem`, `AD`) |
| `-ExportPreset <name>` | Content of the report: `Full` (default), `Access`, `Licenses`, `Migration`, `MFA`, `Directory`, `Custom` or one of your presets. Overrides `AutoExportPreset` |
| `-Columns <list>` | Only these columns of the user list, in this order, e.g. `Username,AD_LastLogon,M365_LastSignIn` (English names or Italian headers) |
| `-ConfigPath <file>` | Uses another configuration file, e.g. one per tenant |

Microsoft 365 sign-in: with `AuthMode = Interactive` (default) a browser prompt opens once; with `Certificate` there is no prompt. If you already ran `Connect-MgGraph` in the same session, that connection is reused and left open.

```powershell
# MFA status only, to Excel in OutputFolder (AD and Exchange are not touched)
.\HybridScope.ps1 -Only MFA

# users without MFA, straight to CSV
.\HybridScope.ps1 -Only MFA -PassThru | Where-Object MFA_Registered -eq 'No' | Export-Csv .\no-mfa.csv -NoTypeInformation

# what will go wrong at the next AD -> Entra ID sync
.\HybridScope.ps1 -Only Sync -PassThru | Where-Object Category -notlike '0 *' | Format-Table Category, Username, UPN, Details -Wrap

# only username, last AD logon and last Microsoft 365 sign-in
.\HybridScope.ps1 -Auto -Columns Username,AD_LastLogon,M365_LastSignIn

# ready-made MFA report
.\HybridScope.ps1 -Auto -ExportPreset MFA

# full report, Microsoft 365 only, without editing config.json
.\HybridScope.ps1 -Auto -Mode Cloud

# merged per-user data as objects: on-prem mailboxes still to migrate, largest first
.\HybridScope.ps1 -PassThru | Where-Object Location -eq 'ON-PREM' | Sort-Object MailboxOnPrem_MB -Descending
```

Exit code: `0` = success, `1` = error (details in `OutputFolder\logs`).

## Scheduled reports (-Auto)

1. **Create the app registration** with a certificate. Run it once, as Global Admin, **with the account that will run the task**:
   ```powershell
   .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com
   ```
   It creates the certificate, the app with read-only permissions and the admin consent, and updates `config.json` (`AuthMode = Certificate`).
2. **Register the daily task**, from an elevated PowerShell:
   ```powershell
   .\tools\Register-HybridScopeTask.ps1 -RunAs CONTOSO\svc-hybridscope -At 06:40
   ```
3. Each run writes to `OutputFolder`:
   - `HybridScope_<date>.xlsx` (with a preset: `HybridScope_<preset>_<date>.xlsx`, see `AutoExportPreset`)
   - `HybridScope_LATEST.xlsx`, always the latest, handy to share
   - `history.csv`, also shown in the HISTORY tab
   - `logs\`

Manual run: `powershell -File .\HybridScope.ps1 -Auto`

## Troubleshooting

<details>
<summary><b>Connected to Microsoft 365 with the wrong account</b></summary>

Click **Change account** to sign out and clear the cached sign-in, then **Connect 365** again. If the browser picks the wrong account by itself, choose *Use another account* on the Microsoft sign-in page, or sign out of that account in the browser first.
</details>

<details>
<summary><b>"WorkgroupCannotUseKerberos" / "Cannot use Kerberos with implicit credentials"</b></summary>

The PC is not joined to the domain. Uncheck **Use Windows credentials** and enter `user@domain.local`. The PC must resolve the domain through internal DNS or VPN: `nslookup -type=SRV _kerberos._tcp.contoso.local`.
</details>

<details>
<summary><b>"The user name or password is incorrect" when connecting to Exchange</b></summary>

Check the exact sAMAccountName, for example `j.doe` vs `jdoe`, and use the `user@domain.local` format. Test outside HybridScope:

```powershell
$cred = Get-Credential user@contoso.local
$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch01.contoso.local/PowerShell/ -Authentication Kerberos -Credential $cred
Invoke-Command -Session $s -ScriptBlock { Get-ExchangeServer }
```
</details>

<details>
<summary><b>The browser asks for a certificate on https://server/powershell</b></summary>

That is expected: `/PowerShell` is a WinRM endpoint, not a web page. Use `http` + `Kerberos` with the internal server name.
</details>

<details>
<summary><b>Microsoft 365 mailbox / OneDrive sizes are empty</b></summary>

Usage reports are anonymized. In the Microsoft 365 admin center go to **Settings > Org settings > Reports** and uncheck *Display concealed user, group, and site names in all reports*. Reports can also lag by 24-48 hours.
</details>

<details>
<summary><b>ActiveDirectory module: error 0x800f0954</b></summary>

Your WSUS blocks Features on Demand. Install *RSAT: Active Directory Domain Services* from **Settings > System > Optional features**, or ask IT to allow it. HybridScope does not change Windows Update settings.
</details>

<details>
<summary><b>Antivirus flags the exe</b></summary>

Executables built with PS2EXE (`tools\Build-Exe.ps1`) are often false positives. Use `Start-HybridScope.cmd`, which runs the plain-text script with Windows PowerShell, or sign the exe with your code-signing certificate.
</details>

<details>
<summary><b>"cannot be loaded because running scripts is disabled" / "not digitally signed"</b></summary>

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Get-ChildItem C:\HybridScope -Recurse | Unblock-File
```
</details>

## FAQ

**Does it change anything?** No. Every call is read-only. The only exception is installing missing PowerShell modules, and only if you agree.

**Where is my data sent?** Nowhere. HybridScope talks only to your AD, your Exchange server and Microsoft Graph. Reports stay on your disk. The only other connection is **Help > Check for updates**, and only when you click it: it reads the latest release number from the public GitHub API (`api.github.com`) and sends nothing about you or your tenant.

**Does it work without Microsoft 365, with only Exchange 2016/2019 and AD?** Yes: use **On-Prem** mode.

**Only Microsoft 365, or Microsoft 365 + AD without Exchange?** Yes: use **Full Cloud** or **Hybrid** mode.

**Only Active Directory?** Yes: use **AD** mode.

**Can I get just the MFA status, without the GUI?** Yes: `.\HybridScope.ps1 -Only MFA` (Excel) or `-Only MFA -PassThru` (objects). See [Command line](#command-line-no-gui).

## Roadmap

- [x] User card: double-click a user to see AD, mailbox, OneDrive, MFA, licenses, groups and roles together
- [x] Sync check: ImmutableId / source anchor, soft-match and conflicts before the next sync
- [ ] Shared / room mailboxes and distribution lists: signed-in shared mailboxes, unneeded licenses, delegations, empty or ownerless lists
- [ ] HTML summary after `-Auto`: overview, attention list and trends, to send by e-mail

## Support the project

HybridScope is free and open source, and it will stay that way. If HybridScope (or one of my other projects) saves you time, you can leave a **one-time donation**: no subscription, no automatic renewal.

| | Amount | For |
|---|---|---|
| ☕ **Coffee** | $5 | A quick thank-you |
| 🍕 **Lunch** | $10 | You use it often and want it to stay free, read-only and open source |
| 🛠️ **Sysadmin** | $25 | It saved you hours: a migration, an audit, a license cleanup |
| 🏢 **Company** | $100 | Companies and MSPs using it in production. If you like, your company name can be listed under [Thanks](#thanks) |

[![Sponsor on GitHub](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=githubsponsors&logoColor=white)](https://github.com/sponsors/Thonno)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/aascione956)
[![Star on GitHub](https://img.shields.io/github/stars/Thonno/hybridscope?style=for-the-badge&logo=github&label=Star)](https://github.com/Thonno/hybridscope/stargazers)

Not in a position to donate? You can still help: leave a ⭐, open an issue with a bug or an idea, or recommend HybridScope to a colleague.

Issues and pull requests are welcome. Please open an issue before large changes.

## Thanks

Thanks to everyone who supports HybridScope.

**Companies and MSPs**

<!-- Add Company-tier sponsors who asked to be listed, one per line, e.g.:
- [Contoso Ltd](https://www.contoso.com)
-->
*No one yet: your company could be the first.*

## License

[MIT](LICENSE). HybridScope is an independent project and is not affiliated with, endorsed by or sponsored by Microsoft. Microsoft 365, Exchange, Entra ID and Active Directory are trademarks of the Microsoft group of companies.
