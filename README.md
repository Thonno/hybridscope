<div align="center">

# HybridScope

**One view of your hybrid identities: Active Directory + Exchange on-premises + Microsoft 365.**

Who is on-prem and who is already in the cloud, how much space every mailbox uses, who has a license, who has MFA, who hasn't signed in for months: in one window, exportable to Excel.

[![Sponsor on GitHub](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=githubsponsors&logoColor=white)](https://github.com/sponsors/Thonno)
[![Star on GitHub](https://img.shields.io/github/stars/Thonno/hybridscope?style=for-the-badge&logo=github&label=Star)](https://github.com/Thonno/hybridscope/stargazers)

![PowerShell 5.1](https://img.shields.io/badge/PowerShell-5.1-5391FE?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011%20%7C%20Server-0078D6?logo=windows&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Languages](https://img.shields.io/badge/UI-English%20%7C%20Italiano-blue)

**English** · [Italiano](README.it.md)

</div>

---

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Quick start](#quick-start)
- [Using the GUI](#using-the-gui)
- [Tabs and columns](#tabs-and-columns)
- [Configuration (config.json)](#configuration-configjson)
- [Permissions](#permissions)
- [Scheduled reports (-Auto)](#scheduled-reports--auto)
- [Troubleshooting](#troubleshooting)
- [FAQ](#faq)
- [Roadmap](#roadmap)
- [Support the project](#support-the-project)
- [License](#license)

## Features

- **RVTools-style GUI**: every dataset in its own tab, a search box that filters the current tab, sortable columns, copy with headers.
- **Three sources in one report**:
  - **Active Directory**: users, OUs, enabled/disabled, last logon, last password change
  - **Exchange Server on-premises**: mailboxes, size, item count, database, quota, last access, and remote mailboxes that were already migrated
  - **Microsoft 365 (Microsoft Graph)**: licenses, last sign-in, Exchange Online mailboxes, OneDrive, MFA registration and methods, per-user MFA state, purchased and free licenses
- **Reliable mailbox location**: `ON-PREM` / `365` / `NO MAILBOX` comes from Exchange (`Get-Mailbox`, `Get-RemoteMailbox`) and Graph (`provisionedPlans`), never from the mailbox size.
- **Real MFA status**: whether at least one method is registered, which methods, the default one, and the legacy per-user state.
- **Overview** per group (OU) with totals, plus an **Attention** list:
  - on-prem mailboxes to migrate, sorted by size
  - licensed users without MFA
  - licenses assigned to disabled accounts
  - inactive accounts
  - AD accounts that are not synced
- **Export**: formatted XLSX (all tabs) or CSV, or only the current tab with the filter applied.
- **Works from a PC outside the domain**: either Windows credentials or a domain user and password entered manually, with Kerberos.
- **Headless mode `-Auto`**: certificate-based sign-in for scheduled reports. Keeps a `HybridScope_LATEST.xlsx` copy and a history of trends (on-prem users, space, % MFA).
- **English / Italian** interface.
- **Read-only**: HybridScope never changes anything in AD, Exchange or Microsoft 365.

> 📸 *Add your screenshots to `docs/` and link them here, e.g. `![Overview](docs/overview.png)`.*

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

`config.json` is created next to the script on first run. `config.sample.json` shows every option.

## Using the GUI

1. **Microsoft 365**: optionally enter the tenant (`contoso.onmicrosoft.com`), then **Connect 365**. A browser window opens. The first time, an administrator must consent to the read-only permissions.
2. **Domain credentials**:
   - leave **Use Windows credentials** checked on a domain-joined PC;
   - otherwise uncheck it and type user and password: `john@contoso.local`, `CONTOSO\john` or just `john`, in which case `@domain` is added for you.

   These credentials are used for both Exchange and AD. The user name is remembered; the password never is.
3. **Exchange on-premises**: enter the **internal** FQDN of an Exchange server (e.g. `exch01.contoso.local`) and choose a mode:
   - `http` + `Kerberos`: recommended, no certificate needed
   - `https`: only with a valid certificate for that name
   - `local`: when HybridScope runs on the Exchange server itself

   Then click **Connect Exchange**.
4. **Active Directory**: check the domain and click **Verify AD**.
5. Click **Load data**. Progress appears in the **Log** panel; a few hundred users take a few minutes.
6. Search, sort, and **export**.

The language can be switched from the top-right menu (restart required).

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
| **ATTENTION** | Action list: mailboxes to migrate, licensed without MFA, licensed but disabled, inactive accounts, AD accounts not synced |
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
| Origin | AD synced / AD only / Cloud only / Guest |
| Notes | Anomalies, e.g. "Licensed without MFA" or "Password never expires" |
</details>

**How Location is decided:**
1. A mailbox exists on Exchange on-prem (`Get-Mailbox`): **ON-PREM**
2. A `RemoteMailbox` exists on Exchange on-prem, i.e. the mailbox was migrated: **365**
3. Exchange Online is provisioned for the user (Graph `provisionedPlans`) or the user appears in the usage report: **365**
4. None of the above: **NO MAILBOX**

## Configuration (config.json)

| Key | Default | Description |
|---|---|---|
| `Language` | system | `en` or `it` |
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
| `LicenseFilter` | `All` | `All` shows every SKU (unknown ones with the raw name); `MappedOnly` shows only SKUs listed in `LicenseNames` |
| `LicenseNames` | common SKUs | `"SKU_PART_NUMBER": "Friendly name"`. Add or rename freely |
| `GroupBy` | `TopLevelOU` | When `GroupRules` is empty: `TopLevelOU`, `ParentOU` or `None` (single group) |
| `GroupRules` | `[]` | `[{ "Group": "EMPLOYEES", "Match": "OU=Employees," }]`. First match on the DistinguishedName wins |
| `UnmappedGroup`, `CloudOnlyGroup` | `OTHER`, `CLOUD_ONLY` | Custom names for those groups |
| `IncludeGuests` | `false` | Include Entra guests |

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
   - `HybridScope_<date>.xlsx`
   - `HybridScope_LATEST.xlsx`, always the latest, handy to share
   - `history.csv`, also shown in the HISTORY tab
   - `logs\`

Manual run: `powershell -File .\HybridScope.ps1 -Auto`

## Troubleshooting

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

**Where is my data sent?** Nowhere. HybridScope talks only to your AD, your Exchange server and Microsoft Graph. Reports stay on your disk.

**Does it work without Exchange on-premises (cloud-only)?** Yes. Leave `ExchangeServer` empty.

**Does it work without Active Directory?** Yes. Leave `DomainFqdn` empty: every Entra user ends up in CLOUD_ONLY.

## Roadmap

- [ ] Clickable overview: double-click a number to open a tab with exactly those users
- [ ] User card: double-click a user to see AD, mailbox, OneDrive, MFA and licenses together
- [ ] Shared / room mailboxes and distribution groups
- [ ] HTML report
- [ ] More languages (PRs welcome: the strings live in `$global:HS_Str` at the top of `HybridScope.ps1`)

## Support the project

If HybridScope saves you time, you can support its development:

[![Sponsor on GitHub](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?style=for-the-badge&logo=githubsponsors&logoColor=white)](https://github.com/sponsors/Thonno)
[![Star on GitHub](https://img.shields.io/github/stars/Thonno/hybridscope?style=for-the-badge&logo=github&label=Star)](https://github.com/Thonno/hybridscope/stargazers)

Issues and pull requests are welcome. Please open an issue before large changes.

## License

[MIT](LICENSE). HybridScope is an independent project and is not affiliated with, endorsed by or sponsored by Microsoft. Microsoft 365, Exchange, Entra ID and Active Directory are trademarks of the Microsoft group of companies.
