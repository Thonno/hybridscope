<#
.SYNOPSIS
    HybridScope - read-only inventory of hybrid identities and mailboxes:
    Active Directory + Exchange on-premises + Microsoft 365 (Microsoft Graph).

.DESCRIPTION
    HybridScope reads users from Active Directory, mailboxes from Exchange Server (2013-SE) and
    licenses, sign-ins, Exchange Online / OneDrive usage and MFA registration from Microsoft Graph,
    merges them into one row per user and exports the result to Excel or CSV.
    It never changes anything in AD, Exchange or Microsoft 365.

    Ways to run it:

      GUI (default)      No parameters. A window to connect to the sources, load the data into
                         RVTools-style tabs, search, sort and export.

      Full report        -Auto. No window: reads config.json, collects everything for the configured
                         Mode, writes the Excel report to OutputFolder and appends to history.csv.

      One dataset        -Only <name>. No window: collects just that dataset and writes
                         HybridScope_<name>_<timestamp>.xlsx, or returns it with -PassThru.
                         <name>: MFA, Licenses, Mailbox365, OneDrive, MailboxOnPrem, Sync
                         (details: Get-Help .\HybridScope.ps1 -Parameter Only)

      Pipeline           -PassThru. Returns objects instead of writing files: the rows of the -Only
                         dataset, or one merged row per user without -Only.

    Microsoft 365 sign-in without the GUI depends on AuthMode in config.json:
      Interactive  -> a browser sign-in opens once (one-off runs from a console)
      Certificate  -> app + certificate, no prompt (scheduled tasks, see tools\New-HybridScopeApp.ps1)
    If you are already connected with Connect-MgGraph in the same session, that connection is reused
    and left open. AD and Exchange use the Windows account running the script.

    Files written in OutputFolder (default: Documents\HybridScope):
      HybridScope_<yyyy-MM-dd_HHmm>.xlsx         full report (-Auto)
      HybridScope_<preset>_<timestamp>.xlsx      report with a preset / custom columns (-ExportPreset, -Columns)
      HybridScope_LATEST.xlsx                    copy of the latest full report
      HybridScope_<dataset>_<timestamp>.xlsx     single dataset (-Only)
      history.csv                                one row per full report (users, on-prem/365, space, % MFA)
      logs\HybridScope_<timestamp>.log           run log
    Reports and logs older than RetentionDays are deleted (0 = keep everything).

.PARAMETER Auto
    Run without the GUI and write the full report to OutputFolder.
    Exit code 0 = success, 1 = error (details in the log file).

.PARAMETER Only
    Collect a single dataset instead of the full report. Implies -Auto.
      MFA            MFA registration per Entra user: registered, capable, default method, methods,
                     per-user state, SSPR, admin
      Licenses       tenant licenses: purchased / assigned / available per SKU
      Mailbox365     Exchange Online mailbox sizes (Microsoft usage report)
      OneDrive       OneDrive usage per account
      MailboxOnPrem  Exchange on-premises mailboxes and sizes (needs ExchangeServer in config.json)
      Sync           AD <-> Entra ID sync check: directory sync status, ImmutableId vs AD anchor
                     (ms-DS-ConsistencyGuid / objectGUID), accounts that will soft-match or conflict at
                     the next sync, provisioning errors, UPN differences (needs DomainFqdn in config.json)
    Only the source that dataset needs is contacted: -Only MFA does not touch AD or Exchange.

.PARAMETER PassThru
    Return the data as objects instead of writing files. Implies -Auto.
    With -Only: the rows of that dataset. Without -Only: one merged row per user (the USERS tab).
    Property names and values are always in English (Yes/No, Enabled/Disabled...), whatever the
    Language setting, so filters like Where-Object MFA_Registered -eq 'No' work everywhere.

.PARAMETER Mode
    Which sources to use for this run, overriding Mode in config.json (not saved):
    Cloud, FullHybrid, Hybrid, OnPrem, AD. In the GUI it only sets the starting mode.

.PARAMETER ExportPreset
    Content of the report written by -Auto, overriding AutoExportPreset in config.json:
      Full       every tab and every column (default)
      Access     sign-ins and inactivity: names, username, status, AD_LastLogon, M365_LastSignIn,
                 PasswordLastSet, licenses
      Licenses   names, username, UPN, status, licenses, M365_LastSignIn, notes
      Migration  group, names, username, status, Location, mailbox type, on-prem / 365 / OneDrive sizes
      MFA        names, username, UPN, status, licenses, MFA status, default method, methods, per-user
      Directory  group, names, username, UPN, OU, status, origin
      Custom     the columns in ExportColumns (config.json, also set by the export window)
      <name>     a preset saved from the export window (ExportPresets in config.json)
    Any value other than Full writes the user list with those columns only (plus, as set in
    config.json: one sheet per group, the overview, enabled accounts only). With -PassThru it
    selects the properties of the returned objects.
    With ExportFastLoad = true (default) only the data those columns need is read: e.g. Access does
    not read Exchange mailboxes, usage reports or MFA. A partial run does not update history.csv.

.PARAMETER Columns
    Exact columns of the user list for this run, in this order, e.g. Username,AD_LastLogon,M365_LastSignIn.
    English names (see the USERS tab) or the Italian headers. Overrides -ExportPreset.

.PARAMETER ConfigPath
    Path of the JSON configuration file. Default: config.json next to the script (or the .exe).
    Created with default values if it does not exist. See config.sample.json for every option.

.EXAMPLE
    .\HybridScope.ps1

    Opens the GUI.

.EXAMPLE
    .\HybridScope.ps1 -Only MFA

    MFA status only, no GUI: signs in to Microsoft 365 (browser prompt with AuthMode = Interactive)
    and writes HybridScope_MFA_<timestamp>.xlsx to OutputFolder. AD and Exchange are not contacted.

.EXAMPLE
    .\HybridScope.ps1 -Only MFA -PassThru | Where-Object MFA_Registered -eq 'No' | Export-Csv .\no-mfa.csv -NoTypeInformation

    Users without MFA, straight to a CSV.

.EXAMPLE
    Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, Reports.Read.All, User.Read.All, AuditLog.Read.All
    $mfa = .\HybridScope.ps1 -Only MFA -PassThru
    $mfa | Format-Table UPN, MFA_Registered, DefaultMethod

    Reuses your own Graph connection (it stays open afterwards) and keeps the result in a variable.

.EXAMPLE
    .\HybridScope.ps1 -Only Sync -PassThru | Where-Object Category -notlike '0 *' | Format-Table Category, Username, UPN, Details -Wrap

    What will go wrong at the next AD -> Entra ID sync: anchors that do not match, accounts that will
    soft-match or conflict, provisioning errors. Reads AD with the current Windows account.

.EXAMPLE
    .\HybridScope.ps1 -Auto

    Full report without the GUI, using Mode and the other settings in config.json.

.EXAMPLE
    .\HybridScope.ps1 -Auto -Mode Cloud -ConfigPath D:\HybridScope\contoso.json

    Full report for Microsoft 365 only, with a different configuration file (e.g. one per tenant).

.EXAMPLE
    .\HybridScope.ps1 -Auto -Columns Username,AD_LastLogon,M365_LastSignIn

    Report with only three columns for every user: username, last AD logon, last Microsoft 365 sign-in.
    Written as HybridScope_Custom_<timestamp>.xlsx (and HybridScope_LATEST.xlsx).

.EXAMPLE
    .\HybridScope.ps1 -Auto -ExportPreset MFA

    Ready-made MFA report: who has registered MFA, default method and methods, per-user state.

.EXAMPLE
    .\HybridScope.ps1 -PassThru | Where-Object Location -eq 'ON-PREM' | Sort-Object MailboxOnPrem_MB -Descending

    Merged per-user data as objects: on-premises mailboxes still to migrate, largest first.

.EXAMPLE
    .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com
    .\tools\Register-HybridScopeTask.ps1 -RunAs CONTOSO\svc-hybridscope -At 06:40

    One-time setup for an unattended daily report: app registration with a certificate
    (sets AuthMode = Certificate), then a scheduled task that runs "HybridScope.ps1 -Auto" every day.

.INPUTS
    None.

.OUTPUTS
    None, unless -PassThru is used: then PSCustomObject rows (see -PassThru).

.NOTES
    Requirements : Windows PowerShell 5.1; modules Microsoft.Graph.Authentication, ImportExcel,
                   ActiveDirectory (RSAT) - each only when a source that needs it is used.
                   The GUI offers to install the missing ones.
    Permissions  : AD - any domain user. Exchange - "View-Only Organization Management" and
                   RemotePowerShellEnabled. Microsoft 365 - Global Reader + one-time admin consent to
                   User.Read.All, AuditLog.Read.All, Reports.Read.All, UserAuthenticationMethod.Read.All,
                   Organization.Read.All, Policy.Read.All (delegated when signing in, application
                   permissions with a certificate). The GUI also asks for GroupMember.Read.All and
                   RoleManagement.Read.Directory, used only by the user card (groups and Entra roles);
                   without consent to them it connects anyway and the card shows no Entra groups.
    Last sign-in and the MFA registration report need Entra ID P1; without it MFA methods are read
    user by user (slower) and a warning is logged.

.LINK
    https://github.com/Thonno/hybridscope
#>
param(
    [switch]$Auto,

    [ValidateSet('MFA', 'Licenses', 'Mailbox365', 'OneDrive', 'MailboxOnPrem', 'Sync')]
    [string]$Only,

    [switch]$PassThru,

    [ValidateSet('Cloud', 'FullHybrid', 'Hybrid', 'OnPrem', 'AD')]
    [string]$Mode,

    [string]$ExportPreset,

    [string[]]$Columns,

    [string]$ConfigPath
)

$HS_Version = '1.6.0'
$HS_Author = 'Thonno'
$HS_Repo = 'Thonno/hybridscope'
$HS_Links = @{
    Repo     = "https://github.com/$HS_Repo"
    Author   = "https://github.com/$HS_Author"
    Releases = "https://github.com/$HS_Repo/releases/latest"
    Issues   = "https://github.com/$HS_Repo/issues/new"
    License  = "https://github.com/$HS_Repo/blob/main/LICENSE"
    Sponsor  = "https://github.com/sponsors/$HS_Author"
    Coffee   = 'https://buymeacoffee.com/aascione956'
}

# Folder of the script / exe (PS2EXE does not set $PSScriptRoot)
$AppDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent ([Environment]::GetCommandLineArgs()[0]) }
if (-not $ConfigPath) { $ConfigPath = Join-Path $AppDir 'config.json' }

# =====================================================================================
#  CORE: localization, data collection, export.
#  Used both by the GUI (inside a background runspace) and by -Auto.
#
#  Shared state lives in $script:HS_* variables (strings, language, log queue, Exchange
#  session...). Script scope is enough everywhere: headless and on the GUI thread it is this
#  file's scope; in the background runspace the engine is loaded at top level, so the state
#  survives between jobs. Nothing is left behind in the caller's session.
# =====================================================================================
$Core = {

    # --------------------------------------------------------------------------- strings
    $script:HS_Str = @{
        en = @{
            'app.subtitle'      = 'Active Directory  ·  Exchange on-premises  ·  Microsoft 365'
            'gb.m365'           = 'Microsoft 365'
            'lbl.tenant'        = 'Tenant'
            'tip.tenant'        = 'Tenant domain or ID. Leave empty to use the tenant of the account you sign in with.'
            'btn.connect365'    = 'Connect 365'
            'st.notconnected'   = 'Not connected'
            'gb.cred'           = 'Domain credentials (Exchange + AD)'
            'chk.wincred'       = 'Use Windows credentials'
            'lbl.user'          = 'User'
            'lbl.password'      = 'Password'
            'tip.user'          = 'e.g. john.doe@contoso.local or CONTOSO\john.doe'
            'tip.notDomain'     = 'This PC is not joined to a domain: enter the credentials manually'
            'gb.exch'           = 'Exchange on-premises'
            'lbl.server'        = 'Server'
            'tip.server'        = 'Internal FQDN of an Exchange server, e.g. exch01.contoso.local'
            'tip.proto'         = 'http = WinRM + Kerberos (standard, no certificate needed) | https = needs a valid certificate | local = HybridScope runs ON the Exchange server'
            'btn.connectExch'   = 'Connect Exchange'
            'gb.ad'             = 'Active Directory'
            'lbl.domain'        = 'Domain'
            'btn.verifyAD'      = 'Verify AD'
            'st.notverified'    = 'Not verified'
            'btn.load'          = 'Load data'
            'btn.xlsx'          = 'Export XLSX...'
            'btn.csvall'        = 'Export CSV...'
            'tip.export'        = 'Exports what is chosen in "Export content"'
            'gb.export'         = 'Export content'
            'btn.expCols'       = 'Columns...'
            'tip.expCols'       = 'Choose a preset or tick the columns of the user list, before or after loading the data'
            'tip.cmbExport'     = 'What Export XLSX / CSV will write. Choose it before Load data.'
            'export.boxFull'    = 'every tab, every column'
            'export.boxCols'    = '{0} columns: {1}'
            'export.summaryCols' = '{0} columns'
            'export.apply'      = 'OK'
            'export.fast'       = 'Read only the data needed'
            'tip.expFast'       = 'With a preset or Custom, Load data reads only what those columns need (much faster, e.g. no Exchange mailboxes or MFA). Overview, tabs and user cards will be partial.'
            'lbl.partial'       = 'PARTIAL LOAD: {0}'
            'tip.partial'       = 'Not read: {0}. Overview, tabs and user cards are incomplete. Load again with FULL, or untick "Read only the data needed", to get everything.'
            'log.partial'       = 'Partial load ({0}): not reading {1}'
            'log.historySkipped' = 'Partial load: history.csv not updated'
            'export.warnPartialFull' = "The data were loaded partially ({0}): the FULL export will miss {1}.`n`nExport anyway?"
            'export.warnMissing' = "These columns were not read in the last load: {0}.`nLoad the data again to include them.`n`nExport without them?"
            'export.title'      = 'Export report'
            'export.subtitle'   = 'Choose before loading the data: the full report, a ready-made preset, or only the columns you tick. Export XLSX / CSV will write this.'
            'export.content'    = 'Content'
            'export.options'    = 'Options'
            'export.columns'    = 'User columns'
            'export.all'        = 'All'
            'export.none'       = 'None'
            'export.optGroups'  = 'One sheet per group'
            'export.optOverview' = 'Include OVERVIEW'
            'export.optEnabled' = 'Enabled accounts only'
            'export.savePreset' = 'Save as preset...'
            'export.deletePreset' = 'Delete preset'
            'export.cancel'     = 'Cancel'
            'export.preset.Full' = 'FULL (all tabs, all columns)'
            'export.preset.Access' = 'Sign-ins and inactivity'
            'export.preset.Licenses' = 'Licenses'
            'export.preset.Migration' = 'Mailbox migration'
            'export.preset.MFA' = 'MFA security'
            'export.preset.Directory' = 'Directory (who and where)'
            'export.preset.Custom' = 'Custom'
            'export.desc.Full'  = 'Every tab (overview, users, groups, mailboxes, OneDrive, MFA, licenses, attention, sync, history) with every column.'
            'export.desc.Access' = 'Last AD logon, last Microsoft 365 sign-in and password age: find the accounts nobody uses.'
            'export.desc.Licenses' = 'Who has which license and when they last signed in: free up unused licenses.'
            'export.desc.Migration' = 'Where each mailbox lives and how big it is: plan the move to Exchange Online.'
            'export.desc.MFA'   = 'MFA status, default method and registered methods of every user.'
            'export.desc.Directory' = 'Names, username, UPN, OU, status and origin: a clean user list.'
            'export.desc.Custom' = 'Tick the columns you want. The selection is remembered for next time.'
            'export.desc.User'  = 'Your preset (saved in config.json). Changing a column switches to Custom.'
            'export.summary'    = '{0} columns  ·  {1} users'
            'export.msgNoCols'  = 'Tick at least one column.'
            'export.askName'    = 'Name of the new preset:'
            'export.badName'    = '"{0}" is the name of a built-in preset: choose another name.'
            'export.confirmDelete' = 'Delete the preset "{0}"?'
            'err.exportNoUsers' = 'There is no user list to export.'
            'err.exportNoColumns' = 'None of these columns exists in the loaded data: {0}'
            'err.exportNoRows'  = 'No user matches the export options.'
            'err.exportPreset'  = 'Unknown export preset "{0}". Available: {1}'
            'log.exportPreset'  = 'Export {0}: {1} columns, {2} users'
            'btn.tab'           = 'Export current tab'
            'btn.folder'        = 'Open reports folder'
            'lbl.search'        = 'Search in tab:'
            'lbl.rows'          = '{0} / {1} rows'
            'lbl.language'      = 'Language'
            'lbl.generated'     = 'Data as of {0}'
            'st.ready'          = 'Ready'
            'exp.log'           = 'Log'
            'empty.hint'        = '1. Connect 365      2. Connect Exchange      3. Verify AD      4. Load data'
            'st.login'          = 'Signing in (browser)...'
            'st.connecting'     = 'Connecting...'
            'st.verifying'      = 'Verifying...'
            'st.error'          = 'Error'
            'st.loaded'         = 'Loaded {0} users - double-click a user to open their card'
            'st.exported'       = 'Exported {0} rows to {1}'
            'job.init'          = 'Initializing...'
            'job.365'           = 'Connecting to Microsoft 365...'
            'job.exch'          = 'Connecting to Exchange on-premises...'
            'job.ad'            = 'Verifying Active Directory...'
            'job.load'          = 'Collecting data...'
            'job.xlsx'          = 'Exporting Excel...'
            'job.csv'           = 'Exporting CSV...'
            'job.modules'       = 'Installing modules...'
            'empty.ready'       = 'Complete the steps above, then click Load data'
            'mode.cloud'        = 'Microsoft 365 only'
            'mode.fullhybrid'   = 'Microsoft 365 + Exchange on-premises + Active Directory'
            'mode.hybrid'       = 'Microsoft 365 + Active Directory (no Exchange on-premises)'
            'mode.onprem'       = 'Exchange on-premises + Active Directory (no Microsoft 365)'
            'mode.ad'           = 'Active Directory only'
            'msg.modeReload'    = 'Infrastructure changed. Reload the data now?'
            'title.mode'        = 'Infrastructure'
            'btn.signout'       = 'Change account'
            'tip.signout'       = 'Signs out of Microsoft 365 and clears the saved sign-in, so you can connect with a different account'
            'job.signout'       = 'Signing out of Microsoft 365...'
            'log.signedOut'     = 'Signed out of Microsoft 365, cached sign-in cleared'
            'btn.stop'          = 'Stop'
            'tip.stop'          = 'Stops the running operation'
            'st.stopping'       = 'Stopping... (waits for the current request to finish)'
            'st.cancelled'      = 'Operation cancelled'
            'lbl.groupBy'       = 'Group by:'
            'tip.groupBy'       = 'How users are split into tabs. Top-level OU = first OU under the domain; Parent OU = the OU that directly contains the user; Custom rules = GroupRules in config.json'
            'grp.opt.top'       = 'Top-level OU'
            'grp.opt.parent'    = 'Parent OU (where the user is)'
            'grp.opt.none'      = 'No grouping'
            'grp.opt.rules'     = 'Custom rules (config.json)'
            'msg.regroup'       = 'Grouping changed. Reload the data now?'
            'msg.noRules'       = 'No GroupRules defined in config.json: "Top-level OU" will be used. See the README for examples.'
            'title.modules'     = 'Missing modules'
            'title.credentials' = 'Credentials'
            'title.warnings'    = 'Warnings'
            'title.error'       = 'Error'
            'title.loadError'   = 'Data collection error'
            'title.language'    = 'Language'
            'msg.modulesMissing' = "Missing PowerShell modules:`n`n - {0}`n`nInstall them now?"
            'msg.rsatNote'      = "For ActiveDirectory (RSAT) a PowerShell window will open and ask for administrator confirmation."
            'msg.modulesOk'     = 'Modules installed.'
            'msg.modulesFail'   = "Installation failed:`n{0}"
            'msg.adMissing'     = "The ActiveDirectory module (RSAT) is not installed.`n`nInstall it now? A PowerShell window will ask for administrator confirmation."
            'msg.notDomain'     = 'This PC is not joined to a domain: Windows credentials cannot be used. Uncheck the box and enter a domain user and password.'
            'msg.credMissing'   = 'Enter a domain user and password (or check "Use Windows credentials").'
            'msg.basicNeedsCred' = 'Basic authentication requires a user and password: uncheck "Use Windows credentials".'
            'msg.exchNoServer'  = 'Enter the Exchange server name (internal FQDN, e.g. exch01.contoso.local).'
            'msg.exchHints'     = "{0}`n`nChecks:`n- use http + Kerberos with the INTERNAL server name`n- from a PC outside the domain uncheck 'Use Windows credentials' and enter user@domain.local`n- the user needs Remote PowerShell enabled (Get-User <user> | fl RemotePowerShellEnabled)`n- port 80 to the Exchange server must be open`n- or run HybridScope on the Exchange server and choose 'local'"
            'msg.loadWarn'      = "Data loaded with warnings:`n`n- {0}"
            'msg.saved'         = "Report saved:`n{0}`n`nOpen it now?"
            'msg.restart'       = 'Restart HybridScope to apply the new language.'
            'dlg.csvFolder'     = 'Folder where one CSV file per tab will be saved'
            'dlg.filterCsv'     = 'CSV (semicolon)|*.csv|Excel (*.xlsx)|*.xlsx'

            'log.ready'         = 'Engine ready (HybridScope {0})'
            'log.connected365'  = 'Connected to Microsoft 365 as {0}'
            'log.exchConnecting' = 'Connecting to Exchange on-premises: {0} ({1})'
            'log.exchLocal'     = 'Loading the local Exchange Management Shell (snap-in)...'
            'log.exchConnected' = 'Connected to Exchange: {0}'
            'log.adOk'          = 'Active Directory OK: {0} (PDC {1})'
            'log.skus'          = 'Tenant licenses (SKU): {0}'
            'log.entraUsers'    = 'Entra users: {0}'
            'warn.signin'       = 'Last Microsoft 365 sign-in not available (requires Entra ID P1 and AuditLog.Read.All): {0}'
            'log.mfaReg'        = 'MFA registrations: {0}'
            'warn.mfaReport'    = 'MFA registration report not available ({0}): reading methods user by user'
            'log.mfaProgress'   = '  MFA methods {0}/{1}'
            'log.perUser'       = 'Reading per-user MFA state...'
            'log.perUserProgress' = '  per-user MFA {0}/{1}'
            'warn.perUser'      = 'Per-user MFA state not readable, skipped ({0})'
            'log.reports'       = 'Microsoft 365 mailboxes: {0} - OneDrive: {1}'
            'warn.concealed'    = "Microsoft 365 usage reports are ANONYMIZED: 365 mailbox and OneDrive sizes cannot be matched to users. Turn it off in Microsoft 365 admin center > Settings > Org settings > Reports > 'Display concealed user, group, and site names in all reports'."
            'log.mbxOnPrem'     = 'On-premises mailboxes: {0}'
            'log.remoteMbx'     = 'Remote mailboxes (already in Exchange Online): {0}'
            'warn.remoteMbx'    = 'Get-RemoteMailbox not available: {0}'
            'log.statsProgress' = '  mailbox statistics {0}/{1}'
            'warn.stats'        = 'Statistics not available for {0}: {1}'
            'warn.exch'         = 'Exchange on-premises not available: {0}'
            'log.exchSkipped'   = 'Exchange on-premises not configured: skipped'
            'log.adUsers'       = 'AD users: {0}'
            'warn.ad'           = 'Active Directory not available: {0}'
            'log.adSkipped'     = 'Active Directory domain not configured: skipped'
            'log.rows'          = 'Users in report: {0} in {1} groups'
            'log.done'          = 'Done: {0} users | 365: {1} | ON-PREM: {2} | MFA: {3}%'
            'log.excelSheet'    = '  sheet {0} ({1} rows)'
            'log.excelSaved'    = 'Excel saved: {0}'
            'log.csvSaved'      = 'CSV saved to {0}'
            'log.rsat'          = 'Installing RSAT Active Directory: confirm the administrator prompt and follow the window that opens...'
            'log.rsatOk'        = 'ActiveDirectory module ready'
            'err.rsat'          = 'RSAT Active Directory installation failed: see the messages in the installation window.'
            'err.rsatFile'      = 'File not found: {0}'
            'log.installing'    = 'Installing {0}...'
            'log.modulesDone'   = 'Module installation completed'
            'err.not365'        = 'Not connected to Microsoft 365: click "Connect 365" (or configure the certificate for -Auto).'
            'err.graphConnect'  = 'Connection to Microsoft Graph failed'
            'warn.latestLocked' = 'HybridScope_LATEST.xlsx not updated (is it open in Excel?)'
            'err.generic'       = 'ERROR: {0}'
            'warn.noRows'       = 'No data returned for {0}: no file written'
            'warn.scopes'       = 'Sign-in with the group permissions (GroupMember.Read.All, RoleManagement.Read.Directory) failed, connected without them: user cards will not show Entra groups and roles until an administrator consents once ({0})'
            'warn.adConsistency' = 'ms-DS-ConsistencyGuid not readable, objectGUID used as the sync anchor ({0})'
            'warn.org'          = 'Tenant directory sync status not readable: {0}'
            'warn.syncAnchor'   = 'Only {0} of {1} synced Entra accounts match an AD account by anchor: the source anchor is probably a custom attribute or another forest is synced. Anchor checks (SYNC categories 1, 2) skipped.'
            'warn.syncStale'    = 'Directory sync: last cycle {0} ago (threshold {1} h)'
            'log.syncStart'     = 'Checking AD <-> Entra ID sync...'
            'log.sync'          = 'Sync checks: {0} findings ({1} s)'
            'warn.sync'         = 'Sync checks skipped: {0}'
            'err.noDomain'      = 'Active Directory domain not configured (DomainFqdn in config.json).'

            'sync.c0'           = '0 - Directory sync status'
            'sync.c1'           = '1 - ImmutableId differs from the AD anchor'
            'sync.c2'           = '2 - Synced account not found in AD'
            'sync.c3'           = '3 - Cloud account with a leftover ImmutableId'
            'sync.c4'           = '4 - Will soft-match a cloud account'
            'sync.c5'           = '5 - Conflicts with a cloud account (different ImmutableId)'
            'sync.c6'           = '6 - Provisioning error in Entra'
            'sync.c7'           = '7 - Entra UPN fell back to onmicrosoft.com'
            'sync.c8'           = '8 - Same anchor on more than one AD account'
            'sync.d.last'       = 'Last sync {0} ({1} ago)'
            'sync.d.stale'      = 'NOT SYNCING: last sync {0} ({1} ago, threshold {2} h). Check the Entra Connect / Cloud Sync server.'
            'sync.d.off'        = 'Directory sync is not enabled in this tenant'
            'sync.d.mismatch'   = 'AD anchor {0} ({1}), Entra ImmutableId {2}. Usually the AD account was deleted and recreated: Entra is still linked to the old object.'
            'sync.d.notInAd'    = 'Entra is linked to {0}, but no AD account has ImmutableId {1}: deleted or recreated in AD, or sync not running.'
            'sync.d.leftover'   = 'Sync is off for this account but ImmutableId {0} is still set: an AD account with this anchor would take it over (hard match).'
            'sync.d.leftoverAd' = 'Sync is off for this account but its ImmutableId belongs to AD account {0}: it will be taken over if {0} comes back into sync scope.'
            'sync.d.soft'       = 'Cloud account {0} has the same {1}: at the first sync of this AD account the two will be merged (soft match), unless soft match is blocked in the tenant.'
            'sync.d.conflict'   = 'Cloud account {0} has the same {1} but ImmutableId {2} (AD anchor {3}): sync will fail with a duplicate attribute error.'
            'sync.d.prov'       = '{0} = {1} ({2}, {3})'
            'sync.d.upn'        = 'Entra: {0} - AD userPrincipalName: {1}, mail: {2}. The attribute the sync uses as UPN has a suffix that is not a verified domain in the tenant.'
            'sync.d.upnSuffix'  = 'Not verified: {0}'
            'sync.d.dup'        = 'Anchor {0} also on: {1}'

            'menu.open'         = '_Open report...'
            'dlg.open'          = 'HybridScope report (*.xlsx;*.csv)|*.xlsx;*.csv|All files (*.*)|*.*'
            'job.import'        = 'Opening report...'
            'log.import'        = 'Opening report {0}'
            'log.importSheet'   = '  {0} ({1} rows)'
            'st.imported'       = 'Report opened: {0} users - double-click a user to open their card'
            'lbl.imported'      = 'Report of {0}  ·  {1}'
            'title.import'      = 'Open report'
            'err.importEmpty'   = 'No HybridScope tables found in {0}.'
            'card.imported'     = 'Groups and roles are read live from AD and Microsoft 365: not available for an imported report.'
            'menu.file'         = '_File'
            'menu.exit'         = 'E_xit'
            'menu.help'         = '_Help'
            'menu.docs'         = 'Documentation (GitHub)'
            'menu.updates'      = 'Check for updates...'
            'menu.issue'        = 'Report a problem or suggest an idea'
            'menu.sponsor'      = 'Sponsor on GitHub'
            'menu.coffee'       = 'Buy me a coffee'
            'menu.about'        = 'About HybridScope'
            'upd.checking'      = 'Checking for updates...'
            'upd.new'           = "HybridScope {0} is available (you have {1}).`n`nOpen the download page?"
            'upd.latest'        = 'You have the latest version ({0}).'
            'upd.ahead'         = 'This version ({0}) is newer than the latest release ({1}).'
            'upd.error'         = "Could not check for updates:`n{0}`n`nYou can check the releases page on GitHub."
            'title.updates'     = 'Updates'
            'about.title'       = 'About HybridScope'
            'about.version'     = 'Version {0}'
            'about.desc'        = 'Read-only inventory of hybrid identities and mailboxes.'
            'about.by'          = 'Made by'
            'about.project'     = 'Project'
            'about.licenseLbl'  = 'License'
            'about.license'     = 'MIT - free to use, modify and share'
            'about.feedback'    = 'Feedback'
            'about.support'     = 'If HybridScope saves you time, you can support its development:'
            'about.star'        = 'Star on GitHub'
            'about.disclaimer'  = 'Independent project, not affiliated with or endorsed by Microsoft.'
            'about.close'       = 'Close'
            'status.by'         = 'by'
            'card.copy'         = 'Copy'
            'card.reload'       = 'Reload'
            'card.search'       = 'Search:'
            'card.groups'       = 'Groups and roles'
            'card.loading'      = 'Loading groups and roles...'
            'card.busy'         = 'Another operation is running: click Reload when it has finished.'
            'card.count'        = '{0} AD groups  ·  {1} Entra groups  ·  {2} Entra roles'
            'card.notFound'     = 'User not found in the loaded data.'
            'card.sec.identity' = 'Identity'
            'card.sec.ad'       = 'Active Directory'
            'card.sec.m365'     = 'Microsoft 365'
            'card.sec.mailbox'  = 'Mailbox'
            'card.sec.onedrive' = 'OneDrive'
            'card.sec.mfa'      = 'MFA'
            'card.sec.sync'     = 'Sync'
            'card.sec.notes'    = 'Notes'
            'card.direct'       = 'Direct'
            'card.primary'      = 'Direct (primary group)'
            'card.inherited'    = 'Inherited (nested)'
            'card.type.security' = 'Security'
            'card.type.distribution' = 'Distribution'
            'card.type.m365'    = 'Microsoft 365'
            'card.type.mailSec' = 'Mail-enabled security'
            'card.type.dl'      = 'Distribution list'
            'card.type.dynamic' = '{0} (dynamic)'
            'card.type.role'    = 'Entra role'
            'card.type.au'      = 'Administrative unit'
            'card.err.ad'       = 'AD groups not readable: {0}'
            'card.err.graph'    = 'Entra groups not readable: {0}'
            'card.err.perm'     = 'Some Entra groups or roles have no name: the {0} permission is missing. An administrator must consent to it once (Change account, then Connect 365 as an admin).'

            'val.enabled'       = 'Enabled'
            'val.disabled'      = 'Disabled'
            'val.none'          = 'None'
            'val.yes'           = 'Yes'
            'val.no'            = 'No'
            'val.mfaYes'        = 'Configured'
            'val.mfaNo'         = 'Not configured'
            'val.mfaNotInEntra' = 'Not in Entra'
            'val.originSynced'  = 'AD synced'
            'val.originAD'      = 'AD only'
            'val.originCloud'   = 'Cloud only'
            'val.originGuest'   = 'Guest'
            'val.originUnlinked' = 'AD + cloud (not linked)'
            'src.getMailbox'    = 'Exchange on-prem (Get-Mailbox)'
            'src.remote'        = 'Exchange on-prem (RemoteMailbox)'
            'src.graphPlan'     = 'Graph (Exchange Online enabled)'
            'src.graphReport'   = 'Graph (usage report)'
            'src.exchNotRead'   = 'Exchange on-prem not read'
            'note.multiOnPrem'  = '{0} on-prem mailboxes linked'
            'note.no365Size'    = '365 mailbox size missing from the Microsoft report (anonymized report or new mailbox)'
            'note.licensedDisabled' = 'Licensed but disabled'
            'note.noMfa'        = 'Licensed without MFA'
            'note.notSynced'    = 'AD account not synced to Entra'
            'note.pwdNeverExpires' = 'Password never expires'
            'quota.dbDefault'   = 'Database default'
            'att.never'         = 'Never'
            'att.last'          = 'Last: {0:yyyy-MM-dd}'
            'att.c1'            = '1 - On-prem mailbox (to migrate)'
            'att.c2'            = '2 - Licensed without MFA'
            'att.c3'            = '3 - Licensed but disabled'
            'att.c4'            = '4 - Licensed, no 365 sign-in for >{0} days'
            'att.c5'            = '5 - Enabled, no AD logon for >{0} days'
            'att.c6'            = '6 - AD account not synced'
            'att.c7'            = '7 - Sync problem (see SYNC)'

            'm.sec.users'       = 'USERS'
            'm.total'           = 'Total users'
            'm.enabled'         = 'Enabled'
            'm.disabled'        = 'Disabled'
            'm.adOnly'          = 'AD only (not synced)'
            'm.cloudOnly'       = 'Cloud only'
            'm.unlinked'        = 'AD + cloud account, not linked'
            'm.syncIssues'      = 'With sync problems'
            'm.sec.location'    = 'MAILBOX LOCATION'
            'm.loc'             = 'Users {0}'
            'm.sec.space'       = 'SPACE (GB)'
            'm.spaceOnPrem'     = 'On-prem mailboxes (GB)'
            'm.space365'        = '365 mailboxes (GB)'
            'm.spaceOD'         = 'OneDrive (GB)'
            'm.avgOnPrem'       = 'Average on-prem mailbox (GB)'
            'm.maxOnPrem'       = 'Largest on-prem mailbox (GB)'
            'm.sec.lic'         = 'ASSIGNED LICENSES'
            'm.noLic'           = 'No license'
            'm.licDisabled'     = 'Licensed but disabled'
            'm.sec.mfa'         = 'MFA'
            'm.mfaYes'          = 'MFA configured'
            'm.mfaNo'           = 'MFA NOT configured'
            'm.mfaNotInEntra'   = 'Not in Entra'
            'm.mfaPct'          = '% MFA configured (Entra users)'
            'm.licNoMfa'        = 'Enabled + licensed WITHOUT MFA'
            'm.perUserEnforced' = 'Per-user MFA enforced'
            'm.sec.inactive'    = 'INACTIVITY'
            'm.inactiveAD'      = 'Enabled, no AD logon for >{0} days'
            'm.inactive365'     = 'Licensed, no 365 sign-in for >{0} days'
            'm.pwdOld'          = 'Password not changed for >{0} days'
            'col.metric'        = 'Metric'
            'col.total'         = 'TOTAL'

            'tab.overview'      = 'OVERVIEW'
            'tab.users'         = 'USERS'
            'tab.mbxOnPrem'     = 'MAILBOX_ONPREM'
            'tab.mbx365'        = 'MAILBOX_365'
            'tab.onedrive'      = 'ONEDRIVE'
            'tab.mfa'           = 'MFA'
            'tab.licenses'      = 'TENANT_LICENSES'
            'tab.attention'     = 'ATTENTION'
            'tab.sync'          = 'SYNC'
            'tab.history'       = 'HISTORY'
            'grp.other'         = 'OTHER'
            'grp.cloud'         = 'CLOUD_ONLY'
            'grp.all'           = 'ALL'

            'xl.generated'      = 'Generated on {0} - {1}'
            'xl.legend'         = 'LOCATION: 365 = mailbox in Exchange Online | ON-PREM = mailbox on Exchange on-premises | NO MAILBOX = no mailbox. Microsoft 365 sizes come from Microsoft usage reports (up to 48h delay).'
            'xl.warning'        = 'WARNING: {0}'
        }

        it = @{
            'app.subtitle'      = 'Active Directory  ·  Exchange on-premises  ·  Microsoft 365'
            'gb.m365'           = 'Microsoft 365'
            'lbl.tenant'        = 'Tenant'
            'tip.tenant'        = "Dominio o ID del tenant. Lascia vuoto per usare il tenant dell'account con cui accedi."
            'btn.connect365'    = 'Connetti 365'
            'st.notconnected'   = 'Non connesso'
            'gb.cred'           = 'Credenziali di dominio (Exchange + AD)'
            'chk.wincred'       = 'Usa credenziali Windows'
            'lbl.user'          = 'Utente'
            'lbl.password'      = 'Password'
            'tip.user'          = 'es. mario.rossi@contoso.local oppure CONTOSO\mario.rossi'
            'tip.notDomain'     = 'Questo PC non è nel dominio: inserisci le credenziali a mano'
            'gb.exch'           = 'Exchange on-premises'
            'lbl.server'        = 'Server'
            'tip.server'        = 'FQDN interno di un server Exchange, es. exch01.contoso.local'
            'tip.proto'         = 'http = WinRM + Kerberos (standard, nessun certificato) | https = serve un certificato valido | local = HybridScope avviato SUL server Exchange'
            'btn.connectExch'   = 'Connetti Exchange'
            'gb.ad'             = 'Active Directory'
            'lbl.domain'        = 'Dominio'
            'btn.verifyAD'      = 'Verifica AD'
            'st.notverified'    = 'Non verificato'
            'btn.load'          = 'Carica dati'
            'btn.xlsx'          = 'Esporta XLSX...'
            'btn.csvall'        = 'Esporta CSV...'
            'tip.export'        = 'Esporta quello che hai scelto in "Contenuto export"'
            'gb.export'         = 'Contenuto export'
            'btn.expCols'       = 'Colonne...'
            'tip.expCols'       = 'Scegli un preset o spunta le colonne della lista utenti, prima o dopo aver caricato i dati'
            'tip.cmbExport'     = 'Cosa scriveranno Esporta XLSX / CSV. Sceglilo prima di Carica dati.'
            'export.boxFull'    = 'tutti i tab, tutte le colonne'
            'export.boxCols'    = '{0} colonne: {1}'
            'export.summaryCols' = '{0} colonne'
            'export.apply'      = 'OK'
            'export.fast'       = 'Leggi solo i dati necessari'
            'tip.expFast'       = 'Con un preset o Personalizzato, Carica dati legge solo cio che serve a quelle colonne (molto piu veloce, es. niente mailbox Exchange o MFA). Panoramica, tab e schede utente saranno parziali.'
            'lbl.partial'       = 'CARICAMENTO PARZIALE: {0}'
            'tip.partial'       = 'Non letti: {0}. Panoramica, tab e schede utente sono incompleti. Ricarica con FULL, oppure togli "Leggi solo i dati necessari", per avere tutto.'
            'log.partial'       = 'Caricamento parziale ({0}): non leggo {1}'
            'log.historySkipped' = 'Caricamento parziale: history.csv non aggiornato'
            'export.warnPartialFull' = "I dati sono stati caricati in modo parziale ({0}): nell'export FULL mancheranno {1}.`n`nEsportare comunque?"
            'export.warnMissing' = "Queste colonne non sono state lette nell'ultimo caricamento: {0}.`nRicarica i dati per includerle.`n`nEsportare senza?"
            'export.title'      = 'Esporta report'
            'export.subtitle'   = 'Scegli prima di caricare i dati: il report completo, un preset pronto oppure solo le colonne che spunti. Esporta XLSX / CSV scriveranno questo.'
            'export.content'    = 'Contenuto'
            'export.options'    = 'Opzioni'
            'export.columns'    = 'Colonne utente'
            'export.all'        = 'Tutte'
            'export.none'       = 'Nessuna'
            'export.optGroups'  = 'Un foglio per gruppo'
            'export.optOverview' = 'Includi PANORAMICA'
            'export.optEnabled' = 'Solo account abilitati'
            'export.savePreset' = 'Salva come preset...'
            'export.deletePreset' = 'Elimina preset'
            'export.cancel'     = 'Annulla'
            'export.preset.Full' = 'FULL (tutti i tab, tutte le colonne)'
            'export.preset.Access' = 'Accessi e inattivita'
            'export.preset.Licenses' = 'Licenze'
            'export.preset.Migration' = 'Migrazione mailbox'
            'export.preset.MFA' = 'Sicurezza MFA'
            'export.preset.Directory' = 'Anagrafica (chi e dove)'
            'export.preset.Custom' = 'Personalizzato'
            'export.desc.Full'  = 'Tutti i tab (panoramica, utenti, gruppi, mailbox, OneDrive, MFA, licenze, da attenzionare, sync, storico) con tutte le colonne.'
            'export.desc.Access' = 'Ultimo logon AD, ultimo accesso Microsoft 365 ed eta della password: trova gli account che nessuno usa.'
            'export.desc.Licenses' = 'Chi ha quale licenza e quando ha fatto l''ultimo accesso: libera le licenze inutilizzate.'
            'export.desc.Migration' = 'Dove si trova ogni mailbox e quanto e grande: pianifica il passaggio a Exchange Online.'
            'export.desc.MFA'   = 'Stato MFA, metodo predefinito e metodi registrati di ogni utente.'
            'export.desc.Directory' = 'Nome, cognome, username, UPN, OU, stato e origine: una lista utenti pulita.'
            'export.desc.Custom' = 'Spunta le colonne che vuoi. La scelta viene ricordata per la prossima volta.'
            'export.desc.User'  = 'Il tuo preset (salvato in config.json). Se cambi una colonna passa a Personalizzato.'
            'export.summary'    = '{0} colonne  ·  {1} utenti'
            'export.msgNoCols'  = 'Spunta almeno una colonna.'
            'export.askName'    = 'Nome del nuovo preset:'
            'export.badName'    = '"{0}" e il nome di un preset predefinito: scegli un altro nome.'
            'export.confirmDelete' = 'Eliminare il preset "{0}"?'
            'err.exportNoUsers' = 'Non c''e nessuna lista utenti da esportare.'
            'err.exportNoColumns' = 'Nessuna di queste colonne esiste nei dati caricati: {0}'
            'err.exportNoRows'  = 'Nessun utente corrisponde alle opzioni di export.'
            'err.exportPreset'  = 'Preset di export sconosciuto "{0}". Disponibili: {1}'
            'log.exportPreset'  = 'Export {0}: {1} colonne, {2} utenti'
            'btn.tab'           = 'Esporta tab corrente'
            'btn.folder'        = 'Apri cartella report'
            'lbl.search'        = 'Cerca nel tab:'
            'lbl.rows'          = '{0} / {1} righe'
            'lbl.language'      = 'Lingua'
            'lbl.generated'     = 'Dati del {0}'
            'st.ready'          = 'Pronto'
            'exp.log'           = 'Log'
            'empty.hint'        = '1. Connetti 365      2. Connetti Exchange      3. Verifica AD      4. Carica dati'
            'st.login'          = 'Accesso in corso (browser)...'
            'st.connecting'     = 'Connessione...'
            'st.verifying'      = 'Verifica...'
            'st.error'          = 'Errore'
            'st.loaded'         = 'Caricati {0} utenti - doppio clic su un utente per aprire la scheda'
            'st.exported'       = 'Esportate {0} righe in {1}'
            'job.init'          = 'Inizializzazione...'
            'job.365'           = 'Connessione a Microsoft 365...'
            'job.exch'          = 'Connessione a Exchange on-premises...'
            'job.ad'            = 'Verifica Active Directory...'
            'job.load'          = 'Raccolta dati in corso...'
            'job.xlsx'          = 'Esportazione Excel...'
            'job.csv'           = 'Esportazione CSV...'
            'job.modules'       = 'Installazione moduli...'
            'empty.ready'       = 'Completa i passaggi qui sopra, poi premi Carica dati'
            'mode.cloud'        = 'Solo Microsoft 365'
            'mode.fullhybrid'   = 'Microsoft 365 + Exchange on-premises + Active Directory'
            'mode.hybrid'       = 'Microsoft 365 + Active Directory (senza Exchange on-premises)'
            'mode.onprem'       = 'Exchange on-premises + Active Directory (senza Microsoft 365)'
            'mode.ad'           = 'Solo Active Directory'
            'msg.modeReload'    = 'Infrastruttura cambiata. Ricaricare i dati ora?'
            'title.mode'        = 'Infrastruttura'
            'btn.signout'       = 'Cambia account'
            'tip.signout'       = "Disconnette da Microsoft 365 e cancella l'accesso salvato, così puoi collegarti con un altro account"
            'job.signout'       = 'Disconnessione da Microsoft 365...'
            'log.signedOut'     = 'Disconnesso da Microsoft 365, accesso salvato cancellato'
            'btn.stop'          = 'Interrompi'
            'tip.stop'          = "Interrompe l'operazione in corso"
            'st.stopping'       = 'Interruzione in corso... (attende la fine della richiesta attuale)'
            'st.cancelled'      = 'Operazione interrotta'
            'lbl.groupBy'       = 'Raggruppa per:'
            'tip.groupBy'       = "Come dividere gli utenti nei tab. OU di primo livello = prima OU sotto il dominio; OU di appartenenza = l'OU che contiene direttamente l'utente; Regole personalizzate = GroupRules in config.json"
            'grp.opt.top'       = 'OU di primo livello'
            'grp.opt.parent'    = 'OU di appartenenza'
            'grp.opt.none'      = 'Nessun raggruppamento'
            'grp.opt.rules'     = 'Regole personalizzate (config.json)'
            'msg.regroup'       = 'Raggruppamento cambiato. Ricaricare i dati ora?'
            'msg.noRules'       = 'Nessuna regola GroupRules in config.json: verrà usata "OU di primo livello". Vedi il README per gli esempi.'
            'title.modules'     = 'Moduli mancanti'
            'title.credentials' = 'Credenziali'
            'title.warnings'    = 'Avvisi'
            'title.error'       = 'Errore'
            'title.loadError'   = 'Errore raccolta dati'
            'title.language'    = 'Lingua'
            'msg.modulesMissing' = "Moduli PowerShell mancanti:`n`n - {0}`n`nLi installo ora?"
            'msg.rsatNote'      = "Per ActiveDirectory (RSAT) si aprirà una finestra PowerShell che chiede la conferma di amministratore."
            'msg.modulesOk'     = 'Moduli installati.'
            'msg.modulesFail'   = "Installazione non riuscita:`n{0}"
            'msg.adMissing'     = "Il modulo ActiveDirectory (RSAT) non è installato.`n`nLo installo ora? Una finestra PowerShell chiederà la conferma di amministratore."
            'msg.notDomain'     = 'Questo PC non è nel dominio: le credenziali Windows non possono essere usate. Togli la spunta e inserisci utente e password di dominio.'
            'msg.credMissing'   = 'Inserisci utente e password di dominio (oppure metti la spunta su "Usa credenziali Windows").'
            'msg.basicNeedsCred' = 'Con autenticazione Basic servono utente e password: togli la spunta "Usa credenziali Windows".'
            'msg.exchNoServer'  = 'Inserisci il nome del server Exchange (FQDN interno, es. exch01.contoso.local).'
            'msg.exchHints'     = "{0}`n`nVerifiche:`n- usa http + Kerberos con il nome INTERNO del server`n- da un PC fuori dominio togli 'Usa credenziali Windows' e inserisci utente@dominio.local`n- l'utente deve avere Remote PowerShell abilitato (Get-User <utente> | fl RemotePowerShellEnabled)`n- porta 80 verso il server Exchange aperta`n- oppure avvia HybridScope sul server Exchange e scegli 'local'"
            'msg.loadWarn'      = "Dati caricati con avvisi:`n`n- {0}"
            'msg.saved'         = "Report salvato:`n{0}`n`nAprirlo ora?"
            'msg.restart'       = 'Riavvia HybridScope per applicare la nuova lingua.'
            'dlg.csvFolder'     = 'Cartella in cui salvare un file CSV per ogni tab'
            'dlg.filterCsv'     = 'CSV (separatore ;)|*.csv|Excel (*.xlsx)|*.xlsx'

            'log.ready'         = 'Motore pronto (HybridScope {0})'
            'log.connected365'  = 'Connesso a Microsoft 365 come {0}'
            'log.exchConnecting' = 'Connessione a Exchange on-premises: {0} ({1})'
            'log.exchLocal'     = 'Caricamento Exchange Management Shell locale (snap-in)...'
            'log.exchConnected' = 'Connesso a Exchange: {0}'
            'log.adOk'          = 'Active Directory OK: {0} (PDC {1})'
            'log.skus'          = 'Licenze tenant (SKU): {0}'
            'log.entraUsers'    = 'Utenti Entra: {0}'
            'warn.signin'       = 'Ultimo accesso a Microsoft 365 non disponibile (servono Entra ID P1 e AuditLog.Read.All): {0}'
            'log.mfaReg'        = 'Registrazioni MFA: {0}'
            'warn.mfaReport'    = 'Report registrazione MFA non disponibile ({0}): leggo i metodi utente per utente'
            'log.mfaProgress'   = '  metodi MFA {0}/{1}'
            'log.perUser'       = 'Lettura stato MFA per-user...'
            'log.perUserProgress' = '  MFA per-user {0}/{1}'
            'warn.perUser'      = 'Stato MFA per-user non leggibile, salto ({0})'
            'log.reports'       = 'Mailbox Microsoft 365: {0} - OneDrive: {1}'
            'warn.concealed'    = "I report di utilizzo di Microsoft 365 sono ANONIMIZZATI: dimensioni di mailbox 365 e OneDrive non associabili agli utenti. Disattiva in Interfaccia di amministrazione M365 > Impostazioni > Impostazioni organizzazione > Report > 'Visualizza nomi utente, gruppi e siti nascosti in tutti i report'."
            'log.mbxOnPrem'     = 'Mailbox on-premises: {0}'
            'log.remoteMbx'     = "Remote mailbox (già su Exchange Online): {0}"
            'warn.remoteMbx'    = 'Get-RemoteMailbox non disponibile: {0}'
            'log.statsProgress' = '  statistiche mailbox {0}/{1}'
            'warn.stats'        = 'Statistiche non disponibili per {0}: {1}'
            'warn.exch'         = 'Exchange on-premises non disponibile: {0}'
            'log.exchSkipped'   = 'Exchange on-premises non configurato: saltato'
            'log.adUsers'       = 'Utenti AD: {0}'
            'warn.ad'           = 'Active Directory non disponibile: {0}'
            'log.adSkipped'     = 'Dominio Active Directory non configurato: saltato'
            'log.rows'          = 'Utenti nel report: {0} in {1} gruppi'
            'log.done'          = 'Completato: {0} utenti | 365: {1} | ON-PREM: {2} | MFA: {3}%'
            'log.excelSheet'    = '  foglio {0} ({1} righe)'
            'log.excelSaved'    = 'Excel salvato: {0}'
            'log.csvSaved'      = 'CSV salvati in {0}'
            'log.rsat'          = 'Installazione RSAT Active Directory: conferma la richiesta di amministratore e segui la finestra che si apre...'
            'log.rsatOk'        = 'Modulo ActiveDirectory pronto'
            'err.rsat'          = 'Installazione RSAT Active Directory non riuscita: vedi i messaggi nella finestra di installazione.'
            'err.rsatFile'      = 'File non trovato: {0}'
            'log.installing'    = 'Installo {0}...'
            'log.modulesDone'   = 'Installazione moduli completata'
            'err.not365'        = 'Non connesso a Microsoft 365: premi "Connetti 365" (o configura il certificato per -Auto).'
            'err.graphConnect'  = 'Connessione a Microsoft Graph non riuscita'
            'warn.latestLocked' = 'HybridScope_LATEST.xlsx non aggiornato (è aperto in Excel?)'
            'err.generic'       = 'ERRORE: {0}'
            'warn.noRows'       = 'Nessun dato restituito per {0}: nessun file scritto'
            'warn.scopes'       = "Accesso con i permessi sui gruppi (GroupMember.Read.All, RoleManagement.Read.Directory) non riuscito, connesso senza: la scheda utente non mostrerà gruppi e ruoli Entra finché un amministratore non dà il consenso una volta ({0})"
            'warn.adConsistency' = "ms-DS-ConsistencyGuid non leggibile, uso objectGUID come ancora di sincronizzazione ({0})"
            'warn.org'          = 'Stato della sincronizzazione del tenant non leggibile: {0}'
            'warn.syncAnchor'   = "Solo {0} account Entra sincronizzati su {1} corrispondono a un account AD tramite l'ancora: probabilmente l'ancora è un attributo personalizzato o viene sincronizzata un'altra foresta. Controlli sull'ancora (categorie SYNC 1, 2) saltati."
            'warn.syncStale'    = 'Sincronizzazione directory: ultimo ciclo {0} fa (soglia {1} h)'
            'log.syncStart'     = 'Controllo sincronizzazione AD <-> Entra ID...'
            'log.sync'          = 'Controlli di sincronizzazione: {0} segnalazioni ({1} s)'
            'warn.sync'         = 'Controlli di sincronizzazione saltati: {0}'
            'err.noDomain'      = 'Dominio Active Directory non configurato (DomainFqdn in config.json).'

            'sync.c0'           = '0 - Stato della sincronizzazione'
            'sync.c1'           = "1 - ImmutableId diverso dall'ancora AD"
            'sync.c2'           = '2 - Account sincronizzato non trovato in AD'
            'sync.c3'           = '3 - Account cloud con ImmutableId residuo'
            'sync.c4'           = '4 - Farà soft-match con un account cloud'
            'sync.c5'           = '5 - In conflitto con un account cloud (ImmutableId diverso)'
            'sync.c6'           = '6 - Errore di provisioning in Entra'
            'sync.c7'           = '7 - UPN in Entra ripiegato su onmicrosoft.com'
            'sync.c8'           = '8 - Stessa ancora su più account AD'
            'sync.d.last'       = 'Ultima sincronizzazione {0} ({1} fa)'
            'sync.d.stale'      = 'NON SINCRONIZZA: ultima sincronizzazione {0} ({1} fa, soglia {2} h). Controlla il server Entra Connect / Cloud Sync.'
            'sync.d.off'        = 'La sincronizzazione della directory non è attiva in questo tenant'
            'sync.d.mismatch'   = "Ancora AD {0} ({1}), ImmutableId in Entra {2}. Di solito l'account AD è stato cancellato e ricreato: Entra è ancora collegato al vecchio oggetto."
            'sync.d.notInAd'    = 'Entra è collegato a {0}, ma nessun account AD ha ImmutableId {1}: cancellato o ricreato in AD, oppure la sincronizzazione non gira.'
            'sync.d.leftover'   = "Sincronizzazione disattivata per questo account ma ImmutableId {0} è ancora impostato: un account AD con questa ancora lo prenderebbe in carico (hard match)."
            'sync.d.leftoverAd' = "Sincronizzazione disattivata per questo account ma il suo ImmutableId appartiene all'account AD {0}: verrà ripreso se {0} torna nell'ambito della sincronizzazione."
            'sync.d.soft'       = "L'account cloud {0} ha lo stesso {1}: alla prima sincronizzazione di questo account AD i due verranno uniti (soft match), salvo che il soft match sia bloccato nel tenant."
            'sync.d.conflict'   = "L'account cloud {0} ha lo stesso {1} ma ImmutableId {2} (ancora AD {3}): la sincronizzazione fallirà con un errore di attributo duplicato."
            'sync.d.prov'       = '{0} = {1} ({2}, {3})'
            'sync.d.upn'        = "Entra: {0} - AD userPrincipalName: {1}, mail: {2}. L'attributo che la sincronizzazione usa come UPN ha un suffisso che non è un dominio verificato nel tenant."
            'sync.d.upnSuffix'  = 'Non verificati: {0}'
            'sync.d.dup'        = 'Ancora {0} presente anche su: {1}'

            'menu.open'         = '_Apri report...'
            'dlg.open'          = 'Report HybridScope (*.xlsx;*.csv)|*.xlsx;*.csv|Tutti i file (*.*)|*.*'
            'job.import'        = 'Apertura report...'
            'log.import'        = 'Apertura report {0}'
            'log.importSheet'   = '  {0} ({1} righe)'
            'st.imported'       = 'Report aperto: {0} utenti - doppio clic su un utente per aprire la scheda'
            'lbl.imported'      = 'Report del {0}  ·  {1}'
            'title.import'      = 'Apri report'
            'err.importEmpty'   = 'Nessuna tabella HybridScope trovata in {0}.'
            'card.imported'     = 'Gruppi e ruoli vengono letti dal vivo da AD e Microsoft 365: non disponibili per un report importato.'
            'menu.file'         = '_File'
            'menu.exit'         = '_Esci'
            'menu.help'         = '_Aiuto'
            'menu.docs'         = 'Documentazione (GitHub)'
            'menu.updates'      = 'Cerca aggiornamenti...'
            'menu.issue'        = "Segnala un problema o proponi un'idea"
            'menu.sponsor'      = 'Sponsorizza su GitHub'
            'menu.coffee'       = 'Offrimi un caffè'
            'menu.about'        = 'Informazioni su HybridScope'
            'upd.checking'      = 'Ricerca aggiornamenti...'
            'upd.new'           = "È disponibile HybridScope {0} (tu hai la {1}).`n`nAprire la pagina di download?"
            'upd.latest'        = 'Hai già la versione più recente ({0}).'
            'upd.ahead'         = "Questa versione ({0}) è più recente dell'ultima release ({1})."
            'upd.error'         = "Impossibile cercare aggiornamenti:`n{0}`n`nPuoi controllare la pagina delle release su GitHub."
            'title.updates'     = 'Aggiornamenti'
            'about.title'       = 'Informazioni su HybridScope'
            'about.version'     = 'Versione {0}'
            'about.desc'        = 'Inventario in sola lettura di identità e mailbox ibride.'
            'about.by'          = 'Creato da'
            'about.project'     = 'Progetto'
            'about.licenseLbl'  = 'Licenza'
            'about.license'     = 'MIT - libero da usare, modificare e condividere'
            'about.feedback'    = 'Feedback'
            'about.support'     = 'Se HybridScope ti fa risparmiare tempo, puoi sostenerne lo sviluppo:'
            'about.star'        = 'Metti una stella su GitHub'
            'about.disclaimer'  = 'Progetto indipendente, non affiliato né approvato da Microsoft.'
            'about.close'       = 'Chiudi'
            'status.by'         = 'di'
            'card.copy'         = 'Copia'
            'card.reload'       = 'Ricarica'
            'card.search'       = 'Cerca:'
            'card.groups'       = 'Gruppi e ruoli'
            'card.loading'      = 'Caricamento gruppi e ruoli...'
            'card.busy'         = "C'è un'altra operazione in corso: premi Ricarica quando è finita."
            'card.count'        = '{0} gruppi AD  ·  {1} gruppi Entra  ·  {2} ruoli Entra'
            'card.notFound'     = 'Utente non trovato nei dati caricati.'
            'card.sec.identity' = 'Identità'
            'card.sec.ad'       = 'Active Directory'
            'card.sec.m365'     = 'Microsoft 365'
            'card.sec.mailbox'  = 'Mailbox'
            'card.sec.onedrive' = 'OneDrive'
            'card.sec.mfa'      = 'MFA'
            'card.sec.sync'     = 'Sincronizzazione'
            'card.sec.notes'    = 'Note'
            'card.direct'       = 'Diretta'
            'card.primary'      = 'Diretta (gruppo primario)'
            'card.inherited'    = 'Ereditata (annidata)'
            'card.type.security' = 'Sicurezza'
            'card.type.distribution' = 'Distribuzione'
            'card.type.m365'    = 'Microsoft 365'
            'card.type.mailSec' = 'Sicurezza abilitata alla posta'
            'card.type.dl'      = 'Lista di distribuzione'
            'card.type.dynamic' = '{0} (dinamico)'
            'card.type.role'    = 'Ruolo Entra'
            'card.type.au'      = 'Unità amministrativa'
            'card.err.ad'       = 'Gruppi AD non leggibili: {0}'
            'card.err.graph'    = 'Gruppi Entra non leggibili: {0}'
            'card.err.perm'     = 'Alcuni gruppi o ruoli Entra sono senza nome: manca il permesso {0}. Un amministratore deve dare il consenso una volta (Cambia account, poi Connetti 365 come amministratore).'

            'val.enabled'       = 'Abilitato'
            'val.disabled'      = 'Disabilitato'
            'val.none'          = 'Nessuna'
            'val.yes'           = 'Sì'
            'val.no'            = 'No'
            'val.mfaYes'        = 'Configurato'
            'val.mfaNo'         = 'Non configurato'
            'val.mfaNotInEntra' = 'Non presente in Entra'
            'val.originSynced'  = 'AD sincronizzato'
            'val.originAD'      = 'Solo AD'
            'val.originCloud'   = 'Solo cloud'
            'val.originGuest'   = 'Guest'
            'val.originUnlinked' = 'AD + cloud (non collegati)'
            'src.getMailbox'    = 'Exchange on-prem (Get-Mailbox)'
            'src.remote'        = 'Exchange on-prem (RemoteMailbox)'
            'src.graphPlan'     = 'Graph (Exchange Online attivo)'
            'src.graphReport'   = 'Graph (report utilizzo)'
            'src.exchNotRead'   = 'Exchange on-prem non letto'
            'note.multiOnPrem'  = '{0} mailbox on-prem associate'
            'note.no365Size'    = 'Dimensione mailbox 365 assente dal report Microsoft (report anonimizzato o mailbox recente)'
            'note.licensedDisabled' = 'Licenziato ma disabilitato'
            'note.noMfa'        = 'Licenziato senza MFA'
            'note.notSynced'    = 'Account AD non sincronizzato in Entra'
            'note.pwdNeverExpires' = 'Password senza scadenza'
            'quota.dbDefault'   = 'Default database'
            'att.never'         = 'Mai'
            'att.last'          = 'Ultimo: {0:dd/MM/yyyy}'
            'att.c1'            = '1 - Mailbox ON-PREM (da migrare)'
            'att.c2'            = '2 - Licenziato senza MFA'
            'att.c3'            = '3 - Licenziato ma disabilitato'
            'att.c4'            = '4 - Licenziato inattivo su 365 da >{0} giorni'
            'att.c5'            = '5 - Abilitato inattivo in AD da >{0} giorni'
            'att.c6'            = '6 - Account AD non sincronizzato'
            'att.c7'            = '7 - Problema di sincronizzazione (vedi SYNC)'

            'm.sec.users'       = 'UTENTI'
            'm.total'           = 'Utenti totali'
            'm.enabled'         = 'Abilitati'
            'm.disabled'        = 'Disabilitati'
            'm.adOnly'          = 'Solo AD (non sincronizzati)'
            'm.cloudOnly'       = 'Solo cloud'
            'm.unlinked'        = 'AD + account cloud, non collegati'
            'm.syncIssues'      = 'Con problemi di sincronizzazione'
            'm.sec.location'    = 'POSIZIONE MAILBOX'
            'm.loc'             = 'Utenti {0}'
            'm.sec.space'       = 'SPAZIO (GB)'
            'm.spaceOnPrem'     = 'Mailbox ON-PREM (GB)'
            'm.space365'        = 'Mailbox 365 (GB)'
            'm.spaceOD'         = 'OneDrive (GB)'
            'm.avgOnPrem'       = 'Media mailbox ON-PREM (GB)'
            'm.maxOnPrem'       = 'Mailbox ON-PREM più grande (GB)'
            'm.sec.lic'         = 'LICENZE ASSEGNATE'
            'm.noLic'           = 'Nessuna licenza'
            'm.licDisabled'     = 'Licenziati ma disabilitati'
            'm.sec.mfa'         = 'MFA'
            'm.mfaYes'          = 'MFA configurato'
            'm.mfaNo'           = 'MFA NON configurato'
            'm.mfaNotInEntra'   = 'Non presenti in Entra'
            'm.mfaPct'          = '% MFA configurato (utenti Entra)'
            'm.licNoMfa'        = 'Abilitati + licenziati SENZA MFA'
            'm.perUserEnforced' = 'MFA per-user enforced'
            'm.sec.inactive'    = 'INATTIVITÀ'
            'm.inactiveAD'      = 'Abilitati senza login AD da >{0} giorni'
            'm.inactive365'     = 'Licenziati senza accesso 365 da >{0} giorni'
            'm.pwdOld'          = 'Password non cambiata da >{0} giorni'
            'col.metric'        = 'Metrica'
            'col.total'         = 'TOTALE'

            'tab.overview'      = 'PANORAMICA'
            'tab.users'         = 'UTENTI'
            'tab.mbxOnPrem'     = 'MAILBOX_ONPREM'
            'tab.mbx365'        = 'MAILBOX_365'
            'tab.onedrive'      = 'ONEDRIVE'
            'tab.mfa'           = 'MFA'
            'tab.licenses'      = 'LICENZE_TENANT'
            'tab.attention'     = 'DA_ATTENZIONARE'
            'tab.sync'          = 'SYNC'
            'tab.history'       = 'STORICO'
            'grp.other'         = 'ALTRI'
            'grp.cloud'         = 'SOLO_CLOUD'
            'grp.all'           = 'TUTTI'

            'xl.generated'      = 'Generato il {0} - {1}'
            'xl.legend'         = 'LOCATION: 365 = mailbox su Exchange Online | ON-PREM = mailbox su Exchange on-premises | NO MAILBOX = nessuna mailbox. Le dimensioni 365 arrivano dai report di utilizzo Microsoft (ritardo fino a 48h).'
            'xl.warning'        = 'ATTENZIONE: {0}'
        }
    }

    # Column headers (internal names are English; Italian translation applied on display/export)
    $script:HS_Cols = @{
        it = @{
            Group = 'Gruppo'; FirstName = 'Nome'; LastName = 'Cognome'; Username = 'AD_Username'; Status = 'Stato'
            AD_LastLogon = 'AD_UltimoLogon'; M365_LastSignIn = 'M365_UltimoAccesso'; Licenses = 'Licenze'
            PasswordLastSet = 'UltimoCambioPassword'; MailboxType = 'TipoMailbox'; MailboxSource = 'FonteMailbox'
            MFA_Status = 'MFA_Stato'; MFA_DefaultMethod = 'MFA_MetodoDefault'; MFA_Methods = 'MFA_Metodi'; Origin = 'Origine'; Notes = 'Note'
            DisplayName = 'Nome'; Type = 'Tipo'; Size_MB = 'Dimensione_MB'; Items = 'Elementi'; Deleted_MB = 'Eliminati_MB'
            LastLogon = 'UltimoAccesso'; Quota = 'Quota'; HiddenFromGAL = 'NascostaGAL'; WarningQuota_GB = 'QuotaAvviso_GB'
            SendQuota_GB = 'QuotaBloccoInvio_GB'; Archive = 'Archivio'; LastActivity = 'UltimaAttivita'; Created = 'Creata'
            Files = 'File'; ActiveFiles = 'FileAttivi'; Enabled = 'Abilitato'; SyncedFromAD = 'SincronizzatoAD'
            MFA_Registered = 'MFA_Registrato'; MFA_Capable = 'MFA_Utilizzabile'; DefaultMethod = 'MetodoDefault'; Methods = 'Metodi'
            PerUserState = 'StatoPerUser'; SSPR_Registered = 'SSPR_Registrato'; Name = 'Nome'; Purchased = 'Acquistate'
            Assigned = 'Assegnate'; Available = 'Libere'; Suspended = 'Sospese'; Warning = 'InScadenza'; Category = 'Categoria'
            Details = 'Dettaglio'; Date = 'Data'; Users = 'Utenti'; OnPrem = 'OnPrem'; On365 = 'Su365'
            OnPremSpace_GB = 'SpazioOnPrem_GB'; Space365_GB = 'Spazio365_GB'; MFA_Percent = 'MFA_Percentuale'; LicensedWithoutMFA = 'LicenziatiSenzaMFA'
            SyncIssues = 'ProblemiSync'; Anchor_AD = 'Ancora_AD'; AnchorSource = 'FonteAncora'; SyncEnabled = 'SyncAttivo'; LastSync = 'UltimoSync'
            Source = 'Fonte'; Membership = 'Appartenenza'
        }
    }

    function T {
        <#
        .SYNOPSIS
            Returns the localized string for Key (current language, falling back to English), formatted with Args2.
        #>
        param([string]$Key, [object[]]$Args2)
        $lang = if ($script:HS_Lang) { $script:HS_Lang } else { 'en' }
        $text = $script:HS_Str[$lang][$Key]
        if (-not $text) { $text = $script:HS_Str['en'][$Key] }
        if (-not $text) { $text = $Key }
        if ($null -ne $Args2 -and $Args2.Count) { $text -f $Args2 } else { $text }
    }

    function Get-HSHeader {
        param([string]$Name)
        $lang = if ($script:HS_Lang) { $script:HS_Lang } else { 'en' }
        if ($script:HS_Cols.ContainsKey($lang) -and $script:HS_Cols[$lang].ContainsKey($Name)) { $script:HS_Cols[$lang][$Name] } else { $Name }
    }

    # --------------------------------------------------------------------------- logging
    function Write-HSLog {
        <#
        .SYNOPSIS
            Writes a timestamped log line to the GUI log queue (or the console) and to the log file. WARN lines are also collected for the report.
        #>
        param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
        $line = '{0:HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message
        if ($Level -eq 'WARN' -and $null -ne $script:HS_Warnings) { $script:HS_Warnings.Add($Message) }
        if ($script:HS_Sync) { $script:HS_Sync.Log.Enqueue($line) } else { [Console]::WriteLine($line) }
        if ($script:HS_LogFile) { Add-Content -Path $script:HS_LogFile -Value $line -Encoding UTF8 }
    }

    # --------------------------------------------------------------------------- configuration
    function Get-HSDefaultConfig {
        [ordered]@{
            Language               = ''
            Mode                   = 'FullHybrid'
            TenantId               = ''
            AuthMode               = 'Interactive'
            ClientId               = ''
            CertificateThumbprint  = ''
            IncludePerUserMfaState = $true
            DomainFqdn             = ''
            SearchBase             = ''
            UseWindowsCredentials  = $true
            DomainUser             = ''
            ExchangeServer         = ''
            ExchangeProtocol       = 'http'
            ExchangeAuth           = 'Kerberos'
            ExchangeSkipCertCheck  = $false
            OutputFolder           = ''
            RetentionDays          = 90
            InactiveDays           = 90
            PasswordAgeDays        = 365
            SyncStaleHours         = 3
            LicenseFilter          = 'All'
            LicenseNames           = [ordered]@{
                O365_BUSINESS_ESSENTIALS = 'Microsoft 365 Business Basic'
                O365_BUSINESS_PREMIUM    = 'Microsoft 365 Business Standard'
                SPB                      = 'Microsoft 365 Business Premium'
                O365_BUSINESS            = 'Microsoft 365 Apps for business'
                OFFICESUBSCRIPTION       = 'Microsoft 365 Apps for enterprise'
                STANDARDPACK             = 'Office 365 E1'
                ENTERPRISEPACK           = 'Office 365 E3'
                ENTERPRISEPREMIUM        = 'Office 365 E5'
                SPE_E3                   = 'Microsoft 365 E3'
                SPE_E5                   = 'Microsoft 365 E5'
                SPE_F1                   = 'Microsoft 365 F3'
                DESKLESSPACK             = 'Office 365 F3'
                EXCHANGESTANDARD         = 'Exchange Online (Plan 1)'
                EXCHANGEENTERPRISE       = 'Exchange Online (Plan 2)'
                EXCHANGEDESKLESS         = 'Exchange Online Kiosk'
                TEAMS_ESSENTIALS_AAD     = 'Teams Essentials'
                Microsoft_Teams_Premium  = 'Teams Premium'
                Microsoft_365_Copilot    = 'Microsoft 365 Copilot'
                POWER_BI_PRO             = 'Power BI Pro'
                POWER_BI_STANDARD        = 'Power BI (free)'
                AAD_PREMIUM              = 'Entra ID P1'
                AAD_PREMIUM_P2           = 'Entra ID P2'
                EMS                      = 'Enterprise Mobility + Security E3'
                PROJECTPROFESSIONAL      = 'Project Plan 3'
                VISIOCLIENT              = 'Visio Plan 2'
                FLOW_FREE                = 'Power Automate (free)'
            }
            GroupBy                = 'TopLevelOU'
            GroupRules             = @()
            UnmappedGroup          = ''
            CloudOnlyGroup         = ''
            IncludeGuests          = $false
            ExportPreset           = 'Full'
            ExportColumns          = @()
            ExportGroupSheets      = $false
            ExportOverview         = $false
            ExportEnabledOnly      = $false
            ExportPresets          = [ordered]@{}
            AutoExportPreset       = 'Full'
            ExportFastLoad         = $true
        }
    }

    function Read-HSConfig {
        param([string]$Path)
        if (-not (Test-Path $Path)) {
            (Get-HSDefaultConfig) | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
        }
        $cfg = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        $defaults = Get-HSDefaultConfig
        foreach ($key in $defaults.Keys) {
            if (-not ($cfg.PSObject.Properties.Name -contains $key)) { $cfg | Add-Member -NotePropertyName $key -NotePropertyValue $defaults[$key] }
        }
        if (-not $cfg.Language) { $cfg.Language = if ((Get-Culture).TwoLetterISOLanguageName -eq 'it') { 'it' } else { 'en' } }
        if (-not $cfg.DomainFqdn -and $env:USERDNSDOMAIN) { $cfg.DomainFqdn = $env:USERDNSDOMAIN.ToLower() }
        $cfg.ExportPresets = ConvertTo-HSPresetObject $cfg.ExportPresets
        $cfg.ExportColumns = @($cfg.ExportColumns | Where-Object { $_ } | ForEach-Object { "$_" })
        $cfg
    }

    function ConvertTo-HSPresetObject {
        # ExportPresets as an object { "Name": ["Col", ...] }, whether it comes from JSON or from the defaults
        param($Value)
        $presets = [pscustomobject]@{}
        if ($Value -is [System.Collections.IDictionary]) { foreach ($key in $Value.Keys) { $presets | Add-Member -NotePropertyName "$key" -NotePropertyValue @($Value[$key] | ForEach-Object { "$_" }) -Force } }
        elseif ($Value) { foreach ($prop in $Value.PSObject.Properties) { $presets | Add-Member -NotePropertyName $prop.Name -NotePropertyValue @($prop.Value | ForEach-Object { "$_" }) -Force } }
        $presets
    }

    function Save-HSConfig {
        param($Cfg, [string]$Path)
        $Cfg | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
    }

    function Get-HSOutputFolder {
        param($Cfg)
        $folder = if ($Cfg.OutputFolder) { $Cfg.OutputFolder } else { Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'HybridScope' }
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $folder
    }

    function Get-HSSources {
        # Which sources each infrastructure mode uses
        param([string]$Mode)
        switch ($Mode) {
            'Cloud' { @{ M365 = $true; Exch = $false; AD = $false } }
            'Hybrid' { @{ M365 = $true; Exch = $false; AD = $true } }
            'OnPrem' { @{ M365 = $false; Exch = $true; AD = $true } }
            'AD' { @{ M365 = $false; Exch = $false; AD = $true } }
            default { @{ M365 = $true; Exch = $true; AD = $true } }
        }
    }

    function Test-HSModules {
        $need = 'Microsoft.Graph.Authentication', 'ImportExcel', 'ActiveDirectory'
        @($need | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
    }

    # --------------------------------------------------------------------------- helpers
    function ConvertTo-LocalDate {
        param($Value)
        if (-not $Value) { return $null }
        if ($Value -is [datetime]) { return $Value.ToLocalTime() }
        try { return ([datetime]::Parse("$Value", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)).ToLocalTime() }
        catch { return $null }
    }

    function ConvertTo-Num { param($Value) if ("$Value" -match '^\s*-?[\d\.]+\s*$') { [double]$Value } else { 0 } }

    function ConvertTo-MB {
        # Accepts a live ByteQuantifiedSize, the deserialized string "194.3 MB (203,734,016 bytes)",
        # localized variants ("byte", . or , separators) and the plain value "194,3 MB".
        param($Size)
        if ($null -eq $Size) { return 0 }
        if ($Size.PSObject.Methods['ToBytes']) { return [math]::Round($Size.ToBytes() / 1MB, 2) }
        if ($Size.Value -and $Size.Value.PSObject.Methods['ToBytes']) { return [math]::Round($Size.Value.ToBytes() / 1MB, 2) }
        $text = "$Size"
        if ($text -match '\(([\d\.,\s'']+)\s*byte') { return [math]::Round(([double]($Matches[1] -replace '[^\d]', '')) / 1MB, 2) }
        if ($text -match '([\d\.,]+)\s*(B|KB|MB|GB|TB)\b') {
            $number = $Matches[1]
            if ($number -match ',\d{1,2}$') { $number = $number -replace '\.', '' -replace ',', '.' } else { $number = $number -replace ',', '' }
            $mult = @{ B = 1 / 1MB; KB = 1 / 1KB; MB = 1; GB = 1KB; TB = 1MB }[$Matches[2]]
            return [math]::Round([double]$number * $mult, 2)
        }
        return 0
    }

    function Invoke-GraphPaged {
        param([string]$Uri)
        $out = [System.Collections.Generic.List[object]]::new()
        while ($Uri) {
            $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
            if ($page.value) { foreach ($item in $page.value) { $out.Add($item) } }
            $Uri = $page.'@odata.nextLink'
        }
        , $out
    }

    function Get-GraphReportCsv {
        param([string]$Uri)
        $tmp = [IO.Path]::GetTempFileName()
        try {
            Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $tmp
            Import-Csv -Path $tmp -Encoding UTF8
        }
        finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
    }

    $script:HS_MethodNames = @{
        microsoftAuthenticatorPush         = 'Microsoft Authenticator (push)'
        microsoftAuthenticatorPasswordless = 'Microsoft Authenticator (passwordless)'
        microsoftAuthenticator             = 'Microsoft Authenticator'
        softwareOneTimePasscode            = 'Authenticator app (OTP)'
        softwareOath                       = 'Authenticator app (OTP)'
        hardwareOneTimePasscode            = 'Hardware token (OTP)'
        hardwareOath                       = 'Hardware token (OTP)'
        mobilePhone                        = 'Mobile phone (SMS/call)'
        phone                              = 'Phone'
        sms                                = 'SMS'
        voiceMobile                        = 'Voice call (mobile)'
        voiceAlternateMobile               = 'Voice call (alternate mobile)'
        voiceOffice                        = 'Voice call (office)'
        alternateMobilePhone               = 'Alternate mobile phone'
        officePhone                        = 'Office phone'
        email                              = 'Email (SSPR only)'
        fido2                              = 'FIDO2 security key'
        passKeyDeviceBound                 = 'Passkey'
        passKeyDeviceBoundAuthenticator    = 'Passkey (Authenticator)'
        windowsHelloForBusiness            = 'Windows Hello for Business'
        temporaryAccessPass                = 'Temporary Access Pass'
        securityQuestion                   = 'Security questions'
        appPassword                        = 'App password'
    }
    function Get-MethodName { param($MethodId) if ($script:HS_MethodNames.ContainsKey("$MethodId")) { $script:HS_MethodNames["$MethodId"] } else { "$MethodId" } }

    # --------------------------------------------------------------------------- connections
    function Connect-HS365 {
        <#
        .SYNOPSIS
            Connects to Microsoft Graph: browser sign-in with delegated scopes, or app + certificate with -Unattended. Returns the signed-in account / app name.
        #>
        param($Cfg, [switch]$Unattended)
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $scopes = 'User.Read.All', 'AuditLog.Read.All', 'Reports.Read.All', 'UserAuthenticationMethod.Read.All', 'Organization.Read.All', 'Policy.Read.All'
        # Group names and Entra roles for the user card. Asked for separately so a tenant where an admin
        # has not consented to them yet still connects with the base permissions.
        $cardScopes = 'GroupMember.Read.All', 'RoleManagement.Read.Directory'
        $connectParams = @{ NoWelcome = $true }
        if ($Cfg.TenantId) { $connectParams.TenantId = $Cfg.TenantId }
        if ($Unattended) {
            $connectParams.ClientId = $Cfg.ClientId; $connectParams.CertificateThumbprint = $Cfg.CertificateThumbprint
        }
        else {
            # WAM needs a window handle: in the GUI the browser sign-in is used instead
            try { Set-MgGraphOption -DisableLoginByWAM $true } catch { }
            $connectParams.Scopes = $scopes + $cardScopes
            # do not persist the sign-in between sessions: no silent re-login with the wrong account
            $connectParams.ContextScope = 'Process'
        }
        try { Connect-MgGraph @connectParams }
        catch {
            # consent missing for the card permissions (non-admin user): connect with the base ones
            if ($Unattended -or "$($_.Exception.Message)" -notmatch 'AADSTS65001|AADSTS65004|AADSTS90094|AADSTS90008|consent|approval|access_denied') { throw }
            Write-HSLog (T 'warn.scopes' $_.Exception.Message) 'WARN'
            $connectParams.Scopes = $scopes
            Connect-MgGraph @connectParams
        }
        $ctx = Get-MgContext
        if (-not $ctx) { throw (T 'err.graphConnect') }
        $who = if ($ctx.Account) { $ctx.Account } else { "app $($ctx.AppName)" }
        Write-HSLog (T 'log.connected365' $who)
        $who
    }

    function Connect-HSExchange {
        <#
        .SYNOPSIS
            Connects to Exchange on-premises (Remote PowerShell over http/https, or the local snap-in) and imports only the read cmdlets HybridScope uses.
        #>
        # ExchangeProtocol:
        #   http  = Remote PowerShell (WinRM) on http://<server>/PowerShell/ with Kerberos - Microsoft's standard,
        #           traffic is encrypted by Kerberos, no certificate needed
        #   https = same over https (requires a valid certificate for that name)
        #   local = local Exchange Management Shell snap-in: only when HybridScope runs ON the Exchange server
        param($Cfg, [pscredential]$Credential)
        if ($script:HS_ExSession) { Remove-PSSession $script:HS_ExSession -ErrorAction SilentlyContinue; $script:HS_ExSession = $null }
        $script:HS_ExMode = $null
        if ($Cfg.ExchangeProtocol -eq 'local') {
            Write-HSLog (T 'log.exchLocal')
            if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
                Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
            }
            $script:HS_ExMode = 'local'
            $info = "local ($env:COMPUTERNAME)"
        }
        else {
            $uri = "$($Cfg.ExchangeProtocol)://$($Cfg.ExchangeServer)/PowerShell/"
            $sessionParams = @{ ConfigurationName = 'Microsoft.Exchange'; ConnectionUri = $uri; Authentication = $Cfg.ExchangeAuth; AllowRedirection = $true }
            if ($Credential) {
                # Kerberos from a non-domain PC needs the full realm: DOMAIN\user -> user@domain.fqdn
                if ($Cfg.ExchangeAuth -eq 'Kerberos' -and $Credential.UserName -match '^[^\\@]+\\(.+)$' -and $Cfg.DomainFqdn) {
                    $Credential = New-Object System.Management.Automation.PSCredential ("$($Matches[1])@$($Cfg.DomainFqdn)", $Credential.Password)
                }
                $sessionParams.Credential = $Credential
            }
            if ($Cfg.ExchangeSkipCertCheck) { $sessionParams.SessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck }
            Write-HSLog (T 'log.exchConnecting' @($uri, $Cfg.ExchangeAuth))
            $session = New-PSSession @sessionParams
            $mod = Import-PSSession $session -CommandName Get-Mailbox, Get-MailboxStatistics, Get-RemoteMailbox, Get-ExchangeServer -AllowClobber -DisableNameChecking
            Import-Module $mod -Global -DisableNameChecking
            $script:HS_ExSession = $session
            $script:HS_ExMode = 'remote'
            $info = "$($Cfg.ExchangeServer)"
        }
        try {
            $servers = @(Get-ExchangeServer -WarningAction SilentlyContinue)
            $info += " - $($servers.Count) server ($(@($servers | ForEach-Object { $_.Name }) -join ', '))"
        }
        catch { }
        Write-HSLog (T 'log.exchConnected' $info)
        $info
    }

    function Clear-HS365Session {
        # Signs out and removes the Microsoft Graph PowerShell token cache, so the next
        # "Connect 365" asks for an account again.
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        $paths = @(
            (Join-Path $env:USERPROFILE '.mg'),
            (Join-Path $env:LOCALAPPDATA '.IdentityService\mg.msal.cache'),
            (Join-Path $env:LOCALAPPDATA '.IdentityService\mg.msal.cache.nocae'),
            (Join-Path $env:LOCALAPPDATA '.IdentityService\mg.msal.cache.cae')
        )
        foreach ($cachePath in $paths) { if (Test-Path $cachePath) { Remove-Item $cachePath -Recurse -Force -ErrorAction SilentlyContinue } }
        Write-HSLog (T 'log.signedOut')
    }

    function Test-HSAD {
        <#
        .SYNOPSIS
            Checks that the AD domain in Cfg.DomainFqdn is reachable (optionally with Credential). Returns the DNS root.
        #>
        param($Cfg, [pscredential]$Credential)
        Import-Module ActiveDirectory -ErrorAction Stop
        $adParams = @{ Server = $Cfg.DomainFqdn }
        if ($Credential) { $adParams.Credential = $Credential }
        $domain = Get-ADDomain @adParams
        Write-HSLog (T 'log.adOk' @($domain.DNSRoot, $domain.PDCEmulator))
        "$($domain.DNSRoot)"
    }

    function Install-HSRsatAD {
        # Runs tools\Install-RsatAD.ps1 (plain text, shipped with HybridScope) in a VISIBLE elevated
        # PowerShell window: Windows shows the UAC prompt.
        $installer = Join-Path $script:HS_AppDir 'tools\Install-RsatAD.ps1'
        if (-not (Test-Path $installer)) { throw (T 'err.rsatFile' $installer) }
        Write-HSLog (T 'log.rsat')
        $proc = Start-Process -FilePath powershell.exe -Verb RunAs -Wait -PassThru -ArgumentList '-NoProfile', '-File', "`"$installer`""
        if ($proc.ExitCode -ne 0 -or -not (Get-Module -ListAvailable -Name ActiveDirectory)) { throw (T 'err.rsat') }
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-HSLog (T 'log.rsatOk')
    }

    # --------------------------------------------------------------------------- collectors
    # One function per data source. Get-HSData calls them all and merges the results;
    # -Only calls just the one it needs (see Get-HSDataset).

    function Format-HSYesNo {
        <#
        .SYNOPSIS
            Localized Yes / No for a true / false value.
        #>
        param($Value)
        if ($Value) { T 'val.yes' } else { T 'val.no' }
    }

    function Get-HSLicenseNameMap {
        <#
        .SYNOPSIS
            Friendly license names from Cfg.LicenseNames, as a hashtable SKU part number -> name.
        #>
        param($Cfg)
        $map = @{}
        if ($Cfg.LicenseNames) { foreach ($entry in $Cfg.LicenseNames.PSObject.Properties) { $map[$entry.Name] = "$($entry.Value)" } }
        $map
    }

    function Get-HSSubscribedSku {
        <#
        .SYNOPSIS
            The tenant's license SKUs (Microsoft Graph subscribedSkus).
        #>
        $skus = Invoke-GraphPaged 'v1.0/subscribedSkus'
        Write-HSLog (T 'log.skus' $skus.Count)
        , $skus
    }

    function Get-HSEntraUser {
        <#
        .SYNOPSIS
            Every Entra ID user with the properties HybridScope uses. Last sign-in needs Entra ID P1:
            without it the users are read again without sign-in data and a warning is logged.
            -NoSignIn skips the sign-in data (and the warning) when it is not needed.
        #>
        param([switch]$NoSignIn)
        $select = 'id,displayName,givenName,surname,userPrincipalName,mail,accountEnabled,assignedLicenses,provisionedPlans,onPremisesSamAccountName,onPremisesSyncEnabled,userType,createdDateTime' +
        ',proxyAddresses,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesDistinguishedName,onPremisesDomainName,onPremisesUserPrincipalName,onPremisesProvisioningErrors'
        if ($NoSignIn) { $users = Invoke-GraphPaged ('v1.0/users?$top=999&$select=' + $select) }
        else {
            try { $users = Invoke-GraphPaged ('v1.0/users?$top=120&$select=' + $select + ',signInActivity') }
            catch {
                Write-HSLog (T 'warn.signin' $_.Exception.Message) 'WARN'
                $users = Invoke-GraphPaged ('v1.0/users?$top=999&$select=' + $select)
            }
        }
        Write-HSLog (T 'log.entraUsers' $users.Count)
        , $users
    }

    function Get-HSMfaRegistration {
        <#
        .SYNOPSIS
            MFA registration of every Entra user. Uses the registration report (Entra ID P1); without it
            reads the methods user by user. With Cfg.IncludePerUserMfaState also reads the legacy
            per-user MFA state.
            Returns @{ Registration = @{ <user id> = details }; PerUser = @{ <user id> = state } }.
        #>
        param($Cfg, $EntraUsers)
        $registration = @{}
        try {
            foreach ($regDetail in (Invoke-GraphPaged 'v1.0/reports/authenticationMethods/userRegistrationDetails')) {
                $registration[$regDetail.id] = [pscustomobject]@{
                    isMfaRegistered  = [bool]$regDetail.isMfaRegistered
                    isMfaCapable     = [bool]$regDetail.isMfaCapable
                    isSsprRegistered = [bool]$regDetail.isSsprRegistered
                    isAdmin          = [bool]$regDetail.isAdmin
                    methods          = @($regDetail.methodsRegistered)
                    defaultMethod    = $regDetail.userPreferredMethodForSecondaryAuthentication
                    systemDefault    = if ($regDetail.isSystemPreferredAuthenticationMethodEnabled) { @($regDetail.systemPreferredAuthenticationMethods)[0] } else { $null }
                }
            }
            Write-HSLog (T 'log.mfaReg' $registration.Count)
        }
        catch {
            Write-HSLog (T 'warn.mfaReport' $_.Exception.Message) 'WARN'
            $members = @($EntraUsers | Where-Object { $_.userType -eq 'Member' }); $i = 0
            foreach ($member in $members) {
                $i++; if ($i % 50 -eq 0) { Write-HSLog (T 'log.mfaProgress' @($i, $members.Count)) }
                try {
                    $authMethods = Invoke-GraphPaged "v1.0/users/$($member.id)/authentication/methods"
                    $methodTypes = @($authMethods | ForEach-Object { ($_.'@odata.type' -replace '#microsoft.graph.', '') -replace 'AuthenticationMethod$', '' } | Where-Object { $_ -ne 'password' })
                    $isMfa = [bool](@($methodTypes | Where-Object { $_ -ne 'email' }).Count)
                    $registration[$member.id] = [pscustomobject]@{ isMfaRegistered = $isMfa; isMfaCapable = $isMfa; isSsprRegistered = $false; isAdmin = $false; methods = $methodTypes; defaultMethod = $null; systemDefault = $null }
                }
                catch { }
            }
        }

        # per-user MFA state (beta endpoint, one call per user)
        $perUser = @{}
        if ($Cfg.IncludePerUserMfaState) {
            Write-HSLog (T 'log.perUser')
            $members = @($EntraUsers | Where-Object { $_.userType -eq 'Member' }); $i = 0
            foreach ($member in $members) {
                $i++; if ($i % 100 -eq 0) { Write-HSLog (T 'log.perUserProgress' @($i, $members.Count)) }
                try { $perUser[$member.id] = (Invoke-MgGraphRequest -Method GET -Uri "beta/users/$($member.id)/authentication/requirements" -OutputType PSObject).perUserMfaState }
                catch { Write-HSLog (T 'warn.perUser' $_.Exception.Message) 'WARN'; break }
            }
        }
        @{ Registration = $registration; PerUser = $perUser }
    }

    function Get-HSDefaultMethodName {
        <#
        .SYNOPSIS
            Display name of a user's default MFA method (system-preferred first, then user-preferred).
        #>
        param($Registration)
        $method = if ($Registration.systemDefault -and $Registration.systemDefault -ne 'none') { $Registration.systemDefault } else { $Registration.defaultMethod }
        if ($method -and $method -ne 'none') { Get-MethodName $method } else { '' }
    }

    function Get-HSMethodList {
        <#
        .SYNOPSIS
            Display names of all MFA methods a user has registered, comma separated.
        #>
        param($Registration)
        (@($Registration.methods) | Where-Object { $_ } | ForEach-Object { Get-MethodName $_ } | Select-Object -Unique) -join ', '
    }

    function ConvertTo-HSMfaRow {
        <#
        .SYNOPSIS
            Rows of the MFA tab: one per Entra member (guests too with Cfg.IncludeGuests).
        #>
        param($Cfg, $EntraUsers, $Mfa)
        @(foreach ($entraUser in ($EntraUsers | Where-Object { $_.userType -eq 'Member' -or $Cfg.IncludeGuests } | Sort-Object userPrincipalName)) {
                $reg = $Mfa.Registration[$entraUser.id]
                [pscustomobject][ordered]@{
                    UPN             = $entraUser.userPrincipalName
                    DisplayName     = $entraUser.displayName
                    Enabled         = Format-HSYesNo $entraUser.accountEnabled
                    SyncedFromAD    = Format-HSYesNo $entraUser.onPremisesSyncEnabled
                    MFA_Registered  = Format-HSYesNo ($reg -and $reg.isMfaRegistered)
                    MFA_Capable     = Format-HSYesNo ($reg -and $reg.isMfaCapable)
                    DefaultMethod   = if ($reg) { Get-HSDefaultMethodName $reg } else { '' }
                    Methods         = if ($reg) { Get-HSMethodList $reg } else { '' }
                    PerUserState    = if ($Mfa.PerUser.ContainsKey($entraUser.id)) { "$($Mfa.PerUser[$entraUser.id])" } else { '' }
                    SSPR_Registered = Format-HSYesNo ($reg -and $reg.isSsprRegistered)
                    Admin           = Format-HSYesNo ($reg -and $reg.isAdmin)
                }
            })
    }

    function ConvertTo-HSLicenseRow {
        <#
        .SYNOPSIS
            Rows of the TENANT_LICENSES tab: purchased / assigned / available per SKU.
        #>
        param($Skus, [hashtable]$LicenseNames)
        @(foreach ($sku in ($Skus | Sort-Object skuPartNumber)) {
                $purchased = [int]$sku.prepaidUnits.enabled; $assigned = [int]$sku.consumedUnits
                [pscustomobject][ordered]@{
                    SKU       = $sku.skuPartNumber
                    Name      = if ($LicenseNames.ContainsKey($sku.skuPartNumber)) { $LicenseNames[$sku.skuPartNumber] } else { '' }
                    Purchased = $purchased
                    Assigned  = $assigned
                    Available = $purchased - $assigned
                    Suspended = [int]$sku.prepaidUnits.suspended
                    Warning   = [int]$sku.prepaidUnits.warning
                    Status    = $sku.capabilityStatus
                }
            })
    }

    function Get-HSMailbox365Usage {
        <#
        .SYNOPSIS
            Exchange Online mailbox usage report (last 7 days).
            Returns @{ SizeByUpn = @{ <upn> = MB }; Rows = <MAILBOX_365 tab>; Concealed = <report anonymized> }.
        #>
        $sizeByUpn = @{}; $rows = [System.Collections.Generic.List[object]]::new(); $concealed = $false
        foreach ($reportRow in (Get-GraphReportCsv "v1.0/reports/getMailboxUsageDetail(period='D7')")) {
            if ($reportRow.'Is Deleted' -eq 'True') { continue }
            $upn = $reportRow.'User Principal Name'; if ($upn -notlike '*@*') { $concealed = $true }
            $sizeMB = [math]::Round((ConvertTo-Num $reportRow.'Storage Used (Byte)') / 1MB, 2)
            $sizeByUpn[$upn] = $sizeMB
            $rows.Add([pscustomobject][ordered]@{
                    UPN             = $upn
                    DisplayName     = $reportRow.'Display Name'
                    Type            = $reportRow.'Recipient Type'
                    Size_MB         = $sizeMB
                    Items           = ConvertTo-Num $reportRow.'Item Count'
                    Deleted_MB      = [math]::Round((ConvertTo-Num $reportRow.'Deleted Item Size (Byte)') / 1MB, 2)
                    WarningQuota_GB = [math]::Round((ConvertTo-Num $reportRow.'Issue Warning Quota (Byte)') / 1GB, 2)
                    SendQuota_GB    = [math]::Round((ConvertTo-Num $reportRow.'Prohibit Send Quota (Byte)') / 1GB, 2)
                    Archive         = $reportRow.'Has Archive'
                    LastActivity    = ConvertTo-LocalDate $reportRow.'Last Activity Date'
                    Created         = ConvertTo-LocalDate $reportRow.'Created Date'
                })
        }
        @{ SizeByUpn = $sizeByUpn; Rows = $rows; Concealed = $concealed }
    }

    function Get-HSOneDriveUsage {
        <#
        .SYNOPSIS
            OneDrive usage report (last 7 days).
            Returns @{ SizeByUpn = @{ <upn> = MB }; Rows = <ONEDRIVE tab>; Concealed = <report anonymized> }.
        #>
        $sizeByUpn = @{}; $rows = [System.Collections.Generic.List[object]]::new(); $concealed = $false
        foreach ($reportRow in (Get-GraphReportCsv "v1.0/reports/getOneDriveUsageAccountDetail(period='D7')")) {
            if ($reportRow.'Is Deleted' -eq 'True') { continue }
            $upn = $reportRow.'Owner Principal Name'; if ($upn -notlike '*@*') { $concealed = $true }
            $sizeMB = [math]::Round((ConvertTo-Num $reportRow.'Storage Used (Byte)') / 1MB, 2)
            $sizeByUpn[$upn] = $sizeMB
            $rows.Add([pscustomobject][ordered]@{
                    UPN          = $upn
                    DisplayName  = $reportRow.'Owner Display Name'
                    Size_MB      = $sizeMB
                    Quota_GB     = [math]::Round((ConvertTo-Num $reportRow.'Storage Allocated (Byte)') / 1GB, 2)
                    Files        = ConvertTo-Num $reportRow.'File Count'
                    ActiveFiles  = ConvertTo-Num $reportRow.'Active File Count'
                    LastActivity = ConvertTo-LocalDate $reportRow.'Last Activity Date'
                    URL          = $reportRow.'Site URL'
                })
        }
        @{ SizeByUpn = $sizeByUpn; Rows = $rows; Concealed = $concealed }
    }

    function Get-HSExchangeMailbox {
        <#
        .SYNOPSIS
            On-premises mailboxes with their statistics, plus remote mailboxes (already moved to Exchange
            Online). Connects to Exchange first when there is no open session.
            Returns @{ BySam = @{ <sam> = size / count / type }; RemoteBySam = @{ <sam> = type }; Rows = <MAILBOX_ONPREM tab> }.
        #>
        param($Cfg, [pscredential]$Credential)
        $sessionOk = ($script:HS_ExMode -eq 'local' -or ($script:HS_ExSession -and $script:HS_ExSession.State -eq 'Opened')) -and (Get-Command Get-MailboxStatistics -ErrorAction SilentlyContinue)
        if (-not $sessionOk) { Connect-HSExchange -Cfg $Cfg -Credential $Credential | Out-Null }

        $bySam = @{}; $remoteBySam = @{}; $rows = [System.Collections.Generic.List[object]]::new()
        $mailboxes = @(Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue)
        Write-HSLog (T 'log.mbxOnPrem' $mailboxes.Count)
        try {
            foreach ($remoteMbx in @(Get-RemoteMailbox -ResultSize Unlimited -WarningAction SilentlyContinue)) { $remoteBySam["$($remoteMbx.SamAccountName)"] = "$($remoteMbx.RecipientTypeDetails)" }
            Write-HSLog (T 'log.remoteMbx' $remoteBySam.Count)
        }
        catch { Write-HSLog (T 'warn.remoteMbx' $_.Exception.Message) 'WARN' }
        $i = 0
        foreach ($mbx in $mailboxes) {
            $i++; if ($i % 50 -eq 0) { Write-HSLog (T 'log.statsProgress' @($i, $mailboxes.Count)) }
            $stats = $null
            try { $stats = Get-MailboxStatistics -Identity "$($mbx.DistinguishedName)" -WarningAction SilentlyContinue -ErrorAction Stop }
            catch { Write-HSLog (T 'warn.stats' @($mbx.SamAccountName, $_.Exception.Message)) 'WARN' }
            $sizeMB = if ($stats) { ConvertTo-MB $stats.TotalItemSize } else { 0 }
            $sam = "$($mbx.SamAccountName)"
            if ($bySam.ContainsKey($sam)) { $bySam[$sam].SizeMB += $sizeMB; $bySam[$sam].Count++ }
            else { $bySam[$sam] = [pscustomobject]@{ SizeMB = $sizeMB; Count = 1; Type = "$($mbx.RecipientTypeDetails)" } }
            $rows.Add([pscustomobject][ordered]@{
                    Username      = $sam
                    DisplayName   = "$($mbx.DisplayName)"
                    Type          = "$($mbx.RecipientTypeDetails)"
                    Email         = "$($mbx.PrimarySmtpAddress)"
                    Database      = "$($mbx.Database)"
                    Size_MB       = $sizeMB
                    Items         = if ($stats) { [double]$stats.ItemCount } else { $null }
                    Deleted_MB    = if ($stats) { ConvertTo-MB $stats.TotalDeletedItemSize } else { $null }
                    LastLogon     = if ($stats) { ConvertTo-LocalDate $stats.LastLogonTime } else { $null }
                    Quota         = if ($mbx.UseDatabaseQuotaDefaults) { T 'quota.dbDefault' } else { "$($mbx.ProhibitSendQuota)" }
                    HiddenFromGAL = Format-HSYesNo ("$($mbx.HiddenFromAddressListsEnabled)" -eq 'True')
                })
        }
        @{ BySam = $bySam; RemoteBySam = $remoteBySam; Rows = $rows }
    }

    function Get-HSADUser {
        <#
        .SYNOPSIS
            Active Directory users of Cfg.DomainFqdn (limited to Cfg.SearchBase when set), with the
            properties HybridScope uses, including the sync anchor (ms-DS-ConsistencyGuid).
            Returns @{ Users = <AD users>; NetBIOS = <domain NetBIOS name> }.
        #>
        param($Cfg, [pscredential]$Credential)
        Import-Module ActiveDirectory -ErrorAction Stop
        $common = @{ Server = $Cfg.DomainFqdn }
        if ($Credential) { $common.Credential = $Credential }
        $netbios = ''
        try { $netbios = (Get-ADDomain @common).NetBIOSName } catch { $netbios = ($Cfg.DomainFqdn -split '\.')[0].ToUpper() }
        $props = 'GivenName', 'Surname', 'CanonicalName', 'LastLogonDate', 'PasswordLastSet', 'PasswordNeverExpires', 'UserPrincipalName', 'mail', 'whenCreated', 'Description', 'proxyAddresses'
        $adParams = @{ Filter = '*'; Properties = @($props + 'mS-DS-ConsistencyGuid') } + $common
        if ($Cfg.SearchBase) { $adParams.SearchBase = $Cfg.SearchBase }
        try { $users = @(Get-ADUser @adParams) }
        catch {
            # very old schema without ms-DS-ConsistencyGuid: objectGUID is the anchor anyway
            Write-HSLog (T 'warn.adConsistency' $_.Exception.Message) 'WARN'
            $adParams.Properties = $props
            $users = @(Get-ADUser @adParams)
        }
        Write-HSLog (T 'log.adUsers' $users.Count)
        @{ Users = $users; NetBIOS = $netbios }
    }

    function Get-HSDirectorySync {
        <#
        .SYNOPSIS
            Tenant directory sync status and verified domains (Microsoft Graph organization). $null when not readable.
        #>
        try {
            $orgs = Invoke-GraphPaged 'v1.0/organization?$select=id,onPremisesSyncEnabled,onPremisesLastSyncDateTime,verifiedDomains'
            if ($orgs.Count) { $orgs[0] } else { $null }
        }
        catch { Write-HSLog (T 'warn.org' $_.Exception.Message) 'WARN'; $null }
    }

    function Get-HSADAnchor {
        <#
        .SYNOPSIS
            The sync anchor of an AD user as Entra stores it in onPremisesImmutableId (base64):
            ms-DS-ConsistencyGuid when set, otherwise objectGUID (Entra Connect / Cloud Sync default).
        #>
        param($ADUser)
        $consistency = $ADUser.'mS-DS-ConsistencyGuid'
        if ($consistency) {
            $bytes = [byte[]]@($consistency)
            if ($bytes.Length -eq 16) { return [pscustomobject]@{ Value = [Convert]::ToBase64String($bytes); Source = 'ms-DS-ConsistencyGuid' } }
        }
        if ($ADUser.ObjectGUID) { return [pscustomobject]@{ Value = [Convert]::ToBase64String(([guid]"$($ADUser.ObjectGUID)").ToByteArray()); Source = 'objectGUID' } }
        $null
    }

    function Format-HSAge {
        param([timespan]$Age)
        if ($Age.TotalMinutes -lt 120) { '{0:N0} min' -f $Age.TotalMinutes }
        elseif ($Age.TotalHours -lt 72) { '{0:N0} h' -f $Age.TotalHours }
        else { '{0:N0} d' -f $Age.TotalDays }
    }

    function Get-HSSyncIssue {
        <#
        .SYNOPSIS
            Compares AD and Entra ID the way Entra Connect / Cloud Sync link them (sourceAnchor ->
            onPremisesImmutableId, soft match on UPN / primary SMTP) and lists what is broken or about
            to break. Returns @{ Rows = <SYNC tab>; BySam / ById = <categories per AD sam / Entra id>;
            LinkedBySam = <Entra user linked to each AD sam by anchor>; Info = <anchor data per AD sam / Entra id> }.
        #>
        param($Cfg, $ADUsers, $EntraUsers, $Org)
        Write-HSLog (T 'log.syncStart')
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $rows = [System.Collections.Generic.List[object]]::new()
        $bySam = @{}; $byId = @{}; $linkedBySam = @{}; $info = @{}
        $now = Get-Date
        $ADUsers = @($ADUsers)
        $adBySam = @{}; foreach ($ad in $ADUsers) { $adBySam["$($ad.SamAccountName)"] = $ad }

        $add = {
            param([string]$CatKey, $Ad, $EntraUser, $Anchor, [string]$Details)
            $category = T $CatKey
            $rows.Add([pscustomobject][ordered]@{
                    Category     = $category
                    Username     = if ($Ad) { "$($Ad.SamAccountName)" } elseif ($EntraUser) { "$($EntraUser.onPremisesSamAccountName)" } else { '' }
                    DisplayName  = if ($EntraUser) { "$($EntraUser.displayName)" } elseif ($Ad) { ("$($Ad.GivenName) $($Ad.Surname)").Trim() } else { '' }
                    UPN          = if ($EntraUser) { "$($EntraUser.userPrincipalName)" } else { '' }
                    AD_UPN       = if ($Ad) { "$($Ad.UserPrincipalName)" } else { '' }
                    Anchor_AD    = if ($Anchor) { $Anchor.Value } else { '' }
                    AnchorSource = if ($Anchor) { $Anchor.Source } else { '' }
                    ImmutableId  = if ($EntraUser) { "$($EntraUser.onPremisesImmutableId)" } else { '' }
                    SyncEnabled  = if ($EntraUser) { Format-HSYesNo ($EntraUser.onPremisesSyncEnabled -eq $true) } else { '' }
                    LastSync     = if ($EntraUser) { ConvertTo-LocalDate $EntraUser.onPremisesLastSyncDateTime } else { $null }
                    Details      = $Details
                })
            $short = $category -replace '^\d+ - ', ''
            if ($Ad) { if (-not $bySam.ContainsKey("$($Ad.SamAccountName)")) { $bySam["$($Ad.SamAccountName)"] = [System.Collections.Generic.List[string]]::new() }; $bySam["$($Ad.SamAccountName)"].Add($short) }
            if ($EntraUser) { if (-not $byId.ContainsKey($EntraUser.id)) { $byId[$EntraUser.id] = [System.Collections.Generic.List[string]]::new() }; $byId[$EntraUser.id].Add($short) }
        }

        # ---- 0: tenant sync status
        if ($Org) {
            if ($Org.onPremisesSyncEnabled -ne $true) { & $add 'sync.c0' $null $null $null (T 'sync.d.off') }
            else {
                $last = ConvertTo-LocalDate $Org.onPremisesLastSyncDateTime
                if ($last) {
                    $age = $now - $last; $ageText = Format-HSAge $age; $lastText = $last.ToString('yyyy-MM-dd HH:mm')
                    $threshold = if ([double]$Cfg.SyncStaleHours -gt 0) { [double]$Cfg.SyncStaleHours } else { 3 }
                    if ($age.TotalHours -gt $threshold) {
                        & $add 'sync.c0' $null $null $null (T 'sync.d.stale' @($lastText, $ageText, $threshold))
                        Write-HSLog (T 'warn.syncStale' @($ageText, $threshold)) 'WARN'
                    }
                    else { & $add 'sync.c0' $null $null $null (T 'sync.d.last' @($lastText, $ageText)) }
                }
            }
        }

        # ---- indexes
        $verified = @{}
        if ($Org) { foreach ($domainInfo in @($Org.verifiedDomains)) { if ($domainInfo.name) { $verified["$($domainInfo.name)".ToLower()] = $true } } }
        $members = @($EntraUsers | Where-Object { $_.userType -ne 'Guest' })
        $entraByAnchor = @{}; $entraBySam = @{}; $entraByAddress = @{}
        $domainFqdn = "$($Cfg.DomainFqdn)".ToLower(); $searchBase = "$($Cfg.SearchBase)"
        $inScope = {
            # Entra user synced from the domain / OU that HybridScope reads (other domains and forests are not in $ADUsers)
            param($EntraUser)
            if ($EntraUser.onPremisesDomainName -and "$($EntraUser.onPremisesDomainName)".ToLower() -ne $domainFqdn) { return $false }
            if ($searchBase -and $EntraUser.onPremisesDistinguishedName -and -not "$($EntraUser.onPremisesDistinguishedName)".EndsWith($searchBase, [StringComparison]::OrdinalIgnoreCase)) { return $false }
            $true
        }
        foreach ($entraUser in $members) {
            if ($entraUser.onPremisesImmutableId) { $entraByAnchor["$($entraUser.onPremisesImmutableId)"] = $entraUser }
            if ($entraUser.onPremisesSyncEnabled -eq $true -and $entraUser.onPremisesSamAccountName -and (& $inScope $entraUser)) { $entraBySam["$($entraUser.onPremisesSamAccountName)"] = $entraUser }
            $addresses = @("$($entraUser.userPrincipalName)", "$($entraUser.mail)")
            foreach ($proxy in @($entraUser.proxyAddresses)) { if ("$proxy" -match '^smtp:') { $addresses += "$proxy".Substring(5) } }
            foreach ($address in $addresses) { if ($address -and -not $entraByAddress.ContainsKey($address.ToLower())) { $entraByAddress[$address.ToLower()] = $entraUser } }
        }

        $anchorBySam = @{}; $samsByAnchor = @{}
        foreach ($ad in $ADUsers) {
            $anchor = Get-HSADAnchor $ad
            if (-not $anchor) { continue }
            $anchorBySam["$($ad.SamAccountName)"] = $anchor
            if (-not $samsByAnchor.ContainsKey($anchor.Value)) { $samsByAnchor[$anchor.Value] = [System.Collections.Generic.List[string]]::new() }
            $samsByAnchor[$anchor.Value].Add("$($ad.SamAccountName)")
        }

        # custom sourceAnchor (employeeID, mail...) or another forest: anchor comparison would flag everybody
        $syncedInScope = @($members | Where-Object { $_.onPremisesSyncEnabled -eq $true -and $_.onPremisesImmutableId -and (& $inScope $_) })
        $anchorHits = @($syncedInScope | Where-Object { $samsByAnchor.ContainsKey("$($_.onPremisesImmutableId)") }).Count
        $customAnchor = ($syncedInScope.Count -ge 5 -and $anchorHits -lt [math]::Ceiling($syncedInScope.Count / 2))
        if ($customAnchor) { Write-HSLog (T 'warn.syncAnchor' @($anchorHits, $syncedInScope.Count)) 'WARN' }

        # ---- AD side
        $linkedIds = @{}; $reportedIds = @{}; $adById = @{}
        foreach ($ad in $ADUsers) {
            $sam = "$($ad.SamAccountName)"
            $anchor = $anchorBySam[$sam]
            $linked = $null
            if (-not $customAnchor -and $anchor -and $entraByAnchor.ContainsKey($anchor.Value)) { $linked = $entraByAnchor[$anchor.Value] }
            elseif ($customAnchor -and $entraBySam.ContainsKey($sam)) { $linked = $entraBySam[$sam] }
            $info[$sam] = [pscustomobject]@{ Anchor = $anchor; EntraUser = $linked }

            if ($linked) {
                $linkedBySam[$sam] = $linked; $linkedIds[$linked.id] = $true; $adById[$linked.id] = $ad
                if ($linked.onPremisesSyncEnabled -ne $true) {
                    & $add 'sync.c3' $ad $linked $anchor (T 'sync.d.leftoverAd' $sam); $reportedIds[$linked.id] = $true
                }
                elseif ("$($linked.userPrincipalName)" -like '*.onmicrosoft.com' -and "$($ad.UserPrincipalName)" -notlike '*.onmicrosoft.com') {
                    # The sync builds the Entra UPN from an AD attribute (userPrincipalName, or mail with
                    # Alternate Login ID): when that suffix is not a verified domain the UPN falls back to
                    # tenant.onmicrosoft.com. A different UPN with a verified domain is a normal setup.
                    $details = T 'sync.d.upn' @($linked.userPrincipalName, $ad.UserPrincipalName, $ad.mail)
                    $unverified = @(@("$($ad.UserPrincipalName)", "$($ad.mail)") | Where-Object { $_ -like '*@*' } | ForEach-Object { ($_ -split '@')[-1].ToLower() } | Where-Object { -not $verified.ContainsKey($_) } | Select-Object -Unique)
                    if ($verified.Count -and $unverified.Count) { $details += ' ' + (T 'sync.d.upnSuffix' ($unverified -join ', ')) }
                    & $add 'sync.c7' $ad $linked $anchor $details
                }
                continue
            }

            # 1: same account (sam) synced in Entra, but bound to another anchor
            if (-not $customAnchor -and $anchor -and $entraBySam.ContainsKey($sam)) {
                $other = $entraBySam[$sam]
                if (-not $samsByAnchor.ContainsKey("$($other.onPremisesImmutableId)")) {
                    & $add 'sync.c1' $ad $other $anchor (T 'sync.d.mismatch' @($anchor.Value, $anchor.Source, $other.onPremisesImmutableId))
                    $reportedIds[$other.id] = $true
                    continue
                }
            }

            # 4 / 5: a cloud-managed account with the same UPN or primary SMTP (what soft match looks at)
            $keys = [System.Collections.Generic.List[object]]::new()
            $keys.Add([pscustomobject]@{ Kind = 'UPN'; Value = "$($ad.UserPrincipalName)" })
            $keys.Add([pscustomobject]@{ Kind = 'SMTP'; Value = "$($ad.mail)" })
            foreach ($proxy in @($ad.proxyAddresses)) { if ("$proxy" -cmatch '^SMTP:') { $keys.Add([pscustomobject]@{ Kind = 'SMTP'; Value = "$proxy".Substring(5) }) } }
            foreach ($key in $keys) {
                if (-not $key.Value) { continue }
                $candidate = $entraByAddress[$key.Value.ToLower()]
                if (-not $candidate -or $candidate.onPremisesSyncEnabled -eq $true -or $reportedIds.ContainsKey($candidate.id)) { continue }
                if (-not $candidate.onPremisesImmutableId) { & $add 'sync.c4' $ad $candidate $anchor (T 'sync.d.soft' @($candidate.userPrincipalName, $key.Kind)) }
                else { & $add 'sync.c5' $ad $candidate $anchor (T 'sync.d.conflict' @($candidate.userPrincipalName, $key.Kind, $candidate.onPremisesImmutableId, $(if ($anchor) { $anchor.Value } else { '' }))) }
                $reportedIds[$candidate.id] = $true
                break
            }
        }

        # 8: two AD accounts with the same anchor (copied ms-DS-ConsistencyGuid, restored objects)
        foreach ($anchorValue in $samsByAnchor.Keys) {
            $sams = $samsByAnchor[$anchorValue]
            if ($sams.Count -lt 2) { continue }
            foreach ($sam in $sams) {
                $ad = $adBySam[$sam]
                & $add 'sync.c8' $ad $null $anchorBySam[$sam] (T 'sync.d.dup' @($anchorValue, (@($sams | Where-Object { $_ -ne $sam }) -join ', ')))
            }
        }

        # ---- Entra side
        foreach ($entraUser in $members) {
            if ($entraUser.onPremisesImmutableId -and -not $linkedIds.ContainsKey($entraUser.id) -and -not $reportedIds.ContainsKey($entraUser.id)) {
                if ($entraUser.onPremisesSyncEnabled -eq $true) {
                    # 2: synced from the domain we read, but no AD account carries its anchor
                    if (-not $customAnchor -and $ADUsers.Count -and (& $inScope $entraUser)) {
                        $where = if ($entraUser.onPremisesDistinguishedName) { "$($entraUser.onPremisesDistinguishedName)" } else { "$($entraUser.onPremisesSamAccountName)" }
                        & $add 'sync.c2' $null $entraUser $null (T 'sync.d.notInAd' @($where, $entraUser.onPremisesImmutableId))
                    }
                }
                else { & $add 'sync.c3' $null $entraUser $null (T 'sync.d.leftover' $entraUser.onPremisesImmutableId) }   # 3
            }
            # 6: errors Entra reports on the synced object (duplicate proxyAddresses / UPN...)
            foreach ($provError in @($entraUser.onPremisesProvisioningErrors)) {
                if (-not $provError) { continue }
                $occurred = ConvertTo-LocalDate $provError.occurredDateTime
                & $add 'sync.c6' $adById[$entraUser.id] $entraUser $(if ($adById[$entraUser.id]) { $anchorBySam["$($adById[$entraUser.id].SamAccountName)"] }) (T 'sync.d.prov' @($provError.propertyCausingError, $provError.value, $provError.category, $(if ($occurred) { $occurred.ToString('yyyy-MM-dd HH:mm') } else { '' })))
            }
        }

        $sorted = @($rows | Sort-Object Category, Username, UPN)
        Write-HSLog (T 'log.sync' @(@($sorted | Where-Object { $_.Category -ne (T 'sync.c0') }).Count, [math]::Round($timer.Elapsed.TotalSeconds, 1)))
        @{ Rows = $sorted; BySam = $bySam; ById = $byId; LinkedBySam = $linkedBySam; Info = $info; CustomAnchor = $customAnchor }
    }

    function Get-HSLatestRelease {
        <#
        .SYNOPSIS
            Latest HybridScope release on GitHub (public API, no sign-in). Only called from Help > Check for updates.
        #>
        param([string]$Repo)
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        # corporate proxies that need the Windows sign-in
        try { [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials } catch { }
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" -Headers @{ 'User-Agent' = 'HybridScope'; Accept = 'application/vnd.github+json' } -TimeoutSec 15
        [pscustomobject]@{
            Version   = "$($release.tag_name)" -replace '^[vV]', ''
            Url       = "$($release.html_url)"
            Published = ConvertTo-LocalDate $release.published_at
        }
    }

    function Compare-HSVersion {
        # -1 / 0 / 1 like [version].CompareTo; falls back to a string comparison for odd tags
        param([string]$Left, [string]$Right)
        $l = $null; $r = $null
        if ([version]::TryParse($Left, [ref]$l) -and [version]::TryParse($Right, [ref]$r)) { return $l.CompareTo($r) }
        [string]::Compare($Left, $Right, [StringComparison]::OrdinalIgnoreCase)
    }

    function ConvertTo-HSLdapValue {
        # escapes a value for an LDAP filter (RFC 4515)
        param([string]$Value)
        $Value -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29' -replace "`0", '\00'
    }

    function Get-HSGroupKind {
        <#
        .SYNOPSIS
            Readable kind of an Entra group: Microsoft 365, security, mail-enabled security, distribution list; dynamic or not.
        #>
        param($Group)
        $kind = if (@($Group.groupTypes) -contains 'Unified') { T 'card.type.m365' }
        elseif ($Group.securityEnabled -and $Group.mailEnabled) { T 'card.type.mailSec' }
        elseif ($Group.securityEnabled) { T 'card.type.security' }
        else { T 'card.type.dl' }
        if (@($Group.groupTypes) -contains 'DynamicMembership') { T 'card.type.dynamic' $kind } else { $kind }
    }

    function Get-HSUserGroup {
        <#
        .SYNOPSIS
            Groups and roles of one user, for the user card: AD groups (direct, nested via
            LDAP_MATCHING_RULE_IN_CHAIN, and the primary group) and Entra groups / roles / administrative
            units (memberOf + transitiveMemberOf). Either part is skipped when Sam / Upn is empty.
            Returns @{ Rows; Errors; AD; Entra; Roles } (counts).
        #>
        param($Cfg, [string]$Sam, [string]$Upn, [pscredential]$ADCredential)
        $rows = [System.Collections.Generic.List[object]]::new()
        $errors = [System.Collections.Generic.List[string]]::new()
        $countAD = 0; $countEntra = 0; $countRoles = 0

        if ($Sam -and $Cfg.DomainFqdn) {
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $common = @{ Server = $Cfg.DomainFqdn }
                if ($ADCredential) { $common.Credential = $ADCredential }
                $adUser = Get-ADUser -Identity $Sam -Properties memberOf, primaryGroupID @common
                $direct = @{}; foreach ($groupDn in @($adUser.memberOf)) { $direct["$groupDn"] = $true }
                $adGroups = @(Get-ADGroup -LDAPFilter "(member:1.2.840.113556.1.4.1941:=$(ConvertTo-HSLdapValue $adUser.DistinguishedName))" -Properties description @common)
                $primary = $null
                try { $primary = Get-ADGroup -Identity "$($adUser.SID.AccountDomainSid.Value)-$($adUser.primaryGroupID)" -Properties description @common } catch { }
                if ($primary) { $adGroups = @($primary) + @($adGroups | Where-Object { $_.DistinguishedName -ne $primary.DistinguishedName }) }
                foreach ($group in $adGroups) {
                    $category = if ("$($group.GroupCategory)" -eq 'Distribution') { T 'card.type.distribution' } else { T 'card.type.security' }
                    $membership = if ($primary -and $group.DistinguishedName -eq $primary.DistinguishedName) { T 'card.primary' } elseif ($direct.ContainsKey("$($group.DistinguishedName)")) { T 'card.direct' } else { T 'card.inherited' }
                    $rows.Add([pscustomobject][ordered]@{
                            Source = 'AD'; Name = "$($group.Name)"; Type = "$category ($($group.GroupScope))"; Membership = $membership
                            SyncedFromAD = ''; Details = "$($group.description)"
                        })
                    $countAD++
                }
            }
            catch { $errors.Add((T 'card.err.ad' $_.Exception.Message)) }
        }

        $hasGraph = $false
        try { $hasGraph = [bool](Get-MgContext) } catch { }
        if ($Upn -and $hasGraph) {
            try {
                $userKey = [uri]::EscapeDataString($Upn)
                $directIds = @{}
                foreach ($object in (Invoke-GraphPaged "v1.0/users/$userKey/memberOf")) { $directIds["$($object.id)"] = $true }
                $missing = [System.Collections.Generic.List[string]]::new()
                foreach ($object in (Invoke-GraphPaged "v1.0/users/$userKey/transitiveMemberOf")) {
                    $odataType = "$($object.'@odata.type')"
                    $name = if ($object.displayName) { "$($object.displayName)" } else { "$($object.id)" }
                    switch -Wildcard ($odataType) {
                        '*group' { $type = Get-HSGroupKind $object; $countEntra++; if (-not $object.displayName -and -not $missing.Contains('GroupMember.Read.All')) { $missing.Add('GroupMember.Read.All') } }
                        '*directoryRole' { $type = T 'card.type.role'; $countRoles++; if (-not $object.displayName -and -not $missing.Contains('RoleManagement.Read.Directory')) { $missing.Add('RoleManagement.Read.Directory') } }
                        '*administrativeUnit' { $type = T 'card.type.au' }
                        default { $type = $odataType -replace '^#microsoft\.graph\.', '' }
                    }
                    $details = if ($object.membershipRule) { "$($object.membershipRule)" } elseif ($object.mail) { "$($object.mail)" } else { "$($object.description)" }
                    $rows.Add([pscustomobject][ordered]@{
                            Source = 'Entra'; Name = $name; Type = $type
                            Membership = if ($directIds.ContainsKey("$($object.id)")) { T 'card.direct' } else { T 'card.inherited' }
                            SyncedFromAD = if ($odataType -like '*group') { Format-HSYesNo ($object.onPremisesSyncEnabled -eq $true) } else { '' }
                            Details = $details
                        })
                }
                if ($missing.Count) { $errors.Add((T 'card.err.perm' ($missing -join ', '))) }
            }
            catch {
                $message = "$($_.Exception.Message)"
                if ($message -match 'Forbidden|Authorization_RequestDenied|Insufficient privileges') { $errors.Add((T 'card.err.perm' 'GroupMember.Read.All')) }
                else { $errors.Add((T 'card.err.graph' $message)) }
            }
        }
        [pscustomobject]@{ Rows = @($rows | Sort-Object Source, Membership, Name); Errors = @($errors); AD = $countAD; Entra = $countEntra; Roles = $countRoles }
    }

    # --------------------------------------------------------------------------- data collection
    function Get-HSData {
        <#
        .SYNOPSIS
            Full report: runs the collectors for the sources of the selected Mode, merges them into one row per user and builds every tab. Returns Tables, Totals, Warnings and metadata. Microsoft 365 needs an existing Graph connection.
        #>
        param($Cfg, [pscredential]$ADCredential, [pscredential]$ExchangeCredential, [hashtable]$Sources, [hashtable]$Parts, [string]$PartialLabel)

        $ErrorActionPreference = 'Stop'
        $Src = if ($Sources) { $Sources } else { Get-HSSources $Cfg.Mode }
        # dynamic load (Get-HSLoadParts): collectors the export does not need are skipped
        if (-not $Parts) { $Parts = Get-HSLoadParts }
        $Skipped = @()
        if ($Src.M365) {
            if (-not $Parts.Mfa) { $Skipped += 'MFA' } elseif (-not $Parts.MfaPerUser -and $Cfg.IncludePerUserMfaState) { $Skipped += 'MFA per-user' }
            if (-not $Parts.Reports) { $Skipped += 'Mailbox 365 / OneDrive' }
        }
        if ($Src.Exch -and -not $Parts.Exch) { $Skipped += 'Exchange on-prem' }
        if ($Src.AD -and $Src.M365 -and -not $Parts.Sync) { $Skipped += 'Sync' }
        if ($Skipped.Count) { Write-HSLog (T 'log.partial' @($PartialLabel, ($Skipped -join ', '))) }
        $LocOk = ((-not $Src.Exch) -or $Parts.Exch) -and ((-not $Src.M365) -or $Parts.Reports)   # mailbox location is reliable
        $script:HS_Warnings = [System.Collections.Generic.List[string]]::new()
        $Now = Get-Date
        $InactiveCut = $Now.AddDays( - [int]$Cfg.InactiveDays)
        $PwdCut = $Now.AddDays( - [int]$Cfg.PasswordAgeDays)
        $Labels = @{
            Enabled = T 'val.enabled'; Disabled = T 'val.disabled'; None = T 'val.none'
            MfaYes = T 'val.mfaYes'; MfaNo = T 'val.mfaNo'; MfaNotInEntra = T 'val.mfaNotInEntra'
            Synced = T 'val.originSynced'; ADOnly = T 'val.originAD'; Cloud = T 'val.originCloud'; Guest = T 'val.originGuest'
            Unlinked = T 'val.originUnlinked'
        }
        $AdOrigins = @($Labels.Synced, $Labels.ADOnly, $Labels.Unlinked)
        $LicenseNames = Get-HSLicenseNameMap $Cfg
        $MappedOnly = ($Cfg.LicenseFilter -eq 'MappedOnly')

        if ($Src.M365 -and -not (Get-MgContext)) { throw (T 'err.not365') }

        $skus = @(); $SkuMap = @{}; $entra = @(); $EntraBySam = @{}; $EntraByUpn = @{}; $EntraByLocal = @{}
        $MfaReg = @{}; $PerUser = @{}; $Mbx365 = @{}; $OneDrive = @{}
        $TabMbx365 = @(); $TabOD = @()

        # ================================================= MICROSOFT GRAPH
        if ($Src.M365) {
            $skus = Get-HSSubscribedSku
            foreach ($sku in $skus) { $SkuMap["$($sku.skuId)"] = $sku.skuPartNumber }

            $entra = Get-HSEntraUser
            foreach ($entraUser in $entra) {
                if ($entraUser.onPremisesSamAccountName -and -not $EntraBySam.ContainsKey($entraUser.onPremisesSamAccountName)) { $EntraBySam[$entraUser.onPremisesSamAccountName] = $entraUser }
                $EntraByUpn[$entraUser.userPrincipalName] = $entraUser
                $local = ($entraUser.userPrincipalName -split '@')[0]
                if (-not $EntraByLocal.ContainsKey($local) -or $entraUser.userPrincipalName -notlike '*.onmicrosoft.com') { $EntraByLocal[$local] = $entraUser }
            }

            if ($Parts.Mfa) {
                $mfaCfg = $Cfg
                if (-not $Parts.MfaPerUser) { $mfaCfg = $Cfg.PSObject.Copy(); $mfaCfg.IncludePerUserMfaState = $false }
                $mfaData = Get-HSMfaRegistration -Cfg $mfaCfg -EntraUsers $entra
                $MfaReg = $mfaData.Registration; $PerUser = $mfaData.PerUser
            }

            if ($Parts.Reports) {
                $mailboxUsage = Get-HSMailbox365Usage
                $oneDriveUsage = Get-HSOneDriveUsage
                $Mbx365 = $mailboxUsage.SizeByUpn; $TabMbx365 = $mailboxUsage.Rows
                $OneDrive = $oneDriveUsage.SizeByUpn; $TabOD = $oneDriveUsage.Rows
                Write-HSLog (T 'log.reports' @($Mbx365.Count, $OneDrive.Count))
                if ($mailboxUsage.Concealed -or $oneDriveUsage.Concealed) { Write-HSLog (T 'warn.concealed') 'WARN' }
            }
        }

        # ================================================= EXCHANGE ON-PREMISES
        $ExchOnPrem = @{}; $ExchRemote = @{}; $ExchAvailable = $false; $TabOnPrem = @()
        if ($Src.Exch -and $Parts.Exch -and ($Cfg.ExchangeServer -or $Cfg.ExchangeProtocol -eq 'local')) {
            try {
                $exCred = if ($ExchangeCredential) { $ExchangeCredential } else { $ADCredential }
                $exchange = Get-HSExchangeMailbox -Cfg $Cfg -Credential $exCred
                $ExchOnPrem = $exchange.BySam; $ExchRemote = $exchange.RemoteBySam; $TabOnPrem = $exchange.Rows
                $ExchAvailable = $true
            }
            catch { Write-HSLog (T 'warn.exch' $_.Exception.Message) 'WARN' }
        }
        elseif ($Src.Exch -and $Parts.Exch) { Write-HSLog (T 'log.exchSkipped') }

        # ================================================= ACTIVE DIRECTORY
        $adUsers = @(); $NetBIOS = ''; $AdRead = $false
        if ($Src.AD -and $Cfg.DomainFqdn) {
            try {
                $adData = Get-HSADUser -Cfg $Cfg -Credential $ADCredential
                $adUsers = @($adData.Users); $NetBIOS = $adData.NetBIOS; $AdRead = $true
            }
            catch { Write-HSLog (T 'warn.ad' $_.Exception.Message) 'WARN' }
        }
        elseif ($Src.AD) { Write-HSLog (T 'log.adSkipped') }

        # ================================================= SYNC (AD <-> Entra ID)
        $Sync = $null
        if ($Src.M365 -and $AdRead -and $Parts.Sync) {
            try { $Sync = Get-HSSyncIssue -Cfg $Cfg -ADUsers $adUsers -EntraUsers $entra -Org (Get-HSDirectorySync) }
            catch { Write-HSLog (T 'warn.sync' $_.Exception.Message) 'WARN' }
        }
        $SyncInfo = @{}

        # ================================================= USER ROWS
        $GrpOther = if ($Cfg.UnmappedGroup) { $Cfg.UnmappedGroup } else { T 'grp.other' }
        $GrpCloud = if ($Cfg.CloudOnlyGroup) { $Cfg.CloudOnlyGroup } else { T 'grp.cloud' }

        $newRow = {
            param($Ad, $EntraUser, [string]$Group)
            $sam = if ($Ad) { $Ad.SamAccountName } elseif ($EntraUser.onPremisesSamAccountName) { $EntraUser.onPremisesSamAccountName } else { ($EntraUser.userPrincipalName -split '@')[0] }
            $upn = if ($EntraUser) { $EntraUser.userPrincipalName } else { $null }
            $notes = [System.Collections.Generic.List[string]]::new()

            # an AD account paired by name with a cloud account that Entra does not sync from AD is not "synced"
            $origin = if ($Ad -and $EntraUser) { if ($EntraUser.onPremisesSyncEnabled -eq $true) { $Labels.Synced } else { $Labels.Unlinked } } elseif ($Ad) { $Labels.ADOnly } elseif ($EntraUser.userType -eq 'Guest') { $Labels.Guest } else { $Labels.Cloud }

            # sync problems (SYNC tab) and anchor data for the user card
            $syncIssues = ''
            if ($Sync) {
                $issueList = @()
                if ($Ad -and $Sync.BySam.ContainsKey("$($Ad.SamAccountName)")) { $issueList += @($Sync.BySam["$($Ad.SamAccountName)"]) }
                if ($EntraUser -and $Sync.ById.ContainsKey($EntraUser.id)) { $issueList += @($Sync.ById[$EntraUser.id]) }
                $syncIssues = (@($issueList | Select-Object -Unique) -join '; ')
            }
            if ($Src.M365 -and $Src.AD) {
                $anchor = if ($Ad -and $Sync -and $Sync.Info.ContainsKey("$($Ad.SamAccountName)")) { $Sync.Info["$($Ad.SamAccountName)"].Anchor } else { $null }
                $SyncInfo["$sam|$upn"] = [pscustomobject][ordered]@{
                    Anchor_AD    = if ($anchor) { $anchor.Value } else { '' }
                    AnchorSource = if ($anchor) { $anchor.Source } else { '' }
                    ImmutableId  = if ($EntraUser) { "$($EntraUser.onPremisesImmutableId)" } else { '' }
                    SyncEnabled  = if ($EntraUser) { Format-HSYesNo ($EntraUser.onPremisesSyncEnabled -eq $true) } else { '' }
                    LastSync     = if ($EntraUser) { ConvertTo-LocalDate $EntraUser.onPremisesLastSyncDateTime } else { $null }
                    AD_UPN       = if ($Ad) { "$($Ad.UserPrincipalName)" } else { '' }
                }
            }
            $enabled = if ($Ad) { [bool]$Ad.Enabled } else { [bool]$EntraUser.accountEnabled }

            $ou = ''
            if ($Ad -and $Ad.CanonicalName) {
                $parts = @($Ad.CanonicalName -split '/')
                $ou = if ($parts.Count -gt 2) { (@($NetBIOS) + $parts[1..($parts.Count - 2)]) -join '\' } else { $NetBIOS }
            }

            # licenses
            $lic = $Labels.None
            if ($EntraUser -and $EntraUser.assignedLicenses) {
                $names = @(@(foreach ($license in $EntraUser.assignedLicenses) {
                            $skuPart = $SkuMap["$($license.skuId)"]
                            if ($skuPart -and $LicenseNames.ContainsKey($skuPart)) { $LicenseNames[$skuPart] }
                            elseif ($skuPart -and -not $MappedOnly) { $skuPart }
                        }) | Select-Object -Unique)
                if ($names.Count) { $lic = $names -join ', ' }
            }

            # where the mailbox lives (Exchange on-prem and Graph decide; size is informational only)
            $inReport365 = [bool]($upn -and $Mbx365.ContainsKey($upn))
            $exch = if ($Ad -and $ExchOnPrem.ContainsKey($sam)) { $ExchOnPrem[$sam] } else { $null }
            $remoteType = if ($Ad -and $ExchRemote.ContainsKey($sam)) { $ExchRemote[$sam] } else { $null }
            $exoPlan = $false
            if ($EntraUser -and $EntraUser.provisionedPlans) {
                $exoPlan = [bool]@($EntraUser.provisionedPlans | Where-Object { $_.service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled' -and $_.provisioningStatus -eq 'Success' }).Count
            }
            if ($exch) {
                $location = 'ON-PREM'; $mbxType = "$($exch.Type)"; $mbxSource = T 'src.getMailbox'
                if ($exch.Count -gt 1) { $notes.Add((T 'note.multiOnPrem' $exch.Count)) }
            }
            elseif ($remoteType) { $location = '365'; $mbxType = $remoteType; $mbxSource = T 'src.remote' }
            elseif ($exoPlan) { $location = '365'; $mbxType = 'Exchange Online'; $mbxSource = T 'src.graphPlan' }
            elseif ($inReport365) { $location = '365'; $mbxType = 'Exchange Online'; $mbxSource = T 'src.graphReport' }
            else { $location = 'NO MAILBOX'; $mbxType = ''; $mbxSource = if ($Src.Exch -and $Ad -and -not $ExchAvailable) { T 'src.exchNotRead' } else { '' } }

            $exchMB = if ($exch) { $exch.SizeMB } else { 0 }
            $mail365 = if ($inReport365) { $Mbx365[$upn] } elseif ($location -eq '365') { $null } else { 0 }
            if ($location -eq '365' -and -not $inReport365 -and $Parts.Reports) { $notes.Add((T 'note.no365Size')) }
            $od = if ($upn -and $OneDrive.ContainsKey($upn)) { $OneDrive[$upn] } else { 0 }

            # last 365 sign-in (max of interactive / non-interactive)
            $azLast = $null
            if ($EntraUser -and $EntraUser.signInActivity) {
                $azLast = @((ConvertTo-LocalDate $EntraUser.signInActivity.lastSignInDateTime), (ConvertTo-LocalDate $EntraUser.signInActivity.lastNonInteractiveSignInDateTime)) |
                    Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
            }

            # MFA
            $mfa = $Labels.MfaNotInEntra; $mfaDef = ''; $mfaMethods = ''; $mfaPerUser = ''
            if ($EntraUser -and $Parts.Mfa) {
                $reg = $MfaReg[$EntraUser.id]
                if ($reg) {
                    $mfa = if ($reg.isMfaRegistered) { $Labels.MfaYes } else { $Labels.MfaNo }
                    $mfaDef = Get-HSDefaultMethodName $reg
                    $mfaMethods = Get-HSMethodList $reg
                }
                else { $mfa = $Labels.MfaNo }
                if ($PerUser.ContainsKey($EntraUser.id)) { $mfaPerUser = "$($PerUser[$EntraUser.id])" }
            }

            if (-not $enabled -and $lic -ne $Labels.None) { $notes.Add((T 'note.licensedDisabled')) }
            if ($enabled -and $lic -ne $Labels.None -and $mfa -eq $Labels.MfaNo) { $notes.Add((T 'note.noMfa')) }
            if ($Src.M365 -and $Ad -and -not $EntraUser -and $enabled) { $notes.Add((T 'note.notSynced')) }
            if ($Ad -and $Ad.PasswordNeverExpires) { $notes.Add((T 'note.pwdNeverExpires')) }

            [pscustomobject][ordered]@{
                Group             = $Group
                FirstName         = if ($Ad) { $Ad.GivenName } else { $EntraUser.givenName }
                LastName          = if ($Ad) { $Ad.Surname } else { $EntraUser.surname }
                Username          = $sam
                OU                = $ou
                Status            = if ($enabled) { $Labels.Enabled } else { $Labels.Disabled }
                AD_LastLogon      = if ($Ad) { $Ad.LastLogonDate } else { $null }
                M365_LastSignIn   = $azLast
                Licenses          = $lic
                Mailbox365_MB     = $mail365
                MailboxOnPrem_MB  = $exchMB
                OneDrive_MB       = $od
                PasswordLastSet   = if ($Ad) { $Ad.PasswordLastSet } else { $null }
                Location          = $location
                MailboxType       = $mbxType
                MailboxSource     = $mbxSource
                MFA_Status        = $mfa
                MFA_DefaultMethod = $mfaDef
                MFA_Methods       = $mfaMethods
                MFA_PerUser       = $mfaPerUser
                UPN               = $upn
                Origin            = $origin
                SyncIssues        = $syncIssues
                Notes             = $notes -join '; '
            }
        }

        $getGroup = {
            param($ad)
            if ($Cfg.GroupBy -eq 'Rules' -and @($Cfg.GroupRules).Count) {
                foreach ($rule in $Cfg.GroupRules) { if ($ad.DistinguishedName -like "*$($rule.Match)*") { return $rule.Group } }
                return $GrpOther
            }
            $parts = @("$($ad.CanonicalName)" -split '/')
            switch ($Cfg.GroupBy) {
                'ParentOU' { if ($parts.Count -gt 2) { return $parts[$parts.Count - 2].ToUpper() } else { return $GrpOther } }
                'None' { return (T 'grp.all') }
                default { if ($parts.Count -gt 2) { return $parts[1].ToUpper() } else { return $GrpOther } }
            }
        }

        $Groups = [ordered]@{}
        foreach ($rule in @($Cfg.GroupRules)) { if ($rule.Group -and -not $Groups.Contains($rule.Group)) { $Groups[$rule.Group] = [System.Collections.Generic.List[object]]::new() } }
        $seenEntra = @{}
        $LinkedIds = @{}
        if ($Sync) { foreach ($linkedUser in $Sync.LinkedBySam.Values) { $LinkedIds[$linkedUser.id] = $true } }
        foreach ($ad in $adUsers) {
            $entraUser = $null
            # the sync anchor is the real link; names are the fallback (AD-only accounts, custom anchors)
            if ($Sync -and $Sync.LinkedBySam.ContainsKey("$($ad.SamAccountName)")) { $entraUser = $Sync.LinkedBySam["$($ad.SamAccountName)"] }
            elseif ($EntraBySam.ContainsKey($ad.SamAccountName)) { $entraUser = $EntraBySam[$ad.SamAccountName] }
            elseif ($ad.UserPrincipalName -and $EntraByUpn.ContainsKey($ad.UserPrincipalName)) { $entraUser = $EntraByUpn[$ad.UserPrincipalName] }
            elseif ($EntraByLocal.ContainsKey($ad.SamAccountName)) { $entraUser = $EntraByLocal[$ad.SamAccountName] }
            # never pair by name with a cloud account that the anchor already links to another AD account
            if ($entraUser -and $Sync -and $LinkedIds.ContainsKey($entraUser.id) -and -not $Sync.LinkedBySam.ContainsKey("$($ad.SamAccountName)")) { $entraUser = $null }
            if ($entraUser) { $seenEntra[$entraUser.id] = $true }
            $groupName = & $getGroup $ad
            if (-not $Groups.Contains($groupName)) { $Groups[$groupName] = [System.Collections.Generic.List[object]]::new() }
            $Groups[$groupName].Add((& $newRow $ad $entraUser $groupName))
        }
        foreach ($entraUser in $entra) {
            if ($seenEntra.ContainsKey($entraUser.id)) { continue }
            if ($entraUser.userType -eq 'Guest' -and -not $Cfg.IncludeGuests) { continue }
            if (-not $Groups.Contains($GrpCloud)) { $Groups[$GrpCloud] = [System.Collections.Generic.List[object]]::new() }
            $Groups[$GrpCloud].Add((& $newRow $null $entraUser $GrpCloud))
        }
        foreach ($groupName in @($Groups.Keys)) {
            if ($Groups[$groupName].Count -eq 0) { $Groups.Remove($groupName) }
            else { $Groups[$groupName] = @($Groups[$groupName] | Sort-Object LastName, FirstName, Username) }
        }
        $All = @($Groups.Values | ForEach-Object { $_ })
        Write-HSLog (T 'log.rows' @($All.Count, $Groups.Count))

        # ================================================= OVERVIEW
        $sumGB = { param($rows, $prop) $total = 0.0; foreach ($row in $rows) { if ($row.$prop) { $total += [double]$row.$prop } }; [math]::Round($total / 1024, 2) }
        $isLic = { param($row) $row.Licenses -ne $Labels.None }
        $Metrics = [System.Collections.Generic.List[object]]::new()
        $addMetric = { param($id, $name, $formula, [switch]$Section, [switch]$GB) $Metrics.Add([pscustomobject]@{ Id = $id; Name = $name; Formula = $formula; IsSection = [bool]$Section; IsGB = [bool]$GB }) }
        & $addMetric 'secUsers' (T 'm.sec.users') $null -Section
        & $addMetric 'total' (T 'm.total') { param($rows) @($rows).Count }
        & $addMetric 'enabled' (T 'm.enabled') { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled }).Count }
        & $addMetric 'disabled' (T 'm.disabled') { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Disabled }).Count }
        if ($Src.AD -and $Src.M365) {
            & $addMetric 'adOnly' (T 'm.adOnly') { param($rows) @($rows | Where-Object { $_.Origin -eq $Labels.ADOnly }).Count }
            & $addMetric 'unlinked' (T 'm.unlinked') { param($rows) @($rows | Where-Object { $_.Origin -eq $Labels.Unlinked }).Count }
            if ($Sync) { & $addMetric 'syncIssues' (T 'm.syncIssues') { param($rows) @($rows | Where-Object { $_.SyncIssues }).Count } }
            & $addMetric 'cloudOnly' (T 'm.cloudOnly') { param($rows) @($rows | Where-Object { $_.Origin -eq $Labels.Cloud -or $_.Origin -eq $Labels.Guest }).Count }
        }
        if (($Src.M365 -or $Src.Exch) -and $LocOk) {
            & $addMetric 'secLoc' (T 'm.sec.location') $null -Section
            if ($Src.M365) { & $addMetric 'loc365' (T 'm.loc' '365') { param($rows) @($rows | Where-Object { $_.Location -eq '365' }).Count } }
            if ($Src.Exch) { & $addMetric 'locOnPrem' (T 'm.loc' 'ON-PREM') { param($rows) @($rows | Where-Object { $_.Location -eq 'ON-PREM' }).Count } }
            & $addMetric 'locNone' (T 'm.loc' 'NO MAILBOX') { param($rows) @($rows | Where-Object { $_.Location -eq 'NO MAILBOX' }).Count }
            & $addMetric 'secSpace' (T 'm.sec.space') $null -Section
            if ($Src.Exch) { & $addMetric 'spaceOnPrem' (T 'm.spaceOnPrem') { param($rows) & $sumGB $rows 'MailboxOnPrem_MB' } -GB }
            if ($Src.M365) {
                & $addMetric 'space365' (T 'm.space365') { param($rows) & $sumGB $rows 'Mailbox365_MB' } -GB
                & $addMetric 'spaceOD' (T 'm.spaceOD') { param($rows) & $sumGB $rows 'OneDrive_MB' } -GB
            }
            if ($Src.Exch) {
                & $addMetric 'avgOnPrem' (T 'm.avgOnPrem') { param($rows) $onPremRows = @($rows | Where-Object { $_.Location -eq 'ON-PREM' }); if ($onPremRows.Count) { [math]::Round((& $sumGB $onPremRows 'MailboxOnPrem_MB') / $onPremRows.Count, 2) } else { 0 } } -GB
                & $addMetric 'maxOnPrem' (T 'm.maxOnPrem') { param($rows) $maxMB = 0.0; foreach ($row in $rows) { if ($row.MailboxOnPrem_MB -gt $maxMB) { $maxMB = [double]$row.MailboxOnPrem_MB } }; [math]::Round($maxMB / 1024, 2) } -GB
            }
        }
        if ($Src.M365) {
        & $addMetric 'secLic' (T 'm.sec.lic') $null -Section
        $licNames = @($All | ForEach-Object { $_.Licenses -split ', ' } | Where-Object { $_ -and $_ -ne $Labels.None } | Sort-Object -Unique)
        foreach ($licName in $licNames) {
            & $addMetric "lic:$licName" $licName ({ param($rows) @($rows | Where-Object { ($_.Licenses -split ', ') -contains $licName }).Count }.GetNewClosure())
        }
        & $addMetric 'noLic' (T 'm.noLic') { param($rows) @($rows | Where-Object { $_.Licenses -eq $Labels.None }).Count }
        & $addMetric 'licDisabled' (T 'm.licDisabled') { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Disabled -and (& $isLic $_) }).Count }
        if ($Parts.Mfa) {
        & $addMetric 'secMfa' (T 'm.sec.mfa') $null -Section
        & $addMetric 'mfaYes' (T 'm.mfaYes') { param($rows) @($rows | Where-Object { $_.MFA_Status -eq $Labels.MfaYes }).Count }
        & $addMetric 'mfaNo' (T 'm.mfaNo') { param($rows) @($rows | Where-Object { $_.MFA_Status -eq $Labels.MfaNo }).Count }
        & $addMetric 'mfaNotInEntra' (T 'm.mfaNotInEntra') { param($rows) @($rows | Where-Object { $_.MFA_Status -eq $Labels.MfaNotInEntra }).Count }
        & $addMetric 'mfaPct' (T 'm.mfaPct') { param($rows) $inEntra = @($rows | Where-Object { $_.MFA_Status -ne $Labels.MfaNotInEntra }); if ($inEntra.Count) { [math]::Round(100 * @($inEntra | Where-Object { $_.MFA_Status -eq $Labels.MfaYes }).Count / $inEntra.Count, 1) } else { 0 } }
        & $addMetric 'licNoMfa' (T 'm.licNoMfa') { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled -and (& $isLic $_) -and $_.MFA_Status -eq $Labels.MfaNo }).Count }
        if ($Parts.MfaPerUser) { & $addMetric 'perUserEnforced' (T 'm.perUserEnforced') { param($rows) @($rows | Where-Object { $_.MFA_PerUser -eq 'enforced' }).Count } }
        }   # end MFA metrics
        }   # end Microsoft 365 metrics
        & $addMetric 'secInactive' (T 'm.sec.inactive') $null -Section
        if ($Src.AD) { & $addMetric 'inactiveAD' (T 'm.inactiveAD' $Cfg.InactiveDays) { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled -and ($AdOrigins -contains $_.Origin) -and (-not $_.AD_LastLogon -or $_.AD_LastLogon -lt $InactiveCut) }).Count } }
        if ($Src.M365) { & $addMetric 'inactive365' (T 'm.inactive365' $Cfg.InactiveDays) { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled -and (& $isLic $_) -and (-not $_.M365_LastSignIn -or $_.M365_LastSignIn -lt $InactiveCut) }).Count } }
        if ($Src.AD) { & $addMetric 'pwdOld' (T 'm.pwdOld' $Cfg.PasswordAgeDays) { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled -and $_.PasswordLastSet -and $_.PasswordLastSet -lt $PwdCut }).Count } }

        $colMetric = T 'col.metric'; $colTotal = T 'col.total'
        $Totals = @{}
        $Overview = @(foreach ($metric in $Metrics) {
                $line = [ordered]@{ $colMetric = $metric.Name }
                foreach ($groupName in $Groups.Keys) { $line[$groupName] = if ($metric.IsSection) { $null } else { & $metric.Formula $Groups[$groupName] } }
                $line[$colTotal] = if ($metric.IsSection) { $null } else { & $metric.Formula $All }
                if (-not $metric.IsSection) { $Totals[$metric.Id] = $line[$colTotal] }
                [pscustomobject]$line
            })

        # ================================================= ATTENTION
        $Attn = [System.Collections.Generic.List[object]]::new()
        $addAttention = {
            param($category, $row, $details, [double]$sortKey = 0)
            $Attn.Add([pscustomobject][ordered]@{
                    Category = $category; Group = $row.Group; Username = $row.Username
                    Name = ("$($row.FirstName) $($row.LastName)").Trim(); Licenses = $row.Licenses; Details = $details; _k = $sortKey
                })
        }
        $lastOrNever = { param($date) if ($date) { T 'att.last' $date } else { T 'att.never' } }
        foreach ($userRow in $All) {
            $isLicensed = & $isLic $userRow
            if ($userRow.Location -eq 'ON-PREM') { & $addAttention (T 'att.c1') $userRow ('{0:N2} GB' -f ($userRow.MailboxOnPrem_MB / 1024)) (-$userRow.MailboxOnPrem_MB) }
            if ($userRow.Status -eq $Labels.Enabled -and $isLicensed -and $userRow.MFA_Status -eq $Labels.MfaNo) { & $addAttention (T 'att.c2') $userRow $userRow.MFA_PerUser }
            if ($userRow.Status -eq $Labels.Disabled -and $isLicensed) { & $addAttention (T 'att.c3') $userRow '' }
            if ($userRow.Status -eq $Labels.Enabled -and $isLicensed -and (-not $userRow.M365_LastSignIn -or $userRow.M365_LastSignIn -lt $InactiveCut)) { & $addAttention (T 'att.c4' $Cfg.InactiveDays) $userRow (& $lastOrNever $userRow.M365_LastSignIn) }
            if ($userRow.Status -eq $Labels.Enabled -and ($AdOrigins -contains $userRow.Origin) -and (-not $userRow.AD_LastLogon -or $userRow.AD_LastLogon -lt $InactiveCut)) { & $addAttention (T 'att.c5' $Cfg.InactiveDays) $userRow (& $lastOrNever $userRow.AD_LastLogon) }
            if ($Src.M365 -and $Src.AD -and $userRow.Origin -eq $Labels.ADOnly -and $userRow.Status -eq $Labels.Enabled) { & $addAttention (T 'att.c6') $userRow $userRow.OU }
            if ($userRow.SyncIssues) { & $addAttention (T 'att.c7') $userRow $userRow.SyncIssues }
        }
        $Attn = @($Attn | Sort-Object Category, _k, Username | Select-Object * -ExcludeProperty _k)

        # ================================================= MFA DETAILS / TENANT LICENSES
        $TabMfa = ConvertTo-HSMfaRow -Cfg $Cfg -EntraUsers $entra -Mfa @{ Registration = $MfaReg; PerUser = $PerUser }
        $TabLic = ConvertTo-HSLicenseRow -Skus $skus -LicenseNames $LicenseNames

        # ================================================= TABLES (localized names)
        # columns that make no sense for the selected infrastructure
        $drop = @()
        if (-not $Src.M365) { $drop += 'M365_LastSignIn', 'Licenses', 'Mailbox365_MB', 'OneDrive_MB', 'MFA_Status', 'MFA_DefaultMethod', 'MFA_Methods', 'MFA_PerUser', 'UPN', 'Origin' }
        if (-not $Src.Exch) { $drop += 'MailboxOnPrem_MB' }
        if (-not $Src.Exch -and -not $Src.M365) { $drop += 'Location', 'MailboxType', 'MailboxSource' }
        if (-not $Src.AD) { $drop += 'AD_LastLogon', 'PasswordLastSet', 'OU' }
        if (-not $Sync) { $drop += 'SyncIssues' }
        # columns of the collectors skipped by a partial load
        if ($Src.M365 -and -not $Parts.Mfa) { $drop += 'MFA_Status', 'MFA_DefaultMethod', 'MFA_Methods', 'MFA_PerUser' }
        elseif ($Src.M365 -and -not $Parts.MfaPerUser) { $drop += 'MFA_PerUser' }
        if ($Src.M365 -and -not $Parts.Reports) { $drop += 'Mailbox365_MB', 'OneDrive_MB' }
        if ($Src.Exch -and -not $Parts.Exch) { $drop += 'MailboxOnPrem_MB' }
        if (-not $LocOk) { $drop += 'Location', 'MailboxType', 'MailboxSource' }
        $strip = { param($rows) if ($drop.Count) { @($rows | Select-Object * -ExcludeProperty $drop) } else { @($rows) } }

        $Tables = [ordered]@{}
        $Tables[(T 'tab.overview')] = $Overview
        $Tables[(T 'tab.users')] = & $strip $All
        foreach ($groupName in $Groups.Keys) { $tabName = if ($Tables.Contains($groupName)) { "OU_$groupName" } else { $groupName }; $Tables[$tabName] = & $strip $Groups[$groupName] }
        if ($TabOnPrem.Count) { $Tables[(T 'tab.mbxOnPrem')] = @($TabOnPrem | Sort-Object Size_MB -Descending) }
        if ($TabMbx365.Count) { $Tables[(T 'tab.mbx365')] = @($TabMbx365 | Sort-Object Size_MB -Descending) }
        if ($TabOD.Count) { $Tables[(T 'tab.onedrive')] = @($TabOD | Sort-Object Size_MB -Descending) }
        if ($Src.M365) {
            if ($Parts.Mfa) { $Tables[(T 'tab.mfa')] = $TabMfa }
            $Tables[(T 'tab.licenses')] = $TabLic
        }
        if ($Sync -and $Sync.Rows.Count) { $Tables[(T 'tab.sync')] = $Sync.Rows }
        if ($Attn.Count) { $Tables[(T 'tab.attention')] = & $strip $Attn }
        $hist = Join-Path (Get-HSOutputFolder $Cfg) 'history.csv'
        if (Test-Path $hist) { $Tables[(T 'tab.history')] = @(Import-Csv -Path $hist -Delimiter ';' -Encoding UTF8) }

        Write-HSLog (T 'log.done' @($All.Count, "$($Totals['loc365'])", "$($Totals['locOnPrem'])", "$($Totals['mfaPct'])"))

        [pscustomobject]@{
            Generated    = $Now
            Tables       = $Tables
            OverviewName = (T 'tab.overview')
            MetricFlags  = @($Metrics | ForEach-Object { [pscustomobject]@{ IsSection = $_.IsSection; IsGB = $_.IsGB } })
            Users        = $All          # unstripped rows and per-user extras, for the user card
            Partial      = if ($Skipped.Count) { if ($PartialLabel) { $PartialLabel } else { 'Custom' } } else { '' }
            Skipped      = @($Skipped)
            SyncInfo     = $SyncInfo
            SourceFlags  = $Src
            Totals       = $Totals
            Warnings     = @($script:HS_Warnings | Select-Object -Unique)
            Sources      = (@($(if ($Src.AD) { "AD $($Cfg.DomainFqdn)" }), $(if ($Src.Exch) { "Exchange $($Cfg.ExchangeServer)" }), $(if ($Src.M365) { "Microsoft 365 $((Get-MgContext).TenantId)" })) | Where-Object { $_ }) -join ' | '
        }
    }

    function Get-HSDataset {
        <#
        .SYNOPSIS
            Collects one dataset on its own, without the rest of the report (what -Only asks for).
            Returns the rows of that tab. The Microsoft 365 datasets need a Graph connection
            (Connect-HS365); MailboxOnPrem connects to Exchange with the current Windows account.
        #>
        param($Cfg, [ValidateSet('MFA', 'Licenses', 'Mailbox365', 'OneDrive', 'MailboxOnPrem', 'Sync')][string]$Name, [pscredential]$ExchangeCredential)
        $ErrorActionPreference = 'Stop'
        $script:HS_Warnings = [System.Collections.Generic.List[string]]::new()
        if ($Name -ne 'MailboxOnPrem' -and -not (Get-MgContext)) { throw (T 'err.not365') }
        switch ($Name) {
            'MFA' {
                $entra = Get-HSEntraUser -NoSignIn
                $mfaData = Get-HSMfaRegistration -Cfg $Cfg -EntraUsers $entra
                ConvertTo-HSMfaRow -Cfg $Cfg -EntraUsers $entra -Mfa $mfaData
            }
            'Licenses' { ConvertTo-HSLicenseRow -Skus (Get-HSSubscribedSku) -LicenseNames (Get-HSLicenseNameMap $Cfg) }
            'Mailbox365' {
                $usage = Get-HSMailbox365Usage
                if ($usage.Concealed) { Write-HSLog (T 'warn.concealed') 'WARN' }
                $usage.Rows | Sort-Object Size_MB -Descending
            }
            'OneDrive' {
                $usage = Get-HSOneDriveUsage
                if ($usage.Concealed) { Write-HSLog (T 'warn.concealed') 'WARN' }
                $usage.Rows | Sort-Object Size_MB -Descending
            }
            'MailboxOnPrem' {
                if (-not $Cfg.ExchangeServer -and $Cfg.ExchangeProtocol -ne 'local') { throw (T 'msg.exchNoServer') }
                (Get-HSExchangeMailbox -Cfg $Cfg -Credential $ExchangeCredential).Rows | Sort-Object Size_MB -Descending
            }
            'Sync' {
                if (-not $Cfg.DomainFqdn) { throw (T 'err.noDomain') }
                $adUsers = (Get-HSADUser -Cfg $Cfg).Users
                $entra = Get-HSEntraUser -NoSignIn
                (Get-HSSyncIssue -Cfg $Cfg -ADUsers $adUsers -EntraUsers $entra -Org (Get-HSDirectorySync)).Rows
            }
        }
    }

    function Update-HSHistory {
        <#
        .SYNOPSIS
            Appends this run's totals to history.csv in the output folder and reloads the HISTORY table.
        #>
        param($Data, $Cfg)
        $totals = $Data.Totals
        $file = Join-Path (Get-HSOutputFolder $Cfg) 'history.csv'
        [pscustomobject][ordered]@{
            Date               = $Data.Generated.ToString('yyyy-MM-dd HH:mm')
            Users              = $totals['total']
            Enabled            = $totals['enabled']
            On365              = $totals['loc365']
            OnPrem             = $totals['locOnPrem']
            OnPremSpace_GB     = $totals['spaceOnPrem']
            Space365_GB        = $totals['space365']
            OneDrive_GB        = $totals['spaceOD']
            MFA_Percent        = $totals['mfaPct']
            LicensedWithoutMFA = $totals['licNoMfa']
        } | Export-Csv -Path $file -Append -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $Data.Tables[(T 'tab.history')] = @(Import-Csv -Path $file -Delimiter ';' -Encoding UTF8)
    }

    # --------------------------------------------------------------------------- export
    function Get-ColLetter {
        param([int]$Number)
        $letters = ''
        while ($Number -gt 0) { $remainder = ($Number - 1) % 26; $letters = [string][char](65 + $remainder) + $letters; $Number = [int][math]::Floor(($Number - $remainder) / 26) }
        $letters
    }

    function ConvertTo-HSLocalized {
        # Renames the (English) property names to the current language for display / export
        param([object[]]$Rows)
        if (-not $Rows -or -not $Rows.Count) { return @() }
        $props = @($Rows[0].PSObject.Properties.Name)
        $sel = foreach ($prop in $props) { $header = Get-HSHeader $prop; if ($header -eq $prop) { $prop } else { @{ Name = $header; Expression = [scriptblock]::Create("`$_.'$prop'") } } }
        @($Rows | Select-Object -Property $sel)
    }

    function Get-HSSheetName {
        param([string]$Name)
        $sheetName = $Name -replace '[\[\]\*\?/\\:]', '_'
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        $sheetName
    }

    function Export-HSExcel {
        <#
        .SYNOPSIS
            Writes every table in Data to a formatted .xlsx at Path (one sheet per table). Returns Path.
        #>
        param($Data, [string]$Path)
        Import-Module ImportExcel -ErrorAction Stop
        if (Test-Path $Path) { Remove-Item $Path -Force }
        $condFormats = @(
            New-ConditionalText -Text (T 'val.mfaNo') -ConditionalTextColor DarkRed -BackgroundColor MistyRose
            New-ConditionalText -Text 'ON-PREM' -ConditionalTextColor DarkGoldenrod -BackgroundColor LightYellow
            New-ConditionalText -Text 'NO MAILBOX' -ConditionalTextColor Gray
        )
        $sheets = [ordered]@{}
        $i = 0
        foreach ($name in $Data.Tables.Keys) {
            $rows = if ($name -eq $Data.OverviewName) { @($Data.Tables[$name]) } else { @(ConvertTo-HSLocalized @($Data.Tables[$name])) }
            if (-not $rows.Count) { continue }
            $i++
            $sheetName = Get-HSSheetName $name
            $sheets[$sheetName] = $rows
            # not "T1": ImportExcel 7.8 refuses table names that read as a cell address and then creates no table
            $excelParams = @{ Path = $Path; WorksheetName = $sheetName; TableName = "HS_$i"; TableStyle = 'Medium2'; AutoSize = $true; BoldTopRow = $true }
            if ($name -eq $Data.OverviewName) { $excelParams.FreezeTopRowFirstColumn = $true } else { $excelParams.FreezeTopRow = $true; $excelParams.ConditionalText = $condFormats }
            $rows | Export-Excel @excelParams
            Write-HSLog (T 'log.excelSheet' @($sheetName, $rows.Count))
        }

        $pkg = Open-ExcelPackage -Path $Path
        try {
            foreach ($sheetName in $sheets.Keys) {
                $sheet = $pkg.Workbook.Worksheets[$sheetName]; if (-not $sheet) { continue }
                $colIndex = 0
                foreach ($prop in $sheets[$sheetName][0].PSObject.Properties) {
                    $colIndex++
                    $sample = @($sheets[$sheetName] | ForEach-Object { $_.($prop.Name) } | Where-Object { $null -ne $_ } | Select-Object -First 1)
                    if ($sample.Count -and $sample[0] -is [datetime]) { $sheet.Column($colIndex).Style.Numberformat.Format = 'yyyy-mm-dd hh:mm' }
                    elseif ($sample.Count -and $sample[0] -is [double] -and $prop.Name -match '_MB$|_GB$') { $sheet.Column($colIndex).Style.Numberformat.Format = '#,##0.00' }
                }
            }
            $sheet = if ($Data.OverviewName) { $pkg.Workbook.Worksheets[(Get-HSSheetName $Data.OverviewName)] }
            if ($sheet) {
                $nCols = @($Data.Tables[$Data.OverviewName])[0].PSObject.Properties.Name.Count
                $last = Get-ColLetter $nCols
                for ($metricIdx = 0; $metricIdx -lt $Data.MetricFlags.Count; $metricIdx++) {
                    $excelRow = $metricIdx + 2; $flags = $Data.MetricFlags[$metricIdx]
                    if ($flags.IsSection) { Set-ExcelRange -Worksheet $sheet -Range "A$($excelRow):$last$($excelRow)" -Bold -BackgroundColor ([System.Drawing.Color]::LightSteelBlue) }
                    elseif ($flags.IsGB) { Set-ExcelRange -Worksheet $sheet -Range "B$($excelRow):$last$($excelRow)" -NumberFormat '#,##0.00' }
                }
                Set-ExcelRange -Worksheet $sheet -Range "$($last)2:$last$($Data.MetricFlags.Count + 1)" -Bold
                $noteRow = $Data.MetricFlags.Count + 3
                $sheet.Cells["A$noteRow"].Value = (T 'xl.generated' @($Data.Generated.ToString('yyyy-MM-dd HH:mm'), $Data.Sources)); $noteRow++
                $sheet.Cells["A$noteRow"].Value = (T 'xl.legend'); $noteRow++
                foreach ($warning in $Data.Warnings) { $sheet.Cells["A$noteRow"].Value = (T 'xl.warning' $warning); $sheet.Cells["A$noteRow"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed); $noteRow++ }
            }
        }
        finally { Close-ExcelPackage $pkg }
        Write-HSLog (T 'log.excelSaved' $Path)
        $Path
    }

    function Export-HSCsv {
        <#
        .SYNOPSIS
            Writes every table in Data as a semicolon-separated CSV file in Folder. Returns Folder.
        #>
        param($Data, [string]$Folder)
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        foreach ($name in $Data.Tables.Keys) {
            $rows = if ($name -eq $Data.OverviewName) { @($Data.Tables[$name]) } else { @(ConvertTo-HSLocalized @($Data.Tables[$name])) }
            if (-not $rows.Count) { continue }
            $rows | Export-Csv -Path (Join-Path $Folder "$(Get-HSSheetName $name).csv") -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        }
        Write-HSLog (T 'log.csvSaved' $Folder)
        $Folder
    }

    function Export-HSUserCsv {
        <#
        .SYNOPSIS
            Writes the user list of Data (USERS tab) as one semicolon-separated CSV file at Path. Returns Path.
        #>
        param($Data, [string]$Path)
        $rows = @(ConvertTo-HSLocalized @($Data.Tables[(T 'tab.users')]))
        $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-HSLog (T 'log.csvSaved' $Path)
        $Path
    }

    # --------------------------------------------------------------------------- export presets (columns of the user list)
    # Built-in presets: English column names of the USERS tab, in export order. Full and Custom are not listed here:
    # Full = every tab and column, Custom = Cfg.ExportColumns. Presets saved by the user live in Cfg.ExportPresets.
    $script:HS_ExportPresets = [ordered]@{
        Access    = @('FirstName', 'LastName', 'Username', 'Status', 'AD_LastLogon', 'M365_LastSignIn', 'PasswordLastSet', 'Licenses')
        Licenses  = @('FirstName', 'LastName', 'Username', 'UPN', 'Status', 'Licenses', 'M365_LastSignIn', 'Notes')
        Migration = @('Group', 'FirstName', 'LastName', 'Username', 'Status', 'Location', 'MailboxType', 'MailboxOnPrem_MB', 'Mailbox365_MB', 'OneDrive_MB')
        MFA       = @('FirstName', 'LastName', 'Username', 'UPN', 'Status', 'Licenses', 'MFA_Status', 'MFA_DefaultMethod', 'MFA_Methods', 'MFA_PerUser')
        Directory = @('Group', 'FirstName', 'LastName', 'Username', 'UPN', 'OU', 'Status', 'Origin')
    }

    function Get-HSExportPresets {
        <#
        .SYNOPSIS
            Built-in presets followed by the ones saved in config.json (ExportPresets). Ordered name -> column list.
        #>
        param($Cfg)
        $presets = [ordered]@{}
        foreach ($key in $script:HS_ExportPresets.Keys) { $presets[$key] = @($script:HS_ExportPresets[$key]) }
        if ($Cfg -and $Cfg.ExportPresets) {
            foreach ($prop in (ConvertTo-HSPresetObject $Cfg.ExportPresets).PSObject.Properties) { if (@($prop.Value).Count) { $presets[$prop.Name] = @($prop.Value) } }
        }
        $presets
    }

    function Get-HSExportColumns {
        <#
        .SYNOPSIS
            Columns of an export preset: nothing for Full (= the whole report), Cfg.ExportColumns for Custom.
        #>
        param($Cfg, [string]$Preset)
        if (-not $Preset -or $Preset -eq 'Full') { return @() }
        if ($Preset -eq 'Custom') { return @($Cfg.ExportColumns) }
        $presets = Get-HSExportPresets $Cfg
        $key = @($presets.Keys | Where-Object { $_ -eq $Preset })[0]
        if (-not $key) { throw (T 'err.exportPreset' @($Preset, ((@('Full') + @($presets.Keys) + @('Custom')) -join ', '))) }
        @($presets[$key])
    }

    # columns of the USERS tab, in order; Get-HSUserColumns drops those the infrastructure mode cannot fill (as Get-HSData does)
    $script:HS_UserColumns = @('Group', 'FirstName', 'LastName', 'Username', 'OU', 'Status', 'AD_LastLogon', 'M365_LastSignIn', 'Licenses',
        'Mailbox365_MB', 'MailboxOnPrem_MB', 'OneDrive_MB', 'PasswordLastSet', 'Location', 'MailboxType', 'MailboxSource',
        'MFA_Status', 'MFA_DefaultMethod', 'MFA_Methods', 'MFA_PerUser', 'UPN', 'Origin', 'SyncIssues', 'Notes')

    function Get-HSUserColumns {
        param([string]$Mode)
        $src = Get-HSSources $Mode
        $drop = @()
        if (-not $src.M365) { $drop += 'M365_LastSignIn', 'Licenses', 'Mailbox365_MB', 'OneDrive_MB', 'MFA_Status', 'MFA_DefaultMethod', 'MFA_Methods', 'MFA_PerUser', 'UPN', 'Origin' }
        if (-not $src.Exch) { $drop += 'MailboxOnPrem_MB' }
        if (-not $src.Exch -and -not $src.M365) { $drop += 'Location', 'MailboxType', 'MailboxSource' }
        if (-not $src.AD) { $drop += 'AD_LastLogon', 'PasswordLastSet', 'OU' }
        if (-not ($src.AD -and $src.M365)) { $drop += 'SyncIssues' }
        @($script:HS_UserColumns | Where-Object { $drop -notcontains $_ })
    }

    function Get-HSLoadParts {
        <#
        .SYNOPSIS
            Optional collectors needed by the export columns (dynamic load). No columns = FULL: everything.
            AD, the Entra user list and the licenses are always read: they make the rows.
            Mfa = MFA registration report, MfaPerUser = legacy per-user state (one call per user),
            Reports = Exchange Online / OneDrive usage, Exch = Exchange on-premises mailboxes, Sync = AD <-> Entra check.
        #>
        param([string[]]$Columns)
        $cols = @($Columns | Where-Object { $_ })
        if (-not $cols.Count) { return @{ Mfa = $true; MfaPerUser = $true; Reports = $true; Exch = $true; Sync = $true } }
        $need = @{ Mfa = $false; MfaPerUser = $false; Reports = $false; Exch = $false; Sync = $false }
        foreach ($col in @(Resolve-HSExportColumns -Columns $cols -Available $script:HS_UserColumns)) {
            switch ($col) {
                { $_ -in 'MFA_Status', 'MFA_DefaultMethod', 'MFA_Methods', 'Notes' } { $need.Mfa = $true }
                'MFA_PerUser' { $need.Mfa = $true; $need.MfaPerUser = $true }
                { $_ -in 'Mailbox365_MB', 'OneDrive_MB' } { $need.Reports = $true }
                'MailboxOnPrem_MB' { $need.Exch = $true }
                { $_ -in 'Location', 'MailboxType', 'MailboxSource' } { $need.Exch = $true; $need.Reports = $true }
                'SyncIssues' { $need.Sync = $true }
            }
        }
        $need
    }

    function Get-HSExportUserRows {
        # the user list of the loaded data: the USERS tab (columns of the infrastructure mode), else the merged rows
        param($Data)
        $rows = $Data.Tables[(T 'tab.users')]
        if (-not $rows) { $rows = $Data.Users }
        @($rows)
    }

    function Get-HSEnabledValues {
        # "Enabled" in every language: an imported report keeps the values of the language it was written in
        @($script:HS_Str.Keys | ForEach-Object { $script:HS_Str[$_]['val.enabled'] } | Where-Object { $_ })
    }

    function Resolve-HSExportColumns {
        <#
        .SYNOPSIS
            English property names for the requested columns (English names or headers in any language,
            any case), in the requested order; columns not in Available are dropped.
        #>
        param([string[]]$Columns, [string[]]$Available)
        $resolved = [System.Collections.Generic.List[string]]::new()
        foreach ($col in @($Columns | ForEach-Object { "$_" -split '[,;]' })) {
            $name = $col.Trim(); if (-not $name) { continue }
            $match = @($Available | Where-Object { $_ -eq $name })[0]
            if (-not $match) {
                foreach ($lang in $script:HS_Cols.Keys) {
                    $match = @($Available | Where-Object { $script:HS_Cols[$lang][$_] -eq $name })[0]
                    if ($match) { break }
                }
            }
            if ($match -and -not $resolved.Contains($match)) { $resolved.Add($match) }
        }
        $resolved.ToArray()
    }

    function Select-HSExportData {
        <#
        .SYNOPSIS
            Copy of Data for a preset / custom export: the user list with only Columns (in that order), optionally
            one sheet per group, the overview, enabled accounts only. Without Columns returns Data unchanged (FULL).
        #>
        param($Data, [string[]]$Columns, [switch]$GroupSheets, [switch]$Overview, [switch]$EnabledOnly, [string]$Label = 'Custom')
        if (-not @($Columns | Where-Object { $_ }).Count) { return $Data }
        $rows = @(Get-HSExportUserRows $Data)
        if (-not $rows.Count) { throw (T 'err.exportNoUsers') }
        $available = @($rows[0].PSObject.Properties.Name)
        $cols = @(Resolve-HSExportColumns -Columns $Columns -Available $available)
        if (-not $cols.Count) { throw (T 'err.exportNoColumns' ($Columns -join ', ')) }
        if ($EnabledOnly -and $available -contains 'Status') {
            $enabledValues = Get-HSEnabledValues
            $rows = @($rows | Where-Object { $enabledValues -contains "$($_.Status)" })
        }
        if (-not $rows.Count) { throw (T 'err.exportNoRows') }

        $tables = [ordered]@{}
        $overviewName = ''
        if ($Overview -and $Data.OverviewName -and $Data.Tables.Contains($Data.OverviewName)) {
            $overviewName = $Data.OverviewName
            $tables[$overviewName] = $Data.Tables[$overviewName]
        }
        $usersTab = T 'tab.users'
        $tables[$usersTab] = @($rows | Select-Object -Property $cols)
        if ($GroupSheets -and $available -contains 'Group') {
            foreach ($bucket in @($rows | Group-Object -Property Group)) {
                if (-not $bucket.Name) { continue }
                $tabName = if ($tables.Contains($bucket.Name)) { "OU_$($bucket.Name)" } else { $bucket.Name }
                $tables[$tabName] = @($bucket.Group | Select-Object -Property $cols)
            }
        }
        Write-HSLog (T 'log.exportPreset' @($Label, $cols.Count, $rows.Count))
        [pscustomobject]@{
            Generated    = $Data.Generated
            Tables       = $tables
            OverviewName = $overviewName
            MetricFlags  = if ($overviewName) { $Data.MetricFlags } else { @() }
            Totals       = $Data.Totals
            Warnings     = @($Data.Warnings)
            Sources      = $Data.Sources
        }
    }

    # --------------------------------------------------------------------------- import
    $script:HS_DateProps = 'AD_LastLogon', 'M365_LastSignIn', 'PasswordLastSet', 'LastLogon', 'LastActivity', 'Created', 'LastSync'
    $script:HS_NumberProps = '_MB$|_GB$|^(Items|Files|ActiveFiles|Purchased|Assigned|Available|Suspended|Warning|Users|Enabled|OnPrem|On365|MFA_Percent|LicensedWithoutMFA)$'

    function Resolve-HSHeader {
        <#
        .SYNOPSIS
            English property name for a column header of an exported report (any language).
            "Nome" is ambiguous in Italian: first name in user tables, license name in TENANT_LICENSES /
            ATTENTION, display name elsewhere.
        #>
        param([string]$Header, [string[]]$SheetHeaders)
        $candidates = @(@(foreach ($lang in $script:HS_Cols.Keys) { foreach ($prop in $script:HS_Cols[$lang].Keys) { if ($script:HS_Cols[$lang][$prop] -ceq $Header) { $prop } } }) | Select-Object -Unique)
        if (-not $candidates.Count) { return $Header }
        if ($candidates.Count -eq 1) { return $candidates[0] }
        if ($SheetHeaders -contains (Get-HSHeaderIn 'LastName')) { return 'FirstName' }
        if ($SheetHeaders -contains 'SKU' -or $SheetHeaders -contains (Get-HSHeaderIn 'Licenses')) { return 'Name' }
        'DisplayName'
    }

    function Get-HSHeaderIn {
        # localized header of Prop in whichever language has a translation (Italian today)
        param([string]$Prop)
        foreach ($lang in $script:HS_Cols.Keys) { if ($script:HS_Cols[$lang].ContainsKey($Prop)) { return $script:HS_Cols[$lang][$Prop] } }
        $Prop
    }

    function ConvertFrom-HSCell {
        <#
        .SYNOPSIS
            Restores the type of one imported value: Excel stores dates as numbers (OADate), CSV stores
            everything as text in the culture of the PC that exported it.
        #>
        param($Value, [string]$Prop, [switch]$FromCsv, [switch]$Number)
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) {
            $text = $Value.Trim()
            if ($text -eq '') { return $null }
            if ($script:HS_DateProps -contains $Prop) {
                $date = [datetime]::MinValue
                if ([datetime]::TryParse($text, [Globalization.CultureInfo]::CurrentCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) { return $date }
                if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) { return $date }
                return $text
            }
            if ($Number -or $Prop -match $script:HS_NumberProps) {
                $num = 0.0
                if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::CurrentCulture, [ref]$num)) { return $num }
                if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$num)) { return $num }
            }
            return $text
        }
        if ($Value -is [double] -and $script:HS_DateProps -contains $Prop) { try { return [datetime]::FromOADate($Value) } catch { return $Value } }
        $Value
    }

    function Import-HSReport {
        <#
        .SYNOPSIS
            Reads a report written by HybridScope back into the same shape Get-HSData returns, so the GUI
            can show it: a full or single-tab .xlsx, a single .csv, or a HybridScope_CSV_* folder (or one
            of its files). Headers in any language are mapped back to the English property names.
        #>
        param([string]$Path)
        $Path = (Resolve-Path -LiteralPath $Path).ProviderPath
        Write-HSLog (T 'log.import' $Path)
        $sheets = [System.Collections.Generic.List[object]]::new()     # @{ Name; Rows (raw) }
        $isCsv = $false

        # tab names of either language -> the current one (OVERVIEW / PANORAMICA, USERS / UTENTI...)
        $tabNames = @{}
        foreach ($lang in $script:HS_Str.Keys) { foreach ($key in $script:HS_Str[$lang].Keys) { if ($key -like 'tab.*') { $tabNames[(Get-HSSheetName $script:HS_Str[$lang][$key])] = $key } } }
        # CSV files carry no order: same order as the report (group tabs between USERS and the mailbox tabs)
        $tabOrder = @{ 'tab.overview' = 0; 'tab.users' = 1; 'tab.mbxOnPrem' = 3; 'tab.mbx365' = 4; 'tab.onedrive' = 5; 'tab.mfa' = 6; 'tab.licenses' = 7; 'tab.sync' = 8; 'tab.attention' = 9; 'tab.history' = 10 }
        $rankOf = { param($baseName) $tabKey = $tabNames[($baseName -replace '_\d{4}-\d{2}-\d{2}$', '')]; if ($tabKey -and $tabOrder.ContainsKey($tabKey)) { $tabOrder[$tabKey] } else { 2 } }

        if ((Test-Path -LiteralPath $Path -PathType Container) -or $Path -like '*.csv') {
            $isCsv = $true
            $folder = if (Test-Path -LiteralPath $Path -PathType Container) { $Path } else { Split-Path -Parent $Path }
            # a file of a HybridScope_CSV_* export opens the whole export; any other CSV opens alone
            $files = if ((Test-Path -LiteralPath $Path -PathType Container) -or (Split-Path -Leaf $folder) -like 'HybridScope_CSV_*') {
                @(Get-ChildItem -LiteralPath $folder -Filter '*.csv' | Sort-Object @{ Expression = { & $rankOf $_.BaseName } }, LastWriteTime, Name)
            }
            else { @(Get-Item -LiteralPath $Path) }
            foreach ($file in $files) {
                $sheets.Add(@{ Name = ($file.BaseName -replace '_\d{4}-\d{2}-\d{2}$', ''); Rows = @(Import-Csv -LiteralPath $file.FullName -Delimiter ';' -Encoding UTF8) })
            }
        }
        else {
            Import-Module ImportExcel -ErrorAction Stop
            # read a copy: the file may be open in Excel
            $copy = Join-Path ([IO.Path]::GetTempPath()) ("HybridScope_import_{0}.xlsx" -f [guid]::NewGuid().ToString('N'))
            Copy-Item -LiteralPath $Path -Destination $copy -Force
            try {
                foreach ($info in @(Get-ExcelSheetInfo -Path $copy | Sort-Object Index)) {
                    $sheets.Add(@{ Name = "$($info.Name)"; Rows = @(Import-Excel -Path $copy -WorksheetName $info.Name) })
                }
            }
            finally { Remove-Item -LiteralPath $copy -Force -ErrorAction SilentlyContinue }
        }

        $notePrefixes = @(foreach ($lang in $script:HS_Str.Keys) { foreach ($key in 'xl.generated', 'xl.legend', 'xl.warning') { ($script:HS_Str[$lang][$key] -split '\{0\}')[0] } }) | Where-Object { $_ } | Select-Object -Unique
        $warningPrefixes = @(foreach ($lang in $script:HS_Str.Keys) { ($script:HS_Str[$lang]['xl.warning'] -split '\{0\}')[0] }) | Select-Object -Unique

        $tables = [ordered]@{}; $overviewName = ''; $metricFlags = @(); $generated = $null; $sources = ''
        $warnings = [System.Collections.Generic.List[string]]::new()
        foreach ($sheet in $sheets) {
            $rows = @($sheet.Rows | Where-Object { $_ })
            if (-not $rows.Count) { continue }
            $key = $tabNames[$sheet.Name]
            $name = if ($key) { T $key } else { $sheet.Name }
            $headers = @($rows[0].PSObject.Properties.Name)

            if ($key -eq 'tab.overview') {
                # metric rows; the lines under the table are notes (generated on / legend / warnings)
                $overviewName = $name
                $kept = [System.Collections.Generic.List[object]]::new(); $flags = [System.Collections.Generic.List[object]]::new()
                $inNotes = $false
                foreach ($row in $rows) {
                    $first = "$($row.($headers[0]))"
                    $others = [System.Collections.Generic.List[object]]::new()
                    for ($i = 1; $i -lt $headers.Count; $i++) { $cellValue = ConvertFrom-HSCell $row.($headers[$i]) '' -FromCsv:$isCsv -Number; $others.Add($cellValue) }
                    $empty = -not @($others | Where-Object { $null -ne $_ }).Count
                    if ($empty -and -not $first) { continue }
                    if ($empty -and @($notePrefixes | Where-Object { $first.StartsWith($_) }).Count) { $inNotes = $true }
                    if ($inNotes) {
                        if ($first -match '(\d{4}-\d{2}-\d{2} \d{2}:\d{2})') { if (-not $generated) { $generated = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm', $null) }; if ($first -match ' - (.+)$') { $sources = $Matches[1] } }
                        foreach ($prefix in $warningPrefixes) { if ($first.StartsWith($prefix)) { $warnings.Add($first.Substring($prefix.Length)) } }
                        continue
                    }
                    $line = [ordered]@{ $headers[0] = $first }
                    for ($i = 1; $i -lt $headers.Count; $i++) { $line[$headers[$i]] = $others[$i - 1] }
                    $kept.Add([pscustomobject]$line)
                    $flags.Add([pscustomobject]@{ IsSection = $empty; IsGB = ($first -like '*(GB)*') })
                }
                $tables[$name] = @($kept); $metricFlags = @($flags)
                Write-HSLog (T 'log.importSheet' @($name, $kept.Count))
                continue
            }

            $props = @(foreach ($header in $headers) { Resolve-HSHeader $header $headers })
            $converted = @(foreach ($row in $rows) {
                    $line = [ordered]@{}
                    for ($i = 0; $i -lt $headers.Count; $i++) { $line[$props[$i]] = ConvertFrom-HSCell $row.($headers[$i]) $props[$i] -FromCsv:$isCsv }
                    if (@($line.Values | Where-Object { $null -ne $_ }).Count) { [pscustomobject]$line }
                })
            if (-not $converted.Count) { continue }
            $tables[$name] = $converted
            Write-HSLog (T 'log.importSheet' @($name, $converted.Count))
        }
        if (-not $tables.Count) { throw (T 'err.importEmpty' $Path) }

        # the merged user list: the USERS tab, or the first table that looks like one (single-tab exports)
        $users = $tables[(T 'tab.users')]
        if (-not $users) { foreach ($tableName in $tables.Keys) { $first = @($tables[$tableName])[0]; if ($first.PSObject.Properties['Username'] -and $first.PSObject.Properties['Status']) { $users = $tables[$tableName]; break } } }
        $users = @($users)
        $sample = if ($users.Count) { @($users[0].PSObject.Properties.Name) } else { @() }

        if (-not $generated -and (Split-Path -Leaf $Path) -match '(\d{4}-\d{2}-\d{2})_(\d{2})(\d{2})') { $generated = [datetime]::ParseExact("$($Matches[1]) $($Matches[2]):$($Matches[3])", 'yyyy-MM-dd HH:mm', $null) }
        if (-not $generated) { $generated = (Get-Item -LiteralPath $Path).LastWriteTime }

        [pscustomobject]@{
            Generated    = $generated
            Tables       = $tables
            OverviewName = $overviewName
            MetricFlags  = $metricFlags
            Totals       = @{ total = $users.Count }
            Warnings     = @($warnings)
            Sources      = $sources
            Users        = $users
            SyncInfo     = @{}
            SourceFlags  = @{ AD = ($sample -contains 'OU'); M365 = ($sample -contains 'UPN'); Exch = ($sample -contains 'MailboxOnPrem_MB') }
            Imported     = $true
            ImportPath   = $Path
        }
    }

    function ConvertTo-HSDataTable {
        param([object[]]$Rows, [switch]$NoLocalize)
        $dt = New-Object System.Data.DataTable
        if (-not $Rows -or $Rows.Count -eq 0) { return , $dt }
        $props = @($Rows[0].PSObject.Properties.Name)
        foreach ($prop in $props) {
            $type = [string]
            $values = @($Rows | ForEach-Object { $_.$prop } | Where-Object { $null -ne $_ -and "$_" -ne '' })
            if ($values.Count) {
                if (-not @($values | Where-Object { $_ -isnot [datetime] }).Count) { $type = [datetime] }
                elseif (-not @($values | Where-Object { -not ($_ -is [int] -or $_ -is [long] -or $_ -is [double] -or $_ -is [decimal]) }).Count) { $type = [double] }
            }
            $hdr = if ($NoLocalize) { $prop } else { Get-HSHeader $prop }
            [void]$dt.Columns.Add($hdr, $type)
        }
        foreach ($item in $Rows) {
            $row = $dt.NewRow()
            for ($i = 0; $i -lt $props.Count; $i++) {
                $value = $item.($props[$i]); $col = $dt.Columns[$i]
                if ($null -eq $value -or ("$value" -eq '' -and $col.DataType -ne [string])) { $row[$i] = [DBNull]::Value }
                elseif ($col.DataType -eq [string]) { $row[$i] = "$value" }
                else { $row[$i] = $value }
            }
            $dt.Rows.Add($row)
        }
        , $dt
    }
}

# =====================================================================================
#  HEADLESS MODE (-Auto, -Only, -PassThru)
# =====================================================================================
if ($Only -or $PassThru -or $Columns -or $ExportPreset) { $Auto = $true }
if ($Auto) {
    . $Core
    $ErrorActionPreference = 'Stop'
    $script:HS_AppDir = $AppDir
    $cfg = Read-HSConfig $ConfigPath
    if ($Mode) { $cfg.Mode = $Mode }          # command-line override, not saved to config.json
    $script:HS_Lang = if ($PassThru) { 'en' } else { $cfg.Language }   # objects: predictable English values
    $out = Get-HSOutputFolder $cfg
    $logDir = Join-Path $out 'logs'
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $script:HS_LogFile = Join-Path $logDir "HybridScope_$stamp.log"
    $exit = 0

    # Reuse a Microsoft Graph connection the caller already opened, and leave it open afterwards
    $graphWasConnected = $false
    try { $graphWasConnected = [bool](Get-MgContext -ErrorAction Stop) } catch { }

    try {
        $needGraph = if ($Only) { $Only -ne 'MailboxOnPrem' } else { (Get-HSSources $cfg.Mode).M365 }
        if ($needGraph -and -not $graphWasConnected) { Connect-HS365 -Cfg $cfg -Unattended:($cfg.AuthMode -eq 'Certificate') | Out-Null }

        if ($Only) {
            $rows = @(Get-HSDataset -Cfg $cfg -Name $Only)
            if ($PassThru) { $rows }
            else {
                $tabKeys = @{ MFA = 'tab.mfa'; Licenses = 'tab.licenses'; Mailbox365 = 'tab.mbx365'; OneDrive = 'tab.onedrive'; MailboxOnPrem = 'tab.mbxOnPrem'; Sync = 'tab.sync' }
                $data = [pscustomobject]@{
                    Generated    = Get-Date
                    Tables       = [ordered]@{ (T $tabKeys[$Only]) = $rows }
                    OverviewName = ''
                    MetricFlags  = @()
                    Totals       = @{}
                    Warnings     = @($script:HS_Warnings | Select-Object -Unique)
                    Sources      = ''
                }
                if ($rows.Count) { Export-HSExcel -Data $data -Path (Join-Path $out "HybridScope_$($Only)_$stamp.xlsx") | Out-Null }
                else { Write-HSLog (T 'warn.noRows' $Only) 'WARN' }
            }
        }
        else {
            # -Columns > -ExportPreset > AutoExportPreset (config.json); -PassThru only follows the parameters
            $Columns = @($Columns | ForEach-Object { "$_" -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            $preset = if ($Columns.Count) { 'Custom' } elseif ($ExportPreset) { $ExportPreset } elseif ($PassThru) { 'Full' } else { "$($cfg.AutoExportPreset)" }
            $exportCols = if ($Columns.Count) { $Columns } else { @(Get-HSExportColumns -Cfg $cfg -Preset $preset) }
            $parts = if (@($exportCols).Count -and $cfg.ExportFastLoad) { Get-HSLoadParts -Columns $exportCols } else { $null }
            $data = Get-HSData -Cfg $cfg -Parts $parts -PartialLabel $preset
            if ($PassThru) {
                if ($exportCols.Count) { (Select-HSExportData -Data $data -Columns $exportCols -Label $preset).Tables[(T 'tab.users')] }
                else { $data.Tables[(T 'tab.users')] }
            }
            else {
                if ($data.Partial) { Write-HSLog (T 'log.historySkipped') } else { Update-HSHistory -Data $data -Cfg $cfg }
                $suffix = if ($exportCols.Count) { '_' + ($preset -replace '[^\w\-]', '') } else { '' }
                $exportData = Select-HSExportData -Data $data -Columns $exportCols -Label $preset -GroupSheets:([bool]$cfg.ExportGroupSheets) -Overview:([bool]$cfg.ExportOverview) -EnabledOnly:([bool]$cfg.ExportEnabledOnly)
                $xl = Join-Path $out "HybridScope$($suffix)_$stamp.xlsx"
                Export-HSExcel -Data $exportData -Path $xl | Out-Null
                try { Copy-Item $xl (Join-Path $out 'HybridScope_LATEST.xlsx') -Force }
                catch { Write-HSLog (T 'warn.latestLocked') 'WARN' }
            }
        }

        if ([int]$cfg.RetentionDays -gt 0) {
            $cutoff = (Get-Date).AddDays( - [int]$cfg.RetentionDays)
            Get-ChildItem $out -Filter 'HybridScope_*.xlsx' | Where-Object { $_.Name -ne 'HybridScope_LATEST.xlsx' -and $_.LastWriteTime -lt $cutoff } | Remove-Item -Force
            Get-ChildItem $logDir -Filter '*.log' | Where-Object LastWriteTime -lt $cutoff | Remove-Item -Force
        }
    }
    catch {
        Write-HSLog (T 'err.generic' "$($_.Exception.Message) $($_.InvocationInfo.PositionMessage)") 'ERROR'
        if ($PassThru) { Write-Error -Message $_.Exception.Message -ErrorAction Continue }
        $exit = 1
    }
    finally {
        if (-not $graphWasConnected) { try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { } }
        if ($script:HS_ExSession) { Remove-PSSession $script:HS_ExSession -ErrorAction SilentlyContinue }
    }
    exit $exit
}

# =====================================================================================
#  GUI
# =====================================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing
. $Core
$script:HS_AppDir = $AppDir
$cfg = Read-HSConfig $ConfigPath
$script:HS_Lang = $cfg.Language

$xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="HybridScope" Height="860" Width="1500"
        WindowStartupLocation="CenterScreen" Background="#F3F4F6" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
    <Style x:Key="PrimaryButton" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#1565C0" CornerRadius="3" Padding="18,5">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#0D47A1"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ModeButton" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#CFE0F5"/>
      <Setter Property="FontWeight" Value="Normal"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="3,0"/>
      <Setter Property="GroupName" Value="mode"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" CornerRadius="15" Padding="16,6" Background="#183858" BorderBrush="#2E5585" BorderThickness="1">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#21486F"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#4FA3F7"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="#8CC4FF"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="12,4"/>
      <Setter Property="Margin" Value="6,0,0,0"/>
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#B8C2D0"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Margin" Value="0,0,10,6"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="BorderBrush" Value="#C9D2DE"/>
      <Setter Property="Foreground" Value="#0E2A47"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="TextBlock"><Setter Property="VerticalAlignment" Value="Center"/></Style>
    <Style TargetType="TextBox"><Setter Property="VerticalContentAlignment" Value="Center"/><Setter Property="Height" Value="24"/><Setter Property="FontWeight" Value="Normal"/></Style>
    <Style TargetType="ComboBox"><Setter Property="Height" Value="24"/><Setter Property="FontWeight" Value="Normal"/></Style>
    <Style TargetType="CheckBox"><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="FontWeight" Value="Normal"/></Style>
  </Window.Resources>
  <DockPanel>
    <Border DockPanel.Dock="Top" Background="#FFFFFF" BorderBrush="#DDE3EA" BorderThickness="0,0,0,1">
      <Menu x:Name="mainMenu" Background="Transparent" Padding="4,1">
        <MenuItem x:Name="miFile" Header="{{menu.file}}">
          <MenuItem x:Name="miOpen" Header="{{menu.open}}" InputGestureText="Ctrl+O"/>
          <Separator/>
          <MenuItem x:Name="miExportXlsx" Header="{{btn.xlsx}}"/>
          <MenuItem x:Name="miExportCsv" Header="{{btn.csvall}}"/>
          <MenuItem x:Name="miExportTab" Header="{{btn.tab}}"/>
          <Separator/>
          <MenuItem x:Name="miFolder" Header="{{btn.folder}}"/>
          <Separator/>
          <MenuItem x:Name="miExit" Header="{{menu.exit}}" InputGestureText="Alt+F4"/>
        </MenuItem>
        <MenuItem x:Name="miHelp" Header="{{menu.help}}">
          <MenuItem x:Name="miDocs" Header="{{menu.docs}}" InputGestureText="F1"/>
          <MenuItem x:Name="miUpdates" Header="{{menu.updates}}"/>
          <MenuItem x:Name="miIssue" Header="{{menu.issue}}"/>
          <Separator/>
          <MenuItem x:Name="miSponsor" Header="{{menu.sponsor}}"/>
          <MenuItem x:Name="miCoffee" Header="{{menu.coffee}}"/>
          <Separator/>
          <MenuItem x:Name="miAbout" Header="{{menu.about}}"/>
        </MenuItem>
      </Menu>
    </Border>
    <Border DockPanel.Dock="Top" Background="#0E2A47" Padding="16,10">
      <DockPanel>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" VerticalAlignment="Center">
          <TextBlock x:Name="lblGenerated" Foreground="#9FB4D9" Margin="0,0,20,0"/>
          <TextBlock Text="{{lbl.language}}" Foreground="#9FB4D9" Margin="0,0,6,0"/>
          <ComboBox x:Name="cmbLang" Width="100">
            <ComboBoxItem Content="English" Tag="en"/><ComboBoxItem Content="Italiano" Tag="it"/>
          </ComboBox>
        </StackPanel>
        <StackPanel DockPanel.Dock="Left" VerticalAlignment="Center">
          <TextBlock x:Name="lblTitle" Text="HybridScope" Foreground="White" FontSize="20" FontWeight="SemiBold"/>
          <TextBlock Text="{{app.subtitle}}" Foreground="#9FB4D9"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
          <RadioButton x:Name="rbCloud" Style="{StaticResource ModeButton}" Content="Full Cloud" Tag="Cloud" ToolTip="{{mode.cloud}}"/>
          <RadioButton x:Name="rbFullHybrid" Style="{StaticResource ModeButton}" Content="Full Hybrid" Tag="FullHybrid" ToolTip="{{mode.fullhybrid}}"/>
          <RadioButton x:Name="rbHybrid" Style="{StaticResource ModeButton}" Content="Hybrid" Tag="Hybrid" ToolTip="{{mode.hybrid}}"/>
          <RadioButton x:Name="rbOnPrem" Style="{StaticResource ModeButton}" Content="On-Prem" Tag="OnPrem" ToolTip="{{mode.onprem}}"/>
          <RadioButton x:Name="rbAD" Style="{StaticResource ModeButton}" Content="AD" Tag="AD" ToolTip="{{mode.ad}}"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <Border DockPanel.Dock="Top" Background="White" BorderBrush="#DDE3EA" BorderThickness="0,0,0,1" Padding="12,8,12,2">
      <WrapPanel>
        <GroupBox x:Name="gbM365" Header="{{gb.m365}}">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="{{lbl.tenant}}" Margin="0,0,6,0" FontWeight="Normal"/>
            <TextBox x:Name="txtTenant" Width="190" ToolTip="{{tip.tenant}}"/>
            <Button x:Name="btn365" Content="{{btn.connect365}}"/>
            <Button x:Name="btnSignOut" Content="{{btn.signout}}" ToolTip="{{tip.signout}}"/>
            <Ellipse x:Name="led365" Width="12" Height="12" Fill="#B0B0B0" Margin="8,0,4,0"/>
            <TextBlock x:Name="lbl365" Text="{{st.notconnected}}" Foreground="#666" FontWeight="Normal" MinWidth="140"/>
          </StackPanel>
        </GroupBox>
        <GroupBox x:Name="gbCred" Header="{{gb.cred}}">
          <StackPanel Orientation="Horizontal">
            <CheckBox x:Name="chkWinCred" Content="{{chk.wincred}}" IsChecked="True"/>
            <TextBlock Text="{{lbl.user}}" Margin="12,0,6,0" FontWeight="Normal"/>
            <TextBox x:Name="txtUser" Width="180" ToolTip="{{tip.user}}"/>
            <TextBlock Text="{{lbl.password}}" Margin="8,0,6,0" FontWeight="Normal"/>
            <PasswordBox x:Name="pwdPass" Width="130" Height="24" VerticalContentAlignment="Center"/>
          </StackPanel>
        </GroupBox>
        <GroupBox x:Name="gbExch" Header="{{gb.exch}}">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="{{lbl.server}}" Margin="0,0,6,0" FontWeight="Normal"/>
            <TextBox x:Name="txtExch" Width="180" ToolTip="{{tip.server}}"/>
            <ComboBox x:Name="cmbExProto" Width="65" Margin="6,0,0,0" ToolTip="{{tip.proto}}">
              <ComboBoxItem Content="http"/><ComboBoxItem Content="https"/><ComboBoxItem Content="local"/>
            </ComboBox>
            <ComboBox x:Name="cmbExAuth" Width="90" Margin="6,0,0,0">
              <ComboBoxItem Content="Kerberos"/><ComboBoxItem Content="Negotiate"/><ComboBoxItem Content="Basic"/>
            </ComboBox>
            <Button x:Name="btnExch" Content="{{btn.connectExch}}"/>
            <Ellipse x:Name="ledExch" Width="12" Height="12" Fill="#B0B0B0" Margin="8,0,4,0"/>
            <TextBlock x:Name="lblExch" Text="{{st.notconnected}}" Foreground="#666" FontWeight="Normal" MinWidth="140"/>
          </StackPanel>
        </GroupBox>
        <GroupBox x:Name="gbAD" Header="{{gb.ad}}">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="{{lbl.domain}}" Margin="0,0,6,0" FontWeight="Normal"/>
            <TextBox x:Name="txtDomain" Width="140"/>
            <Button x:Name="btnAD" Content="{{btn.verifyAD}}"/>
            <Ellipse x:Name="ledAD" Width="12" Height="12" Fill="#B0B0B0" Margin="8,0,4,0"/>
            <TextBlock x:Name="lblAD" Text="{{st.notverified}}" Foreground="#666" FontWeight="Normal" MinWidth="120"/>
          </StackPanel>
        </GroupBox>
        <GroupBox x:Name="gbExport" Header="{{gb.export}}">
          <StackPanel Orientation="Horizontal">
            <ComboBox x:Name="cmbExport" Width="230" ToolTip="{{tip.cmbExport}}"/>
            <Button x:Name="btnExpCols" Content="{{btn.expCols}}" ToolTip="{{tip.expCols}}"/>
            <CheckBox x:Name="chkExpFast" Content="{{export.fast}}" ToolTip="{{tip.expFast}}" Margin="8,0,0,0"/>
            <TextBlock x:Name="lblExport" Margin="8,0,0,0" Foreground="#555" FontWeight="Normal" MaxWidth="420" TextTrimming="CharacterEllipsis"/>
          </StackPanel>
        </GroupBox>
      </WrapPanel>
    </Border>

    <Border DockPanel.Dock="Top" Padding="12,10,12,0">
      <StackPanel x:Name="stepPanel" Orientation="Horizontal" HorizontalAlignment="Center"/>
    </Border>

    <Border DockPanel.Dock="Top" Padding="12,8">
      <DockPanel>
        <StackPanel Orientation="Horizontal">
          <Button x:Name="btnLoad" Content="{{btn.load}}" Margin="0" Style="{StaticResource PrimaryButton}"/>
          <Button x:Name="btnStop" Content="{{btn.stop}}" ToolTip="{{tip.stop}}" IsEnabled="False" Foreground="#B42318" BorderBrush="#E5A3A3"/>
          <TextBlock Text="{{lbl.groupBy}}" Margin="16,0,6,0"/>
          <ComboBox x:Name="cmbGroupBy" Width="215" ToolTip="{{tip.groupBy}}">
            <ComboBoxItem Content="{{grp.opt.top}}" Tag="TopLevelOU"/>
            <ComboBoxItem Content="{{grp.opt.parent}}" Tag="ParentOU"/>
            <ComboBoxItem Content="{{grp.opt.none}}" Tag="None"/>
            <ComboBoxItem Content="{{grp.opt.rules}}" Tag="Rules"/>
          </ComboBox>
          <Button x:Name="btnXlsx" Content="{{btn.xlsx}}" ToolTip="{{tip.export}}" Margin="16,0,0,0" IsEnabled="False"/>
          <Button x:Name="btnCsvAll" Content="{{btn.csvall}}" ToolTip="{{tip.export}}" IsEnabled="False"/>
          <Button x:Name="btnTab" Content="{{btn.tab}}" IsEnabled="False"/>
          <Button x:Name="btnFolder" Content="{{btn.folder}}"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <TextBlock Text="{{lbl.search}}" Margin="0,0,6,0"/>
          <TextBox x:Name="txtFilter" Width="280"/>
          <TextBlock x:Name="lblCount" Margin="10,0,0,0" Foreground="#555" MinWidth="100"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <Border DockPanel.Dock="Bottom" Background="#0E2A47" Padding="10,4">
      <DockPanel>
        <ProgressBar x:Name="prg" Width="180" Height="10" DockPanel.Dock="Right" Foreground="#4FA3F7" Background="#22406E" BorderThickness="0"/>
        <TextBlock DockPanel.Dock="Right" Margin="16,0,16,0" Foreground="#9FB4D9" FontSize="11">
          <Run x:Name="runVersion"/> <Run Text="{{status.by}}"/> <Hyperlink x:Name="lnkStatusAuthor" Foreground="#CFE0F5" TextDecorations="{x:Null}" Cursor="Hand">Thonno</Hyperlink>
        </TextBlock>
        <TextBlock x:Name="lblStatus" Foreground="White" Text="{{st.ready}}" TextTrimming="CharacterEllipsis"/>
      </DockPanel>
    </Border>
    <Expander DockPanel.Dock="Bottom" Header="{{exp.log}}" x:Name="expLog" Margin="8,0,8,4">
      <TextBox x:Name="txtLog" Height="160" IsReadOnly="True" FontFamily="Consolas" FontSize="11" Background="#0E2A47" Foreground="#EEEDF0"
               VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" TextWrapping="NoWrap" VerticalContentAlignment="Top"/>
    </Expander>

    <Grid Margin="8,0,8,6">
      <TabControl x:Name="tabs"/>
      <TextBlock x:Name="lblEmpty" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#8A94A6" FontSize="14" TextAlignment="Center"
                 Text="{{empty.ready}}"/>
    </Grid>
  </DockPanel>
</Window>
'@

# localize XAML placeholders {{key}}
$xamlText = [regex]::Replace($xamlText, '\{\{([\w\.]+)\}\}', [System.Text.RegularExpressions.MatchEvaluator] {
        param($match) [System.Security.SecurityElement]::Escape((T $match.Groups[1].Value))
    })
[xml]$xaml = $xamlText
$Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$ui = @{}
foreach ($node in $xaml.SelectNodes("//*[@*[local-name()='Name']]")) {
    $nodeName = $node.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
    if ($nodeName) { $ui[$nodeName] = $Window.FindName($nodeName) }
}
$Window.Title = "HybridScope $HS_Version"
# window / taskbar icon: tools\HybridScope.ico next to the script (or the exe)
$script:IconPath = Join-Path $AppDir 'tools\HybridScope.ico'
if (Test-Path $script:IconPath) { try { $Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([uri]$script:IconPath) } catch { } }
$Brush = New-Object System.Windows.Media.BrushConverter
function Set-Led($led, [string]$state) {
    $color = switch ($state) { 'ok' { '#2EA043' } 'err' { '#D1242F' } 'busy' { '#E3A008' } default { '#B0B0B0' } }
    $ui[$led].Fill = $Brush.ConvertFromString($color)
    $step = @{ led365 = 'm365'; ledExch = 'exch'; ledAD = 'ad' }[$led]
    if ($step) { Set-Step $step $(switch ($state) { 'ok' { 'done' } 'err' { 'error' } 'busy' { 'busy' } default { 'todo' } }) }
}

# ---------------------------------------------------------------- step chips (1 Connect 365 > 2 Connect Exchange > ...)
$script:StepState = @{ m365 = 'todo'; exch = 'todo'; ad = 'todo'; load = 'todo' }
$script:StepUi = @{}
function New-StepChip([string]$Key, [int]$Num, [string]$Label) {
    $chip = New-Object System.Windows.Controls.Border
    $chip.CornerRadius = New-Object System.Windows.CornerRadius 15
    $chip.Padding = New-Object System.Windows.Thickness 6, 4, 14, 4
    $chip.BorderThickness = New-Object System.Windows.Thickness 1
    $content = New-Object System.Windows.Controls.StackPanel
    $content.Orientation = 'Horizontal'
    $circle = New-Object System.Windows.Controls.Border
    $circle.Width = 22; $circle.Height = 22
    $circle.CornerRadius = New-Object System.Windows.CornerRadius 11
    $circle.Margin = New-Object System.Windows.Thickness 0, 0, 8, 0
    $glyph = New-Object System.Windows.Controls.TextBlock
    $glyph.HorizontalAlignment = 'Center'; $glyph.VerticalAlignment = 'Center'; $glyph.FontWeight = 'Bold'; $glyph.FontSize = 12
    $glyph.Foreground = $Brush.ConvertFromString('#FFFFFF')
    $circle.Child = $glyph
    $caption = New-Object System.Windows.Controls.TextBlock
    $caption.Text = $Label; $caption.VerticalAlignment = 'Center'; $caption.FontWeight = 'SemiBold'
    [void]$content.Children.Add($circle); [void]$content.Children.Add($caption)
    $chip.Child = $content
    $script:StepUi[$Key] = @{ Border = $chip; Circle = $circle; Glyph = $glyph; Text = $caption; Num = $Num }
    $chip
}
function Set-Step([string]$Key, [string]$State) {
    $script:StepState[$Key] = $State
    $chipUi = $script:StepUi[$Key]; if (-not $chipUi) { return }
    # background, border, circle, glyph, text colour
    $colors = switch ($State) {
        'done' { @('#E6F6EA', '#9AD9AE', '#2EA043', [string][char]0x2713, '#1E6B34') }
        'busy' { @('#FFF6DD', '#F0CF7A', '#E3A008', [string][char]0x2026, '#7A5200') }
        'error' { @('#FDECEC', '#F2A9A5', '#D1242F', [string][char]0x2715, '#9F1D1D') }
        default { @('#FFFFFF', '#D3DAE3', '#9AA6B6', "$($chipUi.Num)", '#5B6B80') }
    }
    $chipUi.Border.Background = $Brush.ConvertFromString($colors[0])
    $chipUi.Border.BorderBrush = $Brush.ConvertFromString($colors[1])
    $chipUi.Circle.Background = $Brush.ConvertFromString($colors[2])
    $chipUi.Glyph.Text = $colors[3]
    $chipUi.Text.Foreground = $Brush.ConvertFromString($colors[4])
}
function Update-Steps($src) {
    $ui.stepPanel.Children.Clear()
    $script:StepUi = @{}
    $steps = @()
    if ($src.M365) { $steps += , @('m365', (T 'btn.connect365')) }
    if ($src.Exch) { $steps += , @('exch', (T 'btn.connectExch')) }
    if ($src.AD) { $steps += , @('ad', (T 'btn.verifyAD')) }
    $steps += , @('load', (T 'btn.load'))
    $i = 0
    foreach ($stepDef in $steps) {
        $i++
        if ($i -gt 1) {
            $sep = New-Object System.Windows.Controls.TextBlock
            $sep.Text = [string][char]0x203A; $sep.FontSize = 20; $sep.VerticalAlignment = 'Center'
            $sep.Margin = New-Object System.Windows.Thickness 10, -3, 10, 0
            $sep.Foreground = $Brush.ConvertFromString('#9AA6B6')
            [void]$ui.stepPanel.Children.Add($sep)
        }
        [void]$ui.stepPanel.Children.Add((New-StepChip $stepDef[0] $i $stepDef[1]))
        Set-Step $stepDef[0] $script:StepState[$stepDef[0]]
    }
}
function Show-Msg([string]$text, [string]$title, [string]$buttons = 'OK', [string]$icon = 'Information') {
    [System.Windows.MessageBox]::Show($text, $title, $buttons, $icon)
}

# ---------------------------------------------------------------- initial values
$ui.txtTenant.Text = "$($cfg.TenantId)"
$ui.txtExch.Text = "$($cfg.ExchangeServer)"
$ui.txtDomain.Text = "$($cfg.DomainFqdn)"
$ui.txtUser.Text = "$($cfg.DomainUser)"
$ui.cmbExAuth.SelectedIndex = [math]::Max(0, @('Kerberos', 'Negotiate', 'Basic').IndexOf("$($cfg.ExchangeAuth)"))
$ui.cmbExProto.SelectedIndex = [math]::Max(0, @('http', 'https', 'local').IndexOf("$($cfg.ExchangeProtocol)"))
$ui.cmbExAuth.IsEnabled = ($ui.cmbExProto.SelectedItem.Content -ne 'local')
$ui.cmbExProto.Add_SelectionChanged({ $ui.cmbExAuth.IsEnabled = ($ui.cmbExProto.SelectedItem.Content -ne 'local') })
$ui.cmbLang.SelectedIndex = if ($cfg.Language -eq 'it') { 1 } else { 0 }
$ui.chkWinCred.IsChecked = [bool]$cfg.UseWindowsCredentials

$script:InDomain = $true
try { $script:InDomain = [bool](Get-CimInstance Win32_ComputerSystem).PartOfDomain } catch { }
if (-not $script:InDomain) {
    $ui.chkWinCred.IsChecked = $false
    $ui.chkWinCred.ToolTip = T 'tip.notDomain'
}
function Update-CredFields {
    $manual = -not $ui.chkWinCred.IsChecked
    $ui.txtUser.IsEnabled = $manual
    $ui.pwdPass.IsEnabled = $manual
}
$ui.chkWinCred.Add_Click({ Update-CredFields })
Update-CredFields

# returns: $null = Windows credentials | PSCredential | $false = missing data (already reported)
function Get-DomainCredential {
    if ($ui.chkWinCred.IsChecked) {
        if (-not $script:InDomain) { Show-Msg (T 'msg.notDomain') (T 'title.credentials') 'OK' 'Warning' | Out-Null; return $false }
        return $null
    }
    $userName = $ui.txtUser.Text.Trim()
    if (-not $userName -or $ui.pwdPass.SecurePassword.Length -eq 0) { Show-Msg (T 'msg.credMissing') (T 'title.credentials') 'OK' 'Warning' | Out-Null; return $false }
    if ($userName -notmatch '[\\@]' -and $cfg.DomainFqdn) { $userName = "$userName@$($cfg.DomainFqdn)" }
    New-Object System.Management.Automation.PSCredential ($userName, $ui.pwdPass.SecurePassword.Copy())
}

function Update-CfgFromUI {
    $cfg.TenantId = $ui.txtTenant.Text.Trim()
    $cfg.ExchangeServer = $ui.txtExch.Text.Trim()
    $cfg.ExchangeAuth = $ui.cmbExAuth.SelectedItem.Content
    $cfg.ExchangeProtocol = $ui.cmbExProto.SelectedItem.Content
    $cfg.DomainFqdn = $ui.txtDomain.Text.Trim()
    $cfg.UseWindowsCredentials = [bool]$ui.chkWinCred.IsChecked
    $cfg.DomainUser = $ui.txtUser.Text.Trim()
    try { Save-HSConfig $cfg $ConfigPath } catch { }
}

$ui.cmbLang.Add_SelectionChanged({
        $lang = $ui.cmbLang.SelectedItem.Tag
        if ($lang -and $lang -ne $cfg.Language) {
            $cfg.Language = $lang
            try { Save-HSConfig $cfg $ConfigPath } catch { }
            $script:HS_Lang = $lang
            Show-Msg (T 'msg.restart') (T 'title.language') | Out-Null
        }
    })

# ---------------------------------------------------------------- background runspace (persistent)
$HS_Sync = [hashtable]::Synchronized(@{ Log = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]' })
$script:HS_Sync = $HS_Sync
$script:RS = [runspacefactory]::CreateRunspace()
$script:RS.ApartmentState = 'STA'
$script:RS.ThreadOptions = 'ReuseThread'
$script:RS.Open()
$script:RS.SessionStateProxy.SetVariable('HS_Sync', $HS_Sync)
$script:RS.SessionStateProxy.SetVariable('HS_AppDir', $AppDir)
$script:RS.SessionStateProxy.SetVariable('HS_Lang', $cfg.Language)
$CoreText = $Core.ToString()
$script:ADCred = $null; $script:ExCred = $null; $script:Data = $null; $script:Job = $null; $script:AfterInstall = $null

function Set-Busy([bool]$busy, [string]$text) {
    foreach ($control in 'btn365', 'btnSignOut', 'btnExch', 'btnAD', 'btnLoad', 'cmbGroupBy', 'rbCloud', 'rbFullHybrid', 'rbHybrid', 'rbOnPrem', 'rbAD', 'miUpdates') { $ui[$control].IsEnabled = -not $busy }
    $ui.btnStop.IsEnabled = $busy
    foreach ($control in 'btnXlsx', 'btnCsvAll', 'btnTab') { $ui[$control].IsEnabled = (-not $busy) -and ($null -ne $script:Data) }
    $ui.prg.IsIndeterminate = $busy
    if ($text) { $ui.lblStatus.Text = $text }
    $Window.Cursor = if ($busy) { [System.Windows.Input.Cursors]::AppStarting } else { $null }
    if (-not $busy -and (Get-Command Update-HSExportLabel -ErrorAction SilentlyContinue)) { Update-HSExportLabel }
}

function Start-HSJob {
    param([string]$Title, [string]$Script, [hashtable]$Vars = @{}, [scriptblock]$OnDone, [string]$Led, [string]$Label, [string]$Step)
    if ($script:Job) { return }
    $script:RS.SessionStateProxy.SetVariable('Cfg', $cfg)
    foreach ($varName in $Vars.Keys) { $script:RS.SessionStateProxy.SetVariable($varName, $Vars[$varName]) }
    $ps = [powershell]::Create()
    $ps.Runspace = $script:RS
    [void]$ps.AddScript("`$ErrorActionPreference = 'Stop'`n`$script:HS_Sync = `$HS_Sync`n`$script:HS_AppDir = `$HS_AppDir`n`$script:HS_Lang = `$HS_Lang`n" + $Script)
    $script:Job = @{ PS = $ps; Handle = $ps.BeginInvoke(); OnDone = $OnDone; Led = $Led; Label = $Label; Step = $Step; Cancelled = $false }
    if ($Step) { Set-Step $Step 'busy' }
    Set-Busy $true $Title
}

function Install-MissingModules {
    param([string[]]$Missing, [scriptblock]$AfterOk)
    $gallery = @($Missing | Where-Object { $_ -ne 'ActiveDirectory' })
    $ui.expLog.IsExpanded = $true
    $script:AfterInstall = $AfterOk
    Start-HSJob -Title (T 'job.modules') -Vars @{ ToInstall = $gallery; NeedAD = ($Missing -contains 'ActiveDirectory') } -Script @'
if ($ToInstall.Count) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -Scope CurrentUser -Force | Out-Null }
    foreach ($moduleName in $ToInstall) { Write-HSLog (T 'log.installing' $moduleName); Install-Module $moduleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery }
}
if ($NeedAD) { Install-HSRsatAD }
Write-HSLog (T 'log.modulesDone')
'@ -OnDone {
        param($out, $err)
        if ($err) { Show-Msg (T 'msg.modulesFail' $err.Message) (T 'title.error') 'OK' 'Error' | Out-Null }
        else {
            Show-Msg (T 'msg.modulesOk') (T 'title.modules') | Out-Null
            if ($script:AfterInstall) { $cb = $script:AfterInstall; $script:AfterInstall = $null; & $cb }
        }
    }
}

# first job: load the engine into the runspace, then check modules
Start-HSJob -Title (T 'job.init') -Script ($CoreText + "`nWrite-HSLog (T 'log.ready' '$HS_Version')") -OnDone {
    param($out, $err)
    $missing = @(Test-HSModules)
    if (-not (Get-HSSources $script:CurrentMode).AD) { $missing = @($missing | Where-Object { $_ -ne 'ActiveDirectory' }) }
    if ($missing.Count) {
        $msg = T 'msg.modulesMissing' ($missing -join "`n - ")
        if ($missing -contains 'ActiveDirectory') { $msg += "`n`n" + (T 'msg.rsatNote') }
        if ((Show-Msg $msg (T 'title.modules') 'YesNo' 'Warning') -eq 'Yes') { Install-MissingModules $missing }
    }
}

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$timer.Add_Tick({
        $line = $null; $got = $false
        while ($HS_Sync.Log.TryDequeue([ref]$line)) { $ui.txtLog.AppendText($line + "`r`n"); $ui.lblStatus.Text = $line; $got = $true }
        if ($got) { $ui.txtLog.ScrollToEnd() }
        if ($script:Job -and $script:Job.Handle.IsCompleted) {
            $finishedJob = $script:Job; $script:Job = $null
            $out = $null; $err = $null
            try { $out = $finishedJob.PS.EndInvoke($finishedJob.Handle) }
            catch { $err = if ($_.Exception.InnerException) { $_.Exception.InnerException } else { $_.Exception } }
            $finishedJob.PS.Dispose()
            Set-Busy $false $null
            if ($finishedJob.Cancelled) {
                $ui.txtLog.AppendText((T 'st.cancelled') + "`r`n"); $ui.lblStatus.Text = T 'st.cancelled'
                if ($finishedJob.Led) { Set-Led $finishedJob.Led 'off'; $ui[$finishedJob.Label].Text = T 'st.cancelled' }
                if ($finishedJob.Step) { Set-Step $finishedJob.Step 'todo' }
                return
            }
            if ($err) { $ui.txtLog.AppendText((T 'err.generic' $err.Message) + "`r`n"); $ui.lblStatus.Text = (T 'err.generic' $err.Message) }
            if ($finishedJob.OnDone) {
                try { & $finishedJob.OnDone $out $err }
                catch { Show-Msg $_.Exception.Message (T 'title.error') 'OK' 'Error' | Out-Null }
            }
        }
    })
$timer.Start()

# ---------------------------------------------------------------- tabs
function Set-HSRowFilter([System.Data.DataTable]$Table, [string]$Text) {
    # DataView filter: rows where any column contains Text
    $dv = $Table.DefaultView
    $Text = "$Text".Trim()
    if (-not $Text) { $dv.RowFilter = ''; return }
    $pattern = $Text.Replace("'", "''")
    $pattern = [regex]::Replace($pattern, '[\[\]\*%]', [System.Text.RegularExpressions.MatchEvaluator] { param($match) "[$($match.Value)]" })
    $parts = foreach ($column in $Table.Columns) { "CONVERT([$($column.ColumnName.Replace(']','\]'))], 'System.String') LIKE '%$pattern%'" }
    try { $dv.RowFilter = ($parts -join ' OR ') } catch { }
}

function Set-HSFilter {
    $tab = $ui.tabs.SelectedItem
    if (-not $tab -or -not $tab.Tag) { $ui.lblCount.Text = ''; return }
    Set-HSRowFilter $tab.Tag $ui.txtFilter.Text
    $ui.lblCount.Text = T 'lbl.rows' @($tab.Tag.DefaultView.Count, $tab.Tag.Rows.Count)
}

function Show-HSTables {
    param($Tables, [string]$OverviewName)
    $ui.tabs.Items.Clear()
    foreach ($name in $Tables.Keys) {
        $dt = $Tables[$name]
        $grid = New-Object System.Windows.Controls.DataGrid
        $grid.IsReadOnly = $true
        $grid.AutoGenerateColumns = $true
        $grid.CanUserAddRows = $false
        $grid.HeadersVisibility = 'Column'
        $grid.GridLinesVisibility = 'Horizontal'
        $grid.HorizontalGridLinesBrush = $Brush.ConvertFromString('#E6EAF0')
        $grid.AlternatingRowBackground = $Brush.ConvertFromString('#F5F8FC')
        $grid.SelectionUnit = 'CellOrRowHeader'
        $grid.ClipboardCopyMode = 'IncludeHeader'
        $grid.Add_AutoGeneratingColumn({
                param($source, $evtArgs)
                $hdr = New-Object System.Windows.Controls.TextBlock   # TextBlock: no access-key handling of underscores
                $hdr.Text = $evtArgs.PropertyName
                $evtArgs.Column.Header = $hdr
                if ($evtArgs.PropertyType -eq [datetime]) { $evtArgs.Column.Binding.StringFormat = 'yyyy-MM-dd HH:mm' }
                elseif ($evtArgs.PropertyType -eq [double]) {
                    $evtArgs.Column.Binding.StringFormat = if ($evtArgs.PropertyName -match '_MB$|_GB$|\(GB\)') { 'N2' } else { '#,##0.##' }
                    $alignStyle = New-Object System.Windows.Style ([System.Windows.Controls.TextBlock])
                    $alignStyle.Setters.Add((New-Object System.Windows.Setter ([System.Windows.Controls.TextBlock]::TextAlignmentProperty), ([System.Windows.TextAlignment]::Right)))
                    $evtArgs.Column.ElementStyle = $alignStyle
                }
            })
        if ($name -eq $OverviewName) {
            $grid.Add_LoadingRow({
                    param($source, $evtArgs)
                    $row = $evtArgs.Row.Item
                    $colCount = $row.Row.Table.Columns.Count
                    if ($row -and $colCount -gt 1 -and $row[$colCount - 1] -is [DBNull]) { $evtArgs.Row.FontWeight = 'Bold'; $evtArgs.Row.Background = $Brush.ConvertFromString('#DCE6F2') }
                    else { $evtArgs.Row.FontWeight = 'Normal'; $evtArgs.Row.Background = $null }
                })
        }
        if (@($OverviewName, (T 'tab.licenses'), (T 'tab.history')) -notcontains $name) {
            # double-click on a user row: user card
            $grid.Add_MouseDoubleClick({
                    param($source, $evtArgs)
                    $item = $null
                    try { $item = $evtArgs.OriginalSource.DataContext } catch { }
                    if ($item -is [System.Data.DataRowView]) { Show-HSUserCard $item }
                })
        }
        $grid.ItemsSource = $dt.DefaultView
        $tab = New-Object System.Windows.Controls.TabItem
        $hdr = New-Object System.Windows.Controls.TextBlock   # TextBlock: underscores are not eaten as access keys
        $hdr.Text = if ($name -eq $OverviewName) { $name } else { "$name ($($dt.Rows.Count))" }
        $tab.Header = $hdr
        $tab.Uid = $name
        $tab.Content = $grid
        $tab.Tag = $dt
        [void]$ui.tabs.Items.Add($tab)
    }
    $ui.lblEmpty.Visibility = 'Collapsed'
    $ui.tabs.SelectedIndex = 0
    Set-HSFilter
}

# ---------------------------------------------------------------- user card (double-click on a user)
$script:CardXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Height="760" Width="1240" WindowStartupLocation="CenterOwner" Background="#F3F4F6" FontFamily="Segoe UI" FontSize="12">
  <DockPanel>
    <Border DockPanel.Dock="Top" Background="#0E2A47" Padding="16,10">
      <DockPanel>
        <Button x:Name="btnCopy" Content="{{card.copy}}" DockPanel.Dock="Right" VerticalAlignment="Center" Padding="14,4" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
        <StackPanel VerticalAlignment="Center">
          <TextBlock x:Name="lblName" Foreground="White" FontSize="20" FontWeight="SemiBold"/>
          <TextBlock x:Name="lblSub" Foreground="#9FB4D9" TextTrimming="CharacterEllipsis"/>
        </StackPanel>
      </DockPanel>
    </Border>
    <Grid Margin="10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="460"/>
        <ColumnDefinition Width="10"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border Background="White" BorderBrush="#DDE3EA" BorderThickness="1">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel x:Name="pnlInfo" Margin="14,2,14,14"/>
        </ScrollViewer>
      </Border>
      <Border Grid.Column="2" Background="White" BorderBrush="#DDE3EA" BorderThickness="1" Padding="12,10">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <Button x:Name="btnReload" Content="{{card.reload}}" DockPanel.Dock="Right" Padding="12,3" Margin="8,0,0,0" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
            <TextBox x:Name="txtGrpFilter" Width="200" Height="24" VerticalContentAlignment="Center" DockPanel.Dock="Right"/>
            <TextBlock Text="{{card.search}}" DockPanel.Dock="Right" VerticalAlignment="Center" Margin="0,0,6,0" Foreground="#555"/>
            <TextBlock Text="{{card.groups}}" FontSize="14" FontWeight="SemiBold" Foreground="#0E2A47" VerticalAlignment="Center"/>
          </DockPanel>
          <TextBlock x:Name="lblGrp" DockPanel.Dock="Top" Foreground="#555" TextWrapping="Wrap" Margin="0,0,0,6"/>
          <DataGrid x:Name="gridGrp" IsReadOnly="True" AutoGenerateColumns="True" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                    CanUserAddRows="False" AlternatingRowBackground="#F5F8FC" HorizontalGridLinesBrush="#E6EAF0" ClipboardCopyMode="IncludeHeader"/>
        </DockPanel>
      </Border>
    </Grid>
  </DockPanel>
</Window>
'@
$script:CardTarget = $null

function Format-HSCardValue($Value) {
    if ($null -eq $Value -or $Value -is [DBNull]) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    if ($Value -is [double]) { return $Value.ToString('N2') }
    "$Value".Trim()
}

function Add-HSCardSection {
    # Title + label / value grid; values are read-only text boxes, so they can be selected and copied.
    # Pairs: ordered hashtable label -> value. Empty values are left out; nothing is shown if all are empty.
    param($Card, [string]$Title, $Pairs)
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($label in $Pairs.Keys) { $text = Format-HSCardValue $Pairs[$label]; if ($text) { $items.Add([pscustomobject]@{ Label = $label; Text = $text }) } }
    if (-not $items.Count) { return }

    $header = New-Object System.Windows.Controls.TextBlock
    $header.Text = $Title; $header.FontSize = 13; $header.FontWeight = 'SemiBold'
    $header.Foreground = $Brush.ConvertFromString('#1565C0')
    $header.Margin = New-Object System.Windows.Thickness 0, 12, 0, 4
    [void]$Card.Info.Children.Add($header)

    $grid = New-Object System.Windows.Controls.Grid
    $col0 = New-Object System.Windows.Controls.ColumnDefinition; $col0.Width = New-Object System.Windows.GridLength 165
    $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = New-Object System.Windows.GridLength 1, ([System.Windows.GridUnitType]::Star)
    $grid.ColumnDefinitions.Add($col0); $grid.ColumnDefinitions.Add($col1)
    $alert = @((T 'val.mfaNo'), (T 'val.disabled'), 'ON-PREM')
    [void]$Card.Text.AppendLine($Title)
    $rowIndex = 0
    foreach ($item in $items) {
        $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $item.Label; $label.Foreground = $Brush.ConvertFromString('#5B6B80'); $label.TextWrapping = 'Wrap'
        $label.VerticalAlignment = 'Top'; $label.Margin = New-Object System.Windows.Thickness 0, 3, 8, 3
        $value = New-Object System.Windows.Controls.TextBox
        $value.Text = $item.Text; $value.IsReadOnly = $true; $value.TextWrapping = 'Wrap'
        $value.BorderThickness = New-Object System.Windows.Thickness 0; $value.Background = [System.Windows.Media.Brushes]::Transparent
        $value.Padding = New-Object System.Windows.Thickness 0; $value.Margin = New-Object System.Windows.Thickness 0, 3, 0, 3
        if ($alert -contains $item.Text -or $item.Label -eq (Get-HSHeader 'SyncIssues')) { $value.Foreground = $Brush.ConvertFromString('#B42318'); $value.FontWeight = 'SemiBold' }
        [System.Windows.Controls.Grid]::SetRow($label, $rowIndex); [System.Windows.Controls.Grid]::SetRow($value, $rowIndex)
        [System.Windows.Controls.Grid]::SetColumn($value, 1)
        [void]$grid.Children.Add($label); [void]$grid.Children.Add($value)
        [void]$Card.Text.AppendLine(('  {0}: {1}' -f $item.Label, ($item.Text -replace "`r?`n", "`r`n      ")))
        $rowIndex++
    }
    [void]$Card.Info.Children.Add($grid)
}

function Set-HSCardFilter($Card) {
    if ($Card -and $Card.Table) { Set-HSRowFilter $Card.Table $Card.Filter.Text }
}

function Start-HSCardGroups($Card) {
    # groups and roles are read on demand, in the background runspace (AD + Graph sessions live there)
    if (-not $Card.Sam -and -not $Card.Upn) { $Card.Status.Text = ''; return }
    if ($script:Data.Imported) { $Card.Status.Text = T 'card.imported'; return }
    if ($script:Job) { $Card.Status.Text = T 'card.busy'; return }
    $Card.Status.Text = T 'card.loading'
    $script:CardTarget = $Card
    Start-HSJob -Title (T 'card.loading') -Vars @{ CardSam = $Card.Sam; CardUpn = $Card.Upn; ADCred = $script:ADCred } -Script '$HS_Sync.Card = Get-HSUserGroup -Cfg $Cfg -Sam $CardSam -Upn $CardUpn -ADCredential $ADCred' -OnDone {
        param($out, $err)
        $card = $script:CardTarget; $script:CardTarget = $null
        if (-not $card -or -not $card.Window.IsLoaded) { return }
        if ($err) { $card.Status.Text = T 'err.generic' $err.Message; return }
        $result = $HS_Sync.Card
        $card.Table = ConvertTo-HSDataTable -Rows @($result.Rows)
        $card.Grid.ItemsSource = $card.Table.DefaultView
        Set-HSCardFilter $card
        $card.Status.Text = (@(T 'card.count' @($result.AD, $result.Entra, $result.Roles)) + @($result.Errors)) -join "`n"
    }
}

function Show-HSUserCard($RowView) {
    <#
        One window with everything HybridScope knows about the user of the double-clicked row:
        identity, AD, Microsoft 365, mailbox, OneDrive, MFA, sync; groups and roles load in the background.
    #>
    if (-not $script:Data -or -not $script:Data.Users) { return }
    $table = $RowView.Row.Table
    $cell = { param($prop) $hdr = Get-HSHeader $prop; if ($table.Columns.Contains($hdr) -and -not ($RowView[$hdr] -is [DBNull])) { "$($RowView[$hdr])" } else { '' } }
    $sam = & $cell 'Username'; $upn = & $cell 'UPN'
    if (-not $sam -and -not $upn) { return }
    $user = $null
    if ($upn) { $user = @($script:Data.Users | Where-Object { $_.UPN -eq $upn })[0] }
    if (-not $user -and $sam) { $user = @($script:Data.Users | Where-Object { $_.Username -eq $sam })[0] }
    if (-not $user) { Show-Msg (T 'card.notFound') 'HybridScope' | Out-Null; return }

    $lookup = {
        param([string]$TabKey, [string]$Prop, [string]$Value)
        if (-not $Value) { return $null }
        $rows = $script:Data.Tables[(T $TabKey)]
        if (-not $rows) { return $null }
        @($rows | Where-Object { "$($_.$Prop)" -eq $Value })[0]
    }
    $src = $script:Data.SourceFlags
    $mfaRow = & $lookup 'tab.mfa' 'UPN' $user.UPN
    $onPremRow = & $lookup 'tab.mbxOnPrem' 'Username' $(if ($user.OU) { $user.Username })
    $exoRow = & $lookup 'tab.mbx365' 'UPN' $user.UPN
    $odRow = & $lookup 'tab.onedrive' 'UPN' $user.UPN
    $syncInfo = if ($script:Data.SyncInfo) { $script:Data.SyncInfo["$($user.Username)|$($user.UPN)"] } else { $null }
    $syncRows = @()
    if ($script:Data.Tables.Contains((T 'tab.sync'))) {
        $syncRows = @($script:Data.Tables[(T 'tab.sync')] | Where-Object { $_.Category -ne (T 'sync.c0') -and (($user.OU -and $_.Username -eq $user.Username) -or ($user.UPN -and $_.UPN -eq $user.UPN)) })
    }

    $xamlText = [regex]::Replace($script:CardXaml, '\{\{([\w\.]+)\}\}', [System.Text.RegularExpressions.MatchEvaluator] {
            param($match) [System.Security.SecurityElement]::Escape((T $match.Groups[1].Value))
        })
    [xml]$cardXml = $xamlText
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $cardXml))
    $card = @{
        Window = $win; Info = $win.FindName('pnlInfo'); Grid = $win.FindName('gridGrp'); Status = $win.FindName('lblGrp')
        Filter = $win.FindName('txtGrpFilter'); Text = New-Object System.Text.StringBuilder; Table = $null
        Sam = if ($src.AD -and $user.OU) { "$($user.Username)" } else { '' }
        Upn = if ($src.M365 -and $user.UPN) { "$($user.UPN)" } else { '' }
    }
    $win.Tag = $card
    $win.Owner = $Window
    if ($Window.Icon) { $win.Icon = $Window.Icon }

    $displayName = ("$($user.FirstName) $($user.LastName)").Trim(); if (-not $displayName) { $displayName = if ($user.UPN) { $user.UPN } else { $user.Username } }
    $win.Title = "$displayName - HybridScope"
    $win.FindName('lblName').Text = $displayName
    $sub = (@($user.Username, $user.UPN, $user.Status, $user.Origin, $user.Group) | Where-Object { $_ } | Select-Object -Unique) -join '  ·  '
    $win.FindName('lblSub').Text = $sub
    [void]$card.Text.AppendLine($displayName); [void]$card.Text.AppendLine($sub)

    $h = { param($prop) Get-HSHeader $prop }
    Add-HSCardSection $card (T 'card.sec.identity') ([ordered]@{
            (& $h 'Username') = $user.Username; (& $h 'UPN') = $user.UPN; (& $h 'Status') = $user.Status
            (& $h 'Origin') = $user.Origin; (& $h 'Group') = $user.Group; 'OU' = $user.OU
        })
    if ($src.AD -and $user.OU) {
        Add-HSCardSection $card (T 'card.sec.ad') ([ordered]@{
                (& $h 'AD_LastLogon') = $user.AD_LastLogon; (& $h 'PasswordLastSet') = $user.PasswordLastSet
                'AD_UPN' = $(if ($syncInfo) { $syncInfo.AD_UPN })
            })
    }
    if ($src.M365) {
        Add-HSCardSection $card (T 'card.sec.m365') ([ordered]@{
                (& $h 'M365_LastSignIn') = $user.M365_LastSignIn
                (& $h 'Licenses') = ("$($user.Licenses)" -replace ', ', "`n")
            })
    }
    Add-HSCardSection $card (T 'card.sec.mailbox') ([ordered]@{
            'Location' = $user.Location; (& $h 'MailboxType') = $user.MailboxType; (& $h 'MailboxSource') = $user.MailboxSource
            (& $h 'Mailbox365_MB') = $(if ($user.Mailbox365_MB) { $user.Mailbox365_MB }); (& $h 'MailboxOnPrem_MB') = $(if ($user.MailboxOnPrem_MB) { $user.MailboxOnPrem_MB })
        })
    if ($onPremRow) {
        Add-HSCardSection $card (T 'gb.exch') ([ordered]@{
                'Email' = $onPremRow.Email; 'Database' = $onPremRow.Database; (& $h 'Items') = $onPremRow.Items
                (& $h 'Deleted_MB') = $onPremRow.Deleted_MB; (& $h 'LastLogon') = $onPremRow.LastLogon
                (& $h 'Quota') = $onPremRow.Quota; (& $h 'HiddenFromGAL') = $onPremRow.HiddenFromGAL
            })
    }
    if ($exoRow) {
        Add-HSCardSection $card 'Exchange Online' ([ordered]@{
                (& $h 'Items') = $exoRow.Items; (& $h 'Deleted_MB') = $exoRow.Deleted_MB; (& $h 'Archive') = $exoRow.Archive
                (& $h 'LastActivity') = $exoRow.LastActivity; (& $h 'WarningQuota_GB') = $exoRow.WarningQuota_GB; (& $h 'SendQuota_GB') = $exoRow.SendQuota_GB
            })
    }
    if ($odRow) {
        Add-HSCardSection $card (T 'card.sec.onedrive') ([ordered]@{
                (& $h 'OneDrive_MB') = $odRow.Size_MB; 'Quota_GB' = $odRow.Quota_GB; (& $h 'Files') = $odRow.Files
                (& $h 'ActiveFiles') = $odRow.ActiveFiles; (& $h 'LastActivity') = $odRow.LastActivity; 'URL' = $odRow.URL
            })
    }
    if ($src.M365) {
        Add-HSCardSection $card (T 'card.sec.mfa') ([ordered]@{
                (& $h 'MFA_Status') = $user.MFA_Status; (& $h 'MFA_DefaultMethod') = $user.MFA_DefaultMethod
                (& $h 'MFA_Methods') = ("$($user.MFA_Methods)" -replace ', ', "`n"); (& $h 'MFA_PerUser') = $user.MFA_PerUser
                (& $h 'MFA_Capable') = $(if ($mfaRow) { $mfaRow.MFA_Capable }); (& $h 'SSPR_Registered') = $(if ($mfaRow) { $mfaRow.SSPR_Registered })
                'Admin' = $(if ($mfaRow) { $mfaRow.Admin })
            })
    }
    if ($syncInfo -or $syncRows.Count) {
        Add-HSCardSection $card (T 'card.sec.sync') ([ordered]@{
                (& $h 'SyncIssues') = $user.SyncIssues
                (& $h 'Details') = (@($syncRows | ForEach-Object { $_.Details }) -join "`n")
                (& $h 'Anchor_AD') = $(if ($syncInfo) { $syncInfo.Anchor_AD }); (& $h 'AnchorSource') = $(if ($syncInfo) { $syncInfo.AnchorSource })
                'ImmutableId' = $(if ($syncInfo) { $syncInfo.ImmutableId }); (& $h 'SyncEnabled') = $(if ($syncInfo) { $syncInfo.SyncEnabled })
                (& $h 'LastSync') = $(if ($syncInfo) { $syncInfo.LastSync })
            })
    }
    Add-HSCardSection $card (T 'card.sec.notes') ([ordered]@{ (& $h 'Notes') = ("$($user.Notes)" -replace '; ', "`n") })

    $win.FindName('txtGrpFilter').Add_TextChanged({ param($source, $evtArgs) Set-HSCardFilter ([System.Windows.Window]::GetWindow($source)).Tag })
    $win.FindName('btnReload').Add_Click({ param($source, $evtArgs) Start-HSCardGroups ([System.Windows.Window]::GetWindow($source)).Tag })
    $win.FindName('btnCopy').Add_Click({
            param($source, $evtArgs)
            $c = ([System.Windows.Window]::GetWindow($source)).Tag
            $text = $c.Text.ToString()
            if ($c.Table -and $c.Table.Rows.Count) {
                $text += "`r`n" + (T 'card.groups') + "`r`n"
                foreach ($groupRow in $c.Table.DefaultView) { $text += '  ' + (@($groupRow.Row.ItemArray | ForEach-Object { "$_" }) -join ' | ') + "`r`n" }
            }
            try { [System.Windows.Clipboard]::SetText($text) } catch { }
        })
    $win.Add_KeyDown({ param($source, $evtArgs) if ($evtArgs.Key -eq 'Escape') { $source.Close() } })
    $win.Show()
    Start-HSCardGroups $card
}

$ui.tabs.Add_SelectionChanged({ param($source, $evtArgs) if ($evtArgs.OriginalSource -eq $ui.tabs) { Set-HSFilter } })
$ui.txtFilter.Add_TextChanged({ Set-HSFilter })

# ---------------------------------------------------------------- connection buttons
$ui.btn365.Add_Click({
        Update-CfgFromUI
        Set-Led 'led365' 'busy'; $ui.lbl365.Text = T 'st.login'
        Start-HSJob -Title (T 'job.365') -Led 'led365' -Label 'lbl365' -Script 'Connect-HS365 -Cfg $Cfg' -OnDone {
            param($out, $err)
            if ($err) { Set-Led 'led365' 'err'; $ui.lbl365.Text = T 'st.error'; Show-Msg $err.Message 'Microsoft 365' 'OK' 'Error' | Out-Null }
            else { Set-Led 'led365' 'ok'; $ui.lbl365.Text = "$(@($out)[-1])" }
        }
    })

$ui.btnExch.Add_Click({
        Update-CfgFromUI
        if ($cfg.ExchangeProtocol -ne 'local' -and -not $cfg.ExchangeServer) { Show-Msg (T 'msg.exchNoServer') (T 'gb.exch') 'OK' 'Warning' | Out-Null; return }
        $script:ExCred = $null
        if ($cfg.ExchangeProtocol -ne 'local') {
            $cred = Get-DomainCredential
            if ($cred -is [bool]) { return }
            if (-not $cred -and $cfg.ExchangeAuth -eq 'Basic') { Show-Msg (T 'msg.basicNeedsCred') (T 'gb.exch') 'OK' 'Warning' | Out-Null; return }
            $script:ExCred = $cred
        }
        Set-Led 'ledExch' 'busy'; $ui.lblExch.Text = T 'st.connecting'
        Start-HSJob -Title (T 'job.exch') -Led 'ledExch' -Label 'lblExch' -Vars @{ ExCred = $script:ExCred } -Script 'Connect-HSExchange -Cfg $Cfg -Credential $ExCred' -OnDone {
            param($out, $err)
            if ($err) { Set-Led 'ledExch' 'err'; $ui.lblExch.Text = T 'st.error'; Show-Msg (T 'msg.exchHints' $err.Message) (T 'gb.exch') 'OK' 'Error' | Out-Null }
            else { Set-Led 'ledExch' 'ok'; $ui.lblExch.Text = "$(@($out)[-1])" }
        }
    })

$ui.btnAD.Add_Click({
        Update-CfgFromUI
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            if ((Show-Msg (T 'msg.adMissing') (T 'gb.ad') 'YesNo' 'Question') -eq 'Yes') {
                Install-MissingModules @('ActiveDirectory') { $ui.btnAD.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent))) }
            }
            return
        }
        $cred = Get-DomainCredential
        if ($cred -is [bool]) { return }
        $script:ADCred = $cred
        Set-Led 'ledAD' 'busy'; $ui.lblAD.Text = T 'st.verifying'
        Start-HSJob -Title (T 'job.ad') -Led 'ledAD' -Label 'lblAD' -Vars @{ ADCred = $script:ADCred } -Script 'Test-HSAD -Cfg $Cfg -Credential $ADCred' -OnDone {
            param($out, $err)
            if ($err) { Set-Led 'ledAD' 'err'; $ui.lblAD.Text = T 'st.error'; Show-Msg $err.Message (T 'gb.ad') 'OK' 'Error' | Out-Null }
            else { Set-Led 'ledAD' 'ok'; $ui.lblAD.Text = "$(@($out)[-1])" }
        }
    })

$ui.btnSignOut.Add_Click({
        Start-HSJob -Title (T 'job.signout') -Script 'Clear-HS365Session' -OnDone {
            param($out, $err)
            Set-Led 'led365' 'off'; $ui.lbl365.Text = T 'st.notconnected'
        }
    })

$ui.btnStop.Add_Click({
        if ($script:Job -and -not $script:Job.Cancelled) {
            $script:Job.Cancelled = $true
            $ui.btnStop.IsEnabled = $false
            $ui.lblStatus.Text = T 'st.stopping'
            [void]$script:Job.PS.BeginStop($null, $null)
        }
    })

# ---------------------------------------------------------------- infrastructure mode (not saved: one click to change)
$script:CurrentMode = if ($Mode) { $Mode } elseif ("$($cfg.Mode)") { "$($cfg.Mode)" } else { 'FullHybrid' }   # -Mode: starting mode for this session only
function Set-HSMode([string]$Mode) {
    $script:CurrentMode = $Mode
    $src = Get-HSSources $Mode
    $ui.gbM365.Visibility = if ($src.M365) { 'Visible' } else { 'Collapsed' }
    $ui.gbExch.Visibility = if ($src.Exch) { 'Visible' } else { 'Collapsed' }
    $ui.gbAD.Visibility = if ($src.AD) { 'Visible' } else { 'Collapsed' }
    $ui.gbCred.Visibility = if ($src.Exch -or $src.AD) { 'Visible' } else { 'Collapsed' }
    Update-Steps $src
}
$modeButtons = @{ Cloud = 'rbCloud'; FullHybrid = 'rbFullHybrid'; Hybrid = 'rbHybrid'; OnPrem = 'rbOnPrem'; AD = 'rbAD' }
if (-not $modeButtons.ContainsKey($script:CurrentMode)) { $script:CurrentMode = 'FullHybrid' }
$ui[$modeButtons[$script:CurrentMode]].IsChecked = $true
Set-HSMode $script:CurrentMode
foreach ($rb in $modeButtons.Values) {
    $ui[$rb].Add_Checked({
            param($modeButton, $evtArgs)
            if ($modeButton.Tag -eq $script:CurrentMode) { return }
            Set-HSMode $modeButton.Tag
            if ($script:Data -and -not $script:Job) {
                if ((Show-Msg (T 'msg.modeReload') (T 'title.mode') 'YesNo' 'Question') -eq 'Yes') {
                    $ui.btnLoad.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent)))
                }
            }
        })
}

# group-by selector
$grpTags = @('TopLevelOU', 'ParentOU', 'None', 'Rules')
$ui.cmbGroupBy.SelectedIndex = [math]::Max(0, $grpTags.IndexOf("$($cfg.GroupBy)"))
$ui.cmbGroupBy.Add_SelectionChanged({
        $tag = $ui.cmbGroupBy.SelectedItem.Tag
        if (-not $tag -or $tag -eq $cfg.GroupBy) { return }
        if ($tag -eq 'Rules' -and -not @($cfg.GroupRules).Count) { Show-Msg (T 'msg.noRules') (T 'lbl.groupBy') 'OK' 'Information' | Out-Null }
        $cfg.GroupBy = $tag
        try { Save-HSConfig $cfg $ConfigPath } catch { }
        if ($script:Data -and -not $script:Job) {
            if ((Show-Msg (T 'msg.regroup') (T 'lbl.groupBy') 'YesNo' 'Question') -eq 'Yes') {
                $ui.btnLoad.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent)))
            }
        }
    })

# ---------------------------------------------------------------- load data
$ui.btnLoad.Add_Click({
        Update-CfgFromUI
        $src = Get-HSSources $script:CurrentMode
        $script:ADCred = $null; $script:ExCred = $null
        if ($src.Exch -or $src.AD) {
            $cred = Get-DomainCredential
            if ($cred -is [bool]) { return }
            $script:ADCred = $cred
            if ($cfg.ExchangeProtocol -ne 'local') { $script:ExCred = $cred }
        }
        # dynamic load: with a preset / Custom and "Read only the data needed", only what the export columns need
        $choice = Get-HSExportChoice
        $parts = if ($choice.Preset -ne 'Full' -and $ui.chkExpFast.IsChecked) { Get-HSLoadParts -Columns $choice.Columns } else { $null }
        $ui.expLog.IsExpanded = $true
        Start-HSJob -Title (T 'job.load') -Step 'load' -Vars @{ ADCred = $script:ADCred; ExCred = $script:ExCred; Sources = $src; Parts = $parts; PartialLabel = (Get-HSExportLabel $choice.Preset) } -Script @'
$data = Get-HSData -Cfg $Cfg -ADCredential $ADCred -ExchangeCredential $ExCred -Sources $Sources -Parts $Parts -PartialLabel $PartialLabel
$tables = [ordered]@{}
foreach ($tabName in $data.Tables.Keys) { $tables[$tabName] = ConvertTo-HSDataTable -Rows @($data.Tables[$tabName]) -NoLocalize:($tabName -eq $data.OverviewName) }
$HS_Sync.Data = $data
$HS_Sync.DataTables = $tables
'@ -OnDone {
            param($out, $err)
            if ($err) { Set-Step 'load' 'error'; Show-Msg $err.Message (T 'title.loadError') 'OK' 'Error' | Out-Null; return }
            Set-Step 'load' 'done'
            $script:Data = $HS_Sync.Data
            Show-HSTables $HS_Sync.DataTables $script:Data.OverviewName
            $ui.lblGenerated.Text = T 'lbl.generated' $script:Data.Generated.ToString('yyyy-MM-dd HH:mm')
            $ui.lblGenerated.ToolTip = $null
            $ui.lblGenerated.Foreground = $Brush.ConvertFromString('#9FB4D9')
            if ($script:Data.Partial) {
                $ui.lblGenerated.Text += '   ' + (T 'lbl.partial' $script:Data.Partial)
                $ui.lblGenerated.ToolTip = T 'tip.partial' (@($script:Data.Skipped) -join ', ')
                $ui.lblGenerated.Foreground = $Brush.ConvertFromString('#F7C948')
            }
            Set-Busy $false (T 'st.loaded' $script:Data.Totals['total'])
            if ($script:Data.Warnings.Count) { Show-Msg (T 'msg.loadWarn' ($script:Data.Warnings -join "`n- ")) (T 'title.warnings') 'OK' 'Warning' | Out-Null }
        }
    })

# ---------------------------------------------------------------- export
$script:ExportXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{{gb.export}}" Width="700" Height="610" MinWidth="600" MinHeight="480" ResizeMode="CanResizeWithGrip" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner" Background="White" FontFamily="Segoe UI" FontSize="12">
  <DockPanel>
    <Border DockPanel.Dock="Top" Background="#0E2A47" Padding="20,14">
      <StackPanel>
        <TextBlock Text="{{gb.export}}" Foreground="White" FontSize="18" FontWeight="SemiBold"/>
        <TextBlock Text="{{export.subtitle}}" Foreground="#9FB4D9" TextWrapping="Wrap" Margin="0,2,0,0"/>
      </StackPanel>
    </Border>
    <Border DockPanel.Dock="Bottom" Background="#F3F4F6" BorderBrush="#DDE3EA" BorderThickness="0,1,0,0" Padding="20,10">
      <DockPanel>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
          <Button x:Name="btnExpCancel" Content="{{export.cancel}}" Padding="16,4" Margin="0,0,8,0" IsCancel="True" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
          <Button x:Name="btnExpOk" Content="{{export.apply}}" Padding="20,4" IsDefault="True" FontWeight="SemiBold" Background="#1565C0" Foreground="White" BorderBrush="#0E4A8F" Cursor="Hand"/>
        </StackPanel>
        <TextBlock x:Name="lblExpSummary" VerticalAlignment="Center" Foreground="#5B6B80"/>
      </DockPanel>
    </Border>
    <Grid Margin="20,14,20,10">
      <Grid.ColumnDefinitions><ColumnDefinition Width="250"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <DockPanel Grid.Column="0" Margin="0,0,16,0">
        <StackPanel DockPanel.Dock="Top">
          <TextBlock Text="{{export.content}}" FontWeight="SemiBold" Foreground="#0E2A47" Margin="0,0,0,4"/>
        </StackPanel>
        <StackPanel DockPanel.Dock="Bottom">
          <TextBlock x:Name="lblExpPresetDesc" TextWrapping="Wrap" Foreground="#5B6B80" Margin="0,6,0,12" MinHeight="48"/>
          <TextBlock Text="{{export.options}}" FontWeight="SemiBold" Foreground="#0E2A47" Margin="0,0,0,4"/>
          <CheckBox x:Name="chkExpGroups" Content="{{export.optGroups}}" Margin="0,2"/>
          <CheckBox x:Name="chkExpOverview" Content="{{export.optOverview}}" Margin="0,2"/>
          <CheckBox x:Name="chkExpEnabled" Content="{{export.optEnabled}}" Margin="0,2"/>
        </StackPanel>
        <ListBox x:Name="lstExpPreset" BorderBrush="#D3DAE3"/>
      </DockPanel>
      <DockPanel Grid.Column="1">
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,4">
          <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
            <Button x:Name="btnExpAll" Content="{{export.all}}" Padding="10,1" Margin="0,0,4,0" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
            <Button x:Name="btnExpNone" Content="{{export.none}}" Padding="10,1" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
          </StackPanel>
          <TextBlock Text="{{export.columns}}" FontWeight="SemiBold" Foreground="#0E2A47" VerticalAlignment="Center"/>
        </DockPanel>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,8,0,0">
          <Button x:Name="btnExpSave" Content="{{export.savePreset}}" Padding="10,3" Margin="0,0,6,0" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
          <Button x:Name="btnExpDelete" Content="{{export.deletePreset}}" Padding="10,3" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
        </StackPanel>
        <Border BorderBrush="#D3DAE3" BorderThickness="1">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="pnlExpCols" Margin="10,6"/>
          </ScrollViewer>
        </Border>
      </DockPanel>
    </Grid>
  </DockPanel>
</Window>
'@

function Get-HSExportLabel([string]$Key) {
    $label = T "export.preset.$Key"
    if ($label -eq "export.preset.$Key") { $Key } else { $label }
}
function Test-HSUserPreset([string]$Key) { [bool]($cfg.ExportPresets -and $cfg.ExportPresets.PSObject.Properties[$Key]) }
function Get-HSExportChecked { $st = $script:ExpState; @($st.Boxes.Keys | Where-Object { $st.Boxes[$_].IsChecked }) }

function Update-HSExportSummary {
    $st = $script:ExpState
    $item = $st.lstExpPreset.SelectedItem
    if ($item -and "$($item.Tag)" -eq 'Full') { $st.lblExpSummary.Text = T 'export.desc.Full'; return }
    $nCols = @(Get-HSExportChecked).Count
    if (-not $st.Users.Count) { $st.lblExpSummary.Text = T 'export.summaryCols' @($nCols); return }
    $users = $st.Users
    if ($st.chkExpEnabled.IsChecked) { $enabledValues = Get-HSEnabledValues; $users = @($users | Where-Object { $enabledValues -contains "$($_.Status)" }) }
    $st.lblExpSummary.Text = T 'export.summary' @($nCols, @($users).Count)
}

function Update-HSExportPresetList([string]$Select) {
    $st = $script:ExpState
    $st.Loading = $true
    $st.lstExpPreset.Items.Clear()
    $keys = @('Full') + @((Get-HSExportPresets $cfg).Keys) + @('Custom')
    foreach ($key in $keys) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = if (Test-HSUserPreset $key) { [string][char]0x2605 + '  ' + $key } else { Get-HSExportLabel $key }
        $item.Tag = $key
        $item.Padding = New-Object System.Windows.Thickness 6, 4, 6, 4
        if ($key -eq 'Full') { $item.FontWeight = 'SemiBold' }
        [void]$st.lstExpPreset.Items.Add($item)
    }
    $st.Loading = $false
    $target = @($st.lstExpPreset.Items | Where-Object { "$($_.Tag)" -eq $Select })[0]
    if (-not $target) { $target = $st.lstExpPreset.Items[0] }
    $st.lstExpPreset.SelectedItem = $target
}

function Set-HSExportChecks([string[]]$Cols) {
    $st = $script:ExpState
    foreach ($col in @($st.Boxes.Keys)) { $st.Boxes[$col].IsChecked = ($Cols -contains $col) }
}

function Set-HSExportPreset {
    # the preset chosen in the list decides the ticked columns; FULL locks them (everything is exported)
    $st = $script:ExpState
    if ($st.Loading) { return }
    $item = $st.lstExpPreset.SelectedItem; if (-not $item) { return }
    $key = "$($item.Tag)"
    $isFull = $key -eq 'Full'
    $desc = T "export.desc.$key"
    if (Test-HSUserPreset $key) { $desc = T 'export.desc.User' } elseif ($desc -eq "export.desc.$key") { $desc = '' }
    $st.lblExpPresetDesc.Text = $desc
    if ($isFull) { $st.lblExpPresetDesc.Text = '' }
    if ($isFull) { Set-HSExportChecks @($st.Boxes.Keys) }
    elseif ($key -eq 'Custom') {
        if (-not $st.KeepChecks) {
            $cols = @(Resolve-HSExportColumns -Columns @($cfg.ExportColumns) -Available $st.Available)
            if (-not $cols.Count) { $cols = @(Resolve-HSExportColumns -Columns @('FirstName', 'LastName', 'Username') -Available $st.Available) }
            Set-HSExportChecks $cols
        }
    }
    else { Set-HSExportChecks @(Resolve-HSExportColumns -Columns @(Get-HSExportColumns -Cfg $cfg -Preset $key) -Available $st.Available) }
    $st.KeepChecks = $false
    $st.pnlExpCols.IsEnabled = -not $isFull
    foreach ($name in 'btnExpAll', 'btnExpNone', 'btnExpSave', 'chkExpGroups', 'chkExpOverview', 'chkExpEnabled') { $st[$name].IsEnabled = -not $isFull }
    $st.btnExpDelete.IsEnabled = Test-HSUserPreset $key
    Update-HSExportSummary
}

function Switch-HSExportToCustom {
    # changing a column of a preset turns the selection into a custom one (the ticks stay as they are)
    $st = $script:ExpState
    $item = $st.lstExpPreset.SelectedItem
    if ($item -and "$($item.Tag)" -ne 'Custom') {
        $st.KeepChecks = $true
        $st.lstExpPreset.SelectedItem = @($st.lstExpPreset.Items | Where-Object { "$($_.Tag)" -eq 'Custom' })[0]
    }
    Update-HSExportSummary
}

function Show-HSExportDialog([string]$Select) {
    <#
        What Export XLSX / CSV will write: FULL report, a preset (built-in or saved) or custom columns of the
        user list. Opened from the "Export content" box, also before loading the data (columns of the current
        infrastructure mode). The choice is saved in config.json. Returns $true if applied.
    #>
    $xamlText = [regex]::Replace($script:ExportXaml, '\{\{([\w\.]+)\}\}', [System.Text.RegularExpressions.MatchEvaluator] {
            param($match) [System.Security.SecurityElement]::Escape((T $match.Groups[1].Value))
        })
    [xml]$expXml = $xamlText
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $expXml))
    $win.Owner = $Window
    if ($Window.Icon) { $win.Icon = $Window.Icon }
    $userRows = if ($script:Data) { @(Get-HSExportUserRows $script:Data) } else { @() }
    $st = @{
        Window = $win; Users = $userRows; Boxes = [ordered]@{}; Result = $null; KeepChecks = $false; Loading = $false
        Available = if ($userRows.Count) { @($userRows[0].PSObject.Properties.Name) } else { @(Get-HSUserColumns $script:CurrentMode) }
        Builtin = @('Full') + @($script:HS_ExportPresets.Keys) + @('Custom')
    }
    foreach ($name in 'lstExpPreset', 'lblExpPresetDesc', 'pnlExpCols', 'chkExpGroups', 'chkExpOverview', 'chkExpEnabled', 'btnExpAll', 'btnExpNone',
        'btnExpSave', 'btnExpDelete', 'btnExpOk', 'btnExpCancel', 'lblExpSummary') { $st[$name] = $win.FindName($name) }
    $script:ExpState = $st

    # one checkbox per column of the USERS tab, in the same order
    foreach ($col in $st.Available) {
        $box = New-Object System.Windows.Controls.CheckBox
        $header = Get-HSHeader $col
        $box.Content = $header
        $box.ToolTip = $col
        $box.Tag = $col
        $box.Margin = New-Object System.Windows.Thickness 0, 3, 12, 3
        $box.Add_Click({ Switch-HSExportToCustom })
        [void]$st.pnlExpCols.Children.Add($box)
        $st.Boxes[$col] = $box
    }

    $st.chkExpGroups.IsChecked = [bool]$cfg.ExportGroupSheets
    $st.chkExpOverview.IsChecked = [bool]$cfg.ExportOverview
    $st.chkExpEnabled.IsChecked = [bool]$cfg.ExportEnabledOnly
    $st.chkExpEnabled.Add_Click({ Update-HSExportSummary })
    $st.lstExpPreset.Add_SelectionChanged({ Set-HSExportPreset })
    $st.btnExpAll.Add_Click({ Set-HSExportChecks @($script:ExpState.Boxes.Keys); Switch-HSExportToCustom })
    $st.btnExpNone.Add_Click({ Set-HSExportChecks @(); Switch-HSExportToCustom })

    $st.btnExpSave.Add_Click({
            $st = $script:ExpState
            $cols = Get-HSExportChecked
            if (-not $cols.Count) { Show-Msg (T 'export.msgNoCols') (T 'export.title') 'OK' 'Warning' | Out-Null; return }
            Add-Type -AssemblyName Microsoft.VisualBasic
            $name = "$([Microsoft.VisualBasic.Interaction]::InputBox((T 'export.askName'), (T 'export.title'), ''))".Trim()
            if (-not $name) { return }
            if ($st.Builtin -contains $name) { Show-Msg (T 'export.badName' $name) (T 'export.title') 'OK' 'Warning' | Out-Null; return }
            $cfg.ExportPresets | Add-Member -NotePropertyName $name -NotePropertyValue @($cols) -Force
            try { Save-HSConfig $cfg $ConfigPath } catch { }
            Update-HSExportPresetList $name
        })
    $st.btnExpDelete.Add_Click({
            $st = $script:ExpState
            $key = "$($st.lstExpPreset.SelectedItem.Tag)"
            if (-not (Test-HSUserPreset $key)) { return }
            if ((Show-Msg (T 'export.confirmDelete' $key) (T 'export.title') 'YesNo' 'Question') -ne 'Yes') { return }
            $cfg.ExportPresets.PSObject.Properties.Remove($key)
            try { Save-HSConfig $cfg $ConfigPath } catch { }
            Update-HSExportPresetList 'Custom'
        })
    $st.btnExpOk.Add_Click({
            $st = $script:ExpState
            $key = "$($st.lstExpPreset.SelectedItem.Tag)"
            $cols = @()
            if ($key -ne 'Full') {
                $checked = Get-HSExportChecked
                if (-not $checked.Count) { Show-Msg (T 'export.msgNoCols') (T 'export.title') 'OK' 'Warning' | Out-Null; return }
                # a preset keeps its own column order; custom follows the order of the USERS tab
                $cols = if ($key -eq 'Custom') { $checked } else { @(Resolve-HSExportColumns -Columns @(Get-HSExportColumns -Cfg $cfg -Preset $key) -Available $st.Available) }
            }
            $st.Result = $true
            $cfg.ExportPreset = $key
            if ($key -eq 'Custom') { $cfg.ExportColumns = @($cols) }
            $cfg.ExportGroupSheets = [bool]$st.chkExpGroups.IsChecked; $cfg.ExportOverview = [bool]$st.chkExpOverview.IsChecked; $cfg.ExportEnabledOnly = [bool]$st.chkExpEnabled.IsChecked
            try { Save-HSConfig $cfg $ConfigPath } catch { }
            $st.Window.DialogResult = $true
        })

    Update-HSExportPresetList $(if ($Select) { $Select } else { "$($cfg.ExportPreset)" })
    [void]$win.ShowDialog()
    Update-HSExportBox
    [bool]$script:ExpState.Result
}

# ---------------------------------------------------------------- "Export content" box in the main window (chosen before Load data)
function Get-HSExportChoice {
    # the saved choice as export options; an unknown preset (deleted by hand in config.json) falls back to FULL
    $preset = "$($cfg.ExportPreset)"; if (-not $preset) { $preset = 'Full' }
    $cols = @()
    try { $cols = @(Get-HSExportColumns -Cfg $cfg -Preset $preset) } catch { $preset = 'Full' }
    if ($preset -ne 'Full' -and -not $cols.Count) { $preset = 'Full' }
    @{ Preset = $preset; Columns = $cols; GroupSheets = [bool]$cfg.ExportGroupSheets; Overview = [bool]$cfg.ExportOverview; EnabledOnly = [bool]$cfg.ExportEnabledOnly }
}

function Update-HSExportBox {
    $script:ExpComboLoading = $true
    $choice = Get-HSExportChoice
    $ui.cmbExport.Items.Clear()
    foreach ($key in @('Full') + @((Get-HSExportPresets $cfg).Keys) + @('Custom')) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = if (Test-HSUserPreset $key) { [string][char]0x2605 + '  ' + $key } else { Get-HSExportLabel $key }
        $item.Tag = $key
        [void]$ui.cmbExport.Items.Add($item)
        if ($key -eq $choice.Preset) { $ui.cmbExport.SelectedItem = $item }
    }
    $script:ExpComboLoading = $false
    Update-HSExportLabel
}

function Update-HSExportLabel {
    $choice = Get-HSExportChoice
    if ($choice.Preset -eq 'Full') { $text = T 'export.boxFull' }
    else {
        $rows = if ($script:Data) { @(Get-HSExportUserRows $script:Data) } else { @() }
        $available = if ($rows.Count) { @($rows[0].PSObject.Properties.Name) } else { @(Get-HSUserColumns $script:CurrentMode) }
        $cols = @(Resolve-HSExportColumns -Columns $choice.Columns -Available $available)
        $text = T 'export.boxCols' @($cols.Count, (@($cols | ForEach-Object { Get-HSHeader $_ }) -join ', '))
        $extra = @($(if ($choice.GroupSheets) { T 'export.optGroups' }), $(if ($choice.Overview) { T 'export.optOverview' }), $(if ($choice.EnabledOnly) { T 'export.optEnabled' })) | Where-Object { $_ }
        if ($extra.Count) { $text += '  ·  ' + ($extra -join ', ') }
    }
    $ui.chkExpFast.IsEnabled = $choice.Preset -ne 'Full'
    $ui.lblExport.Text = $text
    $ui.lblExport.ToolTip = $text
    $tip = (T 'tip.export') + "`n" + $text
    $ui.btnXlsx.ToolTip = $tip; $ui.btnCsvAll.ToolTip = $tip
}

$ui.cmbExport.Add_SelectionChanged({
        if ($script:ExpComboLoading) { return }
        $item = $ui.cmbExport.SelectedItem; if (-not $item) { return }
        $key = "$($item.Tag)"
        if ($key -eq 'Custom' -and -not @($cfg.ExportColumns).Count) {
            # nothing ticked yet: choose now (after this event, the dialog rebuilds the list)
            [void]$Window.Dispatcher.BeginInvoke([action] { Show-HSExportDialog 'Custom' | Out-Null })
            return
        }
        $cfg.ExportPreset = $key
        try { Save-HSConfig $cfg $ConfigPath } catch { }
        Update-HSExportLabel
    })
$ui.btnExpCols.Add_Click({ Show-HSExportDialog | Out-Null })
$ui.chkExpFast.IsChecked = [bool]$cfg.ExportFastLoad
$ui.chkExpFast.Add_Click({ $cfg.ExportFastLoad = [bool]$ui.chkExpFast.IsChecked; try { Save-HSConfig $cfg $ConfigPath } catch { } })
# the listed columns follow the infrastructure mode (until data is loaded)
foreach ($rb in 'rbCloud', 'rbFullHybrid', 'rbHybrid', 'rbOnPrem', 'rbAD') { $ui[$rb].Add_Checked({ Update-HSExportLabel }) }
Update-HSExportBox

function Start-HSExport([string]$Format) {
    # writes what was chosen in the "Export content" box (before or after loading the data)
    if (-not $script:Data -or $script:Job) { return }
    $opt = Get-HSExportChoice
    if ($script:Data.Partial) {
        if ($opt.Preset -eq 'Full') { $warn = T 'export.warnPartialFull' @($script:Data.Partial, (@($script:Data.Skipped) -join ', ')) }
        else {
            $loaded = @(@(Get-HSExportUserRows $script:Data)[0].PSObject.Properties.Name)
            $missing = @(Resolve-HSExportColumns -Columns $opt.Columns -Available (Get-HSUserColumns $script:CurrentMode) | Where-Object { $loaded -notcontains $_ } | ForEach-Object { Get-HSHeader $_ })
            $warn = if ($missing.Count) { T 'export.warnMissing' ($missing -join ', ') } else { '' }
        }
        if ($warn -and (Show-Msg $warn (T 'export.title') 'YesNo' 'Warning') -ne 'Yes') { return }
    }
    $stamp = $script:Data.Generated.ToString('yyyy-MM-dd_HHmm')
    $isFull = $opt.Preset -eq 'Full'
    $tag = if ($isFull) { '' } else { '_' + ($opt.Preset -replace '[^\w\-]', '') }
    $exportOpt = @{ Columns = $opt.Columns; GroupSheets = $opt.GroupSheets; Overview = $opt.Overview; EnabledOnly = $opt.EnabledOnly; Label = $opt.Preset }
    $select = '$exportData = Select-HSExportData -Data $HS_Sync.Data -Columns $ExportOpt.Columns -Label $ExportOpt.Label -GroupSheets:($ExportOpt.GroupSheets) -Overview:($ExportOpt.Overview) -EnabledOnly:($ExportOpt.EnabledOnly)' + "`n"

    if ($Format -eq 'xlsx') {
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
        $dlg.InitialDirectory = Get-HSOutputFolder $cfg
        $dlg.FileName = "HybridScope$($tag)_$stamp.xlsx"
        if (-not $dlg.ShowDialog()) { return }
        Start-HSJob -Title (T 'job.xlsx') -Vars @{ OutPath = $dlg.FileName; ExportOpt = $exportOpt } -Script ($select + 'Export-HSExcel -Data $exportData -Path $OutPath') -OnDone {
            param($out, $err)
            if ($err) { Show-Msg $err.Message 'XLSX' 'OK' 'Error' | Out-Null; return }
            $savedPath = "$(@($out)[-1])"
            if ((Show-Msg (T 'msg.saved' $savedPath) 'XLSX' 'YesNo' 'Information') -eq 'Yes') { Invoke-Item $savedPath }
        }
        return
    }

    # CSV: only the user list -> one file; several tables -> one file per tab in a folder
    if (-not $isFull -and -not $opt.GroupSheets -and -not $opt.Overview) {
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'CSV (*.csv)|*.csv'
        $dlg.InitialDirectory = Get-HSOutputFolder $cfg
        $dlg.FileName = "HybridScope$($tag)_$stamp.csv"
        if (-not $dlg.ShowDialog()) { return }
        Start-HSJob -Title (T 'job.csv') -Vars @{ OutPath = $dlg.FileName; ExportOpt = $exportOpt } -Script ($select + 'Export-HSUserCsv -Data $exportData -Path $OutPath') -OnDone {
            param($out, $err)
            if ($err) { Show-Msg $err.Message 'CSV' 'OK' 'Error' | Out-Null; return }
            $savedPath = "$(@($out)[-1])"
            if ((Show-Msg (T 'msg.saved' $savedPath) 'CSV' 'YesNo' 'Information') -eq 'Yes') { Invoke-Item $savedPath }
        }
        return
    }
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = T 'dlg.csvFolder'
    $fb.SelectedPath = Get-HSOutputFolder $cfg
    if ($fb.ShowDialog() -ne 'OK') { return }
    $dest = Join-Path $fb.SelectedPath "HybridScope_CSV$($tag)_$stamp"
    Start-HSJob -Title (T 'job.csv') -Vars @{ OutPath = $dest; ExportOpt = $exportOpt } -Script ($select + 'Export-HSCsv -Data $exportData -Folder $OutPath') -OnDone {
        param($out, $err)
        if ($err) { Show-Msg $err.Message 'CSV' 'OK' 'Error' | Out-Null; return }
        Invoke-Item "$(@($out)[-1])"
    }
}

$ui.btnXlsx.Add_Click({ Start-HSExport 'xlsx' })
$ui.btnCsvAll.Add_Click({ Start-HSExport 'csv' })

$ui.btnTab.Add_Click({
        $tab = $ui.tabs.SelectedItem; if (-not $tab) { return }
        $name = $tab.Uid
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = T 'dlg.filterCsv'
        $dlg.InitialDirectory = Get-HSOutputFolder $cfg
        $dlg.FileName = "$(Get-HSSheetName $name)_$($script:Data.Generated.ToString('yyyy-MM-dd'))"
        if (-not $dlg.ShowDialog()) { return }
        try {
            $view = $tab.Tag.DefaultView.ToTable()          # honours the active filter
            $cols = @($view.Columns | ForEach-Object { $_.ColumnName })
            $rows = @($view.Rows | Select-Object -Property $cols)
            if ($dlg.FileName -like '*.xlsx') {
                Import-Module ImportExcel -ErrorAction Stop
                if (Test-Path $dlg.FileName) { Remove-Item $dlg.FileName -Force }
                $rows | Export-Excel -Path $dlg.FileName -WorksheetName (Get-HSSheetName $name) -TableName 'HS_1' -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
            }
            else { $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';' }
            $ui.lblStatus.Text = T 'st.exported' @($rows.Count, $dlg.FileName)
        }
        catch { Show-Msg $_.Exception.Message (T 'title.error') 'OK' 'Error' | Out-Null }
    })

$ui.btnFolder.Add_Click({ Invoke-Item (Get-HSOutputFolder $cfg) })

# ---------------------------------------------------------------- menu, about, updates
function Open-HSLink([string]$Url) { try { Start-Process $Url } catch { Show-Msg $_.Exception.Message (T 'title.error') 'OK' 'Error' | Out-Null } }
function Invoke-HSButton([string]$Name) { $ui[$Name].RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent))) }

$script:DocsUrl = if ($cfg.Language -eq 'it') { "$($HS_Links.Repo)/blob/main/README.it.md" } else { "$($HS_Links.Repo)#readme" }
$ui.runVersion.Text = "HybridScope $HS_Version"
$ui.lnkStatusAuthor.Add_Click({ Open-HSLink $HS_Links.Author })
$ui.lnkStatusAuthor.ToolTip = $HS_Links.Author

# the export entries follow the state of the matching buttons (enabled once data is loaded)
$script:MenuButtons = @{ miExportXlsx = 'btnXlsx'; miExportCsv = 'btnCsvAll'; miExportTab = 'btnTab' }
$ui.miFile.Add_SubmenuOpened({ foreach ($item in $script:MenuButtons.Keys) { $ui[$item].IsEnabled = $ui[$script:MenuButtons[$item]].IsEnabled } })
$ui.miExportXlsx.Add_Click({ Invoke-HSButton 'btnXlsx' })
$ui.miExportCsv.Add_Click({ Invoke-HSButton 'btnCsvAll' })
$ui.miExportTab.Add_Click({ Invoke-HSButton 'btnTab' })
$ui.miFolder.Add_Click({ Invoke-HSButton 'btnFolder' })
$ui.miExit.Add_Click({ $Window.Close() })
$ui.miDocs.Add_Click({ Open-HSLink $script:DocsUrl })
$ui.miIssue.Add_Click({ Open-HSLink $HS_Links.Issues })
$ui.miSponsor.Add_Click({ Open-HSLink $HS_Links.Sponsor })
$ui.miCoffee.Add_Click({ Open-HSLink $HS_Links.Coffee })
$Window.Add_KeyDown({
        param($source, $evtArgs)
        if ($evtArgs.Key -eq 'F1') { Open-HSLink $script:DocsUrl }
        elseif ($evtArgs.Key -eq 'O' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { $evtArgs.Handled = $true; Open-HSReport }
    })

# ---------------------------------------------------------------- open a saved report (File > Open, Ctrl+O, drag & drop)
function Open-HSReport([string]$Path) {
    if ($script:Job) { return }
    if (-not $Path) {
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = T 'dlg.open'
        $dlg.Title = T 'title.import'
        $dlg.InitialDirectory = Get-HSOutputFolder $cfg
        if (-not $dlg.ShowDialog()) { return }
        $Path = $dlg.FileName
    }
    $ui.expLog.IsExpanded = $true
    Start-HSJob -Title (T 'job.import') -Vars @{ ImportPath = $Path } -Script @'
$data = Import-HSReport -Path $ImportPath
$tables = [ordered]@{}
foreach ($tabName in $data.Tables.Keys) { $tables[$tabName] = ConvertTo-HSDataTable -Rows @($data.Tables[$tabName]) -NoLocalize:($tabName -eq $data.OverviewName) }
$HS_Sync.Data = $data
$HS_Sync.DataTables = $tables
'@ -OnDone {
        param($out, $err)
        if ($err) { Show-Msg $err.Message (T 'title.import') 'OK' 'Error' | Out-Null; return }
        $script:Data = $HS_Sync.Data
        Show-HSTables $HS_Sync.DataTables $script:Data.OverviewName
        $ui.lblGenerated.Text = T 'lbl.imported' @($script:Data.Generated.ToString('yyyy-MM-dd HH:mm'), (Split-Path -Leaf $script:Data.ImportPath))
        $ui.lblGenerated.ToolTip = $script:Data.ImportPath
        $ui.lblGenerated.Foreground = $Brush.ConvertFromString('#9FB4D9')
        Set-Busy $false (T 'st.imported' $script:Data.Totals['total'])
    }
}
$ui.miOpen.Add_Click({ Open-HSReport })
# drop a report (.xlsx / .csv) or a HybridScope_CSV_* folder anywhere on the window
$Window.AllowDrop = $true
$Window.Add_PreviewDragOver({
        param($source, $evtArgs)
        $evtArgs.Effects = if ($evtArgs.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop) -and -not $script:Job) { 'Copy' } else { 'None' }
        $evtArgs.Handled = $true
    })
$Window.Add_PreviewDrop({
        param($source, $evtArgs)
        $dropped = @($evtArgs.Data.GetData([System.Windows.DataFormats]::FileDrop))
        $evtArgs.Handled = $true
        if ($dropped.Count) { Open-HSReport "$($dropped[0])" }
    })

function Start-HSUpdateCheck {
    # GitHub is contacted only here, when the user asks
    if ($script:Job) { return }
    Start-HSJob -Title (T 'upd.checking') -Vars @{ HSRepo = $HS_Repo } -Script 'Get-HSLatestRelease -Repo $HSRepo' -OnDone {
        param($out, $err)
        if ($err) {
            if ((Show-Msg (T 'upd.error' $err.Message) (T 'title.updates') 'YesNo' 'Warning') -eq 'Yes') { Open-HSLink $HS_Links.Releases }
            return
        }
        $latest = @($out)[-1]
        $cmp = Compare-HSVersion $latest.Version $HS_Version
        if ($cmp -gt 0) {
            if ((Show-Msg (T 'upd.new' @($latest.Version, $HS_Version)) (T 'title.updates') 'YesNo' 'Information') -eq 'Yes') { Open-HSLink $(if ($latest.Url) { $latest.Url } else { $HS_Links.Releases }) }
        }
        elseif ($cmp -lt 0) { Show-Msg (T 'upd.ahead' @($HS_Version, $latest.Version)) (T 'title.updates') | Out-Null }
        else { Show-Msg (T 'upd.latest' $HS_Version) (T 'title.updates') | Out-Null }
        $ui.lblStatus.Text = T 'st.ready'
    }
}
$ui.miUpdates.Add_Click({ Start-HSUpdateCheck })

$script:AboutXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{{about.title}}" Width="540" SizeToContent="Height" ResizeMode="NoResize" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner" Background="White" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
    <Style x:Key="Pill" TargetType="Button">
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="14" Padding="14,5">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.85"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Hyperlink">
      <Setter Property="Foreground" Value="#1565C0"/>
      <Setter Property="TextDecorations" Value="{x:Null}"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
  </Window.Resources>
  <StackPanel>
    <Border Background="#0E2A47" Padding="26,20">
      <StackPanel>
        <TextBlock Text="HybridScope" Foreground="White" FontSize="28" FontWeight="SemiBold"/>
        <TextBlock x:Name="lblVersion" Foreground="#9FB4D9" FontSize="13" Margin="0,2,0,0"/>
      </StackPanel>
    </Border>
    <StackPanel Margin="26,18,26,0">
      <TextBlock Text="{{about.desc}}" TextWrapping="Wrap" Foreground="#0E2A47" FontSize="13.5"/>
      <TextBlock Text="{{app.subtitle}}" Foreground="#5B6B80" Margin="0,3,0,16"/>
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="110"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Grid.RowDefinitions><RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="{{about.by}}" Foreground="#5B6B80" Margin="0,3"/>
        <TextBlock Grid.Row="0" Grid.Column="1" Margin="0,3" FontWeight="SemiBold"><Hyperlink x:Name="lnkAuthor">Thonno</Hyperlink><Run x:Name="runAuthorUrl" FontWeight="Normal" Foreground="#8A94A6"/></TextBlock>
        <TextBlock Grid.Row="1" Text="{{about.project}}" Foreground="#5B6B80" Margin="0,3"/>
        <TextBlock Grid.Row="1" Grid.Column="1" Margin="0,3"><Hyperlink x:Name="lnkRepo"><Run x:Name="runRepo"/></Hyperlink></TextBlock>
        <TextBlock Grid.Row="2" Text="{{about.licenseLbl}}" Foreground="#5B6B80" Margin="0,3"/>
        <TextBlock Grid.Row="2" Grid.Column="1" Margin="0,3"><Hyperlink x:Name="lnkLicense">{{about.license}}</Hyperlink></TextBlock>
        <TextBlock Grid.Row="3" Text="{{about.feedback}}" Foreground="#5B6B80" Margin="0,3"/>
        <TextBlock Grid.Row="3" Grid.Column="1" Margin="0,3"><Hyperlink x:Name="lnkIssues">{{menu.issue}}</Hyperlink></TextBlock>
      </Grid>
      <TextBlock Text="{{about.support}}" TextWrapping="Wrap" Foreground="#0E2A47" Margin="0,18,0,8"/>
      <WrapPanel>
        <Button x:Name="btnSponsor" Style="{StaticResource Pill}" Background="#EA4AAA" BorderBrush="#EA4AAA" Foreground="White" Content="&#x2665;  {{menu.sponsor}}"/>
        <Button x:Name="btnCoffee" Style="{StaticResource Pill}" Background="#FFDD00" BorderBrush="#E6C700" Foreground="Black" Content="&#x2615;  {{menu.coffee}}"/>
        <Button x:Name="btnStar" Style="{StaticResource Pill}" Background="White" BorderBrush="#B8C2D0" Foreground="#0E2A47" Content="&#x2605;  {{about.star}}"/>
      </WrapPanel>
      <TextBlock Text="{{about.disclaimer}}" TextWrapping="Wrap" Foreground="#8A94A6" FontSize="11" Margin="0,18,0,0"/>
      <TextBlock x:Name="lblEnv" Foreground="#8A94A6" FontSize="11" Margin="0,2,0,0"/>
    </StackPanel>
    <Border Background="#F3F4F6" BorderBrush="#DDE3EA" BorderThickness="0,1,0,0" Padding="26,10" Margin="0,18,0,0">
      <DockPanel>
        <Button x:Name="btnClose" Content="{{about.close}}" DockPanel.Dock="Right" Padding="18,4" IsCancel="True" IsDefault="True" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
        <Button x:Name="btnAboutUpdates" Content="{{menu.updates}}" HorizontalAlignment="Left" Padding="12,4" Background="White" BorderBrush="#B8C2D0" Cursor="Hand"/>
      </DockPanel>
    </Border>
  </StackPanel>
</Window>
'@

function Show-HSAbout {
    $xamlText = [regex]::Replace($script:AboutXaml, '\{\{([\w\.]+)\}\}', [System.Text.RegularExpressions.MatchEvaluator] {
            param($match) [System.Security.SecurityElement]::Escape((T $match.Groups[1].Value))
        })
    [xml]$aboutXml = $xamlText
    $about = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $aboutXml))
    $about.Owner = $Window
    if ($Window.Icon) { $about.Icon = $Window.Icon }
    $about.FindName('lblVersion').Text = T 'about.version' $HS_Version
    $about.FindName('runAuthorUrl').Text = "   github.com/$HS_Author"
    $about.FindName('runRepo').Text = "github.com/$HS_Repo"
    $about.FindName('lblEnv').Text = 'PowerShell {0}  ·  Windows {1}' -f $PSVersionTable.PSVersion, [Environment]::OSVersion.Version
    $about.FindName('lnkAuthor').Add_Click({ Open-HSLink $HS_Links.Author })
    $about.FindName('lnkRepo').Add_Click({ Open-HSLink $HS_Links.Repo })
    $about.FindName('lnkLicense').Add_Click({ Open-HSLink $HS_Links.License })
    $about.FindName('lnkIssues').Add_Click({ Open-HSLink $HS_Links.Issues })
    $about.FindName('btnSponsor').Add_Click({ Open-HSLink $HS_Links.Sponsor })
    $about.FindName('btnCoffee').Add_Click({ Open-HSLink $HS_Links.Coffee })
    $about.FindName('btnStar').Add_Click({ Open-HSLink $HS_Links.Repo })
    $about.FindName('btnAboutUpdates').Add_Click({ param($source, $evtArgs) ([System.Windows.Window]::GetWindow($source)).Close(); Start-HSUpdateCheck })
    $about.FindName('btnClose').Add_Click({ param($source, $evtArgs) ([System.Windows.Window]::GetWindow($source)).Close() })
    [void]$about.ShowDialog()
}
$ui.miAbout.Add_Click({ Show-HSAbout })

$Window.Add_Closing({
        $timer.Stop()
        try {
            $ps = [powershell]::Create(); $ps.Runspace = $script:RS
            [void]$ps.AddScript('try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}; if ($script:HS_ExSession) { Remove-PSSession $script:HS_ExSession -ErrorAction SilentlyContinue }')
            [void]$ps.Invoke(); $ps.Dispose()
            $script:RS.Close()
        }
        catch { }
    })

[void]$Window.ShowDialog()