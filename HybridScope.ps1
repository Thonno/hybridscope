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

      One dataset        -Only <name>. No window: collects just that dataset (e.g. MFA) and writes
                         HybridScope_<name>_<timestamp>.xlsx, or returns it with -PassThru.

      Pipeline           -PassThru. Returns objects instead of writing files: the rows of the -Only
                         dataset, or one merged row per user without -Only.

    Microsoft 365 sign-in without the GUI depends on AuthMode in config.json:
      Interactive  -> a browser sign-in opens once (one-off runs from a console)
      Certificate  -> app + certificate, no prompt (scheduled tasks, see tools\New-HybridScopeApp.ps1)
    If you are already connected with Connect-MgGraph in the same session, that connection is reused
    and left open. AD and Exchange use the Windows account running the script.

    Files written in OutputFolder (default: Documents\HybridScope):
      HybridScope_<yyyy-MM-dd_HHmm>.xlsx         full report (-Auto)
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
    Only the source that dataset needs is contacted: -Only MFA does not touch AD or Exchange.

.PARAMETER PassThru
    Return the data as objects instead of writing files. Implies -Auto.
    With -Only: the rows of that dataset. Without -Only: one merged row per user (the USERS tab).
    Property names and values are always in English (Yes/No, Enabled/Disabled...), whatever the
    Language setting, so filters like Where-Object MFA_Registered -eq 'No' work everywhere.

.PARAMETER Mode
    Which sources to use for this run, overriding Mode in config.json (not saved):
    Cloud, FullHybrid, Hybrid, OnPrem, AD. In the GUI it only sets the starting mode.

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
    .\HybridScope.ps1 -Auto

    Full report without the GUI, using Mode and the other settings in config.json.

.EXAMPLE
    .\HybridScope.ps1 -Auto -Mode Cloud -ConfigPath D:\HybridScope\contoso.json

    Full report for Microsoft 365 only, with a different configuration file (e.g. one per tenant).

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
                   permissions with a certificate).
    Last sign-in and the MFA registration report need Entra ID P1; without it MFA methods are read
    user by user (slower) and a warning is logged.

.LINK
    https://github.com/Thonno/hybridscope
#>
param(
    [switch]$Auto,

    [ValidateSet('MFA', 'Licenses', 'Mailbox365', 'OneDrive', 'MailboxOnPrem')]
    [string]$Only,

    [switch]$PassThru,

    [ValidateSet('Cloud', 'FullHybrid', 'Hybrid', 'OnPrem', 'AD')]
    [string]$Mode,

    [string]$ConfigPath
)

$HS_Version = '1.4.0'

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
            'btn.xlsx'          = 'Export XLSX (all tabs)'
            'btn.csvall'        = 'Export CSV (all tabs)'
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
            'st.loaded'         = 'Loaded {0} users'
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

            'm.sec.users'       = 'USERS'
            'm.total'           = 'Total users'
            'm.enabled'         = 'Enabled'
            'm.disabled'        = 'Disabled'
            'm.adOnly'          = 'AD only (not synced)'
            'm.cloudOnly'       = 'Cloud only'
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
            'btn.xlsx'          = 'Esporta XLSX (tutti i tab)'
            'btn.csvall'        = 'Esporta CSV (tutti i tab)'
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
            'st.loaded'         = 'Caricati {0} utenti'
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

            'm.sec.users'       = 'UTENTI'
            'm.total'           = 'Utenti totali'
            'm.enabled'         = 'Abilitati'
            'm.disabled'        = 'Disabilitati'
            'm.adOnly'          = 'Solo AD (non sincronizzati)'
            'm.cloudOnly'       = 'Solo cloud'
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
        $cfg
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
        $connectParams = @{ NoWelcome = $true }
        if ($Cfg.TenantId) { $connectParams.TenantId = $Cfg.TenantId }
        if ($Unattended) {
            $connectParams.ClientId = $Cfg.ClientId; $connectParams.CertificateThumbprint = $Cfg.CertificateThumbprint
        }
        else {
            # WAM needs a window handle: in the GUI the browser sign-in is used instead
            try { Set-MgGraphOption -DisableLoginByWAM $true } catch { }
            $connectParams.Scopes = $scopes
            # do not persist the sign-in between sessions: no silent re-login with the wrong account
            $connectParams.ContextScope = 'Process'
        }
        Connect-MgGraph @connectParams
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
        $select = 'id,displayName,givenName,surname,userPrincipalName,mail,accountEnabled,assignedLicenses,provisionedPlans,onPremisesSamAccountName,onPremisesSyncEnabled,userType,createdDateTime'
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
    # --------------------------------------------------------------------------- data collection
    function Get-HSData {
        <#
        .SYNOPSIS
            Full report: runs the collectors for the sources of the selected Mode, merges them into one row per user and builds every tab. Returns Tables, Totals, Warnings and metadata. Microsoft 365 needs an existing Graph connection.
        #>
        param($Cfg, [pscredential]$ADCredential, [pscredential]$ExchangeCredential, [hashtable]$Sources)

        $ErrorActionPreference = 'Stop'
        $Src = if ($Sources) { $Sources } else { Get-HSSources $Cfg.Mode }
        $script:HS_Warnings = [System.Collections.Generic.List[string]]::new()
        $Now = Get-Date
        $InactiveCut = $Now.AddDays( - [int]$Cfg.InactiveDays)
        $PwdCut = $Now.AddDays( - [int]$Cfg.PasswordAgeDays)
        $Labels = @{
            Enabled = T 'val.enabled'; Disabled = T 'val.disabled'; None = T 'val.none'
            MfaYes = T 'val.mfaYes'; MfaNo = T 'val.mfaNo'; MfaNotInEntra = T 'val.mfaNotInEntra'
            Synced = T 'val.originSynced'; ADOnly = T 'val.originAD'; Cloud = T 'val.originCloud'; Guest = T 'val.originGuest'
        }
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

            $mfaData = Get-HSMfaRegistration -Cfg $Cfg -EntraUsers $entra
            $MfaReg = $mfaData.Registration; $PerUser = $mfaData.PerUser

            $mailboxUsage = Get-HSMailbox365Usage
            $oneDriveUsage = Get-HSOneDriveUsage
            $Mbx365 = $mailboxUsage.SizeByUpn; $TabMbx365 = $mailboxUsage.Rows
            $OneDrive = $oneDriveUsage.SizeByUpn; $TabOD = $oneDriveUsage.Rows
            Write-HSLog (T 'log.reports' @($Mbx365.Count, $OneDrive.Count))
            if ($mailboxUsage.Concealed -or $oneDriveUsage.Concealed) { Write-HSLog (T 'warn.concealed') 'WARN' }
        }

        # ================================================= EXCHANGE ON-PREMISES
        $ExchOnPrem = @{}; $ExchRemote = @{}; $ExchAvailable = $false; $TabOnPrem = @()
        if ($Src.Exch -and ($Cfg.ExchangeServer -or $Cfg.ExchangeProtocol -eq 'local')) {
            try {
                $exCred = if ($ExchangeCredential) { $ExchangeCredential } else { $ADCredential }
                $exchange = Get-HSExchangeMailbox -Cfg $Cfg -Credential $exCred
                $ExchOnPrem = $exchange.BySam; $ExchRemote = $exchange.RemoteBySam; $TabOnPrem = $exchange.Rows
                $ExchAvailable = $true
            }
            catch { Write-HSLog (T 'warn.exch' $_.Exception.Message) 'WARN' }
        }
        elseif ($Src.Exch) { Write-HSLog (T 'log.exchSkipped') }

        # ================================================= ACTIVE DIRECTORY
        $adUsers = @(); $NetBIOS = ''
        if ($Src.AD -and $Cfg.DomainFqdn) {
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $common = @{ Server = $Cfg.DomainFqdn }
                if ($ADCredential) { $common.Credential = $ADCredential }
                try { $NetBIOS = (Get-ADDomain @common).NetBIOSName } catch { $NetBIOS = ($Cfg.DomainFqdn -split '\.')[0].ToUpper() }
                $adParams = @{
                    Filter     = '*'
                    Properties = 'GivenName', 'Surname', 'CanonicalName', 'LastLogonDate', 'PasswordLastSet', 'PasswordNeverExpires', 'UserPrincipalName', 'mail', 'whenCreated', 'Description'
                } + $common
                if ($Cfg.SearchBase) { $adParams.SearchBase = $Cfg.SearchBase }
                $adUsers = @(Get-ADUser @adParams)
                Write-HSLog (T 'log.adUsers' $adUsers.Count)
            }
            catch { Write-HSLog (T 'warn.ad' $_.Exception.Message) 'WARN' }
        }
        elseif ($Src.AD) { Write-HSLog (T 'log.adSkipped') }

        # ================================================= USER ROWS
        $GrpOther = if ($Cfg.UnmappedGroup) { $Cfg.UnmappedGroup } else { T 'grp.other' }
        $GrpCloud = if ($Cfg.CloudOnlyGroup) { $Cfg.CloudOnlyGroup } else { T 'grp.cloud' }

        $newRow = {
            param($Ad, $EntraUser, [string]$Group)
            $sam = if ($Ad) { $Ad.SamAccountName } elseif ($EntraUser.onPremisesSamAccountName) { $EntraUser.onPremisesSamAccountName } else { ($EntraUser.userPrincipalName -split '@')[0] }
            $upn = if ($EntraUser) { $EntraUser.userPrincipalName } else { $null }
            $notes = [System.Collections.Generic.List[string]]::new()

            $origin = if ($Ad -and $EntraUser) { $Labels.Synced } elseif ($Ad) { $Labels.ADOnly } elseif ($EntraUser.userType -eq 'Guest') { $Labels.Guest } else { $Labels.Cloud }
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
            if ($location -eq '365' -and -not $inReport365) { $notes.Add((T 'note.no365Size')) }
            $od = if ($upn -and $OneDrive.ContainsKey($upn)) { $OneDrive[$upn] } else { 0 }

            # last 365 sign-in (max of interactive / non-interactive)
            $azLast = $null
            if ($EntraUser -and $EntraUser.signInActivity) {
                $azLast = @((ConvertTo-LocalDate $EntraUser.signInActivity.lastSignInDateTime), (ConvertTo-LocalDate $EntraUser.signInActivity.lastNonInteractiveSignInDateTime)) |
                    Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
            }

            # MFA
            $mfa = $Labels.MfaNotInEntra; $mfaDef = ''; $mfaMethods = ''; $mfaPerUser = ''
            if ($EntraUser) {
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
        foreach ($ad in $adUsers) {
            $entraUser = $null
            if ($EntraBySam.ContainsKey($ad.SamAccountName)) { $entraUser = $EntraBySam[$ad.SamAccountName] }
            elseif ($ad.UserPrincipalName -and $EntraByUpn.ContainsKey($ad.UserPrincipalName)) { $entraUser = $EntraByUpn[$ad.UserPrincipalName] }
            elseif ($EntraByLocal.ContainsKey($ad.SamAccountName)) { $entraUser = $EntraByLocal[$ad.SamAccountName] }
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
            & $addMetric 'cloudOnly' (T 'm.cloudOnly') { param($rows) @($rows | Where-Object { $_.Origin -eq $Labels.Cloud -or $_.Origin -eq $Labels.Guest }).Count }
        }
        if ($Src.M365 -or $Src.Exch) {
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
        & $addMetric 'secMfa' (T 'm.sec.mfa') $null -Section
        & $addMetric 'mfaYes' (T 'm.mfaYes') { param($rows) @($rows | Where-Object { $_.MFA_Status -eq $Labels.MfaYes }).Count }
        & $addMetric 'mfaNo' (T 'm.mfaNo') { param($rows) @($rows | Where-Object { $_.MFA_Status -eq $Labels.MfaNo }).Count }
        & $addMetric 'mfaNotInEntra' (T 'm.mfaNotInEntra') { param($rows) @($rows | Where-Object { $_.MFA_Status -eq $Labels.MfaNotInEntra }).Count }
        & $addMetric 'mfaPct' (T 'm.mfaPct') { param($rows) $inEntra = @($rows | Where-Object { $_.MFA_Status -ne $Labels.MfaNotInEntra }); if ($inEntra.Count) { [math]::Round(100 * @($inEntra | Where-Object { $_.MFA_Status -eq $Labels.MfaYes }).Count / $inEntra.Count, 1) } else { 0 } }
        & $addMetric 'licNoMfa' (T 'm.licNoMfa') { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled -and (& $isLic $_) -and $_.MFA_Status -eq $Labels.MfaNo }).Count }
        & $addMetric 'perUserEnforced' (T 'm.perUserEnforced') { param($rows) @($rows | Where-Object { $_.MFA_PerUser -eq 'enforced' }).Count }
        }   # end Microsoft 365 metrics
        & $addMetric 'secInactive' (T 'm.sec.inactive') $null -Section
        if ($Src.AD) { & $addMetric 'inactiveAD' (T 'm.inactiveAD' $Cfg.InactiveDays) { param($rows) @($rows | Where-Object { $_.Status -eq $Labels.Enabled -and ($_.Origin -eq $Labels.Synced -or $_.Origin -eq $Labels.ADOnly) -and (-not $_.AD_LastLogon -or $_.AD_LastLogon -lt $InactiveCut) }).Count } }
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
            if ($userRow.Status -eq $Labels.Enabled -and ($userRow.Origin -eq $Labels.Synced -or $userRow.Origin -eq $Labels.ADOnly) -and (-not $userRow.AD_LastLogon -or $userRow.AD_LastLogon -lt $InactiveCut)) { & $addAttention (T 'att.c5' $Cfg.InactiveDays) $userRow (& $lastOrNever $userRow.AD_LastLogon) }
            if ($Src.M365 -and $Src.AD -and $userRow.Origin -eq $Labels.ADOnly -and $userRow.Status -eq $Labels.Enabled) { & $addAttention (T 'att.c6') $userRow $userRow.OU }
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
        $strip = { param($rows) if ($drop.Count) { @($rows | Select-Object * -ExcludeProperty $drop) } else { @($rows) } }

        $Tables = [ordered]@{}
        $Tables[(T 'tab.overview')] = $Overview
        $Tables[(T 'tab.users')] = & $strip $All
        foreach ($groupName in $Groups.Keys) { $tabName = if ($Tables.Contains($groupName)) { "OU_$groupName" } else { $groupName }; $Tables[$tabName] = & $strip $Groups[$groupName] }
        if ($TabOnPrem.Count) { $Tables[(T 'tab.mbxOnPrem')] = @($TabOnPrem | Sort-Object Size_MB -Descending) }
        if ($TabMbx365.Count) { $Tables[(T 'tab.mbx365')] = @($TabMbx365 | Sort-Object Size_MB -Descending) }
        if ($TabOD.Count) { $Tables[(T 'tab.onedrive')] = @($TabOD | Sort-Object Size_MB -Descending) }
        if ($Src.M365) {
            $Tables[(T 'tab.mfa')] = $TabMfa
            $Tables[(T 'tab.licenses')] = $TabLic
        }
        if ($Attn.Count) { $Tables[(T 'tab.attention')] = & $strip $Attn }
        $hist = Join-Path (Get-HSOutputFolder $Cfg) 'history.csv'
        if (Test-Path $hist) { $Tables[(T 'tab.history')] = @(Import-Csv -Path $hist -Delimiter ';' -Encoding UTF8) }

        Write-HSLog (T 'log.done' @($All.Count, "$($Totals['loc365'])", "$($Totals['locOnPrem'])", "$($Totals['mfaPct'])"))

        [pscustomobject]@{
            Generated    = $Now
            Tables       = $Tables
            OverviewName = (T 'tab.overview')
            MetricFlags  = @($Metrics | ForEach-Object { [pscustomobject]@{ IsSection = $_.IsSection; IsGB = $_.IsGB } })
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
        param($Cfg, [ValidateSet('MFA', 'Licenses', 'Mailbox365', 'OneDrive', 'MailboxOnPrem')][string]$Name, [pscredential]$ExchangeCredential)
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
            $excelParams = @{ Path = $Path; WorksheetName = $sheetName; TableName = "T$i"; TableStyle = 'Medium2'; AutoSize = $true; BoldTopRow = $true }
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
if ($Only -or $PassThru) { $Auto = $true }
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
                $tabKeys = @{ MFA = 'tab.mfa'; Licenses = 'tab.licenses'; Mailbox365 = 'tab.mbx365'; OneDrive = 'tab.onedrive'; MailboxOnPrem = 'tab.mbxOnPrem' }
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
            $data = Get-HSData -Cfg $cfg
            if ($PassThru) { $data.Tables[(T 'tab.users')] }
            else {
                Update-HSHistory -Data $data -Cfg $cfg
                $xl = Join-Path $out "HybridScope_$stamp.xlsx"
                Export-HSExcel -Data $data -Path $xl | Out-Null
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
          <Button x:Name="btnXlsx" Content="{{btn.xlsx}}" Margin="16,0,0,0" IsEnabled="False"/>
          <Button x:Name="btnCsvAll" Content="{{btn.csvall}}" IsEnabled="False"/>
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
    foreach ($control in 'btn365', 'btnSignOut', 'btnExch', 'btnAD', 'btnLoad', 'cmbGroupBy', 'rbCloud', 'rbFullHybrid', 'rbHybrid', 'rbOnPrem', 'rbAD') { $ui[$control].IsEnabled = -not $busy }
    $ui.btnStop.IsEnabled = $busy
    foreach ($control in 'btnXlsx', 'btnCsvAll', 'btnTab') { $ui[$control].IsEnabled = (-not $busy) -and ($null -ne $script:Data) }
    $ui.prg.IsIndeterminate = $busy
    if ($text) { $ui.lblStatus.Text = $text }
    $Window.Cursor = if ($busy) { [System.Windows.Input.Cursors]::AppStarting } else { $null }
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
function Set-HSFilter {
    $tab = $ui.tabs.SelectedItem
    if (-not $tab -or -not $tab.Tag) { $ui.lblCount.Text = ''; return }
    $dv = $tab.Tag.DefaultView
    $filterText = $ui.txtFilter.Text.Trim()
    if (-not $filterText) { $dv.RowFilter = '' }
    else {
        $pattern = $filterText.Replace("'", "''")
        $pattern = [regex]::Replace($pattern, '[\[\]\*%]', [System.Text.RegularExpressions.MatchEvaluator] { param($match) "[$($match.Value)]" })
        $parts = foreach ($column in $tab.Tag.Columns) { "CONVERT([$($column.ColumnName.Replace(']','\]'))], 'System.String') LIKE '%$pattern%'" }
        try { $dv.RowFilter = ($parts -join ' OR ') } catch { }
    }
    $ui.lblCount.Text = T 'lbl.rows' @($dv.Count, $tab.Tag.Rows.Count)
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
        $ui.expLog.IsExpanded = $true
        Start-HSJob -Title (T 'job.load') -Step 'load' -Vars @{ ADCred = $script:ADCred; ExCred = $script:ExCred; Sources = $src } -Script @'
$data = Get-HSData -Cfg $Cfg -ADCredential $ADCred -ExchangeCredential $ExCred -Sources $Sources
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
            Set-Busy $false (T 'st.loaded' $script:Data.Totals['total'])
            if ($script:Data.Warnings.Count) { Show-Msg (T 'msg.loadWarn' ($script:Data.Warnings -join "`n- ")) (T 'title.warnings') 'OK' 'Warning' | Out-Null }
        }
    })

# ---------------------------------------------------------------- export
$ui.btnXlsx.Add_Click({
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
        $dlg.InitialDirectory = Get-HSOutputFolder $cfg
        $dlg.FileName = "HybridScope_$($script:Data.Generated.ToString('yyyy-MM-dd_HHmm')).xlsx"
        if (-not $dlg.ShowDialog()) { return }
        Start-HSJob -Title (T 'job.xlsx') -Vars @{ OutPath = $dlg.FileName } -Script 'Export-HSExcel -Data $HS_Sync.Data -Path $OutPath' -OnDone {
            param($out, $err)
            if ($err) { Show-Msg $err.Message 'XLSX' 'OK' 'Error' | Out-Null; return }
            $savedPath = "$(@($out)[-1])"
            if ((Show-Msg (T 'msg.saved' $savedPath) 'XLSX' 'YesNo' 'Information') -eq 'Yes') { Invoke-Item $savedPath }
        }
    })

$ui.btnCsvAll.Add_Click({
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = T 'dlg.csvFolder'
        $fb.SelectedPath = Get-HSOutputFolder $cfg
        if ($fb.ShowDialog() -ne 'OK') { return }
        $dest = Join-Path $fb.SelectedPath "HybridScope_CSV_$($script:Data.Generated.ToString('yyyy-MM-dd_HHmm'))"
        Start-HSJob -Title (T 'job.csv') -Vars @{ OutPath = $dest } -Script 'Export-HSCsv -Data $HS_Sync.Data -Folder $OutPath' -OnDone {
            param($out, $err)
            if ($err) { Show-Msg $err.Message 'CSV' 'OK' 'Error' | Out-Null; return }
            Invoke-Item "$(@($out)[-1])"
        }
    })

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
                $rows | Export-Excel -Path $dlg.FileName -WorksheetName (Get-HSSheetName $name) -TableName 'T1' -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
            }
            else { $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';' }
            $ui.lblStatus.Text = T 'st.exported' @($rows.Count, $dlg.FileName)
        }
        catch { Show-Msg $_.Exception.Message (T 'title.error') 'OK' 'Error' | Out-Null }
    })

$ui.btnFolder.Add_Click({ Invoke-Item (Get-HSOutputFolder $cfg) })

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
