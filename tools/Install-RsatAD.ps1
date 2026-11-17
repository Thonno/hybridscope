<#
.SYNOPSIS
    Installs the ActiveDirectory PowerShell module (RSAT). Must run as administrator.
    HybridScope launches it in a visible elevated window; you can also run it manually.
    No registry or Windows Update settings are changed.
#>
$it = (Get-Culture).TwoLetterISOLanguageName -eq 'it'
function Say($en, $itMsg, $color = 'Gray') { Write-Host $(if ($it) { $itMsg } else { $en }) -ForegroundColor $color }

Say 'Installing RSAT - Active Directory (it may take a few minutes)...' 'Installazione RSAT - Active Directory (può richiedere alcuni minuti)...' Cyan
try {
    if ((Get-CimInstance Win32_OperatingSystem).ProductType -ne 1) {
        Install-WindowsFeature RSAT-AD-PowerShell -ErrorAction Stop | Out-Null                                    # Windows Server
    }
    else {
        Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop | Out-Null   # Windows 10/11
    }
    Say 'Installed successfully.' 'Installato correttamente.' Green
    Start-Sleep -Seconds 3
    exit 0
}
catch {
    Say "ERROR: $($_.Exception.Message)" "ERRORE: $($_.Exception.Message)" Red
    Write-Host ''
    Say 'If the error is 0x800f0954 the download is blocked by your WSUS server.' "Se l'errore è 0x800f0954 il download è bloccato dal server WSUS aziendale."
    Say 'Install it from Settings > System > Optional features > "RSAT: Active Directory Domain Services" or ask your IT team.' 'Installalo da Impostazioni > Sistema > Funzionalità facoltative > "RSAT: Strumenti di Servizi di dominio Active Directory" oppure chiedi al reparto IT.'
    Read-Host $(if ($it) { 'Premi INVIO per chiudere' } else { 'Press ENTER to close' })
    exit 1
}
