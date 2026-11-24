<#
.SYNOPSIS
    One-time setup of the HybridScope App Registration in Entra ID (one app per tenant), READ-ONLY permissions,
    admin consent, and config.json updated (TenantId, ClientId, ...).

.DESCRIPTION
    Run it once as a Global Administrator. Two uses, that can be combined on the same app:

    -Defender     Microsoft Defender in the GUI (tabs SECURITY, VULNERABILITIES, SECURITY_FIXES).
                  No certificate and no secret: every HybridScope user signs in in the browser with their own
                  account (as with Connect 365) and sees only what their Defender role allows (e.g. Security Reader).
                  Adds to the app: platform "Mobile and desktop applications" with redirect http://localhost,
                  delegated WindowsDefenderATP permissions (Machine.Read, User.Read.All, Vulnerability.Read,
                  Software.Read, SecurityRecommendation.Read, Alert.Read, Score.Read) + admin consent.
                  Sets Defender = true in config.json.

    (no switch)   Unattended mode (-Auto, scheduled task): creates a certificate in the Windows certificate store
                  of the account that runs the script, adds it to the app, application Microsoft Graph permissions
                  (+ application Defender permissions when Defender is set up) + admin consent.
                  Sets CertificateThumbprint and AuthMode = Certificate. Run it with the SAME Windows account that
                  will run the scheduled task (or -Store LocalMachine from an elevated prompt and give that account
                  read access to the private key).

    If config.json already has a ClientId the same app is updated; otherwise the app is created.

.EXAMPLE
    .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com -Defender
    Defender in the GUI, browser sign-in, no certificate.

.EXAMPLE
    .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com
    Scheduled reports (-Auto) with a certificate.

.EXAMPLE
    .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com -Defender -Certificate
    Both at once.
#>
param(
    [Parameter(Mandatory)][string]$TenantId,
    [switch]$Defender,
    [switch]$Certificate,
    [string]$AppName = 'HybridScope',
    [int]$CertYears = 2,
    [ValidateSet('CurrentUser', 'LocalMachine')][string]$Store = 'CurrentUser',
    [string]$ConfigPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config.json')
)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication

$withCert = $Certificate -or -not $Defender
$graphAppId = '00000003-0000-0000-c000-000000000000'
$graphPerms = 'User.Read.All', 'AuditLog.Read.All', 'Reports.Read.All', 'UserAuthenticationMethod.Read.All', 'Organization.Read.All', 'Policy.Read.All', 'Device.Read.All'
# Microsoft Defender API (WindowsDefenderATP), read-only
$defAppId = 'fc780465-2017-40d4-a0c5-307022471b92'
$defUserPerms = 'Machine.Read', 'User.Read.All', 'Vulnerability.Read', 'Software.Read', 'SecurityRecommendation.Read', 'Alert.Read', 'Score.Read'
$defAppPerms = 'Machine.Read.All', 'User.Read.All', 'Vulnerability.Read.All', 'Software.Read.All', 'SecurityRecommendation.Read.All', 'Alert.Read.All', 'Score.Read.All'

$cfg = $null
if (Test-Path $ConfigPath) { $cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json }
$clientId = if ($cfg) { "$($cfg.ClientId)" } else { '' }
$useDefender = $Defender -or ($cfg -and $cfg.Defender -eq $true)

try { Set-MgGraphOption -DisableLoginByWAM $true } catch { }
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All' -NoWelcome

function Get-Sp([string]$AppId) { Invoke-MgGraphRequest -Method GET -Uri "v1.0/servicePrincipals(appId='$AppId')" -OutputType PSObject }
function Get-AppRole($Sp, [string[]]$Names) {
    foreach ($p in $Names) {
        $r = $Sp.appRoles | Where-Object { $_.value -eq $p -and $_.allowedMemberTypes -contains 'Application' }
        if (-not $r) { throw "Application permission $p not found on $($Sp.displayName)" }
        $r
    }
}
function Get-Scope($Sp, [string[]]$Names) {
    foreach ($p in $Names) {
        $s = $Sp.oauth2PermissionScopes | Where-Object { $_.value -eq $p }
        if ($s) { $s } else { Write-Warning "Delegated permission $p not found on $($Sp.displayName): skipped" }
    }
}
function Grant-AppRole($Sp, $ResourceSp, $Roles) {
    $existing = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/servicePrincipals/$($Sp.id)/appRoleAssignments" -OutputType PSObject).value | ForEach-Object { "$($_.resourceId)|$($_.appRoleId)" })
    foreach ($r in $Roles) {
        if ($existing -contains "$($ResourceSp.id)|$($r.id)") { Write-Host "  already granted: $($r.value)"; continue }
        $b = @{ principalId = $Sp.id; resourceId = $ResourceSp.id; appRoleId = $r.id } | ConvertTo-Json
        Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals/$($Sp.id)/appRoleAssignedTo" -Body $b -ContentType 'application/json' | Out-Null
        Write-Host "  consent (application): $($r.value)"
    }
}
function Grant-Scope($Sp, $ResourceSp, [string[]]$Values) {
    $filter = [uri]::EscapeDataString("clientId eq '$($Sp.id)' and resourceId eq '$($ResourceSp.id)'")
    $existing = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/oauth2PermissionGrants?`$filter=$filter" -OutputType PSObject).value | Where-Object { $_.consentType -eq 'AllPrincipals' })[0]
    $scope = (@($Values) + @("$($existing.scope)" -split ' ') | Where-Object { $_ } | Select-Object -Unique) -join ' '
    if ($existing) {
        Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/oauth2PermissionGrants/$($existing.id)" -Body (@{ scope = $scope } | ConvertTo-Json) -ContentType 'application/json' | Out-Null
    }
    else {
        $b = @{ clientId = $Sp.id; consentType = 'AllPrincipals'; resourceId = $ResourceSp.id; scope = $scope } | ConvertTo-Json
        Invoke-MgGraphRequest -Method POST -Uri 'v1.0/oauth2PermissionGrants' -Body $b -ContentType 'application/json' | Out-Null
    }
    Write-Host "  consent (delegated, all users): $scope"
}

$graphSp = Get-Sp $graphAppId
$defSp = $null
if ($useDefender) {
    try { $defSp = Get-Sp $defAppId }
    catch { throw 'Microsoft Defender (WindowsDefenderATP) is not set up in this tenant: open https://security.microsoft.com once with an administrator, then run this script again.' }
}

# 1. the app: the one in config.json, or a new one
$app = $null; $sp = $null
if ($clientId) {
    try { $app = Invoke-MgGraphRequest -Method GET -Uri "v1.0/applications(appId='$clientId')" -OutputType PSObject; $sp = Get-Sp $clientId; Write-Host "Updating the existing app $($app.displayName) ($clientId) ..." }
    catch { Write-Warning "App $clientId of config.json not found: a new one is created."; $app = $null }
}
if (-not $app) {
    Write-Host "Creating App Registration $AppName ..."
    $app = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/applications' -Body (@{ displayName = $AppName; signInAudience = 'AzureADMyOrg' } | ConvertTo-Json) -ContentType 'application/json' -OutputType PSObject
    $sp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/servicePrincipals' -Body (@{ appId = $app.appId } | ConvertTo-Json) -ContentType 'application/json' -OutputType PSObject
    Write-Host 'Waiting for replication...'
    Start-Sleep -Seconds 15
}

# 2. certificate (unattended mode)
$cert = $null; $cerFile = $null
if ($withCert) {
    Write-Host "Creating certificate CN=$AppName in Cert:\$Store\My ..."
    $cert = New-SelfSignedCertificate -Subject "CN=$AppName" -CertStoreLocation "Cert:\$Store\My" `
        -KeyExportPolicy NonExportable -KeySpec Signature -KeyLength 2048 -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($CertYears)
    if (-not @($app.keyCredentials).Count) {
        $key = @{ keyCredentials = @(@{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($cert.RawData); displayName = "CN=$AppName $env:USERNAME@$env:COMPUTERNAME" }) }
        Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/applications/$($app.id)" -Body ($key | ConvertTo-Json -Depth 5) -ContentType 'application/json' | Out-Null
        Write-Host '  certificate added to the app'
    }
    else {
        # existing certificates cannot be read back to send the full list: the new one is uploaded by hand
        $cerFile = Join-Path (Split-Path $ConfigPath -Parent) "$AppName-$env:USERNAME-$env:COMPUTERNAME.cer"
        Export-Certificate -Cert $cert -FilePath $cerFile | Out-Null
    }
}
$hasCert = $withCert -or @($app.keyCredentials).Count -gt 0

# 3. permissions requested by the app (existing ones are kept)
$access = [ordered]@{}
foreach ($r in @($app.requiredResourceAccess)) {
    if (-not $r) { continue }
    $access["$($r.resourceAppId)"] = [System.Collections.ArrayList]@(@($r.resourceAccess) | ForEach-Object { @{ id = "$($_.id)"; type = "$($_.type)" } })
}
$addAccess = {
    param([string]$ResourceAppId, [string]$Id, [string]$Type)
    if (-not $access.Contains($ResourceAppId)) { $access[$ResourceAppId] = [System.Collections.ArrayList]@() }
    if (-not @($access[$ResourceAppId] | Where-Object { $_.id -eq $Id -and $_.type -eq $Type }).Count) { [void]$access[$ResourceAppId].Add(@{ id = $Id; type = $Type }) }
}
$graphRoles = @(); $defRoles = @(); $defScopes = @()
if ($withCert) { $graphRoles = @(Get-AppRole $graphSp $graphPerms); foreach ($r in $graphRoles) { & $addAccess $graphAppId $r.id 'Role' } }
if ($defSp) {
    $defScopes = @(Get-Scope $defSp $defUserPerms); foreach ($s in $defScopes) { & $addAccess $defAppId $s.id 'Scope' }
    if ($hasCert) { $defRoles = @(Get-AppRole $defSp $defAppPerms); foreach ($r in $defRoles) { & $addAccess $defAppId $r.id 'Role' } }
}
$patch = @{ requiredResourceAccess = @($access.Keys | ForEach-Object { @{ resourceAppId = $_; resourceAccess = @($access[$_]) } }) }
if ($defSp) {
    # browser sign-in from HybridScope: platform "Mobile and desktop applications", redirect http://localhost
    $patch.publicClient = @{ redirectUris = @(@(@($app.publicClient.redirectUris) + 'http://localhost') | Where-Object { $_ } | Select-Object -Unique) }
}
Invoke-MgGraphRequest -Method PATCH -Uri "v1.0/applications/$($app.id)" -Body ($patch | ConvertTo-Json -Depth 8) -ContentType 'application/json' | Out-Null

# 4. admin consent
Write-Host 'Granting admin consent...'
if ($graphRoles.Count) { Grant-AppRole $sp $graphSp $graphRoles }
if ($defScopes.Count) { Grant-Scope $sp $defSp @($defScopes | ForEach-Object { "$($_.value)" }) }
if ($defRoles.Count) { Grant-AppRole $sp $defSp $defRoles }
Disconnect-MgGraph | Out-Null

# 5. config.json
function Set-CfgValue($Name, $Value) { if ($cfg.PSObject.Properties.Name -contains $Name) { $cfg.$Name = $Value } else { $cfg | Add-Member -NotePropertyName $Name -NotePropertyValue $Value } }
if ($cfg) {
    Set-CfgValue 'TenantId' $TenantId
    Set-CfgValue 'ClientId' $app.appId
    if ($cert) { Set-CfgValue 'CertificateThumbprint' $cert.Thumbprint; Set-CfgValue 'AuthMode' 'Certificate' }
    if ($defSp) { Set-CfgValue 'Defender' $true }
    $cfg | ConvertTo-Json -Depth 6 | Set-Content $ConfigPath -Encoding UTF8
    Write-Host 'config.json updated.'
}
else { Write-Host "config.json not found (start HybridScope once to create it): set ClientId = $($app.appId)$(if ($cert) { ", CertificateThumbprint = $($cert.Thumbprint), AuthMode = Certificate" })$(if ($defSp) { ', Defender = true' }) by hand." -ForegroundColor Yellow }

Write-Host ''
Write-Host 'DONE' -ForegroundColor Green
Write-Host "App        : $($app.displayName)"
Write-Host "ClientId   : $($app.appId)"
if ($cert) { Write-Host "Thumbprint : $($cert.Thumbprint)  (expires $($cert.NotAfter.ToString('yyyy-MM-dd')))" }
if ($cerFile) {
    Write-Host ''
    Write-Host "The app already has a certificate: upload $cerFile in" -ForegroundColor Yellow
    Write-Host "Entra admin center > App registrations > $($app.displayName) > Certificates & secrets > Certificates > Upload certificate" -ForegroundColor Yellow
}
if ($defSp) {
    Write-Host ''
    Write-Host 'Defender: tick the Defender box in HybridScope and press Load data; every user signs in in the browser with their own account.'
    Write-Host 'Users need a Defender role, e.g. Security Reader (Entra ID > Roles and administrators).'
    Write-Host 'Other PCs: copy ClientId (and TenantId) into their config.json, or copy the whole config.json. ClientId is not a secret.'
}
if ($cert) { Write-Host 'The GUI keeps using the browser sign-in; the certificate is used only by -Auto.' }
