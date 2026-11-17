<#
.SYNOPSIS
    One-time setup for unattended mode (-Auto): creates an Entra ID App Registration with a
    certificate and READ-ONLY Microsoft Graph application permissions, grants admin consent and
    writes ClientId / Thumbprint / AuthMode=Certificate into config.json.

.DESCRIPTION
    Run it as a Global Administrator, with the SAME Windows account that will run the scheduled
    task (CurrentUser store) - or use -Store LocalMachine from an elevated prompt and give the task
    account read access to the private key.

.EXAMPLE
    .\tools\New-HybridScopeApp.ps1 -TenantId contoso.onmicrosoft.com
#>
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$AppName = 'HybridScope',
    [int]$CertYears = 2,
    [ValidateSet('CurrentUser', 'LocalMachine')][string]$Store = 'CurrentUser',
    [string]$ConfigPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config.json')
)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication

$perms = 'User.Read.All', 'AuditLog.Read.All', 'Reports.Read.All', 'UserAuthenticationMethod.Read.All', 'Organization.Read.All', 'Policy.Read.All'
$graphAppId = '00000003-0000-0000-c000-000000000000'

try { Set-MgGraphOption -DisableLoginByWAM $true } catch { }
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All' -NoWelcome

Write-Host "Creating certificate CN=$AppName in Cert:\$Store\My ..."
$cert = New-SelfSignedCertificate -Subject "CN=$AppName" -CertStoreLocation "Cert:\$Store\My" `
    -KeyExportPolicy NonExportable -KeySpec Signature -KeyLength 2048 -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($CertYears)

$graphSp = Invoke-MgGraphRequest -Method GET -Uri "v1.0/servicePrincipals(appId='$graphAppId')" -OutputType PSObject
$roles = foreach ($p in $perms) {
    $r = $graphSp.appRoles | Where-Object { $_.value -eq $p -and $_.allowedMemberTypes -contains 'Application' }
    if (-not $r) { throw "Application permission $p not found" }
    $r
}

Write-Host "Creating App Registration $AppName ..."
$body = @{
    displayName            = $AppName
    signInAudience         = 'AzureADMyOrg'
    keyCredentials         = @(@{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($cert.RawData); displayName = "CN=$AppName" })
    requiredResourceAccess = @(@{ resourceAppId = $graphAppId; resourceAccess = @($roles | ForEach-Object { @{ id = $_.id; type = 'Role' } }) })
}
$app = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/applications' -Body ($body | ConvertTo-Json -Depth 6) -ContentType 'application/json' -OutputType PSObject
$sp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/servicePrincipals' -Body (@{ appId = $app.appId } | ConvertTo-Json) -ContentType 'application/json' -OutputType PSObject

Write-Host 'Waiting for replication, then granting admin consent...'
Start-Sleep -Seconds 15
foreach ($r in $roles) {
    $b = @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $r.id } | ConvertTo-Json
    Invoke-MgGraphRequest -Method POST -Uri "v1.0/servicePrincipals/$($sp.id)/appRoleAssignedTo" -Body $b -ContentType 'application/json' | Out-Null
    Write-Host "  consent: $($r.value)"
}
Disconnect-MgGraph | Out-Null

if (Test-Path $ConfigPath) {
    $cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $cfg.TenantId = $TenantId
    $cfg.ClientId = $app.appId
    $cfg.CertificateThumbprint = $cert.Thumbprint
    $cfg.AuthMode = 'Certificate'
    $cfg | ConvertTo-Json -Depth 6 | Set-Content $ConfigPath -Encoding UTF8
    Write-Host "config.json updated."
}
else { Write-Host "config.json not found: set ClientId, CertificateThumbprint and AuthMode=Certificate manually." -ForegroundColor Yellow }

Write-Host ''
Write-Host 'DONE' -ForegroundColor Green
Write-Host "ClientId   : $($app.appId)"
Write-Host "Thumbprint : $($cert.Thumbprint)  (expires $($cert.NotAfter.ToString('yyyy-MM-dd')))"
Write-Host 'The GUI keeps using interactive sign-in; the certificate is used only by -Auto.'
