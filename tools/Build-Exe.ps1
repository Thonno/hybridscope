<#
.SYNOPSIS
    OPTIONAL: compiles HybridScope.ps1 into HybridScope.exe with PS2EXE (Windows PowerShell 5.1).

.NOTES
    Executables produced by PS2EXE are frequently flagged by antivirus / sandbox engines
    (false positives). Prefer Start-HybridScope.cmd, or sign the exe with your code-signing certificate.
#>
$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable -Name ps2exe)) { Install-Module ps2exe -Scope CurrentUser -Force }
Import-Module ps2exe
$root = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $root 'HybridScope.ps1'
# the exe CONTAINS the script: rebuild it after every update of HybridScope.ps1
$version = if ((Get-Content -Path $scriptPath -Raw) -match '\$HS_Version = ''(\d+\.\d+\.\d+)''') { "$($Matches[1]).0" } else { '1.0.0.0' }
$exeParams = @{
    inputFile   = $scriptPath
    outputFile  = Join-Path $root 'HybridScope.exe'
    noConsole   = $true
    STA         = $true
    title       = 'HybridScope'
    description = 'Hybrid identity and mailbox inventory: AD + Exchange + Microsoft 365'
    product     = 'HybridScope'
    company     = 'Thonno'
    copyright   = 'MIT License - github.com/Thonno/hybridscope'
    version     = $version
}
$icon = Join-Path $PSScriptRoot 'HybridScope.ico'
if (Test-Path $icon) { $exeParams.iconFile = $icon }
Invoke-ps2exe @exeParams
Write-Host "Created HybridScope.exe (version $version)" -ForegroundColor Green
Write-Host 'Recommended: sign it -> Set-AuthenticodeSignature -FilePath .\HybridScope.exe -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)[0] -TimestampServer http://timestamp.digicert.com'