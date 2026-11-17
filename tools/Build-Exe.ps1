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
Invoke-ps2exe -inputFile (Join-Path $root 'HybridScope.ps1') -outputFile (Join-Path $root 'HybridScope.exe') -noConsole -STA `
    -title 'HybridScope' -description 'Hybrid identity and mailbox inventory: AD + Exchange + Microsoft 365' -product 'HybridScope' -version '1.0.0.0'
Write-Host 'Created HybridScope.exe' -ForegroundColor Green
Write-Host 'Recommended: sign it -> Set-AuthenticodeSignature -FilePath .\HybridScope.exe -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)[0] -TimestampServer http://timestamp.digicert.com'
