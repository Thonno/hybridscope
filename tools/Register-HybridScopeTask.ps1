<#
.SYNOPSIS
    Registers a daily scheduled task that runs HybridScope in headless mode (-Auto).
    Run from an elevated PowerShell on the machine that will produce the report.

.PARAMETER RunAs
    Account that runs the task (e.g. CONTOSO\svc-hybridscope). It must:
      - have the app certificate in its store (see New-HybridScopeApp.ps1)
      - be a member of the Exchange role group "View-Only Organization Management"
      - be able to write to OutputFolder

.EXAMPLE
    .\tools\Register-HybridScopeTask.ps1 -RunAs CONTOSO\svc-hybridscope -At 06:40
#>
param(
    [Parameter(Mandatory)][string]$RunAs,
    [string]$At = '06:40',
    [string]$TaskName = 'HybridScope - daily report'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$ps1 = Join-Path $root 'HybridScope.ps1'
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -File `"$ps1`" -Auto" -WorkingDirectory $root
$trigger = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew

$cred = Get-Credential -UserName $RunAs -Message 'Password of the account that will run the report'
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Limited -Force | Out-Null

Write-Host "Task '$TaskName' registered: every day at $At as $RunAs" -ForegroundColor Green
Write-Host "Test it now with: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host 'Note: the task account needs ExecutionPolicy RemoteSigned (Set-ExecutionPolicy -Scope CurrentUser RemoteSigned).'
