# Registers a Windows scheduled task that runs run_daily.ps1 every day.
# Usage:  .\register_daily_task.ps1 -Time 06:30
param([string]$Time = "06:30")
$script = Join-Path $PSScriptRoot "run_daily.ps1"
$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
Register-ScheduledTask -TaskName "SalesKpi_DailyLoad" -Action $action -Trigger $trigger -Force | Out-Null
Write-Host "Task 'SalesKpi_DailyLoad' registered, runs daily at $Time. Remove with: Unregister-ScheduledTask SalesKpi_DailyLoad"
