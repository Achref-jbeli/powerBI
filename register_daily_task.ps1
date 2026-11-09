# register_daily_task.ps1 - Alternative au job SQL Server Agent.
# Crée une tâche Windows planifiée si SQL Server Agent n'est pas disponible (Express).
# Note : SQL Server Express n'inclut pas SQL Server Agent -> utiliser cette tâche Windows à la place.
param([string]$Time = "06:30")
$script  = Join-Path $PSScriptRoot "run_daily.ps1"
$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
Register-ScheduledTask -TaskName "SalesKpi_DailyLoad" -Action $action -Trigger $trigger -Force | Out-Null
Write-Host "Tâche 'SalesKpi_DailyLoad' créée, se lance chaque jour à $Time."
Write-Host "Pour supprimer : Unregister-ScheduledTask -TaskName SalesKpi_DailyLoad -Confirm:`$false"
