# Daily job: load new date folders into SQL Server, then refresh the Power BI dataset if something new arrived.
# Scheduled by register_daily_task.ps1
Set-Location $PSScriptRoot
$log = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force $log | Out-Null
$logFile = Join-Path $log ("run_{0:yyyy-MM-dd}.log" -f (Get-Date))

python load_daily.py *>> $logFile
$code = $LASTEXITCODE

if ($code -eq 0) {
    "New data loaded -> refreshing Power BI" | Out-File $logFile -Append
    if (Test-Path "$PSScriptRoot\refresh_powerbi.ps1") { & "$PSScriptRoot\refresh_powerbi.ps1" *>> $logFile }
} elseif ($code -eq 2) {
    "Nothing new -> no refresh needed" | Out-File $logFile -Append
} else {
    "LOAD FAILED (see above)" | Out-File $logFile -Append
}
exit $code
