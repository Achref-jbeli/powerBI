# run_daily.ps1 - Lancé par la tâche SQL Server Agent chaque jour.
# Charge les nouveaux dossiers, puis déclenche l'actualisation Power BI si de nouvelles données.
Set-Location $PSScriptRoot
$logDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force $logDir | Out-Null
$logFile = Join-Path $logDir ("run_{0:yyyy-MM-dd}.log" -f (Get-Date))

& "$PSScriptRoot\load_daily.ps1" *>> $logFile
$code = $LASTEXITCODE

if ($code -eq 0) {
    "Nouvelles données chargées -> actualisation Power BI" | Out-File $logFile -Append
    if (Test-Path "$PSScriptRoot\refresh_powerbi.ps1") { & "$PSScriptRoot\refresh_powerbi.ps1" *>> $logFile }
} elseif ($code -eq 2) {
    "Rien de nouveau -> pas d'actualisation" | Out-File $logFile -Append
} else {
    "ECHEC DU CHARGEMENT (voir ci-dessus)" | Out-File $logFile -Append
}
exit $code
