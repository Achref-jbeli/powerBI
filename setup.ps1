# setup.ps1 - Installation complète en une commande. Aucun Python ni pip requis.
# Usage : .\setup.ps1 [-Server "localhost\SQLEXPRESS"]
param([string]$Server = "localhost\SQLEXPRESS")
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "1/3 Création de la base de données et des tables..."
sqlcmd -S $Server -E -C -i "sql\01_schema.sql"
sqlcmd -S $Server -E -C -i "sql\02_daily_load.sql"

Write-Host "2/3 Génération du fichier Excel d'exemple..."
if (-not (Test-Path "data\sales_data.xlsx")) { & ".\generate_sample_excel.ps1" }

Write-Host "3/3 Chargement de l'historique initial..."
& ".\load_excel_to_db.ps1" -Server $Server

Write-Host "`nInstallation terminée."
Write-Host "Testez avec : .\generate_daily_files.ps1 -StartDate '2026-11-01' -Days 7"
Write-Host "Puis lancez : .\load_daily.ps1"
