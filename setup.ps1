# One-shot installation: Python packages + database objects + initial history load.
# Usage:  .\setup.ps1 [-Server "localhost\SQLEXPRESS"]
param([string]$Server = "localhost\SQLEXPRESS")
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

pip install -q -r requirements.txt
sqlcmd -S $Server -E -C -i "sql\01_schema.sql"        # creates DB + tables (wipes existing tables)
sqlcmd -S $Server -E -C -i "sql\02_daily_load.sql"    # LoadLog, usp_LoadDay, vw_DailyKpi
if (-not (Test-Path "data\sales_data.xlsx")) { python generate_sample_excel.py }
python load_excel_to_db.py
Write-Host "`nSetup done. Next: drop daily files in inbox\YYYY-MM-DD\ and run: python load_daily.py"
