# load_excel_to_db.ps1 - Chargement initial de l'historique (remplace load_excel_to_db.py).
# Efface et recharge toutes les tables staging + reconstruit le schéma en étoile.
# Usage : .\load_excel_to_db.ps1 [-XlsxPath "C:\...\sales_data.xlsx"] [-Server "localhost\SQLEXPRESS"]
param(
    [string]$XlsxPath = "$PSScriptRoot\data\sales_data.xlsx",
    [string]$Server   = "localhost\SQLEXPRESS",
    [string]$Database = "SalesKpiDW"
)
$ErrorActionPreference = "Stop"

function Read-ExcelSheet($xlsxPath, [string]$sheet) {
    $cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$xlsxPath;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'"
    $cn = New-Object System.Data.OleDb.OleDbConnection($cs)
    $cn.Open()
    $da = New-Object System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [$sheet`$]", $cn)
    $dt = New-Object System.Data.DataTable
    $da.Fill($dt) | Out-Null
    $cn.Close()
    return $dt
}

$sqlCs = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True"
$cx = New-Object System.Data.SqlClient.SqlConnection($sqlCs)
$cx.Open()

$stgMap = @{ "Customers"="stg_Customers"; "Products"="stg_Products"; "Sales"="stg_Sales"; "Targets"="stg_Targets" }
$stgMap.GetEnumerator() | ForEach-Object {
    $cmd = $cx.CreateCommand()
    $cmd.CommandText = "DELETE FROM $($_.Value)"; $cmd.ExecuteNonQuery() | Out-Null
    $dt = Read-ExcelSheet $XlsxPath $_.Key
    $bcp = New-Object System.Data.SqlClient.SqlBulkCopy($cx)
    $bcp.DestinationTableName = "dbo.$($_.Value)"
    foreach ($col in $dt.Columns) { $bcp.ColumnMappings.Add($col.ColumnName, $col.ColumnName) | Out-Null }
    $bcp.WriteToServer($dt)
    Write-Host "$($_.Key.PadRight(10)) -> $($_.Value.PadRight(14)) $($dt.Rows.Count) lignes"
}

$cmd2 = $cx.CreateCommand()
$cmd2.CommandText = "EXEC usp_BuildWarehouse"; $cmd2.CommandTimeout = 300; $cmd2.ExecuteNonQuery() | Out-Null
Write-Host "Entrepôt reconstruit."
$cx.Close()
