# load_daily.ps1 - Remplace load_daily.py : aucun Python ni pip requis.
# Lit les fichiers Excel via OleDb (.NET natif), charge SQL Server via SqlBulkCopy.
# Prérequis : Microsoft Access Database Engine 2016 (ACE) - gratuit, sans pip :
#   https://www.microsoft.com/en-us/download/details.aspx?id=54920
# Usage : .\load_daily.ps1 [-Inbox "C:\...\inbox"] [-Server "localhost\SQLEXPRESS"]
param(
    [string]$Inbox  = "$PSScriptRoot\inbox",
    [string]$Server = "localhost\SQLEXPRESS",
    [string]$Database = "SalesKpiDW"
)
$ErrorActionPreference = "Stop"
$loaded = 0; $failed = 0

function Is-AlreadyLoaded($cx, [string]$day) {
    $cmd = $cx.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(1) FROM LoadLog WHERE LoadDate=@d AND Status='Success'"
    $cmd.Parameters.AddWithValue("@d", $day) | Out-Null
    return [int]$cmd.ExecuteScalar() -gt 0
}

function Read-ExcelSheet($xlsxPath, [string]$sheet) {
    $cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$xlsxPath;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'"
    $cn = New-Object System.Data.OleDb.OleDbConnection($cs)
    $cn.Open()
    $da = New-Object System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [$sheet`$]", $cn)
    $dt = New-Object System.Data.DataTable
    try { $da.Fill($dt) | Out-Null; return $dt } catch { return $null } finally { $cn.Close() }
}

function BulkLoad($sqlCx, [System.Data.DataTable]$dt, [string]$table) {
    $bcp = New-Object System.Data.SqlClient.SqlBulkCopy($sqlCx)
    $bcp.DestinationTableName = "dbo.$table"
    $bcp.BulkCopyTimeout = 120
    foreach ($col in $dt.Columns) { $bcp.ColumnMappings.Add($col.ColumnName, $col.ColumnName) | Out-Null }
    $bcp.WriteToServer($dt)
}

function Log-Failure($cx, [string]$day, [string]$file, [string]$msg) {
    $cmd = $cx.CreateCommand()
    $cmd.CommandText = "INSERT LoadLog (LoadDate,FileName,RowsLoaded,Status,Message) VALUES (@d,@f,0,'Failed',@m)"
    $cmd.Parameters.AddWithValue("@d", $day) | Out-Null
    $cmd.Parameters.AddWithValue("@f", $file) | Out-Null
    $cmd.Parameters.AddWithValue("@m", $msg.Substring(0, [Math]::Min(900,$msg.Length))) | Out-Null
    $cmd.ExecuteNonQuery() | Out-Null
}

if (-not (Test-Path $Inbox)) { Write-Error "Inbox introuvable : $Inbox"; exit 1 }

$sqlCs = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True"
$mainCx = New-Object System.Data.SqlClient.SqlConnection($sqlCs)
$mainCx.Open()

$stgMap = @{ "Customers"="stg_Customers"; "Products"="stg_Products"; "Sales"="stg_Sales"; "Targets"="stg_Targets" }

Get-ChildItem $Inbox -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } | Sort-Object Name | ForEach-Object {
    $day  = $_.Name
    $xlsx = Get-ChildItem $_.FullName -Filter "*.xlsx" | Select-Object -First 1
    if (-not $xlsx) { Write-Host "[$day] aucun .xlsx -> ignoré"; return }
    if (Is-AlreadyLoaded $mainCx $day) { Write-Host "[$day] déjà chargé -> ignoré"; return }

    try {
        $cx = New-Object System.Data.SqlClient.SqlConnection($sqlCs)
        $cx.Open()
        $tx = $cx.BeginTransaction()
        $stgMap.GetEnumerator() | ForEach-Object {
            $cmd = $cx.CreateCommand(); $cmd.Transaction = $tx
            $cmd.CommandText = "DELETE FROM $($_.Value)"; $cmd.ExecuteNonQuery() | Out-Null
            $dt = Read-ExcelSheet $xlsx.FullName $_.Key
            if ($dt -and $dt.Rows.Count -gt 0) {
                $bcp = New-Object System.Data.SqlClient.SqlBulkCopy($cx, [System.Data.SqlClient.SqlBulkCopyOptions]::Default, $tx)
                $bcp.DestinationTableName = "dbo.$($_.Value)"
                $bcp.BulkCopyTimeout = 120
                foreach ($col in $dt.Columns) { $bcp.ColumnMappings.Add($col.ColumnName, $col.ColumnName) | Out-Null }
                $bcp.WriteToServer($dt)
            }
        }
        $cmd2 = $cx.CreateCommand(); $cmd2.Transaction = $tx
        $cmd2.CommandText = "EXEC usp_LoadDay @d, @f"
        $cmd2.Parameters.AddWithValue("@d", $day) | Out-Null
        $cmd2.Parameters.AddWithValue("@f", $xlsx.Name) | Out-Null
        $cmd2.ExecuteNonQuery() | Out-Null
        $tx.Commit(); $cx.Close()

        $rows = (& sqlcmd -S $Server -E -C -d $Database -Q "SELECT TOP 1 RowsLoaded FROM LoadLog WHERE LoadDate='$day' ORDER BY LogID DESC" -h -1 -W).Trim()
        Write-Host "[$day] chargé : $rows nouvelles commandes depuis $($xlsx.Name)"
        $loaded++
    } catch {
        try { $tx.Rollback() } catch {}
        try { $cx.Close() } catch {}
        Write-Host "[$day] ECHEC : $_"
        Log-Failure $mainCx $day $xlsx.Name "$_"
        $failed++
    }
}
$mainCx.Close()
Write-Host "Terminé. chargé=$loaded echecs=$failed"
exit $(if ($failed -gt 0) { 1 } elseif ($loaded -gt 0) { 0 } else { 2 })
