# generate_sample_excel.ps1 - Crée data\sales_data.xlsx via COM Excel (remplace generate_sample_excel.py).
# Prérequis : Microsoft Excel installé.
param([string]$Out = "$PSScriptRoot\data\sales_data.xlsx")
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null

$xl  = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
$wb  = $xl.Workbooks.Add()

function Add-Sheet($wb, [string]$name, $headers, $rows) {
    $ws = $wb.Worksheets.Add(); $ws.Name = $name
    for ($c=0; $c -lt $headers.Count; $c++) { $ws.Cells(1, $c+1) = $headers[$c] }
    for ($r=0; $r -lt $rows.Count; $r++) { for ($c=0; $c -lt $rows[$r].Count; $c++) { $ws.Cells($r+2, $c+1) = $rows[$r][$c] } }
}

$regions = @{Tunis="North";Sfax="South";Sousse="Center";Bizerte="North";Gabes="South"}
$cities  = @("Tunis","Sfax","Sousse","Bizerte","Gabes")
$segs    = @("Retail","Wholesale","Online")
$custRows = 1..40 | ForEach-Object { $c = $cities[($_ % 5)]; @($_, "Customer $("{0:D3}" -f $_)", $c, $segs[($_ % 3)], $regions[$c]) }
Add-Sheet $wb "Customers" @("CustomerID","CustomerName","City","Segment","Region") $custRows

$cats = @{Laptops=@(900,1800);Phones=@(300,1200);Accessories=@(10,120);Monitors=@(150,600)}
$prodRows = @(); $pid=1
foreach ($cat in $cats.Keys) {
    $lo,$hi = $cats[$cat]
    1..5 | ForEach-Object {
        $p = [Math]::Round($lo + (Get-Random -Max 1000)/1000*($hi-$lo), 2)
        $prodRows += ,@($pid, "$($cat.TrimEnd('s')) $_", $cat, $p, [Math]::Round($p*0.65,2))
        $pid++
    }
}
Add-Sheet $wb "Products" @("ProductID","ProductName","Category","UnitPrice","UnitCost") $prodRows

$start = [datetime]"2025-01-01"; $salesRows = @()
1..3000 | ForEach-Object {
    $p = $prodRows[(Get-Random -Max $prodRows.Count)]
    $salesRows += ,@($_, $start.AddDays((Get-Random -Max 600)).ToString("yyyy-MM-dd"), (Get-Random -Min 1 -Max 41), $p[0], (Get-Random -Min 1 -Max 11), @(0,0,0,0.05,0.10,0.15)[(Get-Random -Max 6)])
}
Add-Sheet $wb "Sales" @("OrderID","OrderDate","CustomerID","ProductID","Quantity","DiscountPct") $salesRows

$tgtRows = @(); $d = [datetime]"2025-01-01"
while ($d -le [datetime]"2026-08-01") { $tgtRows += ,@($d.ToString("yyyy-MM-dd"), (Get-Random -Min 180000 -Max 260001)); $d=$d.AddMonths(1) }
Add-Sheet $wb "Targets" @("MonthStart","RevenueTarget") $tgtRows

$wb.Worksheets.Item(1).Delete()   # supprime la feuille vide par défaut
if (Test-Path $Out) { Remove-Item $Out -Force }
$wb.SaveAs($Out, 51)   # 51 = xlOpenXMLWorkbook
$wb.Close(); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Créé : $Out"
