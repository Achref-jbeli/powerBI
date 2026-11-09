# generate_daily_files.ps1 - Crée des dossiers inbox\AAAA-MM-JJ de test (remplace generate_daily_files.py).
# Usage : .\generate_daily_files.ps1 -StartDate "2026-11-01" -Days 7
param([string]$StartDate="2026-11-01",[int]$Days=7,[string]$Inbox="$PSScriptRoot\inbox")
$start = [datetime]$StartDate
0..($Days-1) | ForEach-Object {
    $day = $start.AddDays($_); $dayStr = $day.ToString("yyyy-MM-dd")
    $xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = "Sales"
    $headers = @("OrderID","OrderDate","CustomerID","ProductID","Quantity","DiscountPct")
    for ($c=0;$c -lt $headers.Count;$c++) { $ws.Cells(1,$c+1) = $headers[$c] }
    $base = [int]($day.ToString("yyMMdd")) * 1000
    $n = Get-Random -Min 60 -Max 141
    for ($r=1;$r -le $n;$r++) {
        $ws.Cells($r+1,1) = $base+$r; $ws.Cells($r+1,2) = $dayStr
        $ws.Cells($r+1,3) = Get-Random -Min 1 -Max 41
        $ws.Cells($r+1,4) = Get-Random -Min 1 -Max 21
        $ws.Cells($r+1,5) = Get-Random -Min 1 -Max 11
        $ws.Cells($r+1,6) = @(0,0,0,0.05,0.10,0.15)[(Get-Random -Max 6)]
    }
    $folder = Join-Path $Inbox $dayStr; New-Item -ItemType Directory -Force $folder | Out-Null
    $path = Join-Path $folder "sales_$dayStr.xlsx"
    if (Test-Path $path) { Remove-Item $path -Force }
    $wb.SaveAs($path,51); $wb.Close(); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    Write-Host "Créé $dayStr : $n commandes"
}
