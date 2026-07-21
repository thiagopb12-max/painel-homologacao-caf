$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open((Join-Path $PSScriptRoot "CAF.xlsx"), 0, $true)
foreach ($ws in $wb.Worksheets) { Write-Host "Aba: $($ws.Name)" }
$ws = $null
foreach ($s in $wb.Worksheets) { if ($s.Name -match "Resumo") { $ws = $s; break } }
if ($ws) {
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count
    Write-Host "---HEADERS (Cols: $lastCol, Rows: $lastRow)---"
    for ($c = 1; $c -le $lastCol; $c++) { Write-Host "Col $c : $($ws.Cells.Item(1,$c).Text)" }
    Write-Host "---PRIMEIRAS LINHAS---"
    for ($r = 2; $r -le [Math]::Min($lastRow, 6); $r++) {
        $line = ""
        for ($c = 1; $c -le $lastCol; $c++) { $line += "$($ws.Cells.Item($r,$c).Text) | " }
        Write-Host "Row $r : $line"
    }
} else { Write-Host "Aba ResumoDiario nao encontrada!" }
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
