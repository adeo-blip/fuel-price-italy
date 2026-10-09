$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK
$outJson = Join-Path $env:FUEL_OUTDIR "national.json"
$xlExternal = 2; $xlRowField = 1; $xlColumnField = 2; $xlDataField = 4

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    $wb = $excel.Workbooks.Open($path)
    $conn = $wb.Connections.Item("ThisWorkbookDataModel")
    $pc = $wb.PivotCaches().Create($xlExternal, $conn)
    $tmpSheet = $wb.Worksheets.Add()
    $tmpSheet.Name = "TmpNational"
    $pt = $pc.CreatePivotTable($tmpSheet.Cells.Item(3,1), "NationalPT")
    $pt.ColumnGrand = $false; $pt.RowGrand = $false
    ($pt.CubeFields("[Fuel price data daily].[Date]")).Orientation = $xlRowField
    ($pt.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlColumnField
    ($pt.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 500
    $pt.PivotFields("[Fuel price data daily].[descCarburante].[descCarburante]").VisibleItemsList = @(
        "[Fuel price data daily].[descCarburante].&[Benzina]",
        "[Fuel price data daily].[descCarburante].&[Gasolio]"
    )
    [void]$pt.RefreshTable()
    Start-Sleep -Milliseconds 500

    $dataRange = $pt.DataBodyRange; $rowHeaders = $pt.RowRange; $colHeaders = $pt.ColumnRange
    $benzinaCol = -1; $gasolioCol = -1
    for ($c = 1; $c -le $colHeaders.Columns.Count; $c++) {
        $v = $colHeaders.Cells.Item($colHeaders.Rows.Count, $c).Value2
        if ($v -eq "Benzina") { $benzinaCol = $c }
        if ($v -eq "Gasolio") { $gasolioCol = $c }
    }
    $rowOffset = $rowHeaders.Rows.Count - $dataRange.Rows.Count
    $rows = @()
    for ($r = 1; $r -le $dataRange.Rows.Count; $r++) {
        $rows += [PSCustomObject]@{
            date = $rowHeaders.Cells.Item($r + $rowOffset, 1).Text
            benzina = $dataRange.Cells.Item($r, $benzinaCol).Value2
            gasolio = $dataRange.Cells.Item($r, $gasolioCol).Value2
        }
    }
    $rows | ConvertTo-Json -Depth 5 | Out-File -FilePath $outJson -Encoding utf8
    Write-Output ("national rows: " + $rows.Count)
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
