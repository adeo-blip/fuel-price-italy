$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK
$outJson = Join-Path $env:FUEL_OUTDIR "trends.json"

$xlExternal = 2
$xlPageField = 3
$xlRowField = 1
$xlColumnField = 2
$xlDataField = 4
$xlHidden = 0
$xlAverage = -4106

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

try {
    $wb = $excel.Workbooks.Open($path, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $false)

    $conn = $wb.Connections.Item("ThisWorkbookDataModel")
    $pc = $wb.PivotCaches().Create($xlExternal, $conn)

    $tmpSheet = $wb.Worksheets.Add()
    $tmpSheet.Name = "TmpExtract"
    $destRange = $tmpSheet.Cells.Item(3,1)
    $pt = $pc.CreatePivotTable($destRange, "ExtractPT")
    $pt.ColumnGrand = $false
    $pt.RowGrand = $false

    # Row field: Date
    $dateField = $pt.CubeFields("[Fuel price data daily].[Date]")
    $dateField.Orientation = $xlRowField

    # Column field: descCarburante, limited to Benzina + Gasolio
    $fuelField = $pt.CubeFields("[Fuel price data daily].[descCarburante]")
    $fuelField.Orientation = $xlColumnField
    Start-Sleep -Milliseconds 500
    $fuelPF = $pt.PivotFields("[Fuel price data daily].[descCarburante].[descCarburante]")
    $fuelPF.VisibleItemsList = @(
        "[Fuel price data daily].[descCarburante].&[Benzina]",
        "[Fuel price data daily].[descCarburante].&[Gasolio]"
    )

    # Page field: Denominazione Regione
    $regionField = $pt.CubeFields("[Fuel price data daily].[Denominazione Regione]")
    $regionField.Orientation = $xlPageField
    Start-Sleep -Milliseconds 500

    # Data field: Average of prezzo
    $avgField = $pt.CubeFields("[Measures].[Average of prezzo]")
    $avgField.Orientation = $xlDataField

    Start-Sleep -Milliseconds 800
    $pt.RefreshTable()
    Start-Sleep -Milliseconds 500

    # Enumerate region items live from the cube (avoids script-file encoding issues with accented names)
    $regionField.Orientation = $xlRowField
    Start-Sleep -Milliseconds 500
    $regionLongName = "[Fuel price data daily].[Denominazione Regione].[Denominazione Regione]"
    $regionPF = $pt.PivotFields($regionLongName)
    $regionItems = @()
    foreach ($it in $regionPF.PivotItems()) {
        $n = $it.Name
        $cap = $n.Substring($n.LastIndexOf("&[") + 2)
        $cap = $cap.Substring(0, $cap.Length - 1)
        $regionItems += [PSCustomObject]@{ uniqueName = $n; caption = $cap }
    }
    $regionField.Orientation = $xlPageField
    Start-Sleep -Milliseconds 500

    $pageFieldObj = $pt.PageFields(1)
    $results = @{}

    foreach ($regionItem in $regionItems) {
        $region = $regionItem.caption
        $itemName = $regionItem.uniqueName
        $pageFieldObj.CurrentPageName = $itemName
        $pt.RefreshTable()
        Start-Sleep -Milliseconds 300

        $rows = @()
        $dataRange = $pt.DataBodyRange
        $rowHeaders = $pt.RowRange
        $colHeaders = $pt.ColumnRange

        # Determine column order for Benzina / Gasolio
        $colVals = $colHeaders.Value2
        $benzinaCol = -1
        $gasolioCol = -1
        if ($colHeaders.Columns.Count -eq 1) {
            # only one visible fuel column, figure out which by checking pivotfield visible item
        }
        for ($c = 1; $c -le $colHeaders.Columns.Count; $c++) {
            $cellVal = $colHeaders.Cells.Item($colHeaders.Rows.Count, $c).Value2
            if ($cellVal -eq "Benzina") { $benzinaCol = $c }
            if ($cellVal -eq "Gasolio") { $gasolioCol = $c }
        }

        $rowOffset = $rowHeaders.Rows.Count - $dataRange.Rows.Count
        $nRows = $dataRange.Rows.Count
        for ($r = 1; $r -le $nRows; $r++) {
            $dateVal = $rowHeaders.Cells.Item($r + $rowOffset, 1).Text
            $benzina = $null
            $gasolio = $null
            if ($benzinaCol -gt 0) { $benzina = $dataRange.Cells.Item($r,$benzinaCol).Value2 }
            if ($gasolioCol -gt 0) { $gasolio = $dataRange.Cells.Item($r,$gasolioCol).Value2 }
            $rows += [PSCustomObject]@{ date = $dateVal; benzina = $benzina; gasolio = $gasolio }
        }
        $results[$region] = $rows
        Write-Output ("done region: " + $region + " rows=" + $rows.Count)
        $results | ConvertTo-Json -Depth 5 | Out-File -FilePath $outJson -Encoding utf8
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "EXTRACTION COMPLETE"
