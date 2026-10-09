$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK

$xlExternal = 2
$xlPageField = 3
$xlRowField = 1
$xlColumnField = 2
$xlDataField = 4
$xlLine = 4
$xlColumnClustered = 51
$xlLocationAsObject = 2

function Say($m) { Write-Output $m }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

try {
    $wb = $excel.Workbooks.Open($path, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $false)
    Say "opened"

    $orig = $wb.Worksheets("Dashboard")
    $orig.Name = "Dashboard (Original)"
    Say "renamed original"

    $lastSheet = $wb.Worksheets.Item($wb.Worksheets.Count)
    $src = $wb.Worksheets.Add([Type]::Missing, $lastSheet)
    $src.Name = "PivotSource"
    Say "added PivotSource"

    $firstSheet = $wb.Worksheets.Item(1)
    $dash = $wb.Worksheets.Add($firstSheet)
    $dash.Name = "Dashboard"
    Say "added Dashboard sheet (first)"

    $conn = $wb.Connections.Item("ThisWorkbookDataModel")

    function New-Pivot($cellRow, $cellCol, $name) {
        $pc = $wb.PivotCaches().Create($xlExternal, $conn)
        $dest = $src.Cells.Item($cellRow, $cellCol)
        return $pc.CreatePivotTable($dest, $name)
    }

    $ptB = New-Pivot 2 1 "pt_kpiBenzina"
    $ptB.ColumnGrand = $false; $ptB.RowGrand = $false
    ($ptB.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPageField
    ($ptB.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 300
    $ptB.PageFields(1).CurrentPageName = "[Fuel price data daily].[descCarburante].&[Benzina]"
    Start-Sleep -Milliseconds 300
    $bAddr = $ptB.DataBodyRange.Cells.Item(1,1).Address()
    Say ("ptB built, value at " + $bAddr)

    $ptG = New-Pivot 2 4 "pt_kpiGasolio"
    $ptG.ColumnGrand = $false; $ptG.RowGrand = $false
    ($ptG.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPageField
    ($ptG.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 300
    $ptG.PageFields(1).CurrentPageName = "[Fuel price data daily].[descCarburante].&[Gasolio]"
    Start-Sleep -Milliseconds 300
    $gAddr = $ptG.DataBodyRange.Cells.Item(1,1).Address()
    Say ("ptG built, value at " + $gAddr)

    $ptC = New-Pivot 2 7 "pt_kpiCount"
    $ptC.ColumnGrand = $false; $ptC.RowGrand = $false
    ($ptC.CubeFields("[Measures].[Stations]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 300
    Start-Sleep -Milliseconds 300
    $cAddr = $ptC.DataBodyRange.Cells.Item(1,1).Address()
    Say ("ptC built, value at " + $cAddr)

    $ptTrend = New-Pivot 10 1 "pt_trend"
    $ptTrend.ColumnGrand = $false; $ptTrend.RowGrand = $false
    ($ptTrend.CubeFields("[Fuel price data daily].[Date]")).Orientation = $xlRowField
    ($ptTrend.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlColumnField
    ($ptTrend.CubeFields("[Fuel price data daily].[Denominazione Regione]")).Orientation = $xlPageField
    ($ptTrend.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 500
    $ptTrend.PivotFields("[Fuel price data daily].[descCarburante].[descCarburante]").VisibleItemsList = @(
        "[Fuel price data daily].[descCarburante].&[Benzina]",
        "[Fuel price data daily].[descCarburante].&[Gasolio]"
    )
    Start-Sleep -Milliseconds 500
    Say "ptTrend built"

    $ptReg = New-Pivot 10 10 "pt_regional"
    $ptReg.ColumnGrand = $false; $ptReg.RowGrand = $false
    ($ptReg.CubeFields("[Fuel price data daily].[Denominazione Regione]")).Orientation = $xlRowField
    ($ptReg.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlColumnField
    ($ptReg.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 500
    $ptReg.PivotFields("[Fuel price data daily].[descCarburante].[descCarburante]").VisibleItemsList = @(
        "[Fuel price data daily].[descCarburante].&[Benzina]",
        "[Fuel price data daily].[descCarburante].&[Gasolio]"
    )
    Start-Sleep -Milliseconds 500
    Say "ptReg built"

    # --- Charts (create before hiding PivotSource) ---
    $chartTrend = $wb.Charts.Add()
    [void]$chartTrend.SetSourceData($ptTrend.TableRange2)
    $chartTrend.ChartType = $xlLine
    $chartTrend.HasTitle = $true
    $chartTrend.ChartTitle.Text = "Price trend - Benzina vs Gasolio"
    [void]$chartTrend.Location($xlLocationAsObject, "Dashboard")
    $coTrend = $dash.ChartObjects($dash.ChartObjects().Count)
    $coTrend.Top = 360; $coTrend.Left = 20; $coTrend.Width = 560; $coTrend.Height = 320
    $coTrend.Chart.ShowReportFilterFieldButtons = $false; $coTrend.Chart.ShowAxisFieldButtons = $false; $coTrend.Chart.ShowLegendFieldButtons = $false; $coTrend.Chart.ShowValueFieldButtons = $false
    Say "chartTrend created and positioned"

    $chartReg = $wb.Charts.Add()
    [void]$chartReg.SetSourceData($ptReg.TableRange2)
    $chartReg.ChartType = $xlColumnClustered
    $chartReg.HasTitle = $true
    $chartReg.ChartTitle.Text = "Average price by region - Benzina vs Gasolio"
    [void]$chartReg.Location($xlLocationAsObject, "Dashboard")
    $coReg = $dash.ChartObjects($dash.ChartObjects().Count)
    $coReg.Top = 360; $coReg.Left = 600; $coReg.Width = 560; $coReg.Height = 320
    $coReg.Chart.ShowReportFilterFieldButtons = $false; $coReg.Chart.ShowAxisFieldButtons = $false; $coReg.Chart.ShowLegendFieldButtons = $false; $coReg.Chart.ShowValueFieldButtons = $false
    Say "chartReg created and positioned"

    # --- Slicers ---
    $scRegion = $wb.SlicerCaches().Add2($ptTrend, "[Fuel price data daily].[Denominazione Regione]")
    $slicerRegion = $scRegion.Slicers.Add($dash)
    $slicerRegion.Caption = "Region"
    $slicerRegion.Top = 40; $slicerRegion.Left = 20; $slicerRegion.Width = 140; $slicerRegion.Height = 300
    [void]$scRegion.PivotTables.AddPivotTable($ptB)
    [void]$scRegion.PivotTables.AddPivotTable($ptG)
    [void]$scRegion.PivotTables.AddPivotTable($ptC)
    Say "region slicer added"

    $scBrand = $wb.SlicerCaches().Add2($ptTrend, "[Fuel price data daily].[Bandiera]")
    $slicerBrand = $scBrand.Slicers.Add($dash)
    $slicerBrand.Caption = "Brand"
    $slicerBrand.Top = 40; $slicerBrand.Left = 170; $slicerBrand.Width = 140; $slicerBrand.Height = 300
    [void]$scBrand.PivotTables.AddPivotTable($ptB)
    [void]$scBrand.PivotTables.AddPivotTable($ptG)
    [void]$scBrand.PivotTables.AddPivotTable($ptC)
    [void]$scBrand.PivotTables.AddPivotTable($ptReg)
    Say "brand slicer added"

    $src.Visible = $false
    Say "hid PivotSource"

    # --- KPI display cells (formulas referencing pivot value cells) ---
    $dash.Activate()
    $dash.Cells.Item(1,1).Value2 = "Fuel Price Dashboard - Italy"
    $dash.Cells.Item(1,1).Font.Size = 22
    $dash.Cells.Item(1,1).Font.Bold = $true
    $dash.Cells.Item(2,1).Value2 = "Live from the MIMIT open-data feed - pick a region and brand to drill in"
    $dash.Cells.Item(2,1).Font.Italic = $true
    $dash.Cells.Item(2,1).Font.Size = 11

    $euro = [char]0x20AC
    $eurofmt = $euro + "0.000"

    $kpiCol = 20
    $labels = @("Avg Benzina", "Avg Gasolio", "Gasolio minus Benzina", "Stations (distinct)")
    for ($i=0; $i -lt 4; $i++) {
        $c = $kpiCol + $i*3
        $lc = $dash.Cells.Item(2, $c)
        $lc.Value2 = $labels[$i]
        $lc.Font.Size = 10
        $lc.Font.Color = 8421504
        $vcell = $dash.Cells.Item(3, $c)
        $vcell.Font.Size = 22
        $vcell.Font.Bold = $true
    }
    $dash.Cells.Item(3, $kpiCol).Formula = "='PivotSource'!" + $bAddr
    $dash.Cells.Item(3, $kpiCol).NumberFormat = $eurofmt
    $dash.Cells.Item(3, $kpiCol+3).Formula = "='PivotSource'!" + $gAddr
    $dash.Cells.Item(3, $kpiCol+3).NumberFormat = $eurofmt
    $gapCellAddr1 = $dash.Cells.Item(3, $kpiCol+3).Address($false,$false)
    $gapCellAddr0 = $dash.Cells.Item(3, $kpiCol).Address($false,$false)
    $dash.Cells.Item(3, $kpiCol+6).Formula = "=" + $gapCellAddr1 + "-" + $gapCellAddr0
    $dash.Cells.Item(3, $kpiCol+6).NumberFormat = $eurofmt
    $dash.Cells.Item(3, $kpiCol+9).Formula = "='PivotSource'!" + $cAddr
    $dash.Cells.Item(3, $kpiCol+9).NumberFormat = "#,##0"
    Say "KPIs written"

    $excel.ActiveWindow.DisplayGridlines = $false

    $wb.Save()
    Say "SAVED"

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Say "DASHBOARD BUILD COMPLETE"
