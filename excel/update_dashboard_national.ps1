$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK

$xlExternal = 2
$xlPageField = 3
$xlRowField = 1
$xlColumnField = 2
$xlDataField = 4
$xlLine = 4
$xlLocationAsObject = 2
$xlValue = 2

function Say($m) { Write-Output $m }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

try {
    $wb = $excel.Workbooks.Open($path, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $false)
    Say "opened"

    # cleanup leftover test sheet
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.Name -eq "TmpDaxTest" -or $ws.Name -eq "TmpNational" -or $ws.Name -eq "TmpProbe" -or $ws.Name -eq "TmpExtract") {
            $ws.Delete()
            Say ("deleted leftover sheet: " + $ws.Name)
        }
    }

    $src = $wb.Worksheets("PivotSource")
    $src.Visible = $true
    $dash = $wb.Worksheets("Dashboard")

    $euro = [char]0x20AC
    $eurofmt = $euro + "0.000"

    $conn = $wb.Connections.Item("ThisWorkbookDataModel")
    function New-Pivot($cellRow, $cellCol, $name) {
        $pc = $wb.PivotCaches().Create($xlExternal, $conn)
        $dest = $src.Cells.Item($cellRow, $cellCol)
        return $pc.CreatePivotTable($dest, $name)
    }

    # --- Benzina: region vs national ---
    # Placed in their own columns (T, Z) so they can never collide with pt_trend / pt_regional as dates accumulate.
    $ptTB = New-Pivot 10 20 "pt_trendBenzinaNat"
    $ptTB.ColumnGrand = $false; $ptTB.RowGrand = $false
    ($ptTB.CubeFields("[Fuel price data daily].[Date]")).Orientation = $xlRowField
    ($ptTB.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPageField
    ($ptTB.CubeFields("[Fuel price data daily].[Denominazione Regione]")).Orientation = $xlPageField
    ($ptTB.CubeFields("[Fuel price data daily].[Bandiera]")).Orientation = $xlPageField
    ($ptTB.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    ($ptTB.CubeFields("[Measures].[NationalAvgPrezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 500
    $ptTB.PageFields("[Fuel price data daily].[descCarburante].[descCarburante]").CurrentPageName = "[Fuel price data daily].[descCarburante].&[Benzina]"
    Start-Sleep -Milliseconds 500
    Say "ptTB built"

    # --- Gasolio: region vs national ---
    $ptTG = New-Pivot 10 26 "pt_trendGasolioNat"
    $ptTG.ColumnGrand = $false; $ptTG.RowGrand = $false
    ($ptTG.CubeFields("[Fuel price data daily].[Date]")).Orientation = $xlRowField
    ($ptTG.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPageField
    ($ptTG.CubeFields("[Fuel price data daily].[Denominazione Regione]")).Orientation = $xlPageField
    ($ptTG.CubeFields("[Fuel price data daily].[Bandiera]")).Orientation = $xlPageField
    ($ptTG.CubeFields("[Measures].[Average of prezzo]")).Orientation = $xlDataField
    ($ptTG.CubeFields("[Measures].[NationalAvgPrezzo]")).Orientation = $xlDataField
    Start-Sleep -Milliseconds 500
    $ptTG.PageFields("[Fuel price data daily].[descCarburante].[descCarburante]").CurrentPageName = "[Fuel price data daily].[descCarburante].&[Gasolio]"
    Start-Sleep -Milliseconds 500
    Say "ptTG built"

    # --- connect existing slicers (Region, Brand) to the new pivots ---
    $scRegion = $null; $scBrand = $null
    foreach ($sc in $wb.SlicerCaches()) {
        if ($sc.Slicers(1).Caption -eq "Region") { $scRegion = $sc }
        if ($sc.Slicers(1).Caption -eq "Brand") { $scBrand = $sc }
    }
    [void]$scRegion.PivotTables.AddPivotTable($ptTB)
    [void]$scRegion.PivotTables.AddPivotTable($ptTG)
    [void]$scBrand.PivotTables.AddPivotTable($ptTB)
    [void]$scBrand.PivotTables.AddPivotTable($ptTG)
    Say "slicers connected to new pivots"

    # --- remove old combined trend chart ---
    $toDelete = $null
    foreach ($co in @($dash.ChartObjects())) {
        if ($co.Chart.HasTitle -and $co.Chart.ChartTitle.Text -eq "Price trend - Benzina vs Gasolio") {
            $toDelete = $co
        }
    }
    if ($toDelete -ne $null) {
        $toDelete.Delete()
        Say "deleted old combined trend chart"
    }

    # --- new charts: Benzina region vs national, Gasolio region vs national ---
    $chartTB = $wb.Charts.Add()
    [void]$chartTB.SetSourceData($ptTB.TableRange2)
    $chartTB.ChartType = $xlLine
    $chartTB.HasTitle = $true
    $chartTB.ChartTitle.Text = "Benzina - Region vs National Average"
    [void]$chartTB.Location($xlLocationAsObject, "Dashboard")
    $coTB = $dash.ChartObjects($dash.ChartObjects().Count)
    $coTB.Top = 360; $coTB.Left = 20; $coTB.Width = 560; $coTB.Height = 155
    $coTB.Chart.ShowReportFilterFieldButtons = $false; $coTB.Chart.ShowAxisFieldButtons = $false; $coTB.Chart.ShowLegendFieldButtons = $false; $coTB.Chart.ShowValueFieldButtons = $false
    $coTB.Chart.Axes($xlValue).TickLabels.NumberFormat = $eurofmt
    Say "chartTB created"

    $chartTG = $wb.Charts.Add()
    [void]$chartTG.SetSourceData($ptTG.TableRange2)
    $chartTG.ChartType = $xlLine
    $chartTG.HasTitle = $true
    $chartTG.ChartTitle.Text = "Gasolio - Region vs National Average"
    [void]$chartTG.Location($xlLocationAsObject, "Dashboard")
    $coTG = $dash.ChartObjects($dash.ChartObjects().Count)
    $coTG.Top = 525; $coTG.Left = 20; $coTG.Width = 560; $coTG.Height = 155
    $coTG.Chart.ShowReportFilterFieldButtons = $false; $coTG.Chart.ShowAxisFieldButtons = $false; $coTG.Chart.ShowLegendFieldButtons = $false; $coTG.Chart.ShowValueFieldButtons = $false
    $coTG.Chart.Axes($xlValue).TickLabels.NumberFormat = $eurofmt
    Say "chartTG created"

    $src.Visible = $false
    Say "re-hid PivotSource"

    $dash.Activate()
    $wb.Save()
    Say "SAVED"

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Say "UPDATE COMPLETE"
