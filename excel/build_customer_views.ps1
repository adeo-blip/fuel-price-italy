$ErrorActionPreference = "Stop"
# Adds three customer-facing views to the workbook (run after add_national_measure.ps1):
#   "Place Trend"      - Benzina/Gasolio/GPL/Metano trend for a Region / Province / Comune, against the Italy average
#   "Cheapest Brand"   - the 10 cheapest brands (bandiere) per fuel for the selected place
#   "Regions vs Italy" - region-by-week heat table per fuel (Italy average = total row) + national average trend
$path = $env:FUEL_WORKBOOK
$xlExternal = 2; $xlRow = 1; $xlColumn = 2; $xlPage = 3; $xlData = 4
$xlLine = 4; $xlBarClustered = 57; $xlLocationAsObject = 2; $xlValue = 2; $xlCategory = 1
$euro = [char]0x20AC
$eurofmt = $euro + "0.00"
$FUELS = @("Benzina", "Gasolio", "GPL", "Metano")
$FUEL_COLOR = @{ "Benzina" = 14055466; "Gasolio" = 8040731; "GPL" = 41453; "Metano" = 33792 }   # BGR of #2a78d6 #1baf7a #eda100 #008300

function Say($m) { Write-Output ("[" + (Get-Date -Format "HH:mm:ss") + "] " + $m) }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    $wb = $excel.Workbooks.Open($path)
    $conn = $wb.Connections.Item("ThisWorkbookDataModel")

    $src = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq "PivotSource") { $src = $w } }
    $src.Visible = $true
    $dash = $wb.Worksheets("Dashboard")

    function New-ViewSheet($name, $after) {
        $s = $wb.Worksheets.Add([Type]::Missing, $after)
        $s.Name = $name
        $s.Activate()
        $excel.ActiveWindow.DisplayGridlines = $false
        return $s
    }
    # Each pivot gets its own cache (a second pivot on one cache fails). ManualUpdate stays on while the
    # fields are placed, otherwise every field change recomputes the DAX. Never call RefreshTable here: it reloads the
    # whole cache and costs minutes per pivot, while switching ManualUpdate off already computes the pivot.
    function New-Pivot($row, $col, $name, $sheet = $src) {
        $pc = $wb.PivotCaches().Create($xlExternal, $conn)
        $pt = $pc.CreatePivotTable($sheet.Cells.Item($row, $col), $name)
        $pt.ManualUpdate = $true
        $pt.ColumnGrand = $false; $pt.RowGrand = $false
        return $pt
    }
    function Set-Fuel($pt, $fuel) {
        $pt.PageFields("[Fuel price data daily].[descCarburante].[descCarburante]").CurrentPageName = "[Fuel price data daily].[descCarburante].&[" + $fuel + "]"
    }
    function Add-Title($sheet, $title, $note) {
        $sheet.Cells.Item(1, 1).Value2 = $title
        $sheet.Cells.Item(1, 1).Font.Size = 22
        $sheet.Cells.Item(1, 1).Font.Bold = $true
        $sheet.Cells.Item(2, 1).Value2 = $note
        $sheet.Cells.Item(2, 1).Font.Italic = $true
        $sheet.Cells.Item(2, 1).Font.Size = 11
    }
    function Move-Chart($chart, $sheetName, $top, $left, $width, $height) {
        [void]$chart.Location($xlLocationAsObject, $sheetName)
        $sheet = $wb.Worksheets($sheetName)
        $co = $sheet.ChartObjects($sheet.ChartObjects().Count)
        $co.Top = $top; $co.Left = $left; $co.Width = $width; $co.Height = $height
        $c = $co.Chart
        $c.ShowReportFilterFieldButtons = $false; $c.ShowAxisFieldButtons = $false
        $c.ShowLegendFieldButtons = $false; $c.ShowValueFieldButtons = $false
        return $co
    }

    function Set-AxisFromData($co, $dataRange, $margin) {
        $mn = $excel.WorksheetFunction.Min($dataRange)
        $mx = $excel.WorksheetFunction.Max($dataRange)
        $axis = $co.Chart.Axes($xlValue)
        $axis.MinimumScale = [Math]::Floor(($mn - $margin) * 10) / 10
        $axis.MaximumScale = [Math]::Ceiling(($mx + $margin) * 10) / 10
    }

    $sheetPlace = New-ViewSheet "Place Trend" $dash
    $sheetCheap = New-ViewSheet "Cheapest Brand" $sheetPlace
    $sheetRegions = New-ViewSheet "Regions vs Italy" $sheetCheap
    Say "view sheets added"

    $placePivots = @(); $cheapPivots = @(); $heatPivots = @(); $natPivots = @()
    $rowBase = 100
    $i = 0
    foreach ($fuel in $FUELS) {
        $col = 1 + $i * 9

        # --- place trend: Date x (This place, Italy average)
        $pt = New-Pivot $rowBase $col ("pt_place_" + $fuel)
        ($pt.CubeFields("[Fuel price data daily].[Date]")).Orientation = $xlRow
        ($pt.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPage
        ($pt.CubeFields("[Measures].[This place]")).Orientation = $xlData
        ($pt.CubeFields("[Measures].[Italy average]")).Orientation = $xlData
        Start-Sleep -Milliseconds 400
        Set-Fuel $pt $fuel
        $pt.ManualUpdate = $false
        $placePivots += $pt
        Say ("place pivot " + $fuel)

        # --- cheapest brands: Bandiera sorted ascending, bottom 10, 3+ stations
        $pt = New-Pivot ($rowBase + 400) $col ("pt_cheap_" + $fuel)
        ($pt.CubeFields("[Fuel price data daily].[Bandiera]")).Orientation = $xlRow
        ($pt.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPage
        ($pt.CubeFields("[Measures].[Avg price 3+ stations]")).Orientation = $xlData
        Start-Sleep -Milliseconds 400
        Set-Fuel $pt $fuel
        $bf = $pt.PivotFields("[Fuel price data daily].[Bandiera].[Bandiera]")
        $bf.AutoSort(1, "[Measures].[Avg price 3+ stations]")
        [void]$bf.PivotFilters.Add2(2, $pt.DataFields(1), 10)
        $pt.ManualUpdate = $false
        $cheapPivots += $pt
        Say ("cheap pivot " + $fuel)

        # --- heat table: Region x Week, Italy average = total row
        $pt = New-Pivot (20 + $i * 30) 1 ("pt_heat_" + $fuel) $sheetRegions
        $pt.ColumnGrand = $true
        ($pt.CubeFields("[Fuel price data daily].[Denominazione Regione]")).Orientation = $xlRow
        ($pt.CubeFields("[Fuel price data daily].[Week of Year]")).Orientation = $xlColumn
        ($pt.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPage
        ($pt.CubeFields("[Measures].[Avg price]")).Orientation = $xlData
        Start-Sleep -Milliseconds 400
        Set-Fuel $pt $fuel
        $pt.ManualUpdate = $false
        $heatPivots += $pt
        Say ("heat pivot " + $fuel)

        # --- national average trend
        $pt = New-Pivot ($rowBase + 1200) $col ("pt_nat_" + $fuel)
        ($pt.CubeFields("[Fuel price data daily].[Date]")).Orientation = $xlRow
        ($pt.CubeFields("[Fuel price data daily].[descCarburante]")).Orientation = $xlPage
        ($pt.CubeFields("[Measures].[Italy average]")).Orientation = $xlData
        Start-Sleep -Milliseconds 400
        Set-Fuel $pt $fuel
        $pt.ManualUpdate = $false
        $natPivots += $pt
        Say ("national pivot " + $fuel)
        $i++
    }

    # ---------------- slicers (shared caches so a choice on one sheet follows to the other) ----------------
    $first = $placePivots[0]
    $slicerDefs = @(
        @{ field = "[Fuel price data daily].[Denominazione Regione]"; caption = "Region"; w = 160 },
        @{ field = "[Fuel price data daily].[Provincia]"; caption = "Province"; w = 160 },
        @{ field = "[Fuel price data daily].[Comune]"; caption = "Comune (municipality)"; w = 200 },
        @{ field = "[Fuel price data daily].[isSelf]"; caption = "Price type (1 self, 0 full)"; w = 190 }
    )
    $caches = @()
    foreach ($d in $slicerDefs) {
        $sc = $wb.SlicerCaches().Add2($first, $d.field)
        foreach ($pt in ($placePivots + $cheapPivots)) {
            if ($pt.Name -ne $first.Name) { [void]$sc.PivotTables.AddPivotTable($pt) }
        }
        $caches += ,@($sc, $d)
    }
    # price type also drives the region heat tables and the national trend
    $scSelf = $caches[3][0]
    foreach ($pt in ($heatPivots + $natPivots)) { [void]$scSelf.PivotTables.AddPivotTable($pt) }
    try { $scSelf.VisibleSlicerItemsList = @("[Fuel price data daily].[isSelf].&[1]") } catch { Say ("isSelf default failed: " + $_.Exception.Message) }

    foreach ($sheet in @($sheetPlace, $sheetCheap)) {
        $left = 15
        foreach ($c in $caches) {
            $sl = $c[0].Slicers.Add($sheet)
            $sl.Caption = $c[1].caption
            $sl.Top = 55; $sl.Left = $left; $sl.Width = $c[1].w; $sl.Height = 175
            $left += $c[1].w + 12
        }
    }
    Say "slicers added"

    # ---------------- charts ----------------
    Add-Title $sheetPlace "Price trend in your place" "Pick a Region, Province or Comune: each fuel is shown against the Italy average. Price type defaults to self-service."
    Add-Title $sheetCheap "Where is fuel cheaper?" "The 10 cheapest brands (bandiere) per fuel for the place picked above - brands with 3+ stations, average of posted prices."
    $positions = @(@(250, 15), @(250, 560), @(570, 15), @(570, 560))
    for ($k = 0; $k -lt 4; $k++) {
        $fuel = $FUELS[$k]
        $ch = $wb.Charts.Add()
        [void]$ch.SetSourceData($placePivots[$k].TableRange2)
        $ch.ChartType = $xlLine
        $ch.HasTitle = $true; $ch.ChartTitle.Text = $fuel + " - this place vs Italy average"
        $co = Move-Chart $ch "Place Trend" $positions[$k][0] $positions[$k][1] 530 300
        $co.Chart.Axes($xlValue).TickLabels.NumberFormat = $eurofmt
        Set-AxisFromData $co $natPivots[$k].DataBodyRange 0.25
        $s1 = $co.Chart.SeriesCollection(1); $s1.Format.Line.ForeColor.RGB = $FUEL_COLOR[$fuel]; $s1.Format.Line.Weight = 2.5
        $s2 = $co.Chart.SeriesCollection(2); $s2.Format.Line.ForeColor.RGB = 8554377; $s2.Format.Line.Weight = 1.75; $s2.Format.Line.DashStyle = 4
        Say ("place chart " + $fuel)

        $ch = $wb.Charts.Add()
        [void]$ch.SetSourceData($cheapPivots[$k].TableRange2)
        $ch.ChartType = $xlBarClustered
        $ch.HasTitle = $true; $ch.ChartTitle.Text = $fuel + " - cheapest brands (average price)"
        $co = Move-Chart $ch "Cheapest Brand" $positions[$k][0] $positions[$k][1] 530 300
        $co.Chart.Axes($xlCategory).ReversePlotOrder = $true
        $co.Chart.Axes($xlValue).TickLabels.NumberFormat = $eurofmt
        $co.Chart.HasLegend = $false
        $co.Chart.Axes($xlValue).MinimumScale = 0
        $sb = $co.Chart.SeriesCollection(1); $sb.Format.Fill.ForeColor.RGB = $FUEL_COLOR[$fuel]
        $sb.HasDataLabels = $true; $sb.DataLabels().NumberFormat = "0.000"
        $co.Chart.ChartGroups(1).GapWidth = 60
        Say ("cheap chart " + $fuel)
    }

    # ---------------- regions vs Italy ----------------
    Add-Title $sheetRegions "Regions vs Italy" "Top: daily Italy average per fuel. Below: average price per region and ISO week; the Total row is the Italy average. Colour scale per week: green = cheapest regions that week, red = dearest."
    for ($k = 0; $k -lt 4; $k++) {
        $fuel = $FUELS[$k]
        $ch = $wb.Charts.Add()
        [void]$ch.SetSourceData($natPivots[$k].TableRange2)
        $ch.ChartType = $xlLine
        $ch.HasTitle = $true; $ch.ChartTitle.Text = $fuel + " - Italy average"
        $co = Move-Chart $ch "Regions vs Italy" 45 (15 + $k * 265) 255 170
        $co.Chart.HasLegend = $false
        $co.Chart.Axes($xlValue).TickLabels.NumberFormat = $eurofmt
        Set-AxisFromData $co $natPivots[$k].DataBodyRange 0.1
        $co.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = $FUEL_COLOR[$fuel]
        $sheetRegions.Cells.Item(16 + $k * 30, 1).Value2 = $fuel + " - average price by region and week"
        $sheetRegions.Cells.Item(16 + $k * 30, 1).Font.Bold = $true
        $sheetRegions.Cells.Item(16 + $k * 30, 1).Font.Size = 13
        $body = $heatPivots[$k].DataBodyRange
        $regionRows = $body.Rows.Count - 1
        for ($c = 1; $c -le $body.Columns.Count; $c++) {
            $col = $body.Resize($regionRows).Columns($c)
            $cs = $col.FormatConditions.AddColorScale(3)
            $cs.ColorScaleCriteria(1).FormatColor.Color = (99 + 190 * 256 + 123 * 65536)
            $cs.ColorScaleCriteria(2).FormatColor.Color = (255 + 235 * 256 + 132 * 65536)
            $cs.ColorScaleCriteria(3).FormatColor.Color = (248 + 105 * 256 + 107 * 65536)
        }
        $body.NumberFormat = "0.000"
        Say ("regions block " + $fuel)
    }

    $src.Visible = $false
    $dash.Activate()
    $wb.Save()
    Say "SAVED"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "CUSTOMER VIEWS COMPLETE"
