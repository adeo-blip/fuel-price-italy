$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK
$outDir = $env:FUEL_OUTDIR
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($path)
    foreach ($name in @("Place Trend", "Cheapest Brand", "Regions vs Italy")) {
        $ws = $wb.Worksheets($name)
        $ws.Activate()
        $ws.PageSetup.PrintArea = "A1:X150"
        if ($name -ne "Regions vs Italy") { $ws.PageSetup.PrintArea = "A1:X60" }
        $ws.PageSetup.Orientation = 2
        $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1
        $ws.PageSetup.FitToPagesTall = $(if ($name -eq "Regions vs Italy") { 2 } else { 1 })
        $f = Join-Path $outDir (($name -replace " ", "_") + ".pdf")
        if (Test-Path $f) { Remove-Item $f -Force }
        $ws.ExportAsFixedFormat(0, $f)
        Write-Output ("exported " + $f)
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
