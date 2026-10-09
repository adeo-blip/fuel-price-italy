$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK
$outPdf = Join-Path $env:FUEL_OUTDIR "dashboard.pdf"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($path)
    $dash = $wb.Worksheets("Dashboard")
    $dash.Activate()
    $dash.PageSetup.PrintArea = "A1:AF62"
    $dash.PageSetup.Orientation = 2
    $dash.PageSetup.Zoom = $false
    $dash.PageSetup.FitToPagesWide = 1
    $dash.PageSetup.FitToPagesTall = 1
    $dash.PageSetup.LeftMargin = 10; $dash.PageSetup.RightMargin = 10
    $dash.PageSetup.TopMargin = 10; $dash.PageSetup.BottomMargin = 10
    if (Test-Path $outPdf) { Remove-Item $outPdf -Force }
    $dash.ExportAsFixedFormat(0, $outPdf)
    Write-Output "exported pdf"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
