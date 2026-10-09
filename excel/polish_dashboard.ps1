$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

try {
    $wb = $excel.Workbooks.Open($path, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $false)
    $dash = $wb.Worksheets("Dashboard")

    $kpiCol = 20
    for ($i=0; $i -lt 4; $i++) {
        $c = $kpiCol + $i*3
        $dash.Cells.Item(2,$c).EntireColumn.ColumnWidth = 17
    }
    Write-Output "widened KPI columns"

    $wb.Save()
    Write-Output "SAVED"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "POLISH COMPLETE"
