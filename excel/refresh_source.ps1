$ErrorActionPreference = "Stop"
$path = $env:FUEL_WORKBOOK
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    $wb = $excel.Workbooks.Open($path)
    foreach ($c in $wb.Connections) {
        try { $c.OLEDBConnection.BackgroundQuery = $false } catch {}
    }
    $wb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()
    $wb.Save()
    Write-Output "REFRESHED"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
