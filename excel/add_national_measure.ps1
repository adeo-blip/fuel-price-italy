$ErrorActionPreference = "Stop"
# Adds the DAX measures used by the dashboard views to the Data Model (skips measures that already exist).
$path = $env:FUEL_WORKBOOK
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    $wb = $excel.Workbooks.Open($path)
    $model = $wb.Model
    $table = $model.ModelTables("Fuel price data daily")
    $fmt = $model.ModelFormatGeneral

    $existing = @{}
    foreach ($m in $model.ModelMeasures) { $existing[$m.Name] = $true }

    $natAvg = "CALCULATE(AVERAGE('Fuel price data daily'[prezzo]), ALL('Fuel price data daily'[Denominazione Regione]), ALL('Fuel price data daily'[Provincia]), ALL('Fuel price data daily'[Comune]))"
    $measures = [ordered]@{
        "NationalAvgPrezzo"      = $natAvg
        "Stations"               = "DISTINCTCOUNT('Fuel price data daily'[idImpianto])"
        "Italy average"          = $natAvg
        "This place"             = "AVERAGE('Fuel price data daily'[prezzo])"
        "Avg price"              = "AVERAGE('Fuel price data daily'[prezzo])"
        "Avg price 3+ stations"  = "IF(DISTINCTCOUNT('Fuel price data daily'[idImpianto]) >= 3, AVERAGE('Fuel price data daily'[prezzo]), BLANK())"
    }
    foreach ($name in $measures.Keys) {
        if ($existing.ContainsKey($name)) { Write-Output ("exists: " + $name); continue }
        [void]$model.ModelMeasures.Add($name, $table, $measures[$name], $fmt, $name)
        Write-Output ("added: " + $name)
    }
    $wb.Save()
    Write-Output "MEASURE ADDED"
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
