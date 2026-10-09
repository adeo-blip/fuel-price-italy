> Copy of the Excel automation scripts. The weekly task runs them from DashboardOptimization/UpdatedFiles/_automation on the local PC; the .xlsx workbooks are not stored in this repository.

# Weekly update automation

Scripts used by the scheduled task `fuel-dashboard-weekly-files` (Mondays 10:00). Nothing in
`..\..\ExcelfilePowerQuery` is ever modified; each run writes new dated versions into `..\`.

Environment variables read by the scripts:

- `FUEL_WORKBOOK` - the workbook (a working copy) that a script opens
- `FUEL_OUTDIR`   - folder for JSON, PDF and PNG intermediates and the built pptx files

Order: `refresh_source.ps1` -> `extract_trends.ps1` -> `extract_national.ps1` (on the refreshed copy);
then on a second copy `add_national_measure.ps1` -> `build_dashboard.ps1` ->
`update_dashboard_national.ps1` -> `polish_dashboard.ps1` -> `build_customer_views.ps1` ->
`export_dashboard.ps1` (and `export_sheets.ps1` for the three customer views);
then `python pdf_to_png.py`, `node build_slide.js`, `node build_deck.js`.

Setup already done: `npm install pptxgenjs` in this folder. Python needs `pymupdf` and `Pillow`.

Notes:
- `add_national_measure.ps1` creates the DAX measures (Italy average, This place, Avg price, Avg price 3+ stations, Stations).
- `build_customer_views.ps1` adds the sheets Place Trend, Cheapest Brand and Regions vs Italy. Pivots are configured with ManualUpdate on and never refreshed with RefreshTable (that costs minutes per pivot); the whole chain runs in about two minutes.
