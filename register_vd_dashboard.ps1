# Registers the daily VD Rebate dashboard refresh (auto-selects the current month
# from the Quantity Targets workbook, recomputes from the sales parquet, deploys if changed).
# Re-run to update the schedule. Task runs daily at 06:45.
$ErrorActionPreference = 'Stop'
$bat = "C:\Users\Administrator\olympic-paints-vd\run_vd_dashboard.bat"
schtasks /Create /TN "Olympic VD Dashboard Refresh" /TR "`"$bat`"" /SC DAILY /ST 06:45 /RL HIGHEST /F
Write-Host "Registered 'Olympic VD Dashboard Refresh' (daily 06:45)."
