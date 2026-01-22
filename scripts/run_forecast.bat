@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "BOOKINGS_FILE=%SCRIPT_DIR%bookings.xlsx"
set "BUDGETS_FILE=%SCRIPT_DIR%budgets.xlsx"
set "OUTPUT_FILE=%SCRIPT_DIR%forecast_result.xlsx"
set "FY_START_YEAR=2025"

skillrunner run budget_matcher --bookings "%BOOKINGS_FILE%" --budgets "%BUDGETS_FILE%" --out "%OUTPUT_FILE%" --mode forecast --fy-start-year %FY_START_YEAR%

echo Forecast-Auswertung wurde nach %OUTPUT_FILE% geschrieben.
endlocal
