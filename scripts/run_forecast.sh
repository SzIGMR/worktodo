#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

BOOKINGS_FILE="${SCRIPT_DIR}/bookings.xlsx"
BUDGETS_FILE="${SCRIPT_DIR}/budgets.xlsx"
OUTPUT_FILE="${SCRIPT_DIR}/forecast_result.xlsx"
FY_START_YEAR=2025

skillrunner run budget_matcher \
  --bookings "${BOOKINGS_FILE}" \
  --budgets "${BUDGETS_FILE}" \
  --out "${OUTPUT_FILE}" \
  --mode forecast \
  --fy-start-year "${FY_START_YEAR}"

echo "Forecast-Auswertung wurde nach ${OUTPUT_FILE} geschrieben."
