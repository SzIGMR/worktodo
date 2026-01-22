#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTPUT_DIR="${SCRIPT_DIR}/templates"

mkdir -p "${OUTPUT_DIR}"

skillrunner run budget_matcher \
  --generate-templates "${OUTPUT_DIR}"

echo "Vorlagen wurden in ${OUTPUT_DIR} erstellt."
