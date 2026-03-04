#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="${ROOT_DIR}/dist"
PACKAGE_NAME="excel_check_tool_package.zip"
OUTPUT_PATH="${DIST_DIR}/${PACKAGE_NAME}"

mkdir -p "${DIST_DIR}"
rm -f "${OUTPUT_PATH}"

cd "${ROOT_DIR}"
zip -r "${OUTPUT_PATH}" . \
  -x ".git/*" \
  -x ".venv/*" \
  -x "*/__pycache__/*" \
  -x "__pycache__/*" \
  -x "*.pyc" \
  -x "dist/*"

echo "Package created: ${OUTPUT_PATH}"
