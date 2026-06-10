#!/usr/bin/env bash
# Builds the library and produces ExcelChartExporter.zip, ready to upload as
# an external library in the ODC Portal.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PUBLISH_DIR="$REPO_ROOT/artifacts/publish"
ZIP_PATH="$REPO_ROOT/artifacts/ExcelChartExporter.zip"

rm -rf "$REPO_ROOT/artifacts"
dotnet publish "$REPO_ROOT/src/ExcelChartLibrary" -c Release -o "$PUBLISH_DIR"

(cd "$PUBLISH_DIR" && zip -r -q "$ZIP_PATH" .)

echo "ODC upload package created: $ZIP_PATH"
