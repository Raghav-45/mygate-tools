#!/usr/bin/env bash
# Build the three extension zips locally into ./dist for manual release/testing.
# The GitHub workflow does the same thing for an actual release.
set -euo pipefail

cd "$(dirname "$0")"

TOOLS=(
  "mygate-dump-tool"
  "mygate-report-tool"
  "mygate-summary-tool"
)

rm -rf dist
mkdir -p dist

for tool in "${TOOLS[@]}"; do
  rm -rf "dist/$tool"
  cp -r "$tool" "dist/$tool"

  find "dist/$tool" -name '.gitignore' -delete
  find "dist/$tool" -name 'KNOWN_ISSUES.md' -delete
  find "dist/$tool" -name '*.zip' -delete
  find "dist/$tool" -name '*.crx' -delete
  find "dist/$tool" -name '*.pem' -delete

  (cd "dist/$tool" && zip -r "../${tool}.zip" .)

  rm -rf "dist/$tool"
  echo "Built: dist/${tool}.zip"
done

ls -la dist/
