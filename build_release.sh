#!/usr/bin/env bash
# Build the extension zips locally into ./dist for manual release/testing.
# The GitHub workflow does the same thing for an actual release.
set -euo pipefail

cd "$(dirname "$0")"

node scripts/sync_handoff_sources.mjs

TOOLS=(
  "mygate-dump-tool"
  "mygate-report-tool"
  "mygate-summary-tool"
  "mygate-analysis-tool"
  "mygate-handoff-tool"
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
  find "dist/$tool" -type d -name 'tests' -prune -exec rm -rf {} +

  (cd "dist/$tool" && zip -r "../${tool}.zip" .)

  rm -rf "dist/$tool"
  echo "Built: dist/${tool}.zip"
done

ls -la dist/
