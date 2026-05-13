#!/usr/bin/env bash
#
# Packages the Teams app manifest + icons into a sideload-ready .zip.
#
# Usage:
#   ./scripts/package-teams-app.sh            # dev
#   ./scripts/package-teams-app.sh prod       # prod

set -euo pipefail

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
APP_DIR="$SCRIPT_DIR/../appPackage"
ENV_NAME="${1:-dev}"
OUT="$APP_DIR/build/appPackage.$ENV_NAME.zip"

mkdir -p "$APP_DIR/build"
rm -f "$OUT"

for f in manifest.json color.png outline.png; do
  if [ ! -f "$APP_DIR/$f" ]; then
    echo "Required file not found: $APP_DIR/$f" >&2
    exit 1
  fi
done

( cd "$APP_DIR" && zip -j "$OUT" manifest.json color.png outline.png )

echo "Wrote $OUT"
