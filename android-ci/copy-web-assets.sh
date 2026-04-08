#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
CI_DIR="$(cd "$(dirname "$0")" && pwd)"
ASSET_DIR="$CI_DIR/app/src/main/assets"

mkdir -p "$ASSET_DIR"
cp "$ROOT_DIR/index.html" "$ASSET_DIR/index.html"
cp "$ROOT_DIR/app.js" "$ASSET_DIR/app.js"
cp "$ROOT_DIR/style.css" "$ASSET_DIR/style.css"

echo "Web assets copied to $ASSET_DIR"
