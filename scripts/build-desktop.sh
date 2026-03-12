#!/bin/bash
# Build desktop distribution: WASM + Web assets into dist-desktop/
set -e

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DIST="$ROOT/dist-desktop"

echo "=== Building WASM ==="
cd "$ROOT/crates/oxi-wasm"
wasm-pack build --target web --release

echo "=== Assembling dist-desktop ==="
rm -rf "$DIST"
mkdir -p "$DIST/pkg"

# Mirror the project structure so relative paths in index.html work unchanged
# index.html is at web/index.html, WASM at crates/oxi-wasm/pkg/
mkdir -p "$DIST/web"
mkdir -p "$DIST/crates/oxi-wasm/pkg"

# Copy web frontend
cp "$ROOT/web/index.html" "$DIST/web/"

# Copy WASM package
cp "$ROOT/crates/oxi-wasm/pkg/oxi_wasm.js" "$DIST/crates/oxi-wasm/pkg/"
cp "$ROOT/crates/oxi-wasm/pkg/oxi_wasm_bg.wasm" "$DIST/crates/oxi-wasm/pkg/"

# Copy docs assets (icons, favicon)
if [ -d "$ROOT/docs" ]; then
    mkdir -p "$DIST/docs"
    cp "$ROOT/docs/"*.png "$DIST/docs/" 2>/dev/null || true
    cp "$ROOT/docs/"*.ico "$DIST/docs/" 2>/dev/null || true
fi

# Copy sample files if they exist
if [ -d "$ROOT/web/samples" ]; then
    cp -r "$ROOT/web/samples" "$DIST/web/samples"
fi

# Redirect index: Tauri loads /index.html, redirect to /web/index.html
cat > "$DIST/index.html" << 'REDIRECT'
<!DOCTYPE html>
<html><head><meta http-equiv="refresh" content="0;url=web/index.html"></head></html>
REDIRECT

echo "=== dist-desktop assembled ==="
ls -la "$DIST/"
