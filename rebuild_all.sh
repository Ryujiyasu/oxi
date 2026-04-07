#!/usr/bin/env bash
# Rebuild all Oxi binaries that depend on oxidocs-core.
# Use this after any change to crates/oxidocs-core/ before running tests.
#
# Builds:
#   1. oxidocs-core (for cargo run --example layout_json — DML diff)
#   2. oxi-wasm + wasm-pack + copy to web/pkg/ (for browser/web tests)
#   3. oxi-gdi-renderer (for pipeline.verify SSIM tests)
set -e

cd "$(dirname "$0")"

echo "[1/3] Building oxidocs-core..."
cargo build --release -p oxidocs-core

echo "[2/3] Building wasm + copying to web/pkg/..."
(cd crates/oxi-wasm && wasm-pack build --release --target web)
cp crates/oxi-wasm/pkg/oxi_wasm.js web/pkg/
cp crates/oxi-wasm/pkg/oxi_wasm_bg.wasm web/pkg/
cp crates/oxi-wasm/pkg/oxi_wasm.d.ts web/pkg/
cp crates/oxi-wasm/pkg/oxi_wasm_bg.wasm.d.ts web/pkg/

echo "[3/3] Building oxi-gdi-renderer..."
(cd tools/oxi-gdi-renderer && cargo build --release)

echo "[OK] All builds complete."
