#!/bin/bash
# Fast build-convert-compare cycle
set -e
cd "$(dirname "$0")/../.."
cargo build --release -p oxi-cli 2>/dev/null
target/release/oxi.exe docx-to-pdf tests/fixtures/comprehensive_test.docx /tmp/oxi_comp.pdf 2>/dev/null
cp /tmp/oxi_comp.pdf tools/golden-test/pixel_output/oxi_output.pdf
python3 tools/golden-test/quick_ssim.py
