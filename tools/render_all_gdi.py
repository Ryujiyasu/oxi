#!/usr/bin/env python3
"""Render all golden test documents with oxi-gdi-renderer, naming pages page_0001.png etc."""
import subprocess
import sys
import os
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
RENDERER = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
DOCX_DIR = REPO / "tools" / "golden-test" / "documents" / "docx"
OUT_DIR = REPO / "pipeline_data" / "oxi_png"

def main():
    docx_files = sorted(DOCX_DIR.glob("*.docx"))
    print(f"Found {len(docx_files)} docx files")

    t0 = time.time()
    ok = 0
    fail = 0

    for i, docx in enumerate(docx_files):
        stem = docx.stem
        doc_out = OUT_DIR / stem
        doc_out.mkdir(parents=True, exist_ok=True)

        try:
            result = subprocess.run(
                [str(RENDERER), str(docx), str(doc_out / "page")],
                capture_output=True, text=True, timeout=60, encoding='utf-8', errors='replace'
            )

            if result.returncode != 0:
                print(f"  [{i+1}/{len(docx_files)}] FAIL {stem[:50]}: {result.stderr[:100]}")
                fail += 1
                continue

            # Rename page_p1.png -> page_0001.png etc.
            for png in sorted(doc_out.glob("page_p*.png")):
                num = png.stem.replace("page_p", "")
                new_name = f"page_{int(num):04d}.png"
                png.rename(doc_out / new_name)

            pngs = list(doc_out.glob("page_*.png"))
            ok += 1
            if (i + 1) % 20 == 0:
                elapsed = time.time() - t0
                print(f"  [{i+1}/{len(docx_files)}] {ok} ok, {fail} fail, {elapsed:.0f}s")

        except subprocess.TimeoutExpired:
            print(f"  [{i+1}/{len(docx_files)}] TIMEOUT {stem[:50]}")
            fail += 1
        except Exception as e:
            print(f"  [{i+1}/{len(docx_files)}] ERROR {stem[:50]}: {e}")
            fail += 1

    elapsed = time.time() - t0
    print(f"\nDone: {ok} ok, {fail} fail in {elapsed:.0f}s")

if __name__ == "__main__":
    main()
