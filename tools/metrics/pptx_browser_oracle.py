# -*- coding: utf-8 -*-
"""Render pptx decks via @silurus/ooxml (browser canvas viewer) headlessly.

The pptx sibling of browser_oracle.py: serves tools/browser-oracle/dist plus
one deck at a time from a local static server, drives index_pptx.html with
Playwright Chromium, and writes p<N>.png per slide in the layout
`pptx_oracle.py silurus` scores (pipeline_data/pptx_benchmark/dev/oracle/
silurus_png/<deck>/). Embedded EOT fonts are NOT loaded by the browser
engine, so font-heavy decks score with fallback faces -- still a layout
oracle, not a glyph one.

Usage:
  python pptx_browser_oracle.py [--decks d09,d20] [--dpi 150] [--rerender]
"""
from __future__ import annotations

import argparse
import base64
import http.server
import os
import shutil
import socketserver
import sys
import threading
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"
PPTX_DIR = DEV / "pptx"
OUT_ROOT = DEV / "oracle" / "silurus_png"
HARNESS_DIST = REPO / "tools" / "browser-oracle" / "dist"


def decks(selector: str | None) -> list[Path]:
    all_decks = sorted(PPTX_DIR.glob("*.pptx"))
    if not selector:
        return all_decks
    wanted = {s.strip() for s in selector.split(",") if s.strip()}
    return [p for p in all_decks if p.stem.split("__")[0] in wanted or p.stem in wanted]


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--rerender", action="store_true")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright

    targets = decks(args.decks)
    os.chdir(HARNESS_DIST)
    httpd = socketserver.TCPServer(
        ("127.0.0.1", 0), http.server.SimpleHTTPRequestHandler)
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    staged = HARNESS_DIST / "target.pptx"
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1400, "height": 1200})
            errors: list[str] = []
            page.on("pageerror", lambda e: errors.append(str(e)))
            for pptx in targets:
                out = OUT_ROOT / pptx.stem
                if args.rerender:
                    shutil.rmtree(out, ignore_errors=True)
                if list(out.glob("p*.png")):
                    print(f"  cached {pptx.stem[:50]}")
                    continue
                out.mkdir(parents=True, exist_ok=True)
                shutil.copy(pptx, staged)
                errors.clear()
                try:
                    page.goto(f"http://127.0.0.1:{port}/index_pptx.html", timeout=60000)
                    page.wait_for_function("window.oracleReady === true", timeout=30000)
                    n = page.evaluate("window.oracleInit('./target.pptx')")
                    for i in range(n):
                        url = page.evaluate(f"window.oraclePage({i}, {args.dpi})")
                        png = base64.b64decode(url.split(",", 1)[1])
                        (out / f"p{i + 1}.png").write_bytes(png)
                    print(f"  {pptx.stem[:50]:50s} {n} slides"
                          + (f"  errors: {errors[:1]}" if errors else ""))
                except Exception as exc:  # keep the sweep going per deck
                    print(f"  FAIL {pptx.stem[:50]}: {exc}")
                    shutil.rmtree(out, ignore_errors=True)
            browser.close()
    finally:
        httpd.shutdown()
        staged.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
