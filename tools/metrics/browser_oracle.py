"""Render .docx pages via @silurus/ooxml (browser WASM viewer) headlessly.

The vite-built harness (tools/browser-oracle/dist) exposes
window.oracleInit(docUrl) -> pageCount and window.oraclePage(i, dpi) -> PNG
data URL. This drives it with Playwright Chromium over a local static
server and writes page_NNNN.png like the word_png convention — a second
independent oracle next to LibreOffice for Word-fidelity comparison.

Usage:
  python browser_oracle.py <doc.docx> <outdir> [dpi=110]
"""
import base64
import http.server
import os
import shutil
import socketserver
import sys
import threading

HARNESS_DIST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                            "browser-oracle", "dist")


def render(docx, outdir, dpi=110):
    from playwright.sync_api import sync_playwright
    os.makedirs(outdir, exist_ok=True)
    # stage: serve dist + the docx from one root
    doc_name = "target.docx"
    shutil.copy(docx, os.path.join(HARNESS_DIST, doc_name))
    os.chdir(HARNESS_DIST)
    httpd = socketserver.TCPServer(
        ("127.0.0.1", 0), http.server.SimpleHTTPRequestHandler)
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1400, "height": 1200})
            errors = []
            pg.on("pageerror", lambda e: errors.append(str(e)))
            pg.goto(f"http://127.0.0.1:{port}/index.html", timeout=60000)
            pg.wait_for_function("window.oracleReady === true", timeout=30000)
            n = pg.evaluate(f"window.oracleInit('./{doc_name}')")
            print(f"pages: {n}")
            for i in range(n):
                url = pg.evaluate(f"window.oraclePage({i}, {dpi})")
                png = base64.b64decode(url.split(",", 1)[1])
                open(os.path.join(outdir, f"page_{i+1:04d}.png"), "wb").write(png)
            if errors:
                print("page errors:", errors[:3])
            b.close()
    finally:
        httpd.shutdown()
        os.remove(os.path.join(HARNESS_DIST, doc_name))
    return True


if __name__ == "__main__":
    doc = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2])
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 110
    render(doc, out, dpi)
