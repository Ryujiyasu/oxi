"""Does the browser measure a face the way the engine's tables do?

The engine now accepts advances the PAGE measured, so a deck naming a face the
build machine never had can still be laid out by the engine's own rules. That
is only worth anything if the page's measuring agrees with the tables where
both know the same face -- otherwise the browser would be feeding confident,
wrong numbers into a rule that quantises to 1/8pt.

So: measure every ASCII character in a few faces both sources carry, and print
the disagreement in EM. The 1/8pt quantum at 12pt is 0.0104 EM, so a
disagreement has to stay well under that to be harmless.

Also checks the guard the whole thing rests on: a family this browser does NOT
have must be reported absent, because a browser asked for a missing font does
not fail -- it substitutes silently.
"""
import http.server
import os
import socketserver
import sys
import threading
from pathlib import Path

from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
FACES = [("Arial", False, False), ("Arial", True, False),
         ("Times New Roman", False, False), ("Calibri", False, False),
         ("Courier New", False, False)]
ASCII = "".join(chr(c) for c in range(32, 127))

JS = """
async ([faces, text]) => {
  const m = await import('./face-metrics.js');
  const w = await import('./oxidocs_wasm.js');
  await w.default();
  const out = [];
  for (const [family, bold, italic] of faces) {
    const present = m.familyPresent(family);
    const em = present ? m.measureFace(family, bold, italic, text) : null;
    const rows = [];
    if (em) {
      for (let i = 0; i < text.length; i++) {
        const table = w.slide_face_advance(family, bold, italic, text[i]);
        if (table === undefined || table === null) continue;
        rows.push([text[i], em[i], table]);
      }
    }
    out.push({ family, bold, italic, present, rows });
  }
  out.push({ family: 'Zzyzx Nonesuch Ghost', bold: false, italic: false,
             present: m.familyPresent('Zzyzx Nonesuch Ghost'), rows: [] });
  return out;
}
"""


def main():
    os.chdir(REPO / "web")
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    bad = 0
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(f"http://127.0.0.1:{port}/pptx-editor.html")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('ready')")
        res = page.evaluate(JS, [FACES, ASCII])
        browser.close()
    for face in res:
        name = (f"{face['family']}{' Bold' if face['bold'] else ''}"
                f"{' Italic' if face['italic'] else ''}")
        if not face["rows"]:
            state = "present" if face["present"] else "ABSENT"
            print(f"{name:22} {state}  (no shared characters)")
            if face["family"].startswith("Zzyzx") and face["present"]:
                print("    a family this browser cannot have was reported PRESENT")
                bad += 1
            continue
        worst = max(face["rows"], key=lambda r: abs(r[1] - r[2]))
        d = [abs(r[1] - r[2]) for r in face["rows"]]
        over = sum(1 for x in d if x > 0.002)
        print(f"{name:22} {len(face['rows']):3} chars  "
              f"mean |d| {sum(d) / len(d):.5f} em  max {max(d):.5f} em "
              f"({worst[0]!r} browser {worst[1]:.4f} vs table {worst[2]:.4f})  "
              f"over 0.002: {over}")
        if max(d) > 0.002:
            bad += 1
    print("\nquantum for scale: 1/8pt at 12pt = 0.01042 em")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
