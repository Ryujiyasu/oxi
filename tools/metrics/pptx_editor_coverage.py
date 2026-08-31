"""How much of the corpus the engine can lay out, and where the answer came from.

Three numbers per deck, over every text shape of every slide:

  tables    -- shapes the compiled tables alone can measure
  browser   -- shapes measurable once the page measures the faces itself
  complete  -- shapes the engine actually laid out end to end

The gap between the first two is what asking the browser bought. The gap
between the second and third is a shape the engine declined for a reason that
is NOT the font -- worth knowing separately, because those are engine gaps, not
coverage gaps.
"""
import http.server
import json
import os
import socketserver
import sys
import threading
from pathlib import Path

from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]

JS = r"""
async (bytes) => {
  const m = await import('./face-metrics.js');
  const w = await import('./oxidocs_wasm.js');
  await w.default();
  const pres = w.parse_presentation(new Uint8Array(bytes));
  const st = pres.master_styles || {};
  const fallback = pres.minor_font || 'Calibri';
  const levelsOf = (sh) => {
    if (sh.ph_levels && sh.ph_levels.length) return sh.ph_levels;
    const t = (sh.ph_type || '').toLowerCase();
    if (t === 'title' || t === 'ctrtitle') return st.title || [];
    if (t) return st.body || [];
    return st.other || [];
  };
  let total = 0, tables = 0, browser = 0, complete = 0;
  const unmet = {};
  for (const slide of pres.slides) {
    for (const sh of slide.shapes) {
      const paras = sh.content?.TextBox?.paragraphs ?? sh.content?.AutoShape?.paragraphs;
      if (!paras || !paras.length) continue;
      if (!paras.some(p => p.runs.some(r => r.text && r.text.trim()))) continue;
      total++;
      const lv = levelsOf(sh);
      const fams = new Set();
      const runs = [];
      paras.forEach(p => {
        const inherited = lv.length
          ? (lv[Math.min(p.lvl || 0, lv.length - 1)] || {}).font_family : null;
        p.runs.forEach(r => {
          const f = r.font_family || inherited || fallback;
          fams.add(f);
          runs.push({ text: r.text, font_family: f, bold: r.bold, italic: r.italic });
        });
      });
      const inTables = [...fams].every(f => w.slide_family_measurable(f));
      const inBrowser = [...fams].every(f => w.slide_family_measurable(f) || m.familyPresent(f));
      if (inTables) tables++;
      if (inBrowser) browser++;
      else for (const f of fams) {
        if (!w.slide_family_measurable(f) && !m.familyPresent(f)) unmet[f] = (unmet[f] || 0) + 1;
      }
      const adv = m.collectAdvances(runs, fallback);
      let out = null;
      try {
        out = w.layout_slide_shape(sh, paras, levelsOf(sh), sh.ph_levels || [], fallback, adv);
      } catch (e) { /* a shape the engine refuses is counted as incomplete */ }
      if (out && out.complete) complete++;
    }
  }
  return { total, tables, browser, complete, unmet };
}
"""


def main():
    decks = []
    for sub in ("dev/pptx", "pptx"):
        decks += sorted((REPO / "pipeline_data" / "pptx_benchmark" / sub).glob("*.pptx"))
    if len(sys.argv) > 1:
        decks = [d for d in decks if any(a in d.name for a in sys.argv[1:])]
    os.chdir(REPO / "web")
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    httpd.RequestHandlerClass.log_message = lambda *a, **k: None
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    tot = {"total": 0, "tables": 0, "browser": 0, "complete": 0}
    unmet = {}
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page()
        page.goto(f"http://127.0.0.1:{port}/pptx-editor.html")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('ready')")
        for d in decks:
            try:
                r = page.evaluate(JS, list(d.read_bytes()))
            except Exception as e:
                print(f"{d.name[:28]:30} refused: {str(e)[:50]}")
                continue
            if not r["total"]:
                continue
            for k in tot:
                tot[k] += r[k]
            for f, n in r["unmet"].items():
                unmet[f] = unmet.get(f, 0) + n
            print(f"{d.name.split('__')[0][:22]:24} {r['total']:4} shapes  "
                  f"tables {r['tables'] * 100 // r['total']:3}%  "
                  f"browser {r['browser'] * 100 // r['total']:3}%  "
                  f"laid out {r['complete'] * 100 // r['total']:3}%")
        b.close()
    n = max(tot["total"], 1)
    print(f"\n{'ALL':24} {tot['total']:4} shapes  "
          f"tables {tot['tables'] * 100 / n:.1f}%  "
          f"browser {tot['browser'] * 100 / n:.1f}%  "
          f"laid out {tot['complete'] * 100 / n:.1f}%")
    if unmet:
        print("\nfamilies neither source has, by shapes blocked:")
        for f, c in sorted(unmet.items(), key=lambda kv: -kv[1])[:15]:
            print(f"   {c:4}  {f}")
    (REPO / "pipeline_data" / "pptx_editor_coverage.json").write_text(
        json.dumps({"totals": tot, "unmet": unmet}, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
