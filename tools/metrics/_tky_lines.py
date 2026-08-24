# -*- coding: utf-8 -*-
"""Count LINES per page for tokyoshugyo, Word truth (PDF) vs Oxi (dump-layout).

S1207 next step: with the bundle ON the first flip is at Word p35 / Oxi p34,
and the cells that the bundle changes are all Word-correct already.  So the
missing line is present in BOTH arms.  Count lines per page and see where
Word has one more line than Oxi.

  python _tky_lines.py [first_page] [last_page]
"""
import json, os, subprocess, sys, tempfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "tokyoshugyo_000599795.docx")
PDF = os.path.join(REPO, "pipeline_data", "_kojin_rowgeom", "tokyoshugyo.pdf")
EXE = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
BUNDLE = {"OXI_CELLLAW": "1", "OXI_YAKUCOMP": "1", "OXI_AUTOSPACE2": "1",
          "OXI_S1201": "1"}

def _argi(i, dflt):
    try:
        return int(sys.argv[i])
    except (IndexError, ValueError):
        return dflt


P0 = _argi(1, 30)
P1 = _argi(2, 38)


def oxi_dump(env_extra, cache):
    if os.path.exists(cache):
        return json.load(open(cache, encoding="utf-8"))
    env = dict(os.environ)
    env.update(env_extra)
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "d.json")
        subprocess.run([EXE, DOCX, os.path.join(td, "p"), "96",
                        "--dump-layout=" + out], env=env, check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        d = json.load(open(out, encoding="utf-8"))
    json.dump(d, open(cache, "w", encoding="utf-8"), ensure_ascii=False)
    return d


def oxi_lines(dump, page):
    for pg in dump.get("pages", []):
        if pg["page"] != page:
            continue
        rows = {}
        for el in pg.get("elements", []):
            if el.get("type") != "text":
                continue
            key = round(float(el.get("y", 0)), 1)
            rows.setdefault(key, []).append(el)
        out = []
        for y in sorted(rows):
            els = sorted(rows[y], key=lambda e: float(e.get("x", 0)))
            txt = "".join(e.get("text", "") for e in els)
            out.append((y, float(els[0].get("x", 0)), txt))
        return out
    return []


def word_lines(page):
    import fitz
    d = fitz.open(PDF)
    pg = d[page - 1]
    out = []
    for blk in pg.get_text("dict")["blocks"]:
        for ln in blk.get("lines", []):
            txt = "".join(sp["text"] for sp in ln["spans"])
            if not txt.strip():
                continue
            out.append((round(ln["bbox"][1], 1), round(ln["bbox"][0], 1), txt))
    out.sort()
    return out


if __name__ == "__main__":
    base = oxi_dump({}, "C:/tmp/wk/tky_base.json")
    bund = oxi_dump(BUNDLE, "C:/tmp/wk/tky_bundle.json")
    print("pages: word=%d oxi_base=%d oxi_bundle=%d"
          % (90, len(base["pages"]), len(bund["pages"])))
    for p in range(P0, P1 + 1):
        w = word_lines(p)
        b = oxi_lines(base, p)
        n = oxi_lines(bund, p)
        print("\n=== page %d ===  word %d lines / base %d / bundle %d"
              % (p, len(w), len(b), len(n)))
        for i in range(max(len(w), len(n))):
            wt = w[i][2][:34] if i < len(w) else ""
            nt = n[i][2][:34] if i < len(n) else ""
            mark = "  " if wt[:12] == nt[:12] else "<<"
            print("%2d %s W %-36s | O %s" % (i, mark, wt, nt))
