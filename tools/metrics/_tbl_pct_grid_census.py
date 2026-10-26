# -*- coding: utf-8 -*-
"""Fixed pct tables with a mismatched grid: table extents, Word against Oxi.

For each document, the first pages' table border strokes: the leftmost and
rightmost x of horizontal rules and the set of vertical rule x positions.
Word from a PDF by a fresh Word process (cached in
pipeline_data/_tbl_pct_grid_census/), Oxi from --dump-layout. Prints, per
page, the extent and the share of Word's vertical rules that each exe matches
within 1pt, plus page counts.

    python tools/metrics/_tbl_pct_grid_census.py <exe>[;<exe>] <doc> [<doc> ...]
"""
import json, os, subprocess, sys, tempfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_tbl_pct_grid_census"
PAGES = 3


def word(docx, pdf):
    if not pdf.exists():
        import win32com.client
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(docx), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
    import pymupdf
    doc = pymupdf.open(pdf)
    out = []
    for pg in list(doc)[:PAGES]:
        ds = pg.get_drawings()
        hs = [d["rect"] for d in ds if d["rect"].height < 2 and d["rect"].width > 30]
        vs = sorted({round(d["rect"].x0) for d in ds if d["rect"].width < 2 and d["rect"].height > 8})
        out.append(((round(min(r.x0 for r in hs)), round(max(r.x1 for r in hs))) if hs else None, vs))
    return len(doc), out


def oxi(exe, docx):
    t = tempfile.mkdtemp(); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True, timeout=1800)
    pages = json.load(open(lay, encoding="utf-8"))["pages"]
    out = []
    for pg in pages[:PAGES]:
        b = [e for e in pg["elements"] if e.get("type") == "border"]
        hs = [e for e in b if e.get("h", 0) < 2 and e.get("w", 0) > 30]
        vs = sorted({round(e["x"]) for e in b if e.get("w", 0) < 2 and e.get("h", 0) > 8})
        out.append(((round(min(e["x"] for e in hs)), round(max(e["x"] + e["w"] for e in hs))) if hs else None, vs))
    return len(pages), out


def match(wv, ov):
    if not wv:
        return "-"
    hit = sum(any(abs(a - b) <= 1 for b in ov) for a in wv)
    return f"{hit}/{len(wv)}"


if __name__ == "__main__":
    exes = [e for e in sys.argv[1].split(";") if e]
    OUT.mkdir(parents=True, exist_ok=True)
    for doc in sys.argv[2:]:
        docx = REPO / "pipeline_data/docx_corpus" / f"{doc}.docx"
        wn, wp = word(docx, OUT / (doc.replace("/", "_") + ".pdf"))
        res = [oxi(e, docx) for e in exes]
        print(f"== {doc}  pages W {wn} " + " ".join(f"O{k} {n}" for k, (n, _) in enumerate(res)))
        for i, (wext, wv) in enumerate(wp):
            cells = [f"p{i + 1} W {wext}"]
            for k, (_, op) in enumerate(res):
                oext, ov = op[i] if i < len(op) else (None, [])
                cells.append(f"O{k} {oext} rules {match(wv, ov)}")
            print("   " + "   ".join(cells), flush=True)
