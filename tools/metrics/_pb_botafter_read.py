# -*- coding: utf-8 -*-
"""Read the botafter sweep. Usage: _pb_botafter_read.py word|oxi [--sweep lo hi step]

One column per (multiplier, own-ref) arm: FINAL's page and, when it stayed,
its baseline. The spacer at which FINAL flips is the keep boundary. Comparing
the boundary across multipliers says which box the page bottom measured; across
own-ref says whether the relief needs an own reference.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_botafter_gen import OUT, AFTERS, parse_sweep

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_page1(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            d.Close(False)
        finally:
            word.Quit()
    doc = fitz.open(pdf)
    pg = doc[0]
    sep = None
    for dr in pg.get_drawings():
        r = dr["rect"]
        if r.height < 4 and 100 < r.width < 200 and r.y0 > 300:
            sep = r.y0 if sep is None else min(sep, r.y0)
    rows = []
    for blk in pg.get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            sp = [s for s in l["spans"] if s["text"].strip()]
            if sp:
                rows.append((round(sp[0]["origin"][1], 2),
                             "".join(s["text"] for s in sp).strip()))
    rows.sort()
    return rows, sep


def oxi_page1(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    pg = d["pages"][0]
    sep = None
    for e in pg["elements"]:
        if (e.get("text") or "").strip():
            continue
        if e.get("y", 0) > 300 and 100 < e.get("w", 0) < 200 and e.get("h", 9) < 4:
            sep = e["y"] if sep is None else min(sep, e["y"])
    rows = {}
    for e in pg["elements"]:
        if not (e.get("text") or "").strip():
            continue
        rows.setdefault(round(e["y"] + 10.5, 2), []).append((e.get("x", 0), e["text"]))
    return ([(y, "".join(t for _, t in sorted(rows[y])).strip())
             for y in sorted(rows)], sep)


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_page1 if mode == "word" else oxi_page1
sw = parse_sweep(sys.argv)
cols = list(AFTERS)
print("%s   fn-boundary page-bottom box (line=auto multiplier x own ref)\n" % mode.upper())
print("  spacer |" + "".join("  after=%-4d pg  final_y |" % a for a in cols))
for x in sw:
    cells = []
    for a in cols:
        docx = os.path.join(OUT, "s%05d_a%04d.docx" % (x, a))
        if not os.path.exists(docx):
            cells.append("      MISSING     |")
            continue
        rows, sep = reader(docx)
        body = [(y, t) for y, t in rows if sep is None or y <= sep]
        fy = next((y for y, t in body if "FINAL" in t), None)
        cells.append("      %d  %8s |" % (1 if fy is not None else 2,
                                          ("%.2f" % fy) if fy is not None else "-"))
    print("  %6d |%s" % (x, "".join(cells)))
