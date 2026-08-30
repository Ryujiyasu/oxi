# -*- coding: utf-8 -*-
"""Read the afterkeep probe. Usage: _pb_afterkeep_read.py word|oxi

Per arm: the gap from HEAD's baseline to NEXT's. Word's gap is line + after;
an engine that drops the after shows the line alone.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_afterkeep_gen import ARMS, OUT

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_rows(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            d.Close(False)
        finally:
            app.Quit()
    doc = fitz.open(pdf)
    rows = []
    for pno, pg in enumerate(doc):
        for blk in pg.get_text("dict")["blocks"]:
            for l in blk.get("lines", []):
                sp = [s for s in l["spans"] if s["text"].strip()]
                if sp:
                    rows.append((pno, round(sp[0]["origin"][1], 2),
                                 "".join(s["text"] for s in sp).strip()))
    doc.close()
    rows.sort()
    return rows


def oxi_rows(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    rows = {}
    for pno, pg in enumerate(d["pages"]):
        for e in pg["elements"]:
            t = (e.get("text") or "")
            if t.strip():
                rows.setdefault((pno, round(e["y"] + e.get("text_y_off", 0.0), 2)),
                                []).append(t)
    return [(k[0], k[1], "".join(rows[k])) for k in sorted(rows)]


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_rows if mode == "word" else oxi_rows
print("%s   HEAD -> NEXT gap (keepNext x after x same-style)\n" % mode.upper())
print("  arm                       keep after same top fill  pre   head_y    next_y      gap  same_pg")
for tag, k, a, s, tp, f, pr in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    rows = reader(docx)
    h = next(((p, y) for p, y, t in rows if t.startswith("HEAD")), None)
    n = next(((p, y) for p, y, t in rows if t.startswith("NEXT")), None)
    hy = h[1] if h else None
    ny = n[1] if n else None
    if hy is None or ny is None:
        print("  %-25s  (HEAD %s NEXT %s)" % (tag, hy, ny))
        continue
    print("  %-25s %4d %5.1f %4d %3d %4d %4.0f %8.2f %9.2f %8.2f %8s"
          % (tag, k, a / 20.0, s, tp, f, pr / 20.0, hy, ny, ny - hy,
             "yes" if h[0] == n[0] else "no"))
