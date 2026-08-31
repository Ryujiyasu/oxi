# -*- coding: utf-8 -*-
"""Read the bodyindent probe. Usage: _pb_bodyindent_read.py word|oxi

Per arm: the within-paragraph line pitch and the advance from one paragraph's
last line to the next paragraph's first. If the advance is one pitch there is
no spacing; if it is two, something is inserting a blank line.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_bodyindent_gen import ARMS, OUT

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
    for blk in doc[0].get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            sp = [s for s in l["spans"] if s["text"].strip()]
            if sp:
                rows.append((round(sp[0]["origin"][1], 2),
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
    for e in d["pages"][0]["elements"]:
        t = (e.get("text") or "")
        if t.strip():
            rows.setdefault(round(e["y"] + e.get("text_y_off", 0.0), 2), []).append((e.get("x", 0), t))
    return [(y, " ".join(t for _, t in sorted(rows[y]))) for y in sorted(rows)]


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_rows if mode == "word" else oxi_rows
print("%s   line pitch vs paragraph advance\n" % mode.upper())
print("  arm               grid snap num  jc pprdef  pitch   advance")
for tag, g, s, n, j, dd in ARMS:
    rows = reader(os.path.join(OUT, tag + ".docx"))
    ys = [y for y, t in rows]
    starts = [y for y, t in rows if "P1 " in t or "P2 " in t or "P3 " in t]
    if len(ys) < 3 or len(starts) < 2:
        print("  %-17s %4d %4d %3d %3d %5d  (rows %d starts %d)" % (tag, g, s, n, j, dd, len(ys), len(starts)))
        continue
    inner = [b - a for a, b in zip(ys, ys[1:]) if b not in starts and b - a > 1]
    pitch = min(inner) if inner else float("nan")
    adv = starts[1] - starts[0]
    print("  %-17s %4d %4d %3d %3d %5d %7.2f %9.2f"
          % (tag, g, s, n, j, dd, pitch, adv))
