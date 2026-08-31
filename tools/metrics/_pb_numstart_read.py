# -*- coding: utf-8 -*-
"""Read the numstart probe. Usage: _pb_numstart_read.py word|oxi

Prints the marker each engine put in front of every numbered paragraph.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_numstart_gen import ARMS, OUT

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_lines(docx):
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
    out = []
    for blk in doc[0].get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                out.append((round(l["bbox"][1], 2), t))
    doc.close()
    out.sort()
    return [t for _, t in out]


def oxi_lines(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    rows = {}
    for e in d["pages"][0]["elements"]:
        t = (e.get("text") or "")
        if t.strip():
            rows.setdefault(round(e["y"], 2), []).append((e.get("x", 0), t))
    return [" ".join(t for _, t in sorted(rows[y])) for y in sorted(rows)]


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_lines if mode == "word" else oxi_lines
print("%s   markers per arm\n" % mode.upper())
for tag, p0, p1, ch, used in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    lines = reader(docx)
    print("  %-14s starts %d/%d/%d used=%s" % (tag, p0, p1, ch, used))
    for t in lines:
        print("       %s" % t[:60])
