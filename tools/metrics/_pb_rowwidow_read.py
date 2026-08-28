# -*- coding: utf-8 -*-
"""Read the row-vs-paragraph widow probe. Usage: _pb_rowwidow_read.py word|oxi"""
import os, sys, json, subprocess, re
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_rowwidow_gen import ARMS, OUT, SHAPES, FILLS, _shape

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
TOK = re.compile(r"(A{13}|B{13})(\d\d)")


def word_lines(docx):
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
    out = []
    for pi in range(doc.page_count):
        for b in doc[pi].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                out.append((pi + 1, "".join(s["text"] for s in l["spans"])))
    return out


def oxi_lines(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    return [(pi, e.get("text") or "")
            for pi, pg in enumerate(d["pages"], 1) for e in pg["elements"]]


def counts(lines):
    c = {("A", 1): 0, ("A", 2): 0, ("B", 1): 0, ("B", 2): 0}
    for p, t in lines:
        for m in TOK.finditer(t):
            key = (m.group(1)[0], p)
            if key in c:
                c[key] += 1
    return c


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
res = {}
for tag in [a[0] for a in ARMS]:
    docx = os.path.join(OUT, tag + ".docx")
    res[tag] = counts(word_lines(docx) if mode == "word" else oxi_lines(docx))
print("%s   cellA p1/p2   cellB p1/p2" % mode.upper())
for s in SHAPES:
    na, nb, pre, sp = _shape(SHAPES[s])
    print("\n--- %s  A=%d B=%d ---" % (s, na, nb))
    for n in FILLS:
        c = res["%s%d" % (s, n)]
        print("   fill=%3d  A %d/%d   B %d/%d"
              % (n, c[("A", 1)], c[("A", 2)], c[("B", 1)], c[("B", 2)]))
