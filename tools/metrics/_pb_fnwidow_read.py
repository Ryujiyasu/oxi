# -*- coding: utf-8 -*-
"""Read the footnote + widow composite probe. Usage: _pb_fnwidow_read.py word|oxi

Per arm on page 1: how many lines of the test paragraph stayed, how many notes
landed, and the gap from the last body line to the separator. Lines of the test
paragraph are counted by the filler word `wwww`, which appears nowhere else.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fnwidow_gen import ARMS, OUT, NFILL, NREFS, PLINES

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
    sep = None
    for dr in doc[0].get_drawings():
        r = dr["rect"]
        if r.height < 4 and 100 < r.width < 200 and r.y0 > 300:
            sep = r.y0 if sep is None else min(sep, r.y0)
    body, notes = [], 0
    for blk in doc[0].get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            sp = [s for s in l["spans"] if s["text"].strip()]
            if not sp:
                continue
            t = "".join(s["text"] for s in sp).strip()
            y = round(sp[0]["origin"][1], 2)
            if sep is not None and y > sep:
                notes += 1
            else:
                body.append((y, t))
    body.sort()
    return body, notes, sep, 0.0


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
        rows.setdefault(round(e["y"], 2), []).append((e.get("x", 0), e["text"]))
    body, notes = [], 0
    for y in sorted(rows):
        t = "".join(t for _, t in sorted(rows[y])).strip()
        if sep is not None and y > sep:
            notes += 1
        else:
            body.append((y, t))
    return body, notes, sep, 10.5      # box-top -> baseline convention


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_page1 if mode == "word" else oxi_page1
print("%s   page 1: paragraph lines kept / notes / gap to separator\n" % mode.upper())
print("  arm         plines_kept  notes  last_body    sep      gap")
for tag, nf, nr, pl in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    if not os.path.exists(docx):
        print("  %-11s MISSING" % tag)
        continue
    body, notes, sep, base = reader(docx)
    if not body or sep is None:
        print("  %-11s (body %d sep %s)" % (tag, len(body), sep))
        continue
    kept = sum(1 for _, t in body if "wwww" in t)
    ly = body[-1][0] + base
    print("  %-11s %8d %8d %10.2f %8.2f %8.2f"
          % (tag, kept, notes, ly, sep, sep - ly))
