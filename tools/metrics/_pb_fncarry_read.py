# -*- coding: utf-8 -*-
"""Read the footnote-carry probe. Usage: _pb_fncarry_read.py word|oxi

Per arm on page 1: the last body line, the footnote separator's y, the gap
between them, and how many notes landed. A note is CARRIED when fewer notes are
on page 1 than there are references on it -- that is the case the real failing
document is in, and the gap is what Oxi gets wrong there.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fncarry_gen import ARMS, OUT

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
    body, notes, refs = [], 0, 0
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
                body.append((y, t[:22]))
                if t.startswith("REF"):
                    refs += 1
    body.sort()
    return body, notes, refs, sep, doc.page_count


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
    body, notes, refs = [], 0, 0
    for y in sorted(rows):
        t = "".join(t for _, t in sorted(rows[y])).strip()
        if sep is not None and y > sep:
            notes += 1
        else:
            body.append((y, t[:22]))
            if t.startswith("REF"):
                refs += 1
    return body, notes, refs, sep, len(d["pages"])


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_page1 if mode == "word" else oxi_page1
BASE = 10.5 if mode == "oxi" else 0.0   # box-top vs baseline convention
print("%s   page 1  (gap is normalised to Word's baseline convention)\n" % mode.upper())
print("  arm      last_body            y      sep      gap   refs  notes  carried")
for tag, nb, nr in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    if not os.path.exists(docx):
        print("  %-8s MISSING" % tag)
        continue
    body, notes, refs, sep, npages = reader(docx)
    if not body or sep is None:
        print("  %-8s (body %d, sep %s)" % (tag, len(body), sep))
        continue
    ly, lt = body[-1]
    print("  %-8s %-20s %7.2f %7.2f %7.2f %5d %6d   %s"
          % (tag, lt, ly + BASE, sep, sep - (ly + BASE), refs, notes,
             "YES" if notes < refs else "no"))
