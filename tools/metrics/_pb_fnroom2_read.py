# -*- coding: utf-8 -*-
"""Read the fnroom2 probe. Usage: _pb_fnroom2_read.py word|oxi

Per arm: which page the FINAL line landed on, how many notes sit in page 1's
note area and how many in page 2's, and page 1's last body line / separator.
A note in page 2's area while FINAL is on page 1 = the note was rolled.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fnroom2_gen import ARMS, OUT

REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")


def word_pages(docx):
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
    for pno in range(min(2, doc.page_count)):
        pg = doc[pno]
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
        out.append((rows, sep, 0.0))
    return out


def oxi_pages(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pg in d["pages"][:2]:
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
        out.append(([(y, "".join(t for _, t in sorted(rows[y])).strip())
                     for y in sorted(rows)], sep, 10.5))
    return out


def split(rows, sep):
    body = [(y, t) for y, t in rows if sep is None or y <= sep]
    notes = [(y, t) for y, t in rows if sep is not None and y > sep]
    return body, notes


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_pages if mode == "word" else oxi_pages
print("%s   FINAL line page / notes per page area\n" % mode.upper())
print("  arm             final_pg p1seg p1notes p2notes p1_last_body  p1_sep   gap  rolled")
for tag, nf, sg, wd, no in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    pages = reader(docx)
    p1rows, p1sep, _ = pages[0]
    p1body, p1notes = split(p1rows, p1sep)
    if len(pages) > 1:
        p2rows, p2sep, _ = pages[1]
        p2body, p2notes = split(p2rows, p2sep)
    else:
        p2body, p2notes = [], []
    final_pg = 1 if any("FINAL" in t for _, t in p1body) else 2
    p1seg = sum(1 for _, t in p1body if "wwww" in t)
    ly = p1body[-1][0] if p1body else float("nan")
    gap = (p1sep - ly) if (p1sep is not None and p1body) else float("nan")
    # a note rolled = FINAL on page 1 yet notes sit in page 2's area
    rolled = "YES" if (final_pg == 1 and len(p2notes) > 0) else ""
    print("  %-15s %5d %5d %7d %7d %11.2f %8.2f %7.2f  %s"
          % (tag, final_pg, p1seg, len(p1notes), len(p2notes), ly,
             p1sep if p1sep else float("nan"), gap, rolled))
