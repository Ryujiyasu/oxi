# -*- coding: utf-8 -*-
"""Read the footnote-deferral probe. Usage: _pb_fndefer_read.py word|oxi

Per arm: the last BODY line on page 1, the footnote lines on page 1, and how much
blank column is left below the last body line. If Word stops the body at the
reference when the note cannot fit, the last body line reads REFLINE and the
blank is large.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import ARMS, OUT

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
    body, fn = [], []
    for blk in doc[0].get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            sp = [s for s in l["spans"] if s["text"].strip()]
            if not sp:
                continue
            t = "".join(s["text"] for s in sp)
            (body if max(s["size"] for s in sp) >= 10.5 else fn).append(
                (round(sp[0]["origin"][1], 2), t.strip()[:24]))
    body.sort()
    fn.sort()
    return body, fn, doc.page_count


def oxi_page1(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    rows = {}
    for e in d["pages"][0]["elements"]:
        if not (e.get("text") or "").strip():
            continue
        big = (e.get("font_size") or 0) >= 10.5
        rows.setdefault((big, round(e["y"], 2)), []).append((e.get("x", 0), e["text"]))
    body = sorted((y, "".join(t for _, t in sorted(v))[:24]) for (b, y), v in rows.items() if b)
    fn = sorted((y, "".join(t for _, t in sorted(v))[:24]) for (b, y), v in rows.items() if not b)
    return body, fn, len(d["pages"])


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_page1 if mode == "word" else oxi_page1
print("%s   page 1: where the body stops and which notes landed\n" % mode.upper())
print("  arm         pages  last_body_line            y      fn_lines  blank_below")
for tag, nref, fnlen in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    if not os.path.exists(docx):
        print("  %-11s MISSING" % tag)
        continue
    body, fn, npages = reader(docx)
    if not body:
        print("  %-11s (no body)" % tag)
        continue
    ly, lt = body[-1]
    fnly = fn[0][0] if fn else 769.92
    print("  %-11s %5d  %-24s %7.2f %6d      %7.2f"
          % (tag, npages, lt, ly, len(fn), fnly - ly))
