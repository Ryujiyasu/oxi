# -*- coding: utf-8 -*-
"""Read the keepNext-chain table probe v2. Usage: _pb_kntbl2_read.py word | oxi"""
import os, re, sys, json, subprocess, time
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_kntbl2_gen import ARMS, OUT, SHAPES, FILLS
REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")

def word_pages(docx):
    import fitz, win32com.client
    pdf = docx[:-5] + ".pdf"
    if not os.path.exists(pdf):
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False; word.DisplayAlerts = 0
        try:
            d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.SaveAs2(os.path.abspath(pdf), FileFormat=17); d.Close(False)
        finally:
            word.Quit()
    doc = fitz.open(pdf); out = []
    for pi in range(doc.page_count):
        for b in doc[pi].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if t: out.append((pi + 1, round(l["bbox"][1], 1), t))
    return sorted(out)

def oxi_pages(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8")); out = []
    for pi, pg in enumerate(d["pages"], 1):
        rows = {}
        for e in pg["elements"]:
            if not (e.get("text") or "").strip(): continue
            rows.setdefault(round(e["y"], 1), []).append((round(e["x"], 1), e["text"]))
        for y, frs in rows.items():
            frs.sort(); out.append((pi, y, "".join(t for _, t in frs)))
    return sorted(out)

TALL_WORDS = ("Non contentious or common form probate business and other "
              "proceedings of a like kind").split()

def analyse(lines):
    """caption page, header-row page, per-page line count of the tall cell, R3 page.

    v1's detector keyed on three fixed substrings, which only covered 3 of the
    tall cell's 4 lines -- with the faithful package the cell wraps to 5 and two
    of the keys land on the same line, so a split between lines 3 and 4 reads as
    "not split".  Match any contiguous word-slice of the tall string instead, so
    every line of the cell is counted wherever it falls.
    """
    cap = next((p for p, y, t in lines if "CAPTION" in t), None)
    hdr = next((p for p, y, t in lines if "HDR-A" in t), None)
    r3  = next((p for p, y, t in lines if "R3-A" in t), None)
    tall = {}
    for p, y, t in lines:
        w = t.replace("R2-B", "").split()
        if not w or len(w) > len(TALL_WORDS):
            continue
        n = len(w)
        if any(TALL_WORDS[i:i + n] == w for i in range(len(TALL_WORDS) - n + 1)):
            tall[p] = tall.get(p, 0) + 1
    return cap, hdr, tall, r3

mode = sys.argv[1] if len(sys.argv) > 1 else "word"
print(f"{mode.upper()}   cap=caption page, hdr=header row page, "
      f"tall=pages the 3-line row occupies, r3=first non-keepNext row page")
res = {}
for tag, n, ck, rk in ARMS:
    docx = os.path.join(OUT, tag + ".docx")
    lines = word_pages(docx) if mode == "word" else oxi_pages(docx)
    res[tag] = analyse(lines)
for s in SHAPES:
    print(f"\n--- shape {s} ---")
    for n in FILLS:
        cap, hdr, tall, r3 = res[f"{s}{n}"]
        split = "SPLIT" if len(tall) > 1 else "     "
        shape = " ".join(f"p{p}x{c}" for p, c in sorted(tall.items()))
        print(f"   fill={n:3d}  cap=p{cap}  hdr=p{hdr}  tall=[{shape}] {split}  r3=p{r3}")
