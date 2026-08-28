# -*- coding: utf-8 -*-
"""Read the cell-margin default probe. Usage: _pb_cellmar_read.py [word|oxi|both]"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_cellmar_gen import ARMS, OUT
REND = os.path.abspath("tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
BODY_X = None  # body-text origin, filled from the first arm

def word_lines(docx):
    import fitz, win32com.client, time
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
    for b in doc[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = s["text"].strip()
                if t: out.append((round(s["origin"][1], 2), round(s["origin"][0], 3),
                                  round(s["bbox"][2] - s["bbox"][0], 2), t))
    return sorted(out)

def oxi_lines(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8")); out = []
    for e in d["pages"][0]["elements"]:
        t = (e.get("text") or "").strip()
        if t: out.append((round(e["y"], 2), round(e["x"], 3), round(e.get("width", 0), 2), t))
    return sorted(out)

def report(name, rows):
    body = next((x for _, x, _, t in rows if t.startswith("#")), None)
    cell = [(y, x, w, t) for y, x, w, t in rows
            if not t.startswith("#") and t != "AFTER" and t != "B"]
    padl = (cell[0][1] - body) if cell and body is not None else float("nan")
    print(f"  {name:5s} body_x={body:8.3f}  pad_l={padl:6.2f}pt ({padl*20:6.1f}tw)  "
          f"lines={len(cell)}")
    for y, x, w, t in cell:
        print(f"          y={y:8.2f} x={x:8.3f} w={w:7.2f}  {t[:44]!r}")

mode = sys.argv[1] if len(sys.argv) > 1 else "both"
for tag in ARMS:
    tn, tw, tind = ARMS[tag]
    print(f"\n=== {tag}  TableNormal={tn}  tblPr cellMar={tw} tblInd={tind} ===")
    if mode in ("word", "both"): report("WORD", word_lines(os.path.join(OUT, tag + ".docx")))
    if mode in ("oxi", "both"):  report("OXI ", oxi_lines(os.path.join(OUT, tag + ".docx")))
