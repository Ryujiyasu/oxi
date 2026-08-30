# -*- coding: utf-8 -*-
"""Read the fnkeep sweep. Usage: _pb_fnkeep_read.py word|oxi [--sweep lo hi step]

Per arm: FINAL's page, how many notes sit in page 1's area and page 2's, and
FINAL's baseline when it stayed. The spacer value where FINAL flips to page 2
is the keep boundary; comparing that boundary across nown gives the weight the
keep test puts on the line's own notes.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fnkeep_gen import OUT, NOWN, NPRIOR, parse_sweep

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
        out.append((rows, sep))
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
                     for y in sorted(rows)], sep))
    return out


def one(reader, docx):
    pages = reader(docx)
    p1rows, p1sep = pages[0]
    p1b = [(y, t) for y, t in p1rows if p1sep is None or y <= p1sep]
    p1n = [1 for y, t in p1rows if p1sep is not None and y > p1sep]
    if len(pages) > 1:
        p2rows, p2sep = pages[1]
        p2n = [1 for y, t in p2rows if p2sep is not None and y > p2sep]
    else:
        p2n = []
    fy = next((y for y, t in p1b if "FINAL" in t), None)
    lasty = p1b[-1][0] if p1b else float("nan")
    return (1 if fy is not None else 2), len(p1n), len(p2n), fy, lasty


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_pages if mode == "word" else oxi_pages
sw = parse_sweep(sys.argv)
print("%s   keep boundary sweep (nprior=%d)  pg=FINAL page, n1/n2=notes per area\n"
      % (mode.upper(), NPRIOR))
print("  spacer |" + "".join("   o%d pg n1 n2 final_y |" % o for o in NOWN))
for x in sw:
    cells = []
    for o in NOWN:
        docx = os.path.join(OUT, "s%05d_o%d.docx" % (x, o))
        if not os.path.exists(docx):
            cells.append("      MISSING        |")
            continue
        pg, n1, n2, fy, _ = one(reader, docx)
        cells.append("    %d %2d %2d %7s |"
                     % (pg, n1, n2, ("%.2f" % fy) if fy is not None else "-"))
    print("  %6d |%s" % (x, "".join(cells)))
