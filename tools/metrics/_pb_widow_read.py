# -*- coding: utf-8 -*-
"""Read the widow/orphan probe. Usage: _pb_widow_read.py word|oxi

Per arm: how many lines of the test paragraph stayed on page 1. With
widowControl ON, Word should never leave exactly one of them behind (nor carry
exactly one over); with it OFF the split should follow the room alone.
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_widow_gen import ARMS, OUT, FILLERS, PLINES, WIDOW

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
    for pno in range(doc.page_count):
        lines = []
        for blk in doc[pno].get_text("dict")["blocks"]:
            for l in blk.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if t:
                    lines.append(t)
        out.append(lines)
    return out


def oxi_pages(docx):
    dump = docx[:-5] + ".layout.json"
    subprocess.run([REND, docx, docx[:-5] + "_r", "96", "--dump-layout=" + dump],
                   capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pg in d["pages"]:
        rows = {}
        for e in pg["elements"]:
            if not (e.get("text") or "").strip():
                continue
            rows.setdefault(round(e["y"], 2), []).append((e.get("x", 0), e["text"]))
        out.append(["".join(t for _, t in sorted(v)).strip() for _, v in sorted(rows.items())])
    return out


def pcount(page_lines, plines):
    """How many LINES of the test paragraph are on this page.

    ★Counting the P01/P02 markers counts the wrong thing: they are words inside
    one wrapping paragraph and do not start lines, so a 3-line paragraph could
    report 2 with all three lines present. The filler word `wwww` appears only in
    that paragraph, so a line carrying it IS one of its lines.
    """
    return sum(1 for t in page_lines if "wwww" in t)


mode = sys.argv[1] if len(sys.argv) > 1 else "word"
reader = word_pages if mode == "word" else oxi_pages
print("%s   lines of the test paragraph kept on page 1\n" % mode.upper())
print("  widow  plines   " + "  ".join("fill=%d" % f for f in FILLERS))
for w in WIDOW:
    for p in PLINES:
        cells = []
        for f in FILLERS:
            docx = os.path.join(OUT, "%s_f%d_p%d.docx" % (w, f, p))
            if not os.path.exists(docx):
                cells.append("  --   ")
                continue
            pages = reader(docx)
            on1 = pcount(pages[0], p) if pages else 0
            cells.append(" %d/%d   " % (on1, p))
        print("  %-6s %-7d  %s" % (w, p, "".join(cells)))
