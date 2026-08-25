# -*- coding: utf-8 -*-
"""Does the 約物 pool depend on HOW FAR the mark is from the squeeze?

S1209 fixed the pool at half an em for any line carrying a mark, but its 40 arms
placed the marks every third character, so the mark was never more than three
characters from the line end. b35123fe8efc p2 is the case that does not fit: its
、 sits TEN characters before the break, Oxi grants the half em and packs one
more character in, and Word does not (the same sentence one page earlier, 3.6pt
further left, Word DOES break one character later -- so the line sits exactly on
the boundary).

One arm per distance, in a cell, ＭＳ 明朝 10.5pt, jc=both, right indent swept in
0.25pt steps. The credit is the arm's flip point minus the mark-free control's.

    python _pb_yakudist_gen.py gen
    python _pb_yakudist_gen.py pdf
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_yakudist")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CELL_TW = 8000                     # 400pt
EM = 10.5
NCH = 36
DISTS = [1, 2, 3, 5, 8, 10, 12, 16]
# CS=<n> gives the section a docGrid with that charSpace. b35123fe8efc's grid is
# NEGATIVE (-2714), so its characters are already squeezed to 9.837pt before any
# 約物 compression -- the arms below run without a grid, where the pool is a flat
# half em at every distance.
CS = int(os.environ.get("CS") or 0)
# COMPAT15=1 rewrites the base's settings.xml to the modern engine. The base is
# compat 11; b35123fe8efc, whose corpus lines CONTRADICT the CS=-2714 reading, is
# compat 15.
COMPAT15 = os.environ.get("COMPAT15") == "1"
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')


def text_for(d):
    """`d` = characters between the mark and the last character of the line."""
    t = ["甲"] + ["亜"] * (NCH - 1)
    if d:
        t[NCH - 1 - d] = "、"
    return "".join(t)


def window():
    mid = CELL_TW / 20.0 - NCH * EM
    return list(range(max(0, int((mid - 6) * 20)), int((mid + 10) * 20), 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    blocks, index = [], []
    for d in [0] + DISTS:
        txt = text_for(d)
        for r in window():
            index.append((d, r))
            blocks.append(
                '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
                '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
                '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
                '</w:tblCellMar></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>'
                % (CELL_TW, CELL_TW, CELL_TW, r, txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    if CS:
        sect = re.sub(r"<w:docGrid[^>]*/>",
                      '<w:docGrid w:type="linesAndChars" w:linePitch="360" '
                      'w:charSpace="%d"/>' % CS, sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (W_NS, "".join(blocks), sect))
    dst = os.path.join(OUT, "yakudist%s.docx" % ((("_cs%d" % CS) if CS else "") + ("_c15" if COMPAT15 else "")))
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        elif COMPAT15 and item.filename == "word/settings.xml":
            t = data.decode("utf-8").replace("<w:useAltKinsokuLineBreakRules/>", "")
            data = re.sub(r'(w:name="compatibilityMode"[^>]*w:val=")[0-9]+',
                          r"\g<1>15", t).encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "arms%s.txt" % ((("_cs%d" % CS) if CS else "") + ("_c15" if COMPAT15 else ""))), "w", encoding="utf-8") as fh:
        for d, r in index:
            fh.write("%d %d\n" % (d, r))
    print("built %s (%d arms)" % (dst, len(index)))


def to_pdf():
    import win32com.client as wc
    docx = os.path.join(OUT, "yakudist%s.docx" % ((("_cs%d" % CS) if CS else "") + ("_c15" if COMPAT15 else "")))
    pdf = os.path.join(OUT, "yakudist%s.pdf" % ((("_cs%d" % CS) if CS else "") + ("_c15" if COMPAT15 else "")))
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def measure():
    import fitz
    arms = [tuple(int(v) for v in l.split()) for l in
            open(os.path.join(OUT, "arms%s.txt" % ((("_cs%d" % CS) if CS else "") + ("_c15" if COMPAT15 else ""))), encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "yakudist%s.pdf" % ((("_cs%d" % CS) if CS else "") + ("_c15" if COMPAT15 else ""))))
    heads = []
    for page in doc:
        rows = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch:
                    rows.append((round(l["bbox"][1], 1), ch))
        for _, ch in sorted(rows, key=lambda t: t[0]):
            if ch[0]["c"] == "甲":
                heads.append(len(ch))
    if len(heads) != len(arms):
        print("%d heads for %d arms" % (len(heads), len(arms)))
        return
    by = {}
    for (d, r), n in zip(arms, heads):
        by.setdefault(d, []).append((r / 20.0, n))
    top = max(n for v in by.values() for _, n in v)
    base = None
    print("   charSpace=%d  engine=%s" % (CS, "compat 15" if COMPAT15 else "compat 11"))
    print("   marks at distance   flip r      credit (em)")
    for d in [0] + DISTS:
        full = [r for r, n in by[d] if n >= top]
        if not full:
            print("   %-18s never fits" % d)
            continue
        f = max(full)
        if base is None:
            base = f
        print("   %-18s %7.2f     %+.3f" % ("no mark" if d == 0 else "%d chars" % d,
                                            f, (f - base) / EM))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
