# -*- coding: utf-8 -*-
"""Whose font size is the "character" of a *Chars indent?

`_pb_indchars_gen.py` varied the paragraph mark's size and the run's size
TOGETHER, so it could only say that `firstLineChars` shrank with the size while
`leftChars` did not. Three candidates are still tangled: the STYLE's size, the
PARAGRAPH MARK's size (the rPr inside pPr) and the FIRST RUN's size.

So vary them one at a time. Style "a" is 10.5pt in the base document; every arm
carries a single *Chars indent so the unit can be read straight off the first
glyph's origin.

    python _pb_charunit_gen.py
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_charunit")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CELL_TW, MAR_TW = 6000, 108
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
# (mark size, run size) in half-points; None = the attribute is absent
COMBOS = [(18, 18), (18, 21), (21, 18), (None, 18), (None, 21)]
# 04b88e7e renders `firstLineChars` at 6.00pt per character with runs at sz=16
# (8pt) and w:spacing=-20 (-1pt tracking): 8 - 1 = 7, but 8 + 2x(-1) = 6. So put
# the tracking in the arms and read the unit off it. (run size, spacing twips)
TRACK = [(21, 0), (21, -20), (21, -40), (16, -20), (16, 0)]
INDS = [("leftChars", '<w:ind w:leftChars="100"/>'),
        ("firstLineChars", '<w:ind w:firstLineChars="100"/>')]


def sz(v, track=0):
    out = '<w:sz w:val="%d"/>' % v if v else ""
    return ('<w:spacing w:val="%d"/>' % track if track else "") + out


def build():
    os.makedirs(OUT, exist_ok=True)
    blocks, index = [], []
    arms = [(k, i, m, r, 0) for k, i in INDS for m, r in COMBOS]
    arms += [(k, i, None, r, t) for k, i in INDS for r, t in TRACK]
    for kind, ind, mark, run, track in arms:
            index.append((kind, mark, run, track))
            blocks.append(
                '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
                '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
                '<w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/>'
                '</w:tblCellMar><w:tblBorders>'
                '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                '</w:tblBorders></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="left"/>%s'
                '<w:rPr><w:rFonts w:hint="eastAsia"/>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/>%s</w:rPr>'
                '<w:t>甲亜亜亜</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>'
                % (CELL_TW, MAR_TW, MAR_TW, CELL_TW, CELL_TW, ind, sz(mark),
                   sz(run, track)))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    body = "<w:body>" + "".join(blocks) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "charunit.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst, index


def export(docx):
    import win32com.client as wc
    pdf = os.path.splitext(docx)[0] + ".pdf"
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()
    return pdf


def measure(pdf, index):
    import fitz
    doc = fitz.open(pdf)
    heads, rules = [], []
    for page in doc:
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > 3:
                    rules.append(round((it[1].x + it[2].x) / 2, 2))
                elif it[0] == "re" and it[1].width < 0.9 and it[1].height > 3:
                    rules.append(round(it[1].x0, 2))
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
                heads.append(ch[0]["origin"][0])
    inner = (min(rules) if rules else 0.0) + MAR_TW / 20.0
    print("cell inner edge %.2f   style size 10.5" % inner)
    print("   attribute        mark   run   track    x0      unit   size+track  size+2track")
    for (kind, mark, run, track), x in zip(index, heads):
        u = x - inner - 0.24            # 0.24 = the zero-indent baseline
        base = run / 2.0
        t = track / 20.0
        print("   %-15s %-6s %-5s %-6.2f %7.2f  %6.2f   %6.2f      %6.2f"
              % (kind, mark / 2 if mark else "-", base, t, x, u, base + t,
                 base + 2 * t))


if __name__ == "__main__":
    docx, index = build()
    measure(export(docx), index)
