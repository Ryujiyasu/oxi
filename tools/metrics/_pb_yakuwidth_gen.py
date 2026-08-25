# -*- coding: utf-8 -*-
"""How wide is a mark, and does a RUN of marks share one em?

b35123fe8efc's 「…（ファイルの保管を含む。）…、…」 line fits one more character in
Word than Oxi, and 576 pool arms say the 約物 pool cannot be the reason (a
compressing grid gives no credit at any distance, size or compat). The remaining
candidate is the marks' own advances: JIS X 4051 pairs a period with a following
closing bracket into one em, and Oxi may be paying two.

Arms are LEFT-aligned short lines (nothing to stretch), read glyph by glyph out
of Word's PDF: the advance of each mark is the next origin minus its own.

    python _pb_yakuwidth_gen.py            # build, export, report
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_yakuwidth")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CS = int(os.environ.get("CS") or 0)
SZ = int(os.environ.get("SZ") or 21)
# CELL=1 puts each arm in a fixed-width cell. The pair-compression code lives in
# break_into_lines and in the justify pass -- both BODY paths -- so ask whether a
# cell compresses adjacent marks at all.
CELL = os.environ.get("CELL") == "1"
EM = SZ / 2.0
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
ARMS = [
    ("plain", "甲亜亜亜亜亜"),
    ("period", "甲亜。亜亜亜"),
    ("comma", "甲亜、亜亜亜"),
    ("open", "甲亜（亜亜亜"),
    ("close", "甲亜）亜亜亜"),
    ("period_close", "甲亜。）亜亜"),
    ("close_period", "甲亜）。亜亜"),
    ("two_periods", "甲亜。。亜亜"),
    ("open_open", "甲亜（（亜亜"),
    ("bracket_pair", "甲（亜）亜亜"),
]


def build():
    os.makedirs(OUT, exist_ok=True)
    paras = []
    for _, txt in ARMS:
        para = (
            '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="left"/>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (SZ, SZ, txt))
        if CELL:
            para = ('<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>'
                    '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
                    '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
                    '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>'
                    '<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
                    + para + '</w:tc></w:tr></w:tbl>'
                    '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
        paras.append(para)
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    if CS:
        sect = re.sub(r"<w:docGrid[^>]*/>",
                      '<w:docGrid w:type="linesAndChars" w:linePitch="360" '
                      'w:charSpace="%d"/>' % CS, sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (W_NS, "".join(paras), sect))
    tag = ("_cs%d" % CS if CS else "") + ("_sz%d" % SZ) + ("_cell" if CELL else "")
    dst = os.path.join(OUT, "yakuw%s.docx" % tag)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst


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


def measure(pdf):
    import fitz
    doc = fitz.open(pdf)
    rows = []
    for page in doc:
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch and ch[0]["c"] == "甲":
                    rows.append(ch)
    print("charSpace=%d  size=%.1fpt  (one em = %.2f, grid pitch = %.4f)"
          % (CS, EM, EM, EM + CS / 4096.0))
    for (name, txt), ch in zip(ARMS, rows):
        adv = [round(ch[i + 1]["origin"][0] - ch[i]["origin"][0], 3)
               for i in range(len(ch) - 1)]
        print("   %-14s %-8s %s" % (name, txt[1:6], adv))


if __name__ == "__main__":
    measure(export(build()))
