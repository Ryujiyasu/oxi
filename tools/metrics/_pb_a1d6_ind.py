# -*- coding: utf-8 -*-
"""Does a1d6e4ef's note keep its TWIP indents when the row is sliced out?

`_pb_indchars_gen.py` says a non-zero *Chars beats the twip beside it, and the
character is the grid pitch. a1d6e4ef's own note carries
`leftChars=50 left=489 hangingChars=203 hanging=380` and renders at 24.45/19.00
-- the twips. Everything the probe can vary (compat, grid, style file, style
name, hanging, magnitude) leaves the probe on the *Chars side, so the difference
must be in the document around the paragraph.

Slice the row out verbatim and read the same two numbers. Twips in the slice =>
the cause is inside the row; *Chars => it is outside it.

    python _pb_a1d6_ind.py
"""
import glob
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_a1d6_ind")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "a1d6*.docx"))
       if not os.path.basename(p).startswith("~$")][0]
MARK = "提供依頼申出"
# The slice reproduced the twips (59.90 / 78.98 against 59.69 / 78.69 predicted).
# So the cause is inside the row. Replace ONLY the note paragraph's <w:ind> and
# watch which unit Word uses in THIS cell.
INDS = [
    ("as_written", '<w:ind w:leftChars="50" w:left="489" w:rightChars="50"'
                   ' w:right="109" w:hangingChars="203" w:hanging="380"/>'),
    ("no_leftchars", '<w:ind w:left="489" w:rightChars="50" w:right="109"'
                     ' w:hangingChars="203" w:hanging="380"/>'),
    ("no_left_tw", '<w:ind w:leftChars="50" w:rightChars="50" w:right="109"'
                   ' w:hangingChars="203" w:hanging="380"/>'),
    ("probe_arm", '<w:ind w:leftChars="100" w:left="81"/>'),
    ("lc50_l489", '<w:ind w:leftChars="50" w:left="489"/>'),
]


def build():
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    rs = max((x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i)))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    i2 = row.index(MARK)
    ps = max((row.rfind("<w:p ", 0, i2), row.rfind("<w:p>", 0, i2)))
    pe = row.index("</w:p>", i2) + len("</w:p>")
    para = row[ps:pe]
    blocks = []
    for name, ind in INDS:
        p2 = re.sub(r"<w:ind[^>]*/>", ind, para, count=1)
        blocks.append(tbl_head + row[:ps] + p2 + row[pe:] + "</w:tbl>"
                      + '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
    doc = head + "".join(blocks) + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "a1d6_row.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = doc.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    print("built", dst)
    return dst


def to_pdf(docx):
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
    print("arms:", [n for n, _ in INDS])
    p = doc[0]
    rules = []
    for d in p.get_drawings():
        for it in d["items"]:
            if it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > 3:
                rules.append(round((it[1].x + it[2].x) / 2, 2))
            elif it[0] == "re" and it[1].width < 0.9 and it[1].height > 3:
                rules.append(round(it[1].x0, 2))
    left_rule = min(rules) if rules else 0.0
    inner = left_rule + 12 / 20.0          # the table's own tblCellMar left = 12tw
    lines = []
    for b in p.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            ch = sorted([c for s in l["spans"] for c in s["chars"]],
                        key=lambda c: c["origin"][0])
            if ch and ch[0]["origin"][0] < inner + 60 and MARK[:4] in "".join(
                    c["c"] for c in ch):
                lines.append((round(l["bbox"][1], 1), ch[0]["origin"][0],
                              "".join(c["c"] for c in ch)[:12]))
    lines.sort()
    print("rules:", sorted(set(rules))[:6])
    print("cell inner left %.2f" % inner)
    print("  twips would give   first %.2f  cont %.2f" % (inner + 489 / 20.0 - 380 / 20.0,
                                                          inner + 489 / 20.0))
    print("  *Chars would give  first %.2f  cont %.2f  (1 char = 10.8547 grid pitch)"
          % (inner + 0.5 * 10.8547 - 2.03 * 10.8547, inner + 0.5 * 10.8547))
    for y, x, t in lines[:20]:
        print("   y=%7.1f x0=%7.2f %s" % (y, x, t))


if __name__ == "__main__":
    measure(to_pdf(build()))
