# -*- coding: utf-8 -*-
"""Sweep the （エ） cell's own line and read what the 。 is billed there.

The paragraph lives in an auto-width one-column table (8458tw). The slice keeps
the row verbatim, pins the layout (tblLayout fixed) so the sweep is not eaten by
autofit, and sweeps the paragraph's RIGHT indent both ways: negative arms widen
the line until 「と。」 joins line 2, and that flip r* is the deficit Word sees.
Arm r=0 must reproduce Word's own break (line 2 ends 確認するこ).

    python _pb_e_budget.py
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_e_budget")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "tokyoshugyo*.docx"))
       if "~$" not in os.path.basename(p)][0]
MARK = "自己申告した労働時間を超えて"
R_TW = list(range(-300, 301, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    tbl_head = tbl_head.replace("<w:tblPr>", "<w:tblPr>", 1)
    if "tblLayout" not in tbl_head:
        tbl_head = tbl_head.replace("</w:tblPr>", '<w:tblLayout w:type="fixed"/></w:tblPr>', 1)
    rs = max(x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    # locate the paragraph inside the row and template its w:ind
    j = row.index(MARK)
    ps = max(row.rfind("<w:p ", 0, j), row.rfind("<w:p>", 0, j))
    pe = row.index("</w:p>", j) + len("</w:p>")
    para = row[ps:pe]
    blocks, index = [], []
    for r in R_TW:
        ind = re.search(r"<w:ind[^>]*/>", para).group(0)
        ind2 = ind[:-2] + ' w:right="%d"/>' % r
        p2 = para.replace(ind, ind2, 1)
        blocks.append(tbl_head + row[:ps] + p2 + row[pe:] + "</w:tbl>"
                      + '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
        index.append(r)
    doc = head + "".join(blocks) + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "eb.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = doc.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst, index


def main():
    import win32com.client as wc
    import fitz
    docx, index = build()
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
    doc = fitz.open(pdf)
    tails = []
    for page in doc:
        rows = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch:
                    rows.append((round(l["bbox"][1], 1),
                                 "".join(c["c"] for c in ch).strip()))
        for _, t in sorted(rows, key=lambda x: x[0]):
            if "確認するこ" in t or ("確認する" in t and t.endswith(("こ", "こと。", "と。"))):
                tails.append(t[-6:])
    if len(tails) != len(index):
        print("%d matched lines for %d arms" % (len(tails), len(index)))
    cur = None
    for r, t in zip(index, tails):
        if t != cur:
            print("   r=%7.2f  line ends ...%s" % (r / 20.0, t))
            cur = t


if __name__ == "__main__":
    main()
