# -*- coding: utf-8 -*-
"""Measure the width deficit of b35123's 「…不正ア|クセス」 line, 0.25pt-grade.

The bundle's last SSIM residual (p2 −0.0099) is one wrap: Word breaks after
不正ア where Oxi (with the pool) packs ク. Slice the row verbatim (two-column
table 1197+7870, pinned fixed), sweep the paragraph's right indent NEGATIVE and
positive, and read the flip where ク joins Word's line. -r* = how much wider the
line must be for Word to agree with Oxi.

    python _pb_b35_budget.py
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_b35_budget")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "b35123*.docx"))
       if "~$" not in os.path.basename(p)][0]
MARK = "不正アクセス行為"
R_TW = list(range(-300, 301, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    if "tblLayout" not in tbl_head:
        tbl_head = tbl_head.replace("</w:tblPr>", '<w:tblLayout w:type="fixed"/></w:tblPr>', 1)
    rs = max(x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    j = row.index(MARK)
    ps = max(row.rfind("<w:p ", 0, j), row.rfind("<w:p>", 0, j))
    pe = row.index("</w:p>", j) + len("</w:p>")
    para = row[ps:pe]
    ind = re.search(r"<w:ind[^>]*/>", para).group(0)
    blocks, index = [], []
    for r in R_TW:
        if 'w:right="' in ind:
            ind2 = re.sub(r'w:right="-?\d+"', 'w:right="%d"' % r, ind)
        else:
            ind2 = ind[:-2] + ' w:right="%d"/>' % r
        p2 = para.replace(ind, ind2, 1)
        blocks.append(tbl_head + row[:ps] + p2 + row[pe:] + "</w:tbl>"
                      + '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
        index.append(r)
    doc = head + "".join(blocks) + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "bb.docx")
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
    lines = []
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
        lines += [t for _, t in sorted(rows, key=lambda x: x[0]) if t]
    hits = [k for k, t in enumerate(lines) if "電気通信回線に接続している場合" in t]
    print("anchors %d for %d arms" % (len(hits), len(index)))
    cur = None
    for r, k in zip(index, hits):
        t = lines[k]
        key = t[-4:]
        if key != cur:
            print("   r=%7.2f  line ends ...%s" % (r / 20.0, t[-6:]))
            cur = key


if __name__ == "__main__":
    main()
