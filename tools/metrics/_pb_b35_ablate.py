# -*- coding: utf-8 -*-
"""Which feature of b35123's real cell carries the ~4.56pt break credit?

The real slice holds 39 characters up to r=3.50 (credit T in (4.51,4.66) --
numerically 9.837 - 5.28, "the mid-line comma compresses to its UN-gridded half
width"), while the synthetic cs=-2714 cell probe measured credit 0.000. One of
the conditions between them switches it. Each variant repeats the full table
with the target paragraph's right indent swept; the readout is the flip r*.

    python _pb_b35_ablate.py
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_b35_ablate")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "b35123*.docx"))
       if "~$" not in os.path.basename(p)][0]
MARK = "不正アクセス行為"
R_TW = list(range(0, 161, 5))


def variants(tbl, ps, pe):
    para = tbl[ps:pe]
    out = [("verbatim", tbl, para)]
    p2 = re.sub(r'<w:ind[^>]*/>', '<w:ind w:left="197"/>', para, count=1)
    out.append(("no_hanging", tbl[:ps] + p2 + tbl[pe:], p2))
    p3 = para.replace("□", "亜").replace("　", "亜", 1)
    out.append(("plain_head", tbl[:ps] + p3 + tbl[pe:], p3))
    p4 = para.replace("場合、不正", "場合亜不正")
    out.append(("no_comma", tbl[:ps] + p4 + tbl[pe:], p4))
    t5 = tbl.replace("<w:tblPr>", "<w:tblPr>", 1)
    t5 = t5.replace("</w:tblPr>",
                    '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
                    '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
                    '</w:tblCellMar></w:tblPr>', 1)
    out.append(("fixed_mar0", t5, para))
    # split the discriminator: fixed alone, mar0 alone
    t6 = tbl.replace("</w:tblPr>", '<w:tblLayout w:type="fixed"/></w:tblPr>', 1)
    out.append(("fixed_only", t6, para))
    t7 = tbl.replace("</w:tblPr>",
                     '<w:tblCellMar><w:left w:w="0" w:type="dxa"/>'
                     '<w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>', 1)
    out.append(("mar0_only", t7, para))
    return out


def build():
    os.makedirs(OUT, exist_ok=True)
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    ts = x.index("<w:tbl>")
    te = x.index("</w:tbl>", ts) + len("</w:tbl>")
    tbl = x[ts:te]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    i = tbl.index(MARK)
    ps = max(tbl.rfind("<w:p ", 0, i), tbl.rfind("<w:p>", 0, i))
    pe = tbl.index("</w:p>", i) + len("</w:p>")
    blocks, index = [], []
    for name, t2, para in variants(tbl, ps, pe):
        j = t2.index(para)
        ind = re.search(r"<w:ind[^>]*/>", para).group(0)
        for r in R_TW:
            ind2 = ind[:-2] + ' w:right="%d"/>' % r
            blocks.append(t2[:j] + para.replace(ind, ind2, 1) + t2[j + len(para):]
                          + '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>')
            index.append((name, r))
    doc = head + "".join(blocks) + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "ab.docx")
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
    hits = []
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
            t = "".join(c["c"] for c in ch).strip()
            if "電気通信回線に接続している場合" in t:
                hits.append((len(t), round(ch[0]["origin"][0], 2)))
    if len(hits) != len(index):
        print("%d hits for %d arms" % (len(hits), len(index)))
        return
    by = {}
    for (name, r), (n, x0) in zip(index, hits):
        by.setdefault(name, []).append((r / 20.0, n, x0))
    print("   variant      flip r* (last r holding the LONG line)   x0 range")
    for name, rows in by.items():
        top = max(n for _, n, _ in rows)
        hold = [r for r, n, _ in rows if n >= top]
        xs = sorted(set(x for _, _, x in rows))
        print("   %-11s top n=%d holds to r=%s   x0=%s"
              % (name, top, "%.2f" % max(hold) if hold else "-",
                 xs if len(xs) <= 3 else xs[:3] + ["..."]))


if __name__ == "__main__":
    main()
