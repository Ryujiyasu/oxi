# -*- coding: utf-8 -*-
"""The same release rule, asked of a TABLE CELL line.

The line that started this (tokyoshugyo p76's marked item) is a cell paragraph, and
Oxi's 約物 pool lives on the cell path, so the body sweep has to be repeated inside
a cell before the rule can be coded there. Same arms, same 0.25pt right-indent
sweep; every paragraph is alone in a one-cell table of fixed width.

    python _pb_cellyaku_gen.py gen
    python _pb_cellyaku_gen.py pdf
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_bodyyaku_gen as B  # noqa: E402
import _pb_bodyyaku6_gen as V6  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_cellyaku")
R_TW = list(range(0, 601, 5))
TBLW = 7876          # twips: 393.8pt - 2*5.4pt cell margin = 383.0pt of measure


def build():
    os.makedirs(OUT, exist_ok=True)
    rows, index = [], []
    for name, txt in V6.ARMS:
        for r in R_TW:
            index.append((name, r))
            rows.append(
                '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
                '<w:tblLayout w:type="fixed"/></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:pStyle w:val="a"/>'
                '<w:ind w:leftChars="0" w:left="0" w:right="%d"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                '</w:tc></w:tr></w:tbl><w:p/>' % (TBLW, TBLW, TBLW, r, txt))
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(rows), sect))
    dst = os.path.join(OUT, "cellyaku.docx")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8").write(
        "".join("%s\t%d\n" % a for a in index))
    print("built %s (%d cells, %d arms)" % (dst, len(rows), len(V6.ARMS)))


def to_pdf():
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.join(OUT, "cellyaku.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "cellyaku.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def measure():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "cellyaku.pdf"))
    lines = []
    for page in doc:
        rs = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rs.append((round(ln["bbox"][1], 1), t))
        rs.sort(); lines.extend(rs)
    paras, cur = [], None
    for y, t in lines:
        if t.startswith("火亜"):
            if cur:
                paras.append(cur)
            cur = [t]
        elif cur is not None:
            cur.append(t)
    paras.append(cur)
    print("arms %d paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch"); return
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p)))
    keep = {}
    for name, _ in V6.ARMS:
        rr = sorted(res[name])
        one = [r for r, k in rr if k == 1]
        spl = [r for r, k in rr if k > 1]
        keep[name] = max(one) if one else None
        mono = "" if (one and spl and min(spl) == max(one) + 5) else "  (NON-MONOTONE)"
        print("%-13s keep<=%s split>=%s%s" % (name,
              ("%4d (%5.2fpt)" % (keep[name], keep[name] / 20.0)) if one else "  none    ",
              ("%4d" % min(spl)) if spl else "  -", mono))
    z = keep.get("n0_bare")
    print("\ncredit against n0_bare")
    for name, _ in V6.ARMS:
        if keep.get(name) is None or z is None:
            continue
        d = (keep[name] - z) / 20.0
        print("  %-13s %6.3f pt  %.4f em" % (name, d, d / B.EM))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf(); measure()
    else:
        measure()
