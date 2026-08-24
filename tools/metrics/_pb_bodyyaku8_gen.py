# -*- coding: utf-8 -*-
"""Does an OPENING bracket lend its half em from anywhere on the line?

S1198 gives a line 0.5em per opening bracket plus 0.5em if it also carries a
closing mark, capped at one em, whenever the LAST mark on the line is an opening
one. 3a4f9fbe p80's 第５９条 cell is such a line -- one 、 and one （ fourteen
characters from the end -- and Word wraps it where Oxi, holding a full em of
credit, packs one more character in.

The fullwidth-space release (S1208) turned out to depend on the mark's DISTANCE
from the line end, so ask the same of the bracket: sweep its position.

    python _pb_bodyyaku8_gen.py gen
    python _pb_bodyyaku8_gen.py pdf
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

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku8")
R_TW = list(range(0, 601, 5))
IND = V6.IND


def t(nmark, fills):
    return V6.text_of(nmark, fills)


ARMS = [
    ("c0_none", t(0, {})),
    ("c1_mark", t(1, {})),
    ("op_24", t(0, {24: "（"})),
    ("op_30", t(0, {30: "（"})),
    ("op_33", t(0, {33: "（"})),
    ("op_34", t(0, {34: "（"})),
    ("m_op24", t(1, {24: "（"})),
    ("m_op30", t(1, {30: "（"})),
    ("m_op33", t(1, {33: "（"})),
    ("m_op34", t(1, {34: "（"})),
    ("m_cl24", t(1, {24: "）"})),
    ("m_cl30", t(1, {30: "）"})),
    ("mark_24", t(0, {24: "、"})),
    ("mark_30", t(0, {30: "、"})),
    ("mark_33", t(0, {33: "、"})),
    ("mark_34", t(0, {34: "、"})),
]


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for name, txt in ARMS:
        for r in R_TW:
            index.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/>'
                '<w:ind w:leftChars="0" w:left="%d" w:right="%d"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (IND, r, txt))
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "bodyyaku8.docx")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8").write(
        "".join("%s\t%d\n" % a for a in index))
    print("built %s (%d paragraphs, %d arms)" % (dst, len(paras), len(ARMS)))


def to_pdf():
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.join(OUT, "bodyyaku8.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "bodyyaku8.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def rows():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku8.pdf"))
    lines = []
    for page in doc:
        rs = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                tt = "".join(c["c"] for c in cs).rstrip()
                if tt:
                    rs.append((round(ln["bbox"][1], 1), tt, cs))
        rs.sort()
        lines.extend(rs)
    paras, cur = [], None
    for y, tt, cs in lines:
        if tt.startswith("火亜"):
            if cur:
                paras.append(cur)
            cur = [(tt, cs)]
        elif cur is not None:
            cur.append((tt, cs))
    paras.append(cur)
    return index, paras


def measure():
    index, paras = rows()
    print("arms %d paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch")
        return
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p)))
    keep = {}
    for name, _ in ARMS:
        rr = sorted(res[name])
        one = [r for r, k in rr if k == 1]
        spl = [r for r, k in rr if k > 1]
        keep[name] = max(one) if one else None
        mono = "" if (one and spl and min(spl) == max(one) + 5) else "  (NON-MONOTONE)"
        print("%-9s keep<=%s split>=%s%s" % (name,
              ("%4d (%5.2fpt)" % (keep[name], keep[name] / 20.0)) if one else "  none    ",
              ("%4d" % min(spl)) if spl else "  -", mono))
    z = keep.get("c0_none")
    print("\ncredit against c0_none")
    for name, _ in ARMS:
        if keep.get(name) is None or z is None:
            continue
        d = (keep[name] - z) / 20.0
        print("  %-9s %6.3f pt  %.4f em" % (name, d, d / B.EM))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
