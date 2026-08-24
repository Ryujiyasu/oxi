# -*- coding: utf-8 -*-
"""Two 、 lend half an em; four 　 lend nothing; together they lend a whole em.

_pb_bodyyaku3 established that on the ⑧ line. But the same two marks with the
same four spaces lent only half an em in _pb_bodyyaku2, where the spaces sat in
the MIDDLE of the line and the paragraph carried no numbering marker. So sweep
the three things that differ: how many spaces, where they sit, and whether the
marker is there (the no-marker arms get the marker's measured indent instead, so
only the marker itself differs).

    python _pb_bodyyaku4_gen.py gen
    python _pb_bodyyaku4_gen.py pdf
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_bodyyaku_gen as B  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku4")
R_TW = list(range(0, 501, 5))
HEAD = "火災等非常災害の発生を発見したときは、直ちに臨機の措置をとり、"   # 31 ch, 2 marks
MARKER = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="37"/></w:numPr>'
IND_M = 'w:leftChars="0" w:left="884" w:hanging="425"'
IND_N = 'w:leftChars="0" w:left="1090"'


def text_of(marks, nsp, where):
    """36 characters: 31 head + 4 filler + 1 tail, nsp of the filler as 　."""
    h = HEAD if marks else HEAD.replace("、", "亜")
    if marks == 1:
        h = HEAD.replace("、", "亜", 1)
    mid = ["亜"] * 4
    if where == "end":
        for i in range(nsp):
            mid[4 - nsp + i] = "　"
        return h + "".join(mid) + "に"
    # "mid": the spaces sit six characters earlier, inside the head
    hl = list(h)
    for i in range(nsp):
        hl[20 + i] = "　"
    return "".join(hl) + "".join(mid) + "に"


ARMS = []
for j in range(0, 5):
    ARMS.append(("M2_e%d" % j, True, 2, j, "end"))
ARMS += [("M2_m4", True, 2, 4, "mid"),
         ("M0_e4", True, 0, 4, "end"),
         ("M1_e4", True, 1, 4, "end"),
         ("M0_e0", True, 0, 0, "end"),
         ("M1_e0", True, 1, 0, "end"),
         ("N2_e4", False, 2, 4, "end"),
         ("N2_e0", False, 2, 0, "end"),
         ("N0_e4", False, 0, 4, "end"),
         ("N0_e0", False, 0, 0, "end")]


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for name, marker, marks, nsp, where in ARMS:
        txt = text_of(marks, nsp, where)
        assert len(txt) == 36, (name, len(txt))
        for r in R_TW:
            index.append((name, r))
            runs = ('<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                    '<w:t xml:space="preserve">%s</w:t></w:r>' % txt)
            if marker:
                ppr = ('<w:pPr><w:pStyle w:val="a7"/>' + MARKER
                       + '<w:ind %s w:right="%d"/></w:pPr>' % (IND_M, r))
            else:
                ppr = ('<w:pPr><w:pStyle w:val="a7"/>'
                       '<w:ind %s w:right="%d"/></w:pPr>' % (IND_N, r))
            paras.append("<w:p>" + ppr + runs + "</w:p>")
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "bodyyaku4.docx")
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
        d = app.Documents.Open(os.path.join(OUT, "bodyyaku4.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "bodyyaku4.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def measure():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku4.pdf"))
    lines = []
    for page in doc:
        rows = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rows.append((round(ln["bbox"][1], 1), t))
        rows.sort()
        lines.extend(rows)
    paras, cur = [], None
    for y, t in lines:
        if "非常災害" in t:
            if cur:
                paras.append(cur)
            cur = [t]
        elif cur is not None:
            cur.append(t)
    if cur:
        paras.append(cur)
    print("arms %d paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch"); return
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p)))
    keep = {}
    print("\narm      keep<=r         split>=r")
    for name, _, _, _, _ in ARMS:
        rows = sorted(res[name])
        one = [r for r, k in rows if k == 1]
        spl = [r for r, k in rows if k > 1]
        keep[name] = max(one) if one else None
        print("%-8s %s   %s%s" % (name,
              ("%4d (%5.2fpt)" % (keep[name], keep[name] / 20.0)) if one else "   none     ",
              ("%4d" % min(spl)) if spl else "   -",
              "" if (one and spl and min(spl) == max(one) + 5) else "  (NON-MONOTONE)"))
    print("\ncredit against the same family's mark-free, space-free control")
    for name, marker, marks, nsp, where in ARMS:
        ctrl = "M0_e0" if marker else "N0_e0"
        if keep.get(name) is None or keep.get(ctrl) is None:
            continue
        d = (keep[name] - keep[ctrl]) / 20.0
        print("  %-8s marks=%d spaces=%d(%s)  %6.3f pt  %.4f em"
              % (name, marks, nsp, where, d, d / B.EM))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf(); measure()
    else:
        measure()
