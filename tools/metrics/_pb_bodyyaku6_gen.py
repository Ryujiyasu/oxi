# -*- coding: utf-8 -*-
"""The trailing 　 rule, measured without Word's auto-numbering in the way.

_pb_bodyyaku5 used a numbered paragraph to make the measure narrow enough to
wrap, and its later arms silently died: with two thousand numbered paragraphs the
marker grows to four digits, no longer fits the hanging indent and pushes the text
onto its own line. Same geometry here, but from a plain left indent.

Arms answer three questions:
  * how many 約物 the trailing 　 releases (n = 0..4)
  * whether the releasing character must be 　, or any compressible, or a
    halfwidth space
  * how close to the line's last character it has to sit (34, 33, 31, 28; and as
    the last character itself)
  * whether it matters that the character being squeezed in is the PARAGRAPH's
    last (an arm with text continuing after it)

    python _pb_bodyyaku6_gen.py gen
    python _pb_bodyyaku6_gen.py pdf
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_bodyyaku_gen as B  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku6")
R_TW = list(range(0, 601, 5))
NCH = 36
IND = 844                      # twips: measure 425.2 - 42.2 = 383.0, natural 378
MARK_POS = [6, 12, 18, 24]


def text_of(nmark, fills, tail=""):
    t = ["火"] + ["亜"] * (NCH - 2) + ["に"]
    for i in range(nmark):
        t[MARK_POS[i]] = "、"
    for idx, ch in fills.items():
        t[idx] = ch
    return "".join(t) + tail


ARMS = []
for n in range(0, 5):
    ARMS.append(("n%d_bare" % n, text_of(n, {})))
    ARMS.append(("n%d_sp34" % n, text_of(n, {34: "　"})))
ARMS += [
    ("n2_sp33", text_of(2, {33: "　"})),
    ("n2_sp31", text_of(2, {31: "　"})),
    ("n2_sp28", text_of(2, {28: "　"})),
    ("n2_sp3334", text_of(2, {33: "　", 34: "　"})),
    ("n2_yk34", text_of(2, {34: "、"})),
    ("n2_pd34", text_of(2, {34: "。"})),
    ("n2_cl34", text_of(2, {34: "）"})),
    ("n2_op34", text_of(2, {34: "（"})),
    ("n2_asc34", text_of(2, {34: " "})),
    ("n2_sp34_more", text_of(2, {34: "　"}, tail="亜亜亜亜亜")),
    ("n2_bare_more", text_of(2, {}, tail="亜亜亜亜亜")),
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
    dst = os.path.join(OUT, "bodyyaku6.docx")
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
        d = app.Documents.Open(os.path.join(OUT, "bodyyaku6.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "bodyyaku6.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def rows():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku6.pdf"))
    lines = []
    for page in doc:
        rs = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rs.append((round(ln["bbox"][1], 1), t, cs))
        rs.sort(); lines.extend(rs)
    paras, cur = [], None
    for y, t, cs in lines:
        if t.startswith("火亜"):
            if cur:
                paras.append(cur)
            cur = [(t, cs)]
        elif cur is not None:
            cur.append((t, cs))
    paras.append(cur)
    return index, paras


def measure():
    index, paras = rows()
    print("arms %d paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch"); return
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
        print("%-13s keep<=%s split>=%s%s" % (name,
              ("%4d (%5.2fpt)" % (keep[name], keep[name] / 20.0)) if one else "  none    ",
              ("%4d" % min(spl)) if spl else "  -", mono))
    z = keep.get("n0_bare")
    print("\ncredit against n0_bare")
    for name, _ in ARMS:
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
