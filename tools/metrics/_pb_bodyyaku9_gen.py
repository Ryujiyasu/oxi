# -*- coding: utf-8 -*-
"""Falsification test of one rule for the whole 約物 pool.

Every reading taken in _pb_bodyyaku4 / 6 / 8 (35 arms, body and cell alike) fits a
single statement:

    a line's 約物 credit is HALF AN EM, flat, however many marks it carries and
    whatever their class -- EXCEPT when the character immediately before the one
    being squeezed in is a FULLWIDTH SPACE or an OPENING BRACKET, and then every
    mark on the line lends its own half em, up to one and a half.
    (a fullwidth space is not itself a mark; an opening bracket is.)

That subsumes S1198's "last mark type A -> 0.5em / type B -> min(1em, ...)", whose
arms put the marks every third character and so could not tell "last" from
"adjacent to the squeeze".

This module states the prediction for each arm BEFORE measuring and prints both.

    python _pb_bodyyaku9_gen.py gen
    python _pb_bodyyaku9_gen.py pdf
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

OUT = os.path.join(B.REPO, "pipeline_data", "_pb_bodyyaku9")
R_TW = list(range(0, 601, 5))
NCH = 36
A_POS = [6, 12, 18]          # closing marks
B_POS = [9, 15, 21]          # opening brackets
RELEASE = {"sp": "　", "op": "（", "pd": "。", "no": "亜"}


def text_of(na, nb, rel):
    t = ["火"] + ["亜"] * (NCH - 2) + ["に"]
    for i in range(na):
        t[A_POS[i]] = "、"
    for i in range(nb):
        t[B_POS[i]] = "（"
    t[34] = RELEASE[rel]
    return "".join(t)


def predict(na, nb, rel):
    """Marks on the line, and the release test, in ems."""
    marks = na + nb + (1 if rel == "op" else 0) + (1 if rel == "pd" else 0)
    if rel in ("sp", "op"):
        return min(0.5 * marks, 1.5)
    return 0.5 if marks else 0.0


ARMS = []
for na, nb in ((0, 0), (1, 0), (2, 0), (3, 0), (0, 1), (1, 1), (2, 1),
               (1, 2), (2, 2), (3, 1)):
    for rel in ("no", "sp", "op", "pd"):
        ARMS.append(("a%db%d_%s" % (na, nb, rel), text_of(na, nb, rel),
                     predict(na, nb, rel)))


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for name, txt, _ in ARMS:
        assert len(txt) == NCH, (name, len(txt))
        for r in R_TW:
            index.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/>'
                '<w:ind w:leftChars="0" w:left="%d" w:right="%d"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (V6.IND, r, txt))
    src = zipfile.ZipFile(B.SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s><w:body>%s%s</w:body></w:document>'
           % (B.W_NS, "".join(paras), sect))
    dst = os.path.join(OUT, "bodyyaku9.docx")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8").write(
        "".join("%s\t%d\n" % (a, b) for (a, b) in index))
    print("built %s (%d paragraphs, %d arms)" % (dst, len(paras), len(ARMS)))


def to_pdf():
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.join(OUT, "bodyyaku9.docx"), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.join(OUT, "bodyyaku9.pdf"),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def measure():
    import fitz
    index = [l.split("\t") for l in open(os.path.join(OUT, "arms.txt"),
             encoding="utf-8").read().splitlines()]
    doc = fitz.open(os.path.join(OUT, "bodyyaku9.pdf"))
    lines = []
    for page in doc:
        rs = []
        for blk in page.get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                tt = "".join(c["c"] for c in cs).rstrip()
                if tt:
                    rs.append((round(ln["bbox"][1], 1), tt))
        rs.sort()
        lines.extend(rs)
    paras, cur = [], None
    for y, tt in lines:
        if tt.startswith("火亜"):
            if cur:
                paras.append(cur)
            cur = [tt]
        elif cur is not None:
            cur.append(tt)
    paras.append(cur)
    print("arms %d paragraphs %d" % (len(index), len(paras)))
    if len(index) != len(paras):
        print("!! grouping mismatch")
        return
    res = {}
    for (name, r), p in zip(index, paras):
        res.setdefault(name, []).append((int(r), len(p)))
    keep = {}
    for name, _, _ in ARMS:
        rr = sorted(res[name])
        one = [r for r, k in rr if k == 1]
        spl = [r for r, k in rr if k > 1]
        keep[name] = max(one) if one else None
        if not one or (spl and min(spl) != max(one) + 5):
            print("%-12s NON-MONOTONE or never one line" % name)
    z = keep.get("a0b0_no")
    bad = 0
    print("\narm          predicted   measured    verdict")
    for name, _, pred in ARMS:
        if keep.get(name) is None or z is None:
            print("%-12s %6.3f em   --" % (name, pred))
            bad += 1
            continue
        got = (keep[name] - z) / 20.0 / B.EM
        ok = abs(got - pred) < 0.03
        bad += 0 if ok else 1
        print("%-12s %6.3f em   %6.3f em   %s"
              % (name, pred, got, "ok" if ok else "MISMATCH"))
    print("\n%d of %d arms match the single rule" % (len(ARMS) - bad, len(ARMS)))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
