# -*- coding: utf-8 -*-
"""How much can a BODY line borrow from its 約物, and does the count matter?

The `s1174_yakucomp` pool was derived on TABLE CELL lines and caps a line's
type-A credit at half an em however many 、 it carries. tokyoshugyo p76's ⑧ item
refutes that on a BODY line: Word compresses BOTH 、 by ~3.9pt (0.371em each,
7.8pt together) so that the paragraph's last character 「に」 stays on line 1,
and Oxi with OXI_YAKUCOMP=1 breaks before it because 0.5em is not enough.

So sweep the body side directly. Every arm is one paragraph of `NCH` fullwidth
characters carrying `n` marks of one class, and the paragraph's RIGHT INDENT is
swept in 0.25pt steps. The natural line is NCH*em wide, so at right indent r the
last character overflows by r - (content - NCH*em); the largest r that still
leaves the paragraph on ONE line is the credit the line was granted.

A paragraph starts with 甲 so paragraph starts are identifiable in the PDF, and
the character AFTER every mark is a normal one so no mark ever sits at the line
end (the hang rule is a different measurement).

    python _pb_bodyyaku_gen.py gen     # build the probe docx
    python _pb_bodyyaku_gen.py pdf     # export through Word + measure
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_bodyyaku")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")

EM = 10.5
NCH = 40                      # 40 * 10.5 = 420.0pt against a 425.2pt measure
CLASSES = {"C1": "、", "C2": "。", "C3": "）", "C4": "（"}
COUNTS = [0, 1, 2, 3, 4]
R_TW = list(range(0, 301, 5))  # 0..15pt in 0.25pt steps

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')


def text_for(mark, n):
    """甲 + NCH-1 body characters holding `n` marks, never at the tail."""
    body = ["亜"] * (NCH - 1)
    if n:
        # spread the marks over the middle, each followed by a normal char
        step = (NCH - 6) // n
        for i in range(n):
            body[3 + i * step] = mark
    return "甲" + "".join(body)


def build():
    os.makedirs(OUT, exist_ok=True)
    paras = []
    arms = []
    for cname, mark in CLASSES.items():
        for n in COUNTS:
            if n == 0 and cname != "C1":
                continue          # the n=0 control is class-free
            txt = text_for(mark, n)
            for r in R_TW:
                arms.append((cname, n, r))
                paras.append(
                    '<w:p><w:pPr><w:pStyle w:val="a"/>'
                    '<w:ind w:left="0" w:right="%d"/>'
                    '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
                    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                    '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (r, txt))
    src = zipfile.ZipFile(SRC)
    doc = src.read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r'<w:footerReference[^>]*/>', "", sect)
    body = "<w:body>" + "".join(paras) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "bodyyaku.docx")
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8") as fh:
        for a in arms:
            fh.write("%s\t%d\t%d\n" % a)
    print("built %s  (%d paragraphs)" % (dst, len(paras)))


def to_pdf():
    import win32com.client as wc
    docx = os.path.join(OUT, "bodyyaku.docx")
    pdf = os.path.join(OUT, "bodyyaku.pdf")
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


def measure():
    import fitz
    pdf = os.path.join(OUT, "bodyyaku.pdf")
    arms = [l.split("\t") for l in
            open(os.path.join(OUT, "arms.txt"), encoding="utf-8").read().splitlines()]
    doc = fitz.open(pdf)
    lines = []
    for page in doc:
        raw = page.get_text("rawdict")
        rows = []
        for blk in raw.get("blocks", []):
            for ln in blk.get("lines", []):
                cs = [c for sp in ln["spans"] for c in sp.get("chars", [])]
                t = "".join(c["c"] for c in cs).rstrip()
                if t:
                    rows.append((round(ln["bbox"][1], 1), t, cs))
        rows.sort()
        lines.extend(rows)
    # group into paragraphs: a line starting with 甲 opens one
    paras, cur = [], None
    for y, t, cs in lines:
        if t.startswith("甲"):
            if cur:
                paras.append(cur)
            cur = [(y, t, cs)]
        elif cur is not None:
            cur.append((y, t, cs))
    if cur:
        paras.append(cur)
    print("arms %d, paragraphs found %d" % (len(arms), len(paras)))
    if len(paras) != len(arms):
        print("!! count mismatch -- the 甲 grouping missed something")
    res = {}
    for arm, p in zip(arms, paras):
        cname, n, r = arm[0], int(arm[1]), int(arm[2])
        res.setdefault((cname, n), []).append((r, len(p), p))
    content = 425.2
    print("\nclass count | last r (twips) that stays ONE line | credit pt | em")
    for key in sorted(res):
        rows = sorted(res[key])
        one = [r for r, k, _ in rows if k == 1]
        if not one:
            print("%s n=%d  never one line" % (key[0], key[1]))
            continue
        rmax = max(one)
        # first r that split, for the bracket
        split = [r for r, k, _ in rows if k > 1]
        rmin_split = min(split) if split else None
        base = content - NCH * EM          # 5.2pt of natural slack
        credit = rmax / 20.0 - base
        print("%s n=%d  keep<=%4d (%.2fpt) split>=%s | credit %6.3f pt  %.4f em%s"
              % (key[0], key[1], rmax, rmax / 20.0,
                 ("%d" % rmin_split) if rmin_split is not None else "-",
                 credit, credit / EM,
                 "" if rmin_split == rmax + 5 else "   (NON-MONOTONE)"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "gen":
        build()
    elif cmd == "pdf":
        to_pdf()
        measure()
    else:
        measure()
