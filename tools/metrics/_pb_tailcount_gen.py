# -*- coding: utf-8 -*-
"""What width does the BREAK charge the marks that end a CONTINUATION line?

`_pb_hangline2_gen.py` established that a continuation line's line-final solo
mark gets NO hang, and back-computing the flip radii said Word charges the solo
mark a FULL EM (ignoring its narrower Ｐ明朝 natural) while a PAIR is charged
its NATURAL width each. This probe states those laws as predictions and reads
every arm's flip against all three candidates (full em / natural / half em).

Texts are 72 characters (two lines), the tail run varies, the right indent is
swept 0.25pt; flip = the largest r that still holds two lines. The charged
width of the tail run W = (850.4 - 2*r_flip) - 10.5*(#fullwidth chars).

    python _pb_tailcount_gen.py            # builds, exports, reports all four
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_tailcount")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
NCH = 72
MEASURE = 425.2
ARMS = [
    ("none", ""),
    ("solo_period", "。"),
    ("solo_comma", "、"),
    ("solo_close", "）"),
    ("pair_pc", "。）"),
    ("pair_cc", "」）"),
    ("pair_pp", "。。"),
    ("triple", "。」）"),
    # discriminators for the run law: a fourth mark; an ordinary char between the
    # fullwidth body and the pair (the pair still ends the line); and a mark that
    # does NOT end the line (control -- expect one em like any fullwidth char).
    ("quad", "。、」）"),
    ("ord_pair", "亜。）"),
    ("mark_ord", "。亜"),
]
R_TW = list(range(600, 1501, 5))


def build(face, compat15):
    os.makedirs(OUT, exist_ok=True)
    rpr = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
           % (face, face, face))
    paras, index = [], []
    for name, tail in ARMS:
        txt = "甲" + "亜" * (NCH - 1 - len(tail)) + tail
        for r in R_TW:
            index.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr>%s</w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (r, rpr, rpr, txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
           '<w:body>%s%s</w:body></w:document>' % ("".join(paras), sect))
    tag = ("pm" if "Ｐ" in face else "m") + ("15" if compat15 else "11")
    dst = os.path.join(OUT, "tc_%s.docx" % tag)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        elif compat15 and item.filename == "word/settings.xml":
            t = data.decode("utf-8").replace("<w:useAltKinsokuLineBreakRules/>", "")
            data = re.sub(r'(w:name="compatibilityMode"[^>]*w:val=")[0-9]+',
                          "\g<1>15", t).encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst, index


def export(docx):
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


def natural_advances(face):
    """Word's own natural advances for the marks, measured once per face."""
    import _pb_pmincho as PM
    # ★ISOLATED between 甲 -- the first calibration put the four marks adjacent
    # and the pair compression contaminated every "natural" (明朝 。 read 5.28).
    marks = "。、）」"
    probe = "甲" + "甲".join(marks) + "甲"
    adv = PM.advances(probe, face=face)
    return {m: adv[2 * i + 1] for i, m in enumerate(marks)}


def report(face, compat15, pdf, index):
    import fitz
    doc = fitz.open(pdf)
    counts, cur = [], 0
    for page in doc:
        rows = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch:
                    rows.append((round(l["bbox"][1], 1), ch[0]["c"]))
        for _, c0 in sorted(rows, key=lambda t: t[0]):
            if c0 == "甲":
                if cur:
                    counts.append(cur)
                cur = 1
            elif cur:
                cur += 1
    if cur:
        counts.append(cur)
    if len(counts) != len(index):
        print("   %d paragraphs for %d arms" % (len(counts), len(index)))
        return
    nat = natural_advances(face)
    by = {}
    for (name, r), n in zip(index, counts):
        by.setdefault(name, []).append((r / 20.0, n))
    base = None
    print("face=%s compat=%s" % (face, "15" if compat15 else "11+alt"))
    print("   %-12s flip r    charged W of tail   full-em   natural   half-em" % "arm")
    for name, tail in ARMS:
        rows = by[name]
        two = [r for r, n in rows if n <= 2]
        flip = max(two) if two else None
        if flip is None:
            print("   %-12s (no 2-line window)" % name)
            continue
        marks_only = [c for c in tail if c != "亜"]
        nfull = NCH - len(marks_only)
        w_tail = (2 * MEASURE - 2 * flip) - 10.5 * nfull
        if name == "none":
            base = flip
            print("   %-12s %6.2f    (baseline)" % (name, flip))
            continue
        full = 10.5 * len(marks_only)
        natw = sum(nat[c] for c in marks_only)
        half = sum(nat[c] / 2 for c in marks_only)
        tags = []
        for lbl, v in (("full-em", full), ("natural", natw), ("half", half)):
            if abs(w_tail - v) <= 0.6:
                tags.append(lbl)
        print("   %-12s %6.2f    %7.2f            %6.2f    %6.2f    %6.2f   -> %s"
              % (name, flip, w_tail, full, natw, half, ",".join(tags) or "?"))


def main():
    sys.path.insert(0, HERE)
    for face in ("ＭＳ Ｐ明朝", "ＭＳ 明朝"):
        for compat15 in (True, False):
            docx, index = build(face, compat15)
            report(face, compat15, export(docx), index)


if __name__ == "__main__":
    main()
