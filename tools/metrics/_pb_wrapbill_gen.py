# -*- coding: utf-8 -*-
"""The same regime reading, on a WRAPPED continuation line.

`_pb_line2bill_gen.py` showed a <w:br/> line bills like a first line (all marks
half an em, the line-final one free). The real cases (kojin, tokyoshugyo) are
WRAPPED lines, and the 2-line flip probes suggested those bill differently --
but their algebra assumed both lines full. Here line 1 is left to wrap
naturally and its char count is READ off the PDF, so line 2's inequality
stands alone:

    width(line-2 content) <= 425.2 - r,  content known per regime.

    python _pb_wrapbill_gen.py
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_wrapbill")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
FACE = os.environ.get("FACE") or "ＭＳ 明朝"
COMPAT15 = (os.environ.get("COMPAT15") or "1") == "1"
ARMS = [
    ("none", ""),
    ("solo_period", "。"),
    ("pair_pc", "。）"),
]
# text: 甲 + 69x亜 + run + 6x亜 (3 lines at moderate r)
LEAD = 69
TRAIL = 6
R_TW = list(range(0, 2401, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    rpr = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
           % (FACE, FACE, FACE))
    paras, index = [], []
    for name, run in ARMS:
        txt = "甲" + "亜" * LEAD + run + "亜" * TRAIL
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
    tag = ("pm" if "Ｐ" in FACE else "m") + ("15" if COMPAT15 else "11")
    dst = os.path.join(OUT, "wb_%s.docx" % tag)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        elif COMPAT15 and item.filename == "word/settings.xml":
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


def main():
    import fitz
    docx, index = build()
    doc = fitz.open(export(docx))
    paras, cur = [], None
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
            if not t:
                continue
            if t.startswith("甲"):
                if cur is not None:
                    paras.append(cur)
                cur = [t]
            elif cur is not None:
                cur.append(t)
    if cur is not None:
        paras.append(cur)
    if len(paras) != len(index):
        print("%d paragraphs for %d arms" % (len(paras), len(index)))
        return
    print("face=%s compat=%s  lead=%d trail=%d (wrapped; line-1 count read)"
          % (FACE, "15" if COMPAT15 else "11+alt", LEAD, TRAIL))
    seen = None
    for (name, r), lines in zip(index, paras):
        n1 = len(lines[0])
        t2 = lines[1] if len(lines) > 1 else ""
        key = (name, n1, len(t2), t2[-3:])
        if key != seen:
            print("   %-12s r=%6.2f  L1=%2d  L2 n=%2d ...%s  (lines=%d)"
                  % (name, r / 20.0, n1, len(t2), t2[-4:], len(lines)))
            seen = key
    # dump for offline algebra
    import io
    w = io.open(os.path.join(OUT, "regimes_%s.txt" % (("pm" if "Ｐ" in FACE else "m") + ("15" if COMPAT15 else "11"))), "w", encoding="utf-8")
    for (name, r), lines in zip(index, paras):
        w.write("%s\t%d\t%s\n" % (name, r, "|".join(lines)))
    w.close()


if __name__ == "__main__":
    main()
