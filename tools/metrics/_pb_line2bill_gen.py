# -*- coding: utf-8 -*-
"""Read line 2's break position directly, and bill the mark run off it.

The 2-line flip probes convolved both lines' capacities with kinsoku drag; this
one pins line 1 with an explicit soft break (<w:br/>) and watches WHERE the
second line ends as the right indent grows. For each arm the report groups the
sweep into r-ranges by line 2's tail; the width of the range in which the mark
run IS the line end measures (midline bill) - (line-final bill) + one em.

    python _pb_line2bill_gen.py
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_line2bill")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
FACE = os.environ.get("FACE") or "ＭＳ 明朝"
COMPAT15 = (os.environ.get("COMPAT15") or "1") == "1"
K = 30                # 亜 before the run on line 2
TRAIL = 6             # 亜 after the run
ARMS = [
    ("none", ""),
    ("solo_period", "。"),
    ("pair_pc", "。）"),
    ("pair_cc", "」）"),
    ("pair_pp", "。。"),
    ("triple", "。」）"),
    ("quad", "。、」）"),
]
R_TW = list(range(0, 2401, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    rpr = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
           % (FACE, FACE, FACE))
    paras, index = [], []
    for name, run in ARMS:
        line2 = "亜" * K + run + "亜" * TRAIL
        for r in R_TW:
            index.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr>%s</w:rPr>'
                '<w:t>甲甲甲</w:t><w:br/>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (r, rpr, rpr, line2))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
           '<w:body>%s%s</w:body></w:document>' % ("".join(paras), sect))
    dst = os.path.join(OUT, "l2b.docx")
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
    # per paragraph: the LINE-2 text (the line right after the 甲甲甲 line)
    l2 = []
    grab = 0
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
                grab = 1
                continue
            if grab:
                l2.append(t)
                grab = 0
    if len(l2) != len(index):
        print("%d line-2 rows for %d arms" % (len(l2), len(index)))
        return
    print("face=%s compat=%s  K=%d trail=%d  (line 2 pinned by <w:br/>)"
          % (FACE, "15" if COMPAT15 else "11+alt", K, TRAIL))
    cur = None
    for (name, r), t in zip(index, l2):
        key = (name, len(t), t[-3:] if len(t) >= 3 else t)
        if key != cur:
            print("   %-12s r=%6.2f  n=%2d  ...%s" % (name, r / 20.0, len(t), t[-4:]))
            cur = key


if __name__ == "__main__":
    main()
