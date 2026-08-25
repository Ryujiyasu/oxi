# -*- coding: utf-8 -*-
"""Does the body's line-final hang survive onto a CONTINUATION line?

Every hang arm so far overflowed on LINE 1. tokyoshugyo's two refusals both
overflow on line 2 of their paragraphs, in ＭＳ Ｐ明朝, and Word refuses a
1.8pt overflow the line-1 arms accept up to the mark's whole advance.

Text = 甲 + 亜×70 + 。 (72 chars, two lines), right indent swept. The readout is
the LINE COUNT of the paragraph: 2 lines while the 。 still fits (hangs) on
line 2, 3 lines once it is pushed (dragging と-like kinsoku with it).

    python _pb_hangline2_gen.py
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_hangline2")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
FACE = os.environ.get("FACE") or "ＭＳ Ｐ明朝"
# COMPAT15=1 rewrites the base (tokyoshugyo, compat 11 + altKinsoku) to the
# modern engine -- kojin and nedocontract, which the first-line-only gate broke,
# are the docs to explain.
COMPAT15 = os.environ.get("COMPAT15") == "1"
RPR = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
       % (FACE, FACE, FACE))
NCH = 72
ARMS = {
    "tail_mark": "甲" + "亜" * (NCH - 2) + "。",
    "no_mark": "甲" + "亜" * (NCH - 1),
}
R_TW = list(range(700, 1301, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    paras, index = [], []
    for name, txt in ARMS.items():
        for r in R_TW:
            index.append((name, r))
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                '<w:ind w:left="0" w:right="%d"/>'
                '<w:rPr>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr>%s</w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (r, RPR, RPR, txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
           '<w:body>%s%s</w:body></w:document>' % ("".join(paras), sect))
    dst = os.path.join(OUT, "hangline2%s.docx" % ("_c15" if COMPAT15 else ""))
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
                          r"\g<1>15", t).encode("utf-8")
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
    # line count per paragraph: rows from one 甲 head to the next
    counts = []
    cur = 0
    for page in fitz.open(pdf):
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
        print("%d paragraphs for %d arms" % (len(counts), len(index)))
        return
    by = {}
    for (name, r), n in zip(index, counts):
        by.setdefault(name, []).append((r / 20.0, n))
    print("face=%s  NCH=%d  engine=%s"
          % (FACE, NCH, "compat 15" if COMPAT15 else "compat 11 + altKinsoku"))
    for name, rows in by.items():
        two = [r for r, n in rows if n <= 2]
        flip = max(two) if two else None
        print("   %-10s stays 2 lines up to r=%s   counts=%s"
              % (name, "%.2f" % flip if flip is not None else "-",
                 {n: (min(r for r, m in rows if m == n), max(r for r, m in rows if m == n))
                  for n in sorted(set(m for _, m in rows))}))


if __name__ == "__main__":
    main()
