# -*- coding: utf-8 -*-
"""Does a tab on LINE 1 kill the line-final hang on LINE 2?

The unified billing law (marks bill their compressed width, the line-final one
is free) explains every measured case but one: tokyoshugyo's （エ） paragraph
wraps 「と。」 although the free 。 would fit. That paragraph's line 1 carries a
numbering marker AND A TAB; S1213 measured (in cells) that a line with a tab
does not hang its final mark -- but there the tab was on the measured line.
These arms put the tab on line 1 only.

    python _pb_tabhang_gen.py
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_tabhang")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
FACE = os.environ.get("FACE") or "ＭＳ Ｐ明朝"
COMPAT15 = (os.environ.get("COMPAT15") or "0") == "1"
LEAD = 69
TRAIL = 6
# (name, first-run XML prefix, w:ind)
HANG_IND = '<w:ind w:left="884" w:hanging="425"/>'
ARMS = [
    ("plain", "", ""),
    ("tab", "<w:t>カ</w:t><w:tab/>", ""),
    ("hang", "", HANG_IND),
    ("hang_tab", "<w:t>カ</w:t><w:tab/>", HANG_IND),
]
R_TW = list(range(0, 2401, 5))


def build():
    os.makedirs(OUT, exist_ok=True)
    rpr = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
           % (FACE, FACE, FACE))
    paras, index = [], []
    body_txt = "甲" + "亜" * LEAD + "。" + "亜" * TRAIL
    for name, prefix, ind in ARMS:
        for r in R_TW:
            index.append((name, r))
            rind = ind.replace("/>", ' w:right="%d"/>' % r) if ind \
                else '<w:ind w:left="0" w:right="%d"/>' % r
            paras.append(
                '<w:p><w:pPr><w:pStyle w:val="a"/><w:jc w:val="both"/>'
                + rind +
                '<w:rPr>%s</w:rPr></w:pPr>'
                '<w:r><w:rPr>%s</w:rPr>%s'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                % (rpr, rpr, prefix, body_txt))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
           '<w:body>%s%s</w:body></w:document>' % ("".join(paras), sect))
    dst = os.path.join(OUT, "th.docx")
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
            if t.startswith("甲") and cur is not None and len(cur) == 1                     and cur[0].startswith("カ") and len(cur[0]) <= 2:
                # the marker and the text of line 1 come back as two rows
                # (PyMuPDF splits at the tab) -- merge them
                cur[0] = cur[0] + t
            elif t.startswith(("甲", "カ")):
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
    print("face=%s compat=%s" % (FACE, "15" if COMPAT15 else "11+alt"))
    print("   the question: does the line-2-final 。 stay free (regime ends only")
    print("   when the PRECEDING 亜 stop fitting) or is it billed?")
    seen = None
    for (name, r), lines in zip(index, paras):
        n1 = len(lines[0])
        t2 = lines[1] if len(lines) > 1 else ""
        key = (name, n1, len(t2), t2[-2:])
        if key != seen:
            print("   %-9s r=%6.2f L1=%2d L2 n=%2d ...%s (lines=%d)"
                  % (name, r / 20.0, n1, len(t2), t2[-4:], len(lines)))
            seen = key


if __name__ == "__main__":
    main()
