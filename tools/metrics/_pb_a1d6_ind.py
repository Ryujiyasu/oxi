# -*- coding: utf-8 -*-
"""Why does a1d6e4ef's note read its indent in POINTS where a synthetic cell
reads the same attributes in CHARACTERS?

`_pb_indchars_gen.py` (20 arms) says a non-zero *Chars beats the twip beside it,
and one character is the grid pitch. a1d6e4ef's note carries
`leftChars=50 left=489 hangingChars=203 hanging=380` and Word renders 24.45 /
19.00 -- the twips. Slicing the row out reproduces that, so the cause is inside
the row; compat, the grid, styles.xml, settings.xml and the style name are all
ruled out.

One arm per FILE (the first attempt put every arm in one document, but each arm
is a whole table row and they cross pages, so the readings could not be matched
to their arms). Each arm removes exactly one thing from the row.

    python _pb_a1d6_ind.py
"""
import glob
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_a1d6_ind")
SRC = [p for p in glob.glob(os.path.join(REPO, "tools", "golden-test", "documents",
                                         "docx", "a1d6*.docx"))
       if not os.path.basename(p).startswith("~$")][0]
MARK = "提供依頼申出"
CELLMAR_PT = 12 / 20.0          # this table's tblCellMar left


def parts():
    x = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    i = x.index(MARK)
    tbl_start = x.rindex("<w:tbl>", 0, i)
    tbl_head = x[tbl_start:x.index("</w:tblGrid>", tbl_start) + len("</w:tblGrid>")]
    rs = max((x.rfind("<w:tr ", 0, i), x.rfind("<w:tr>", 0, i)))
    row = x[rs:x.index("</w:tr>", i) + len("</w:tr>")]
    head = x[:x.index("<w:body>") + len("<w:body>")]
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", x, re.S).group(0)
    sect = re.sub(r"<w:(headerReference|footerReference)[^>]*/>", "", sect)
    return head, tbl_head, row, sect


def note_para(row):
    i = row.index(MARK)
    ps = max((row.rfind("<w:p ", 0, i), row.rfind("<w:p>", 0, i)))
    pe = row.index("</w:p>", i) + len("</w:p>")
    return ps, pe


def sub_para(row, fn):
    ps, pe = note_para(row)
    return row[:ps] + fn(row[ps:pe]) + row[pe:]


ARMS = [
    ("base", lambda h, r: (h, r)),
    ("no_vmerge", lambda h, r: (h, re.sub(r"<w:vMerge[^>]*/>", "", r))),
    ("no_tcw", lambda h, r: (h, re.sub(r'<w:tcW w:w="\d+" w:type="\w+"/>', "", r))),
    ("style_a", lambda h, r: (h, sub_para(r, lambda p: p.replace(
        '<w:pStyle w:val="ac"/>', '<w:pStyle w:val="a"/>')))),
    ("no_pstyle", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:pStyle[^>]*/>", "", p)))),
    ("no_wordwrap", lambda h, r: (h, sub_para(r, lambda p: p.replace(
        "<w:wordWrap/>", "")))),
    ("no_spacing", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:spacing[^>]*/>", "", p, count=1)))),
    ("no_ppr_rpr", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:rPr>.*?</w:rPr></w:pPr>", "</w:pPr>", p, count=1, flags=re.S)))),
    ("ind_probe", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:ind[^>]*/>", '<w:ind w:leftChars="100" w:left="81"/>', p, count=1)))),
    ("no_grid", lambda h, r: (h, r)),        # handled in build(): drops the docGrid
    # ind_probe reads as CHARS in this very cell while the note's own ind reads as
    # TWIPS -- so vary the ind alone and find which member flips it.
    ("ind_lc50_l489", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:ind[^>]*/>", '<w:ind w:leftChars="50" w:left="489"/>', p, count=1)))),
    ("ind_lc50_l489_h", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:ind[^>]*/>",
        '<w:ind w:leftChars="50" w:left="489" w:hanging="380"/>', p, count=1)))),
    ("ind_lc50_l489_hc", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:ind[^>]*/>",
        '<w:ind w:leftChars="50" w:left="489" w:hangingChars="203"/>', p, count=1)))),
    ("ind_lc100_l489", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:ind[^>]*/>", '<w:ind w:leftChars="100" w:left="489"/>', p, count=1)))),
    ("ind_lc300_l489", lambda h, r: (h, sub_para(r, lambda p: re.sub(
        r"<w:ind[^>]*/>", '<w:ind w:leftChars="300" w:left="489"/>', p, count=1)))),
]


def build(name, fn):
    os.makedirs(OUT, exist_ok=True)
    head, tbl_head, row, sect = parts()
    tbl_head, row = fn(tbl_head, row)
    if name == "no_grid":
        sect = re.sub(r"<w:docGrid[^>]*/>", '<w:docGrid w:type="lines" w:linePitch="292"/>',
                      sect)
    doc = head + tbl_head + row + "</w:tbl>" + sect + "</w:body></w:document>"
    dst = os.path.join(OUT, "arm_%s.docx" % name)
    shutil.copyfile(SRC, dst)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = doc.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return dst


def export(paths):
    import win32com.client as wc
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    out = []
    try:
        for docx in paths:
            pdf = os.path.splitext(docx)[0] + ".pdf"
            d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                                  ExportFormat=17, OpenAfterExport=False)
            d.Close(False)
            out.append(pdf)
    finally:
        app.Quit()
    return out


def read(pdf):
    """(cell inner left, first-line x, continuation x) for the ※1 note."""
    import fitz
    doc = fitz.open(pdf)
    rules, first, cont = [], None, None
    for page in doc:
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > 3:
                    rules.append(round((it[1].x + it[2].x) / 2, 2))
                elif it[0] == "re" and it[1].width < 0.9 and it[1].height > 3:
                    rules.append(round(it[1].x0, 2))
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if not ch:
                    continue
                t = "".join(c["c"] for c in ch)
                if first is None and t.startswith("※１"):
                    first = ch[0]["origin"][0]
                elif first is not None and cont is None and t.startswith("者及び"):
                    cont = ch[0]["origin"][0]
    inner = (min(rules) if rules else 0.0) + CELLMAR_PT
    return inner, first, cont


def main():
    paths = [build(n, f) for n, f in ARMS]
    print("built %d arms" % len(paths))
    pdfs = export(paths)
    print("   arm          inner    first     cont   |  first ind   cont ind   reads as")
    for (name, _), pdf in zip(ARMS, pdfs):
        inner, first, cont = read(pdf)
        if first is None or cont is None:
            print("   %-12s %7.2f   (note not found)" % (name, inner))
            continue
        fi, ci = first - inner, cont - inner
        # twips: left 24.45, hanging 19.00 ; chars: 0.5 and 2.03 of the grid pitch
        pitch = 10.8547
        reads = ("TWIP" if abs(ci - 24.45) < 1.2 else
                 "CHARS" if abs(ci - 0.5 * pitch) < 1.2 else "?")
        print("   %-12s %7.2f  %7.2f  %7.2f  |  %+7.2f   %+7.2f   %s"
              % (name, inner, first, cont, fi, ci, reads))


if __name__ == "__main__":
    main()
