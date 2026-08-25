# -*- coding: utf-8 -*-
"""Which indent does Word obey when a paragraph carries BOTH twips and *Chars?

29dc6e8943fe's cell is the case. Its ウ paragraph carries `<w:ind w:left="81"/>`
(4.05pt) and Word puts it at the cell's inner edge -- as if `left` were zero --
while Oxi indents it 4.05pt, and every paragraph in that cell ends up 4.45pt to
the right of Word's, which costs the ③ line its last character.

Arms (all in one fixed-width cell, one paragraph each, so the first glyph's
origin can be read against the cell's own left rule):

    none        no <w:ind> at all                     -- the zero reference
    left        left=81                               -- twips only
    left_fl     left=81 firstLine=199                 -- twips only, both
    l_flc_fl    left=81 firstLineChars=100 firstLine=199   -- 29dc6e's shape
    lc_left     leftChars=100 left=81                 -- *Chars and twips fight
    flc_only    firstLineChars=100                    -- *Chars alone
    lc0_left    leftChars=0 left=81                   -- an explicit zero *Chars

Declared before measuring: if Word simply prefers *Chars per attribute, then
`l_flc_fl` sits at 4.05 + one character (10.50 at this size) = 14.55pt, and
`lc_left` at 10.50. If a *Chars attribute anywhere makes Word read the whole
w:ind in character units (absent = zero), `l_flc_fl` sits at one character alone
and `lc_left` at one character too. 29dc6e's render fits NEITHER cleanly -- it
reads as `firstLine` (twips) with `left` dropped -- so the arms decide it.

    python _pb_indchars_gen.py            # build, export, report
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_indchars")
SRC = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                   "tokyoshugyo_000599795.docx")
CELL_TW = 6000                     # 300pt
MAR_TW = 108                       # Word's default cell margin, 5.4pt
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
ARMS = [
    ("none", ""),
    ("left", '<w:ind w:left="81"/>'),
    ("left_fl", '<w:ind w:left="81" w:firstLine="199"/>'),
    ("l_flc_fl", '<w:ind w:left="81" w:firstLineChars="100" w:firstLine="199"/>'),
    ("lc_left", '<w:ind w:leftChars="100" w:left="81"/>'),
    ("flc_only", '<w:ind w:firstLineChars="100"/>'),
    ("lc0_left", '<w:ind w:leftChars="0" w:left="81"/>'),
    # ★a1d6e4ef carries leftChars="50" left="489" and Word renders the TWIP
    # (24.45pt), while the arm above (leftChars=100 left=81) renders the CHARS.
    # The two differ in which value is larger -- so ask whether Word takes the
    # bigger of the pair.
    ("lc_bigleft", '<w:ind w:leftChars="100" w:left="600"/>'),
    ("flc_bigfl", '<w:ind w:firstLineChars="100" w:firstLine="600"/>'),
    ("lc50_left489", '<w:ind w:leftChars="50" w:left="489"/>'),
    # a1d6e4ef's ind VERBATIM: the same leftChars/left pair as the arm above, but
    # with a hanging indent alongside. Its render says the TWIPS won there.
    ("a1d6_full", '<w:ind w:leftChars="50" w:left="489" w:rightChars="50"'
                  ' w:right="109" w:hangingChars="203" w:hanging="380"/>'),
    ("lc_hang", '<w:ind w:leftChars="50" w:left="489" w:hanging="380"/>'),
    ("lc_hangchars", '<w:ind w:leftChars="50" w:left="489" w:hangingChars="203"/>'),
]
SIZES = [21, 18]                   # half-points: 10.5 and 9.0
# COMPAT15=1 rewrites the base's settings.xml to compatibilityMode 15 without
# <w:useAltKinsokuLineBreakRules/>. The base (tokyoshugyo) is compat 11 and is
# the ONLY compat-11 document in the corpus; a1d6e4ef, whose render says the
# TWIPS win where these arms say the *Chars do, is compat 15.
COMPAT15 = os.environ.get("COMPAT15") == "1"
# GRID=1 gives the section a CHARACTER grid (linesAndChars, charSpace=1453 --
# a1d6e4ef's own grid). a1d6's note carries leftChars=50 left=489 and renders at
# the TWIP; these arms render at the *Chars. The base has no character grid.
GRID = os.environ.get("GRID") == "1"
# PARTS=styles|settings|both swaps those parts in from a1d6e4ef, whose own render
# resolves `leftChars=50 left=489` to the TWIP while these arms resolve the same
# attributes to the *Chars. Everything about the paragraph is already identical,
# so the difference has to live in the document around it.
PARTS = os.environ.get("PARTS") or ""
# STYLE names the pStyle the arms carry (a1d6e4ef's note uses "ac", whose own
# pPr turns off wordWrap and autoSpace).
STYLE = os.environ.get("STYLE") or "a"
A1D6 = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "a1d6e4efa2e7_tokumei_08_01-4.docx")


def modern(x):
    x = x.replace("<w:useAltKinsokuLineBreakRules/>", "")
    return re.sub(r'(w:name="compatibilityMode"[^>]*w:val=")[0-9]+', r"\g<1>15", x)


def tagname():
    return (("_c15" if COMPAT15 else "") + ("_grid" if GRID else "")
            + (("_" + PARTS) if PARTS else "")
            + (("_" + STYLE) if STYLE != "a" else ""))


def build():
    os.makedirs(OUT, exist_ok=True)
    blocks, index = [], []
    for sz in SIZES:
        for name, ind in ARMS:
            index.append((name, sz))
            blocks.append(
                '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
                '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
                '<w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/>'
                '</w:tblCellMar><w:tblBorders>'
                '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                '</w:tblBorders></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:pStyle w:val="%s"/><w:jc w:val="left"/>%s'
                '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
                '<w:t>甲亜亜亜亜</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>'
                % (CELL_TW, MAR_TW, MAR_TW, CELL_TW, CELL_TW, STYLE, ind, sz, sz))
    doc = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    sect = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc, re.S).group(0)
    sect = re.sub(r"<w:footerReference[^>]*/>", "", sect)
    if GRID:
        sect = re.sub(r"<w:docGrid[^>]*/>",
                      '<w:docGrid w:type="linesAndChars" w:linePitch="360"'
                      ' w:charSpace="1453"/>', sect)
    body = "<w:body>" + "".join(blocks) + sect + "</w:body>"
    new = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document %s>%s</w:document>' % (W_NS, body))
    dst = os.path.join(OUT, "indchars%s.docx" % tagname())
    shutil.copyfile(SRC, dst)
    swap = {}
    if PARTS:
        src2 = zipfile.ZipFile(A1D6)
        want = []
        if PARTS in ("styles", "both"):
            want.append("word/styles.xml")
        if PARTS in ("settings", "both"):
            want.append("word/settings.xml")
        for n in want:
            swap[n] = src2.read(n)
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "word/document.xml":
            data = new.encode("utf-8")
        elif COMPAT15 and item.filename == "word/settings.xml":
            data = modern(data.decode("utf-8")).encode("utf-8")
        elif PARTS and item.filename in swap:
            data = swap[item.filename]
        zout.writestr(item, data)
    zout.close()
    with open(os.path.join(OUT, "arms.txt"), "w", encoding="utf-8") as fh:
        for name, sz in index:
            fh.write(name + "," + str(sz) + os.linesep)
    print("built %s (%d arms)" % (dst, len(index)))


def to_pdf():
    import win32com.client as wc
    docx = os.path.join(OUT, "indchars%s.docx" % tagname())
    pdf = os.path.join(OUT, "indchars%s.pdf" % tagname())
    app = wc.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf),
                              ExportFormat=17, OpenAfterExport=False)
        d.Close(False)
    finally:
        app.Quit()


def measure():
    import fitz
    index = [l.split(",") for l in open(os.path.join(OUT, "arms.txt"),
                                        encoding="utf-8").read().split()]
    doc = fitz.open(os.path.join(OUT, "indchars%s.pdf" % tagname()))
    heads, rules = [], []
    for page in doc:
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > 3:
                    rules.append(round((it[1].x + it[2].x) / 2, 2))
                elif it[0] == "re" and it[1].width < 0.9 and it[1].height > 3:
                    rules.append(round(it[1].x0, 2))
        rows = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = sorted([c for s in l["spans"] for c in s["chars"]],
                            key=lambda c: c["origin"][0])
                if ch:
                    rows.append((round(l["bbox"][1], 1), ch))
        for _, ch in sorted(rows, key=lambda t: t[0]):
            if ch[0]["c"] == "甲":
                heads.append(ch[0]["origin"][0])
    left_rule = min(rules) if rules else 0.0
    inner = left_rule + MAR_TW / 20.0
    print("cell left rule %.2f  ->  inner edge %.2f  (margin %.2f)"
          % (left_rule, inner, MAR_TW / 20.0))
    print("   arm        size   x0       indent    left(tw)=4.05  1 char   fl(tw)=9.95")
    for (name, sz), x in zip(index, heads):
        one = int(sz) / 2.0
        print("   %-9s %5.1f  %7.2f  %+7.2f" % (name, one, x, x - inner))


if __name__ == "__main__":
    build()
    to_pdf()
    measure()
