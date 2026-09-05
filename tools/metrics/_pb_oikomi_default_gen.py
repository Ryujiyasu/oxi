# -*- coding: utf-8 -*-
"""Does Word pull the next character in by compressing mid-line 、。 when the run
sits AT the grid default size (legacy linesAndChars + compressPunctuation)?

correspondence__03ca64d7 (ＭＳ 明朝 10.5 = docDefaults 10.5, compat 14, kern 2,
balanceSingleByteDoubleByteWidth, linesAndChars 298, adjustRightInd default):
Word's first body line holds 40 cells and wraps 「ん」 although compressing its
two marks by 5.25 each would fit it; Oxi (S568 cap 6.0 per mark) pulls it in and
the last paragraph spills to a second page. Reproduce the line and sweep the
discriminators: number of marks (demand per mark), size regime, compat, kern,
balance flag, body vs. table cell. Read Word's own PDF: first-line character
count and the advance of every mark.

    python _pb_oikomi_default_gen.py gen
    python _pb_oikomi_default_gen.py pdf      # Word truth (COM -> PDF)
    python _pb_oikomi_default_gen.py oxi      # Oxi, same arms (--dump-layout)
"""
import collections
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_oikomi_default")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
# 40 cells available (floor of 425.2 / 10.5). Texts are 41 characters so that
# packing the 41st needs 10.5pt = one cell, split over the marks present.
BASE = "日本文化政策学会では標記の研究会を下記により開催いたします会員の方はもちろん"  # 39 chars, no marks
def with_marks(k, total=41):
    # insert k marks at even spacing into BASE, then pad/trim to `total` chars
    t = list(BASE)
    positions = [int(len(t) * (i + 1) / (k + 1)) for i in range(k)]
    for j, p in enumerate(sorted(positions, reverse=True)):
        t.insert(p, "。" if j % 2 else "、")
    s = "".join(t)
    filler = "研究会開催案内文化財保護政策"
    while len(s) < total:
        s += filler[len(s) % len(filler)]
    return s[:total] + "以下省略"

# (label, sz half-points of the run, docDefaults sz, marks, compat, kern, balance, cell)
ARMS = []
for k in (1, 2, 3, 4):
    ARMS.append(("at_m%d" % k, 21, 21, k, 14, True, True, False))
ARMS += [
    ("at_m2_c15", 21, 21, 2, 15, True, True, False),
    ("at_m2_nokern", 21, 21, 2, 14, False, True, False),
    ("at_m2_nobal", 21, 21, 2, 14, True, False, False),
    ("at_m2_cell", 21, 21, 2, 14, True, True, True),
    ("below_m2", 21, 22, 2, 14, True, True, False),   # run 10.5 under default 11
    ("above_m2", 24, 21, 2, 14, True, True, False),   # run 12 over default 10.5
    ("at_m2_c11", 21, 21, 2, 11, True, True, False),
    # 2026-09-05 second sweep: the four JA-blind docs where Word DOES pack at the
    # default size are all two-column (cols num=2 space=440); the refusing docs
    # and the arms above are single-column. Column measure = (425.2-22)/2 =
    # 201.6pt = 19 cells; the text is 20 chars so packing needs one cell.
    # The label prefix selects the section: 2col_ / narrow_ (single column of
    # the same 201.6pt measure via margins) / anything else = full body.
    ("2col_m1", 21, 21, 1, 14, True, True, False),
    ("2col_m2", 21, 21, 2, 14, True, True, False),
    ("2col_m3", 21, 21, 3, 14, True, True, False),
    ("2col_m2_c11", 21, 21, 2, 11, True, True, False),
    ("narrow_m2", 21, 21, 2, 14, True, True, False),
    # third sweep: the packed lines in the two-column docs sit in hanging-indent
    # paragraphs (46-69 w:hanging per doc). hang_: ind left=210 hanging=210 (the
    # first line keeps the full 40 cells); first_: firstLine=210 (39 cells, the
    # text is 40 chars); hang2col_: hanging inside the two-column section.
    ("hang_m2", 21, 21, 2, 14, True, True, False),
    ("hang_m3", 21, 21, 3, 14, True, True, False),
    ("first_m2", 21, 21, 2, 14, True, True, False),
    ("hang2col_m2", 21, 21, 2, 14, True, True, False),
    # fourth sweep: the packing docs (0ea3ec86 / 167853 / 0b6f3b32 / 13e1b7fc) all
    # carry a POSITIVE charSpace on their linesAndChars grid (3194 / 2048 =
    # +0.78 / +0.50pt per cell over the 11pt glyph); 03ca64d7 and every arm
    # above have none (cell = glyph). cs<N>_ arms: docGrid charSpace=N, text
    # length = floor(425.2 / pitch) + 1 so packing still needs one cell.
    ("cs3194_m1", 21, 21, 1, 14, True, True, False),
    ("cs3194_m2", 21, 21, 2, 14, True, True, False),
    ("cs3194_m3", 21, 21, 3, 14, True, True, False),
    ("cs2048_m2", 21, 21, 2, 14, True, True, False),
    ("cs3194_m2_c11", 21, 21, 2, 11, True, True, False),
    ("cs-2880_m2", 21, 21, 2, 14, True, True, False),
    # fifth sweep: in 0ea3ec86 every compressed at-default mark sits on a line
    # whose OVERFLOWING character is itself a line-final 、 (kinsoku forbids it
    # at a line start): Word compresses a mid-line mark to ~0.52em to pull the
    # trailing mark in. end<M>_ arms: the 41st character is 、 with M mid marks.
    ("end_m0", 21, 21, 0, 14, True, True, False),
    ("end_m1", 21, 21, 1, 14, True, True, False),
    ("end_m2", 21, 21, 2, 14, True, True, False),
    ("end_m3", 21, 21, 3, 14, True, True, False),
    ("end_m1_c11", 21, 21, 1, 11, True, True, False),
    ("end_m1_c15", 21, 21, 1, 15, True, True, False),
    ("endkuten_m1", 21, 21, 1, 14, True, True, False),
]


def docx(label):
    return os.path.join(OUT, "oikomi_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             'relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             "</Relationships>")
    for label, sz, dsz, k, compat, kern, bal, cell in ARMS:
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                    '<w:characterSpacingControl w:val="compressPunctuation"/>'
                    "<w:compat>" + ("<w:balanceSingleByteDoubleByteWidth/>" if bal else "")
                    + "<w:useFELayout/>"
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/>'
                    "</w:compat>"
                    '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>' % compat)
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>'
                  + ('<w:kern w:val="2"/>' if kern else "")
                  + '<w:sz w:val="%d"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>') % (MINCHO, dsz)
        rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/>' + ('' if sz == dsz else '<w:sz w:val="%d"/>' % sz) + "</w:rPr>"
        kind = ("2col" if label.startswith("2col") or label.startswith("hang2col") else
                ("narrow" if label.startswith("narrow") else "body"))
        total = 41 if kind == "body" else 20
        cs = None
        if label.startswith("cs"):
            cs = int(label.split("_")[0][2:])
            pitch = sz / 2.0 + cs / 4096.0
            total = int(425.2 // pitch) + 1
        ppr = ""
        if label.startswith("hang"):
            ppr = '<w:pPr><w:ind w:left="210" w:hanging="210"/></w:pPr>'
        elif label.startswith("first"):
            ppr = '<w:pPr><w:ind w:firstLine="210"/></w:pPr>'
            total = 40
        text = with_marks(k, total)
        if label.startswith("end"):
            # 40 body characters (with k mid marks) + a line-final mark as the 41st
            mark = "。" if label.startswith("endkuten") else "、"
            text = with_marks(k, 40)[:40] + mark + "以下省略"
        para = "<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>" % (ppr, rpr, text)
        if cell:
            # one-cell table whose cell text width equals the body: 425.2pt = 8504tw + margins 108x2
            body = ('<w:tbl><w:tblPr><w:tblW w:w="8720" w:type="dxa"/><w:tblCellMar>'
                    '<w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                    '<w:tblGrid><w:gridCol w:w="8720"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8720" w:type="dxa"/></w:tcPr>'
                    + para + "</w:tc></w:tr></w:tbl><w:p/>")
        else:
            body = para
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               + ('<w:pgMar w:top="1985" w:right="3937" w:bottom="1701" w:left="3937" w:header="851" w:footer="992"/>'
                  if kind == "narrow" else
                  '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>')
               + ('<w:cols w:num="2" w:space="440"/>' if kind == "2col" else "")
               + ('<w:docGrid w:type="linesAndChars" w:linePitch="298" w:charSpace="%d"/>' % cs if cs is not None
                  else '<w:docGrid w:type="linesAndChars" w:linePitch="298"/>')
               + '</w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def first_line_pdf(path):
    import fitz
    pg = fitz.open(path)[0]
    rows = []
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"].strip()]
            if chars:
                rows.append((round(chars[0]["origin"][1], 1), chars, l["spans"][0]["size"]))
    rows.sort(key=lambda r: r[0])
    y, chars, sz = rows[0]
    marks = [round((chars[i + 1]["origin"][0] - chars[i]["origin"][0]) / sz, 2)
             for i in range(len(chars) - 1) if chars[i]["c"] in "、。"]
    return len(chars), marks, "".join(c["c"] for c in chars)


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, *_ in ARMS:
            out = docx(label)[:-5] + ".word.pdf"
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(out, 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF export): first-line chars (41 = packed, 40 = refused), mark advances (em) ==")
    for label, sz, dsz, k, compat, kern, bal, cell in ARMS:
        n, marks, t = first_line_pdf(docx(label)[:-5] + ".word.pdf")
        print("%-14s run=%.1f dflt=%.1f M=%d c%d kern=%d bal=%d cell=%d -> n=%d marks=%s" % (label, sz / 2, dsz / 2, k, compat, kern, bal, cell, n, marks))


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        key, _, v = kv.partition("=")
        env[key] = v or "1"
    print("== OXI %s: first-line chars ==" % (envs or "(default)"))
    for label, sz, dsz, k, compat, kern, bal, cell in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "oikomi_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "oik"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        pg = json.load(open(dump, encoding="utf-8"))["pages"][0]["elements"]
        txt = [e for e in pg if e["type"] == "text" and (e.get("text") or "").strip()]
        y0 = min(round(e["y"], 1) for e in txt)
        first = sorted([e for e in txt if round(e["y"], 1) == y0], key=lambda e: e["x"])
        n = sum(len(e["text"].strip()) for e in first)
        marks = [round(e["w"] / (sz / 2.0), 2) for e in first if e["text"] in ("、", "。")]
        print("%-14s -> n=%d marks=%s" % (label, n, marks))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
