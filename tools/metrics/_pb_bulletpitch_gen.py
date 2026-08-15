# -*- coding: utf-8 -*-
"""Bullet-paragraph pitch inside a styled table cell (NDIS -1 root probe).

technical__0043bfe0 (NDIS price guide) p81+: every price row's name cell is
[name para, then 1-2 ListParagraph bullets].  Word's measured pitches vs Oxi:

    wrapped line      9.24/9.36 vs 9.20      (-0.1, px alternation)
    para -> bullet    11.64     vs 11.20     (-0.44)
    bullet -> bullet  11.76     vs 11.20     (-0.56)
    row boundary      13.68     vs 13.70     (OK)

Spec inputs: table style GridTable4-Accent1 pPr spacing before=40 after=40
line=240 auto; bullets are ListParagraph + numPr (marker rFonts=Symbol,
lvlText F0B7) with contextualSpacing w:val="0"; runs carry NO w:sz -- the 8pt
comes from the table style rPr sz=16 via overrideTableStyleFontSizeAndJustification.
docDefaults pPr is before=100 after=100 line=300 atLeast (overridden in-table).

Word's observed inter-para extra is 2.40pt, which is neither max(2,2)=2 nor
2+2=4.  Two decompositions are degenerate on the NDIS data alone:
  (i)  gap=2.4 and the bullet line is Arial-height
  (ii) gap=2.0 and the marker's Symbol metrics lift the bullet line +0.4
This probe breaks the tie: asymmetric spacing arms (40/0, 0/40, 80/40, 40/80)
pin the gap rule; marker-font arms (Symbol vs Arial marker vs inline glyph)
pin the line-height contribution.  A second identical row measures the
row-boundary pitch; a body arm asks whether the gap rule is table-only.

  python _pb_bulletpitch_gen.py gen
  python _pb_bulletpitch_gen.py pdf     # Word truth via ExportAsFixedFormat
  python _pb_bulletpitch_gen.py oxi     # same pitches from --dump-layout
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_bulletpitch")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# (name, table-style spacing (before,after) or None for body arm, marker kind, sz half-points)
#   marker kind: "sym" = numPr Symbol F0B7, "ari" = numPr Arial U+2022,
#                "inl" = no numPr, inline Symbol F0B7 run + text,
#                "non" = no numPr, plain text only
ARMS = [
    ("A_base_sym4040",  (40, 40),  "sym", 16),
    ("B_marker_arial",  (40, 40),  "ari", 16),
    ("C_inline_glyph",  (40, 40),  "inl", 16),
    ("D_after_only",    (40, 0),   "sym", 16),
    ("E_before_only",   (0, 40),   "sym", 16),
    ("F_b80_a40",       (80, 40),  "sym", 16),
    ("G_b40_a80",       (40, 80),  "sym", 16),
    ("H_16pt",          (40, 40),  "sym", 32),
    ("I_body",          None,      "sym", 16),
    ("J_no_marker",     (40, 40),  "non", 16),
    # K/L (2026-08-15): the Wingdings marker. 182 corpus docs drive a bullet
    # through a numbering.xml rFonts="Wingdings", and until S1140 the registry
    # had no Wingdings entry, so those markers were measured with the default
    # (Calibri) metrics. Word's own figure decides whether the real 1.10986em
    # face is what lifts the line.
    ("K_marker_wing",   (40, 40),  "win", 16),
    ("L_marker_wing16", (40, 40),  "win", 32),
]


def docx():
    return os.path.join(OUT, "bulletpitch.docx")


def styles():
    s = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">",
         # docDefaults exactly as NDIS writes them (the atLeast300 default is
         # part of the input space -- the table style must be what defeats it)
         "<w:docDefaults><w:rPrDefault><w:rPr>"
         '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/>'
         "</w:rPr></w:rPrDefault>"
         '<w:pPrDefault><w:pPr><w:spacing w:before="100" w:after="100" w:line="300" w:lineRule="atLeast"/></w:pPr>'
         "</w:pPrDefault></w:docDefaults>",
         '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
         '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
         '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>',
         '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
         '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar>'
         '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
         '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
         "</w:tblCellMar></w:tblPr></w:style>"]
    for name, sp, _mk, sz in ARMS:
        if sp is None:
            continue
        s.append(
            '<w:style w:type="table" w:styleId="TS%s"><w:name w:val="TS%s"/>'
            '<w:basedOn w:val="TableNormal"/>'
            '<w:pPr><w:spacing w:before="%d" w:after="%d" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:rPr><w:sz w:val="%d"/></w:rPr>'
            "<w:tblPr><w:tblBorders>"
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="95B3D7"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="95B3D7"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="95B3D7"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="95B3D7"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="95B3D7"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="95B3D7"/>'
            "</w:tblBorders></w:tblPr></w:style>" % (name, name, sp[0], sp[1], sz))
    s.append("</w:styles>")
    return "".join(s)


NUMBERING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
    '<w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
    "</w:abstractNum>"
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
    '<w:lvlText w:val="&#x2022;"/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr></w:lvl>'
    "</w:abstractNum>"
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
    "</w:numbering>")

# compat15 + overrideTableStyleFontSizeAndJustification, as NDIS
SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
    "<w:compat>"
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"'
    ' w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    "</w:compat></w:settings>")


def rfonts():
    return ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            '<w:szCs w:val="16"/></w:rPr>')


def name_para():
    # NDIS name para: pPr carries only rPr (all spacing from the table style)
    return ("<w:p><w:pPr>" + rfonts() + "</w:pPr>"
            "<w:r>" + rfonts() + '<w:t xml:space="preserve">Name line sample </w:t></w:r></w:p>')


def bullet_para(mk, txt):
    if mk in ("sym", "ari", "win"):
        num = {"sym": "1", "ari": "2", "win": "3"}[mk]
        return ('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
                '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%s"/></w:numPr>'
                '<w:contextualSpacing w:val="0"/>' % num + rfonts() + "</w:pPr>"
                "<w:r>" + rfonts() + '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % txt)
    lead = ""
    if mk == "inl":
        lead = ('<w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/>'
                '<w:szCs w:val="16"/></w:rPr><w:t>&#xF0B7;</w:t></w:r>')
    return ('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
            '<w:ind w:left="284" w:hanging="284"/>'
            '<w:contextualSpacing w:val="0"/>' + rfonts() + "</w:pPr>" + lead +
            "<w:r>" + rfonts() + '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % txt)


def body_para(mk, txt, spacing=True):
    # I_body arm: same paragraph stack, direct spacing instead of a table style
    sp = '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>' if spacing else ""
    if mk == "name":
        return ('<w:p><w:pPr>' + sp + '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
                '<w:sz w:val="16"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % txt)
    return ('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' + sp +
            '<w:contextualSpacing w:val="0"/>'
            '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % txt)


def cell(mk):
    return ('<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
            + name_para()
            + bullet_para(mk, "Must be a support one.")
            + bullet_para(mk, "Must be a support two.")
            + "</w:tc>")


def marker(tag, pbb=False):
    return ('<w:p><w:pPr>%s'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr>'
            "<w:t>%s</w:t></w:r></w:p>"
            % ("<w:pageBreakBefore/>" if pbb else "", tag))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, sp, mk, _sz) in enumerate(ARMS):
        body.append(marker("M%02dS" % ai, pbb=ai > 0))
        if sp is None:
            body.append(body_para("name", "Name line sample "))
            body.append(body_para("b", "Must be a support one."))
            body.append(body_para("b", "Must be a support two."))
        else:
            # two identical rows: within-cell pitches + the row-boundary pitch
            body.append(
                '<w:tbl><w:tblPr><w:tblStyle w:val="TS%s"/>'
                '<w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0"'
                ' w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>'
                "<w:tr>%s</w:tr><w:tr>%s</w:tr></w:tbl>" % (name, cell(mk), cell(mk)))
        body.append(marker("M%02dE" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/numbering.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                    '<Override PartName="/word/settings.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                          '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles())
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


LINES = ["name1", "bullet1", "bullet2", "name2(row2)", "bullet3", "bullet4"]


def report(rows, who):
    print("== %s ==" % who)
    print("%-16s %8s %8s %8s %8s %8s" %
          ("arm", "n->b1", "b1->b2", "b2->NAME2", "n2->b3", "b3->b4"))
    for name, ys in rows:
        if ys is None or len(ys) < 2:
            print("%-16s MISSING (%s)" % (name, ys))
            continue
        ds = ["%8.2f" % (ys[i + 1] - ys[i]) for i in range(len(ys) - 1)]
        print("%-16s %s" % (name, " ".join(ds)))


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    rows = []
    for ai, (name, _sp, _mk, _sz) in enumerate(ARMS):
        pg = doc[ai]
        ys = []
        for bl in pg.get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                for sp2 in ln["spans"]:
                    t = sp2["text"]
                    if t.startswith(("Name line", "Must be a")):
                        ys.append(round(sp2["bbox"][1], 2))
        ys.sort()
        rows.append((name, ys))
    report(rows, "WORD pdf")
    # marker glyph anatomy on the base arm
    pg = doc[0]
    for bl in pg.get_text("dict")["blocks"]:
        if bl["type"] != 0:
            continue
        for ln in bl["lines"]:
            for sp2 in ln["spans"]:
                if "Symbol" in sp2["font"] or sp2["text"] in ("•", ""):
                    print("  marker span font=%s size=%.2f bbox_h=%.2f y0=%.2f"
                          % (sp2["font"], sp2["size"],
                             sp2["bbox"][3] - sp2["bbox"][1], sp2["bbox"][1]))


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "bulletpitch_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "bp"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    rows = []
    for ai, (name, _sp, _mk, _sz) in enumerate(ARMS):
        ys = []
        if ai < len(pages):
            for e in pages[ai]["elements"]:
                t = e.get("text") or ""
                if t.startswith(("Name line", "Must be a")):
                    ys.append(round(e["y"], 2))
        ys.sort()
        rows.append((name, ys))
    report(rows, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
