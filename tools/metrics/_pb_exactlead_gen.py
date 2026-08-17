# -*- coding: utf-8 -*-
"""Where in an `exact` line box does the glyph sit -- and does a shape differ?

1ec1's floor (0.8091) decomposes into two small rules. The first: its box's
opening paragraph carries `<w:spacing w:line="480" w:lineRule="exact"/>` at 14pt,
and Word puts the glyph top 10.90pt below the shape top (i.e. the baseline sits
near the BOTTOM of the 24pt box) while Oxi puts it 1.70pt below (the TOP). 9.2pt.

Before touching it, find out whether that bottom-anchoring is a SHAPE rule or the
general `exact` rule -- if the body path already matches Word, the fix is
shape-scoped; if the body is wrong too, the blast radius is every exact line.
Each arm is a marker paragraph (plain, single spacing) followed by the exact
paragraph, so the offset is measured against a known reference in both engines.

    python _pb_exactlead_gen.py gen
    python _pb_exactlead_gen.py pdf      # Word truth
    python _pb_exactlead_gen.py oxi      # Oxi, same arms
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_exactlead")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = os.environ.get("OXI_PB_FACE", "ＭＳ 明朝")
# OXI_PB_SZ sweeps the font size so the 0.800 slope can be checked for size
# dependence (MS Mincho's ascent fraction is 0.88, so 0.800 is not simply
# ascent/(ascent+descent) and needs pinning).
SZ_HP = int(os.environ.get("OXI_PB_SZ", "28"))   # half-points; 28 = 14pt (1ec1)
COMPAT = os.environ.get("OXI_PB_COMPAT", "15")
# (label, in_shape, line twips)
ARMS = [("body_%d" % v, False, v) for v in (240, 360, 480, 600)]
# The hand-written wps shape below is not accepted by Word yet (it refuses to
# open the file; the recorded trap is that xmlns:wps has to sit on
# mc:AlternateContent). The BODY arms already answer the question that matters —
# whether `exact` bottom-anchoring is the general rule or shape-scoped — so keep
# them separable and default to body-only until the shape XML is fixed.
if not os.environ.get("OXI_PB_SHAPE"):
    pass
else:
    ARMS += [("shape_%d" % v, True, v) for v in (240, 360, 480, 600)]


def docx():
    return os.path.join(OUT, "exactlead.docx")


def para(text, sz=SZ_HP, ppr=""):
    return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, FACE, FACE, FACE, sz, FACE, FACE, FACE, sz, sz, text))


def shape(ai, inner):
    """An inline rect textbox with tIns/bIns = 0, like 1ec1's boxes."""
    return (
        '<w:p><w:pPr><w:rPr><w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
        '<w:sz w:val="%d"/></w:rPr><mc:AlternateContent '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" '
        'distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/'
        'drawingml/2006/wordprocessingDrawing">'
        '<wp:extent cx="5000000" cy="900000"/>'
        '<wp:docPr id="%d" name="box%d"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/'
        'wordprocessingShape">'
        '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/'
        'wordprocessingShape">'
        '<wps:spPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="5000000" cy="900000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        "</wps:spPr>"
        # `inner` goes through a %s placeholder, NOT string concatenation: `%`
        # binds tighter than `+`, so "a" + x + "b" % args formats only "b" and
        # raises "not all arguments converted".
        "<wps:txbx><w:txbxContent>%s</w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="0" '
        'rIns="91440" bIns="0" anchor="t" anchorCtr="0"/>'
        "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>"
        "</mc:Choice></mc:AlternateContent></w:r></w:p>"
        % (SZ_HP, SZ_HP, 100 + ai, ai, inner))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, in_shape, line) in enumerate(ARMS):
        body.append(para("A%02dZ" % ai, SZ_HP,
                         "<w:pageBreakBefore/>" if ai else ""))
        exact = para("M%02dマーカー行" % ai, SZ_HP,
                     '<w:snapToGrid w:val="0"/>'
                     '<w:spacing w:line="%d" w:lineRule="exact"/>' % line)
        if in_shape:
            body.append(shape(ai, exact + para("T%02dあと" % ai, SZ_HP)))
        else:
            body.append(exact)
            body.append(para("T%02dあと" % ai, SZ_HP))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/'
           'wordprocessingShape"><w:body>' + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
              "</w:styles>" % (FACE, FACE, FACE, SZ_HP))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS +
                '><w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word"'
                ' w:val="%s"/></w:compat></w:settings>' % COMPAT)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms; sz %.1fpt; compat %s"
          % (SZ_HP / 2.0, COMPAT))


def report(rows, who):
    print("== %s ==" % who)
    print("%-11s %-7s %-9s %-9s %-9s %s"
          % ("arm", "line_pt", "A_y", "M_y", "T_y", "A->M / M->T"))
    for (label, in_shape, line), r in rows:
        if not r or r.get("m") is None:
            print("%-11s %-7.1f MISSING" % (label, line / 20.0))
            continue
        a, m, t = r.get("a"), r.get("m"), r.get("t")
        print("%-11s %-7.1f %-9s %-9.2f %-9s %s"
              % (label, line / 20.0,
                 "%.2f" % a if a is not None else "-", m,
                 "%.2f" % t if t is not None else "-",
                 "%s / %s"
                 % ("%.2f" % (m - a) if a is not None else "-",
                    "%.2f" % (t - m) if t is not None else "-")))


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
    for ai, arm in enumerate(ARMS):
        r = {}
        for pi in range(doc.page_count):
            found = False
            for b in doc[pi].get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"])
                    for key, pat in (("a", "A%02dZ" % ai), ("m", "M%02d" % ai),
                                     ("t", "T%02d" % ai)):
                        if pat in t and key not in r:
                            r[key] = round(ln["bbox"][1], 2)
                            found = True
            if found:
                break
        rows.append((arm, r))
    report(rows, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "exactlead_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "el"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    rows = []
    for ai, arm in enumerate(ARMS):
        r = {}
        for pg in pages:
            rowsy = {}
            for e in pg["elements"]:
                if e["type"] == "text":
                    rowsy.setdefault(round(e["y"], 2), []).append(e)
            hit = False
            for y, v in sorted(rowsy.items()):
                t = "".join(x.get("text") or "" for x in sorted(v, key=lambda e: e["x"]))
                for key, pat in (("a", "A%02dZ" % ai), ("m", "M%02d" % ai),
                                 ("t", "T%02d" % ai)):
                    if pat in t and key not in r:
                        r[key] = y
                        hit = True
            if hit:
                break
        rows.append((arm, r))
    report(rows, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
