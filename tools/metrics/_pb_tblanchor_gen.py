# -*- coding: utf-8 -*-
"""Where does a MARGIN-ANCHORED floating table put its left edge?

tokyoshugyo p17 is a full-page floating table (`w:tblpPr horzAnchor="margin"`,
no `tblpX`) and Word and Oxi disagree about its left edge by exactly the default
cell margin:

    Word  border x = 79.7   text x0 = 85.1   (= the page margin, 1701tw)
    Oxi   border x = 85.05  text x0 = 90.5

i.e. Word pulls the table LEFT by w:tblCellMar/left (108tw = 5.4pt) so the cell
TEXT lands on the margin, and Oxi puts the BORDER on the margin instead. One
character less fits per line, the block runs long, and p18-19 inherit +21pt
(`_kojin_rowgeom.py scan`).

Sweep the cell margin to see whether the offset tracks it, and sweep tblpX and
the non-floating case to bound the rule.

    python _pb_tblanchor_gen.py gen
    python _pb_tblanchor_gen.py pdf      # Word truth
    python _pb_tblanchor_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_tblanchor")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
SZ_HP = 21                 # 10.5pt
LEFT_TW = 1701             # page left margin, as in tokyoshugyo
MARGIN_PT = LEFT_TW / 20.0
# ★A docx with NO settings.xml does not mean "current Word": the absorption this
# probe measures is a compatibilityMode <= 14 behaviour (S621), and the first run
# here shipped no settings.xml at all, so every Word arm was really an old-mode
# arm. tokyoshugyo itself declares compatibilityMode 11. Always state it.
COMPAT = os.environ.get("OXI_PB_COMPAT", "11")

# (label, floating?, cellMar_left_tw, tblpX_tw or None, tblInd_tw or None)
ARMS = [
    ("float_cm108", True, 108, None, None),
    ("float_cm0", True, 0, None, None),
    ("float_cm200", True, 200, None, None),
    ("float_cm400", True, 400, None, None),
    ("float_x0", True, 108, 0, None),
    ("float_x567", True, 108, 567, None),
    ("plain_cm108", False, 108, None, None),
    ("plain_cm400", False, 400, None, None),
    ("plain_ind0", False, 108, None, 0),
    ("plain_ind567", False, 108, None, 567),
]


def docx():
    return os.path.join(OUT, "tblanchor.docx")


def para(text):
    return ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (FACE, FACE, FACE, SZ_HP, FACE, FACE, FACE, SZ_HP, SZ_HP, text))


def table(ai, floating, cm, tblpx, tblind):
    pr = ["<w:tblPr>"]
    if floating:
        pr.append('<w:tblpPr w:leftFromText="142" w:rightFromText="142"'
                  ' w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="1"'
                  + (' w:tblpX="%d"' % tblpx if tblpx is not None else "") + "/>")
    if tblind is not None:
        pr.append('<w:tblInd w:w="%d" w:type="dxa"/>' % tblind)
    pr.append('<w:tblW w:w="0" w:type="auto"/>')
    pr.append('<w:tblBorders>' + "".join(
        '<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % s
        for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>")
    pr.append('<w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
              '<w:left w:w="%d" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
              '<w:right w:w="%d" w:type="dxa"/></w:tblCellMar>' % (cm, cm))
    pr.append("</w:tblPr>")
    return ("<w:tbl>" + "".join(pr) +
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>' +
            para("M%02dX" % ai) + "</w:tc></w:tr></w:tbl>")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, fl, cm, px, ind) in enumerate(ARMS):
        body.append(para("A%02dZ" % ai))
        body.append(table(ai, fl, cm, px, ind))
        body.append('<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="%d" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/>'
           "</w:sectPr></w:body></w:document>" % LEFT_TW)
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
    print("wrote", docx(), len(ARMS), "arms; margin %.2fpt; compat %s"
          % (MARGIN_PT, COMPAT))


def report(per, who):
    print("== %s == (page margin %.2fpt)" % (who, MARGIN_PT))
    print("%-13s %-5s %-7s %-7s %-9s %-9s %-9s %s"
          % ("arm", "float", "cellMar", "tblpX", "border_x", "text_x", "b-margin", "predict"))
    for ai, (label, fl, cm, px, ind) in enumerate(ARMS):
        g = per.get(ai) or {}
        bx, tx = g.get("border"), g.get("text")
        # the hypothesis: border = margin + (tblpX or tblInd or 0) - cellMar_left
        pred = MARGIN_PT + (px or ind or 0) / 20.0 - cm / 20.0
        print("%-13s %-5s %-7d %-7s %-9s %-9s %-9s %.2f"
              % (label, "yes" if fl else "no", cm,
                 str(px if px is not None else ("ind%d" % ind if ind is not None else "-")),
                 "%.2f" % bx if bx is not None else "-",
                 "%.2f" % tx if tx is not None else "-",
                 "%+.2f" % (bx - MARGIN_PT) if bx is not None else "-", pred))


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
    per = {}
    for pi in range(doc.page_count):
        pg = doc[pi]
        ai = None
        txt = pg.get_text()
        import re
        m = re.search(r"M(\d\d)X", txt)
        if not m:
            continue
        ai = int(m.group(1))
        xs = []
        for dr in pg.get_drawings():
            for it in dr["items"]:
                if it[0] == "l" and abs(it[1].x - it[2].x) < 0.3 \
                        and abs(it[2].y - it[1].y) > 5:
                    xs.append(it[1].x)
                elif it[0] == "re" and it[1].width < 2.0 and it[1].height > 5:
                    xs.append((it[1].x0 + it[1].x1) / 2)
        tx = None
        for b in pg.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                if "M%02dX" % ai in "".join(s["text"] for s in ln["spans"]):
                    tx = ln["bbox"][0]
        per[ai] = {"border": min(xs) if xs else None, "text": tx}
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "tblanchor_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "ta"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    per = {}
    import re
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        txt = "".join(e.get("text") or "" for e in pg["elements"] if e["type"] == "text")
        m = re.search(r"M(\d\d)X", txt)
        if not m:
            continue
        ai = int(m.group(1))
        xs = [e["x"] for e in pg["elements"]
              if e["type"] == "border" and (e.get("w") or 0) <= 2.0]
        tx = min((e["x"] for e in pg["elements"]
                  if e["type"] == "text" and (e.get("text") or "").startswith("M")),
                 default=None)
        per[ai] = {"border": min(xs) if xs else None, "text": tx}
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
