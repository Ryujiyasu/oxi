# -*- coding: utf-8 -*-
"""The line box of an EMPTY body paragraph with snapToGrid=0 in a Japanese document whose
section HAS a typed docGrid (linesAndChars).

ja/forms__02157d72 p1: three empty paragraphs (mark sz=18, ascii Century via minorHAnsi,
eastAsia ＭＳ 明朝, snapToGrid=0, docGrid linesAndChars 290) each sit between a floating
table and a heading. Oxi gives each 10.50 (Century 9pt GDI pixel line 14px); the COM
Info6 steps and the heading baselines say Word gives ~10.84 (= Century hhea 1.204 x 9,
the S1574 cell law and the no-type-grid law of _pb_emptypara_lh_gen). Three such
paragraphs = -1.0 of the -2.5 that doc carries at its page bottom.

Sheet: «前» (ＭＳ 明朝 10.5, makes the document CJK), N empty paragraphs, «後». The
per-empty height is (gap(N=9) - gap(N=1)) / 8 with gap = baseline(後) - baseline(前), so
the text lines cancel. Arms: ascii face (Century / Times New Roman / ＭＳ 明朝) x mark
sz (18 / 21) x grid (lc290_snap0: linesAndChars 290 with snapToGrid=0 on every paragraph
/ nogrid: no docGrid / lc290_snap: linesAndChars 290, snapping on -- the control).

    python tools/metrics/_pb_empty_snap0_gen.py            (PB_OXI=<renderer.exe> adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_empty_snap0")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat><w:useFELayout/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="%%s" w:hAnsi="%%s" w:eastAsia="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>') % W
GRIDS = {"lc290_snap0": ('<w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="-1209"/>', True),
         "nogrid": ("", True), "lc290_snap": ('<w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="-1209"/>', False)}


def build(name, ascii_face, sz, grid, n):
    dg, snap0 = GRIDS[grid]
    ppr = '<w:pPr>%s<w:rPr><w:sz w:val="%d"/></w:rPr></w:pPr>' % ('<w:snapToGrid w:val="0"/>' if snap0 else "", sz)
    tppr = '<w:pPr>%s</w:pPr>' % ('<w:snapToGrid w:val="0"/>' if snap0 else "")
    body = ('<w:p>%s<w:r><w:t>前</w:t></w:r></w:p>' % tppr + ('<w:p>%s</w:p>' % ppr) * n
            + '<w:p>%s<w:r><w:t>後</w:t></w:r></w:p>' % tppr)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="567" w:right="567" w:bottom="284" w:left="1134" w:header="851" w:footer="992"/>%s</w:sectPr></w:body></w:document>') % (W, body, dg)
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES % (ascii_face, ascii_face)); z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return p


def gap_word(pdf):
    import pymupdf
    ys = {}
    for b in pymupdf.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"] in "前後":
                        ys.setdefault(c["c"], c["origin"][1])
    return ys["後"] - ys["前"]


def gap_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    ys = {}
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        if c["char"] in "前後":
            ys.setdefault(c["char"], c["baseline"])
    return ys["後"] - ys["前"]


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    for face in ("Century", "Times New Roman", "ＭＳ 明朝"):
        for sz in (18, 21):
            for grid in GRIDS:
                gw, go = {}, {}
                for n in (1, 9):
                    name = "%s_sz%d_%s_n%d" % (face.replace(" ", "")[:7], sz, grid, n)
                    p = build(name, face, sz, grid, n); pdf = p[:-5] + ".pdf"
                    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
                    try:
                        d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                    finally:
                        w.Quit()
                    gw[n] = gap_word(pdf)
                    if exe:
                        go[n] = gap_oxi(exe, p)
                print("%-16s sz%d %-12s per-empty W %.3f%s   (gap1 W %.2f%s)" % (
                    face, sz, grid, (gw[9] - gw[1]) / 8, (" O %.3f" % ((go[9] - go[1]) / 8)) if exe else "",
                    gw[1], (" O %.2f" % go[1]) if exe else ""), flush=True)
