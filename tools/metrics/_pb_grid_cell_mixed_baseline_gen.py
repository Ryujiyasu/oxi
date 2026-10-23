# -*- coding: utf-8 -*-
"""Where does Word put the baseline of a table-cell line that mixes a CJK face and a Latin face?

ja/administrative__13c71e15 (docGrid lines 360): a cell «9月21日（土）» (Century digits + ＭＳ 明朝)
sits at 472.03 in Word for every glyph, the same as the pure-CJK cell «第１回» beside it;
Oxi set the Century digits at 472.88 and the ＭＳ 明朝 at 471.62 (each face's own ascent).

One row per arm: [pure CJK «第一回»] [mixed «9月21日（土）»] [pure Latin «2024»], all cells
vAlign top. Arms: grid (lines 360 / none) x Latin face (Century / Arial / Times New Roman) x
size (10 / 10.5 / 12pt). Prints each cell's baseline (PDF origin y) minus the row's
pure-CJK baseline, for Word and for Oxi (PB_OXI).

    python tools/metrics/_pb_grid_cell_mixed_baseline_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_grid_cell_mixed_baseline")
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
LATIN = {"century": "Century", "arial": "Arial", "tnr": "Times New Roman"}
SIZES = [20, 21, 24]


SPLIT = os.environ.get("PB_SPLIT") == "1"  # one run per character class (as 13c71e15)
VALIGN = os.environ.get("PB_VALIGN")  # e.g. center


def cell(text, latin, sz, tag):
    rpr = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>' % (latin, latin, sz)
    pieces = []
    if SPLIT:
        cur = ""
        for ch in text:
            if cur and (ch.isascii() != cur[-1].isascii()):
                pieces.append(cur); cur = ""
            cur += ch
        pieces.append(cur)
    else:
        pieces = [text]
    runs = "".join('<w:r>%s<w:t>%s</w:t></w:r>' % (rpr, t) for t in pieces)
    va = '<w:vAlign w:val="%s"/>' % VALIGN if VALIGN else ""
    return ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>%s</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
            '%s<w:r>%s<w:t>%s</w:t></w:r></w:p></w:tc>' % (va, runs, rpr, tag))


def build(grid):
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>') % W
    body = ""
    for lk, lf in LATIN.items():
        for sz in SIZES:
            t = "%s%d" % (lk[0].upper(), sz)
            body += ('<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
                     '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>'
                     + cell("第一回", lf, sz, "Ａ") + cell("9月21日（土）", lf, sz, "Ｂ") + cell("2024", lf, sz, "C")
                     + '</w:tr></w:tbl><w:p><w:r><w:t>%s</w:t></w:r></w:p>' % t)
    sect = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>%s</w:sectPr></w:body></w:document>'
           % (W, body, sect))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "gcm_%s%s%s.docx" % ("grid" if grid else "nogrid", "_split" if SPLIT else "", "_" + VALIGN if VALIGN else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return path


def rows_word(pdf):
    import pymupdf
    chars = []
    for pg in pymupdf.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        chars.append((c["c"], c["origin"][0], c["origin"][1]))
    return chars


def rows_oxi(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    chars = []
    for p in json.load(open(g, encoding="utf-8"))["pages"]:
        for c in p["glyphs"]:
            chars.append((c["char"], c["x"], c["baseline"]))
    return chars


def report(chars):
    """per row (marked by the tag paragraph) -> baselines of 第, 9, 月, 2 (Latin cell)"""
    out = []
    tags = [(c, y) for c, x, y in chars if c in "CTA" and False]
    # rows: find the A tag glyph (Ａ) per row; group by its y
    rows = sorted({round(y, 1) for c, x, y in chars if c == "Ａ"})
    for ry in rows:
        near = [(c, x, y) for c, x, y in chars if abs(y - ry) < 8]
        def base(ch, xmin=0, xmax=1e9):
            v = [y for c, x, y in near if c == ch and xmin <= x <= xmax]
            return v[0] if v else None
        cjk = base("第")
        out.append(tuple(None if v is None or cjk is None else round(v - cjk, 2)
                         for v in (base("9"), base("月"), base("2", 300, 1e9))))
    return out


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    for grid in (True, False):
        p = build(grid); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        wr = report(rows_word(pdf))
        orr = report(rows_oxi(exe, p)) if exe else [None] * len(wr)
        labels = ["%s%d" % (lk, sz) for lk in LATIN for sz in SIZES]
        print("grid" if grid else "nogrid", "(baseline - pure-CJK cell baseline: mixed 9, mixed 月, Latin cell 2)")
        for lab, a, b in zip(labels, wr, orr):
            print("  %-10s Word %s   Oxi %s" % (lab, a, b), flush=True)
