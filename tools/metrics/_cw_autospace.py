# -*- coding: utf-8 -*-
"""Is the gap Word puts between kana and Latin compressible aki?

04b88e's line fits its cell to 0.00pt, and its two CJK/Latin gaps are 1.08 and 0.85
where a controlled sweep at the same size, spacing, balanceSBDB and ascii face says
1.556. The shortfall the line would otherwise carry is about the same 1.13pt it
appears to have saved there, which reads as the demand-driven compression already
derived for 約物, applied to the gap instead. That is a reading off one line, so
measure it: same text, no 約物 anywhere, cell width closing a twip at a time.

If the gap is fixed, the line wraps the moment it is short, exactly as the 約物-free
control arm did. If it is aki, the gap gives way first and the line holds.

    python _cw_autospace.py
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
OUT = os.path.join(REPO, "pipeline_data", "_cw_law")
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TEXT = "甲亜亜H亜亜"          # two CJK/Latin joints, no 約物 at all
SZ = 16                        # 8pt, as 04b88e
CS = -20                       # w:spacing, as 04b88e
FONT = "ＭＳ 明朝"
ASCII_FACE = "Century"
W0, W1, STEP = 700.0, 1200.0, 1.0      # cell width in twips: 35..60pt


def widths():
    ws, w = [], W0
    while w <= W1 + 1e-6:
        ws.append(w)
        w += STEP
    return ws


def build(out):
    rpr = ('<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
           '<w:spacing w:val="%d"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
           % (ASCII_FACE, FONT, ASCII_FACE, CS, SZ, SZ))
    tbls = []
    for w in widths():
        cw = int(round(w))
        tbls.append(
            '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/>'
            '<w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
            '<w:p><w:pPr><w:jc w:val="left"/><w:rPr>%s</w:rPr></w:pPr>'
            '<w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>'
            '</w:tc></w:tr></w:tbl><w:p/>' % (cw, cw, cw, rpr, rpr, TEXT))
    sectpr = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
              'w:header="851" w:footer="992" w:gutter="0"/>'
              '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>' % ("".join(tbls), sectpr))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              '<w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>'
              % (ASCII_FACE, FONT, ASCII_FACE, SZ))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:balanceSingleByteDoubleByteWidth/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
                '</w:compat></w:settings>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    docrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", docrels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)


def main():
    os.makedirs(OUT, exist_ok=True)
    import _cw_law as L
    import fitz
    docx = os.path.join(OUT, "cw_autospace.docx")
    if not os.path.exists(docx):
        build(docx)
    cells, cur = [], None
    for pg in fitz.open(L.export(docx)):
        rows = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                cs = [(c["c"], c["bbox"][0]) for sp in ln["spans"] for c in sp["chars"]]
                if cs:
                    rows.append((round(ln["bbox"][1], 1), cs))
        rows.sort()
        for _, cs in rows:
            if cs[0][0] == TEXT[0]:
                cur = []
                cells.append(cur)
            if cur is not None:
                cur.append(cs)
    ws = widths()
    print("%d cells / %d widths;  %r at %.1fpt, w:spacing %d, balanceSBDB, ascii %s"
          % (len(cells), len(ws), TEXT, SZ / 2.0, CS, ASCII_FACE))
    if len(cells) != len(ws):
        print("cell count mismatch -- stopping")
        return
    print("    %8s %6s %9s %10s %8s" % ("inner", "held", "CJK adv", "gap kana-H", "line w"))
    prev = None
    for w, lines in zip(ws, cells):
        inner = w / 20.0 - 10.8
        first = lines[0]
        if len(first) < len(TEXT):
            key = ("wrap", len(first))
            if key != prev:
                print("    %8.2f %6s %9s %10s   wrapped after %d"
                      % (inner, "no", "-", "-", len(first)))
                prev = key
            continue
        x = [p[1] for p in first]
        ch = [p[0] for p in first]
        i = ch.index("H")
        cjk = x[2] - x[1]
        gap = x[i] - x[i - 1] - cjk
        key = ("hold", round(gap, 2))
        if key != prev:
            print("    %8.2f %6s %9.3f %10.3f %8.2f"
                  % (inner, "yes", cjk, gap, x[-1] - x[0]))
            prev = key


if __name__ == "__main__":
    main()
