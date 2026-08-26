# -*- coding: utf-8 -*-
"""Negative-tblInd table absorption law (S1239 / R21).

technical__0009d767 (compat undeclared, tblInd=-162, TableGrid style
cellMar.left=108tw): Word draws BOTH tables' borders 5.76pt left of Oxi.
Law probed here: a LEGACY (compat <= 14 OR undeclared) non-floating table's
grid_x = margin + tblInd - cellMar.left(effective, style-inherited), applied
to the WHOLE table (borders included), for ANY tblInd sign.  cm15 absorbs
nothing on any arm.

Arms: one doc per compat {none, 15}; five tables each:
tblInd {-162, -500, 0, +162} with style cellMar 108, plus -162 with a DIRECT
tblCellMar left=72 (absorbs its own 3.6, not the style 5.4).
Read the left border x from PDF drawings (thin tall rects).

Measured 2026-08-27 (margin 36pt, border sz12):
  nocm: -162 -> 21.72  -500 -> 4.92  0 -> 29.88  +162 -> 38.04  -162/72 -> 23.52
        (= margin + ind - cellMar - border/2, +-0.1)
  cm15: 27.96 / 11.04 / 36.0 / 44.16 / 27.96 (= margin + ind, no absorption)
"""
import os, sys, time, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_tblneg"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

def settings(cm):
    c = (f'<w:compat><w:compatSetting w:name="compatibilityMode" '
         f'w:uri="http://schemas.microsoft.com/office/word" w:val="{cm}"/></w:compat>'
         if cm else '<w:compat/>')
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">""" + c + "</w:settings>")

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>
<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
</w:style>
</w:styles>"""

def tbl(ind, cellmar):
    cm = (f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="{cellmar}" w:type="dxa"/>'
          f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{cellmar}" w:type="dxa"/></w:tblCellMar>'
          if cellmar is not None else '')
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="8000" w:type="dxa"/>'
            f'<w:tblInd w:w="{ind}" w:type="dxa"/>'
            f'<w:tblBorders><w:top w:val="single" w:sz="12" w:color="auto"/><w:left w:val="single" w:sz="12" w:color="auto"/>'
            f'<w:bottom w:val="single" w:sz="12" w:color="auto"/><w:right w:val="single" w:sz="12" w:color="auto"/>'
            f'<w:insideH w:val="single" w:sz="12" w:color="auto"/><w:insideV w:val="single" w:sz="12" w:color="auto"/></w:tblBorders>{cm}</w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>AA</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>BB</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>')

ARMS_T = [(-162, None), (-500, None), (0, None), (162, None), (-162, 72)]
DOCS = [("nocm", None), ("cm15", 15)]

def doc(arms):
    body = ''.join(tbl(i, c) for i, c in arms)
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>""" + body + """
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""")

if __name__ == "__main__":
    import win32com.client, fitz
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    def retry(fn, tries=10):
        for i in range(tries):
            try:
                return fn()
            except Exception:
                if i == tries - 1:
                    raise
                time.sleep(1.5)
    try:
        for tag, cm in DOCS:
            p = os.path.join(OUT, tag + ".docx")
            with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", CT)
                z.writestr("_rels/.rels", RELS)
                z.writestr("word/_rels/document.xml.rels", DRELS)
                z.writestr("word/styles.xml", STYLES)
                z.writestr("word/settings.xml", settings(cm))
                z.writestr("word/document.xml", doc(ARMS_T))
            pdf = p[:-5] + ".pdf"
            d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
            try:
                retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
            finally:
                retry(lambda: d.Close(False))
            dd = fitz.open(pdf)
            rows = {}
            for pg in dd:
                for dr in pg.get_drawings():
                    r = dr["rect"]
                    if r.height > 10 and r.width < 3:
                        rows.setdefault(round(r.y0 / 40) * 40, []).append(round(r.x0, 2))
            print("==", tag)
            for y in sorted(rows):
                print("  y~%d left_x=%.2f all=%s" % (y, min(rows[y]), sorted(set(rows[y]))[:4]))
    finally:
        word.Quit()
