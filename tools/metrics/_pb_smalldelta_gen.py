# -*- coding: utf-8 -*-
"""Small-|Δ| empty cell ¶ pricing (the S610/S636/S1231 wall, R24).

administrative__00018048: a size-less EMPTY ListParagraph ¶ between bullet
cells prices 1.15pt shorter in Oxi than Word — chain (Normal sz=24 → 12pt)
vs engine default 11.0, |Δ|=1.0 sits under S1231's ≥1.5 guard.  kyodoken
(Normal sz=21 → 10.5, Δ=0.5) was the S610-era falsifier — but that call was
an SSIM lottery, never a COM row-height measurement.  This probe prices the
empty ¶ directly per (docDefaults sz, Normal sz, ascii font) arm.

Table: [label][cell with ONE size-less empty ¶] / [MARKER][empty] — row
height gap read via COM Information(6) with collapsed starts.
"""
import os, sys, time, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_smalldelta"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

def styles(dd_sz, normal_sz, ascii_font):
    dd = f'<w:sz w:val="{dd_sz}"/>' if dd_sz else ''
    nm = f'<w:rPr><w:sz w:val="{normal_sz}"/></w:rPr>' if normal_sz else ''
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii=\"""" + ascii_font + """\" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi=\"""" + ascii_font + """\"/>
""" + dd + """<w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>""" + nm + """</w:style>
</w:styles>""")

DOC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>TOP</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>
<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>LB1</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>MARKER</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"""

# tag, docDefaults sz halves, Normal sz halves, ascii font
ARMS = [
    ("admin_dd22_n24_arial", 22, 24, "Arial"),
    ("kyodo_dd22_n21_century", 22, 21, "Century"),
    ("noddsz_n24_arial", None, 24, "Arial"),
    ("dd22_nonone_arial", 22, None, "Arial"),
    ("big_dd22_n28_arial", 22, 28, "Arial"),
    ("dd22_n24_century", 22, 24, "Century"),
]

if __name__ == "__main__":
    import win32com.client
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
        for tag, dd, nm, font in ARMS:
            p = os.path.join(OUT, tag + ".docx")
            with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", CT)
                z.writestr("_rels/.rels", RELS)
                z.writestr("word/_rels/document.xml.rels", DRELS)
                z.writestr("word/styles.xml", styles(dd, nm, font))
                z.writestr("word/document.xml", DOC)
            d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
            try:
                ys = {}
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    t = ''.join(ch for ch in r.Text if ch.isalnum())
                    cr = d.Range(r.Start, r.Start)
                    if t in ('LB1', 'MARKER'):
                        ys[t] = cr.Information(6)
                gap = ys.get('MARKER', 0) - ys.get('LB1', 0)
                print(f"{tag:26s} empty_row_h={gap:7.3f}")
            finally:
                retry(lambda: d.Close(False))
    finally:
        word.Quit()
