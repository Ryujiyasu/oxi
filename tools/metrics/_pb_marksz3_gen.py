# -*- coding: utf-8 -*-
"""Empty cell ¶ + adjustLineHeightInTable grid snap — round 3: the N-sweep.

R2 found: adj+grid snaps empty cell lines to the pitch (360→18.375, 230→11.625)
but 272 gave 12.188/line (4 lines, snap on) and snap0 gave 11.25/line — neither
pitch nor natural. Hypothesis: the snap is CUMULATIVE with a phase (line i ends
at cell_top + snap_i), so per-line averages mislead. Sweep N=1..6 at pitch
{272,360} × snap0 {0,1}, plus text-line-phase arms; read TOTALS.
One Word session for all arms.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_marksz3"
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
SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:adjustLineHeightInTable/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>"""
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="Century"/>
<w:kern w:val="2"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:sz w:val="21"/></w:rPr></w:style>
</w:styles>"""

def doc(pitch, n_empty, snap0, lead_text):
    sn = '<w:snapToGrid w:val="0"/>' if snap0 else ''
    lead = ''
    if lead_text:
        lead = ('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
                '<w:t>\u3042\u3044\u3046</w:t></w:r></w:p>')
    empt = (f'<w:p><w:pPr>{sn}<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr></w:p>') * n_empty
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6000"/></w:tblGrid>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
           '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>\u30e9\u30d9\u30eb</w:t></w:r></w:p></w:tc>'
           '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>' + lead + empt + '</w:tc></w:tr>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
           '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>MARKER</w:t></w:r></w:p></w:tc>'
           '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>')
    g = f'<w:docGrid w:type="linesAndChars" w:linePitch="{pitch}"/>'
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>TOP</w:t></w:r></w:p>""" + tbl + """
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>""" + g + """</w:sectPr></w:body></w:document>""")

ARMS = []
for pitch in (272, 360):
    for snap0 in (False, True):
        for n in (1, 2, 3, 4, 6):
            ARMS.append((f"p{pitch}_s{int(snap0)}_n{n}", pitch, n, snap0, False))
ARMS.append(("p272_s0_n3_lead", 272, 3, False, True))
ARMS.append(("p272_s1_n3_lead", 272, 3, True, True))

def build(tag, pitch, n, snap0, lead):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc(pitch, n, snap0, lead))
    return p

if __name__ == "__main__":
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for tag, pitch, n, snap0, lead in ARMS:
            p = build(tag, pitch, n, snap0, lead)
            d = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
            try:
                ys = {}
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    t = ''.join(ch for ch in r.Text if ch.isalnum())
                    cr = d.Range(r.Start, r.Start)
                    if t in ('ラベル', 'MARKER'):
                        ys[t] = cr.Information(6)
                gap = ys.get('MARKER', 0) - ys.get('ラベル', 0)
                print(f"{tag}: row_h={gap:7.3f}  /n={gap/max(n,1):7.3f}")
            finally:
                d.Close(False)
    finally:
        word.Quit()
