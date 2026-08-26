# -*- coding: utf-8 -*-
"""Size-less empty cell ¶ height — the RICH input space (_pb_marksz_gen round 2).

Round 1 (6 arms, plain package): the style CHAIN wins everywhere. But real docs
split: kyotei (chain 8pt ✓) vs tokumei_08_07/kyodoken08 (COM gaps 13.5 ≈ NOT
their Normal 10.5). Missing inputs: settings.xml adjustLineHeightInTable,
sectPr docGrid (tokumei/kyodoken08: linesAndChars 272tw=13.6 ≈ the observed
13.5!), and per-para snapToGrid=0. Arms cross those over the dd22_n21 base
(docDefaults 11pt, Normal 10.5pt — the tokumei shape).

Read: 4 size-less empties in a bordered cell; row height from the COM y of the
ラベル/MARKER rows (collapsed-start Info6, R30).
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_marksz2"
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

def settings(adj):
    a = '<w:adjustLineHeightInTable/>' if adj else ''
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>""" + a + """<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>""")

def styles(dd_sz, normal_sz):
    dd = f'<w:sz w:val="{dd_sz}"/>' if dd_sz else ''
    ns = f'<w:sz w:val="{normal_sz}"/>' if normal_sz else ''
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="Century"/>
<w:kern w:val="2"/>""" + dd + """<w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr>""" + ns + """</w:rPr></w:style>
</w:styles>""")

def doc(grid_pitch, snap0):
    sn = '<w:snapToGrid w:val="0"/>' if snap0 else ''
    empt = (f'<w:p><w:pPr>{sn}<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr></w:p>') * 4
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6000"/></w:tblGrid>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
           '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>\u30e9\u30d9\u30eb</w:t></w:r></w:p></w:tc>'
           '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>' + empt + '</w:tc></w:tr>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
           '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>MARKER</w:t></w:r></w:p></w:tc>'
           '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>')
    g = f'<w:docGrid w:type="linesAndChars" w:linePitch="{grid_pitch}"/>' if grid_pitch else '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid_pitch == 0 else ''
    if grid_pitch is None:
        g = ''
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>TOP</w:t></w:r></w:p>""" + tbl + """
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>""" + g + """</w:sectPr></w:body></w:document>""")

ARMS = [
    # (tag, dd_sz, normal_sz, adjustLineHeightInTable, grid_pitch, snap0)
    ("base",              22, 21, False, None, False),  # round-1 control -> chain 10.5 (12.75)
    ("adj",               22, 21, True,  None, False),
    ("grid272",           22, 21, False, 272,  False),
    ("adj_grid272",       22, 21, True,  272,  False),
    ("adj_grid272_snap0", 22, 21, True,  272,  True),   # FULL tokumei shape
    ("grid272_snap0",     22, 21, False, 272,  True),
    ("adj_snap0",         22, 21, True,  None, True),
    ("kyotei_shape",      None, 16, False, 230, False), # dd absent, Normal 8pt, grid 230, no adj
    ("kyotei_adj",        None, 16, True,  230, False),
    ("adj_grid360",       22, 21, True,  360,  False),  # pitch scaling check (18pt)
]

def build(tag, dd, ns, adj, gp, sn):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(dd, ns))
        z.writestr("word/settings.xml", settings(adj))
        z.writestr("word/document.xml", doc(gp, sn))
    return p

def measure(path):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        try:
            ys = {}
            for i in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(i).Range
                t = ''.join(ch for ch in r.Text if ch.isalnum())
                cr = d.Range(r.Start, r.Start)
                if t in ('ラベル', 'MARKER'):
                    ys[t] = cr.Information(6)
            gap = ys.get('MARKER', 0) - ys.get('ラベル', 0)
            print(f"    row1_h={gap:.2f} -> per_empty={gap/4:.3f}")
        finally:
            d.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    for tag, dd, ns, adj, gp, sn in ARMS:
        p = build(tag, dd, ns, adj, gp, sn)
        print(f"== {tag}")
        measure(p)
