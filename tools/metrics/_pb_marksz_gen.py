# -*- coding: utf-8 -*-
"""Empty cell-paragraph ¶-mark SIZE law: rPrDefault w:sz vs the style chain.

kyotei36spec (docDefaults sz ABSENT, Normal sz=16): Word prices size-less empty
cell paras at 8pt (Normal). tokumei_08_07 (docDefaults sz=22, Normal sz=21):
Word prices them at 11pt (docDefaults), NOT Normal's 10.5. Hypothesis: the
empty ¶ mark takes rPrDefault's size when DECLARED, else the paragraph style
chain. Arms isolate (docDefaults sz?, Normal sz?, pStyle sz?).

Each arm: a 1-row 2-col table; col2 holds N=4 size-less empty paragraphs.
Row height read via Word COM (cell paragraph Info6 of the FOLLOWING marker row).
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_emptymark"
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

def doc():
    empt = '<w:p><w:pPr><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr></w:p>' * 4
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
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>TOP</w:t></w:r></w:p>""" + tbl + """
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/></w:sectPr></w:body></w:document>""")

ARMS = [
    # (tag, docDefaults sz halves, Normal sz halves)
    ("dd22_n21", 22, 21),   # tokumei shape: hypothesis -> 11pt lines
    ("ddNONE_n16", None, 16),  # kyotei shape: hypothesis -> 8pt lines
    ("dd22_nNONE", 22, None),  # -> 11pt either way (control)
    ("ddNONE_n21", None, 21),  # -> 10.5 (chain)
    ("dd16_n28", 16, 28),   # inverted big Normal: docDefaults-first -> 8pt; chain-first -> 14pt (max separation)
    ("dd28_n16", 28, 16),   # docDefaults-first -> 14pt; chain -> 8pt
]

def build(tag, dd, ns):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(dd, ns))
        z.writestr("word/document.xml", doc())
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
                y = cr.Information(6)
                if t in ('ラベル', 'MARKER'):
                    ys[t] = y
            gap = ys.get('MARKER', 0) - ys.get('ラベル', 0)
            print(f"    label_y={ys.get('ラベル'):.2f} marker_y={ys.get('MARKER'):.2f} row1_h={gap:.2f} -> per_empty={(gap)/4:.3f}")
        finally:
            d.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    for tag, dd, ns in ARMS:
        p = build(tag, dd, ns)
        print(f"== {tag}")
        measure(p)
