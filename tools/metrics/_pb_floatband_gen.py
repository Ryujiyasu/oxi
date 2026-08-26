# -*- coding: utf-8 -*-
"""Floating-table (tblpPr vertAnchor=text) anchor + collision-displacement law.

kyotei36spec p3: doc order = [title][para X 成立年月日][FLOAT big form][para Y...].
Word renders the float at ~title_top + tblpY and pushes para X BELOW the float
(431.2); Oxi lays X at flow position (59.5) and the float after it. Questions:
  Q1 which paragraph anchors the float's tblpY (prev para? next para? own flow pos)?
  Q2 where does a COLLIDING preceding paragraph land (float bottom + gap?)
  Q3 the gap (topFromText/bottomFromText defaults)
Arms sweep tblpY and the presence/width of X. COM Info6 (collapsed start, R30)
per paragraph + table row 1 cell paragraph Info6 for the table top.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_floatanchor"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

def para(text):
    if not text:
        return "<w:p/>"
    return ("<w:p><w:r><w:rPr><w:rFonts w:ascii=\"\uff2d\uff33 \u660e\u671d\" w:eastAsia=\"\uff2d\uff33 \u660e\u671d\" w:hint=\"eastAsia\"/></w:rPr>"
            f"<w:t>{text}</w:t></w:r></w:p>")

def float_tbl(tblpy, wide=True, xspec="center"):
    w = 9000 if wide else 2000
    xs = f' w:tblpXSpec="{xspec}"' if xspec else ""
    rows = ""
    for r in range(2):
        rows += (f'<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
                 '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>'
                 '<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/></w:tcBorders></w:tcPr>'
                 + para(f"\u8868\u306e\u884c{r+1}") + "</w:tc></w:tr>")
    return (f'<w:tbl><w:tblPr><w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" w:horzAnchor="margin"{xs} w:tblpY="{tblpy}"/>'
            f'<w:tblW w:w="{w}" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="{w}"/></w:tblGrid>{rows}</w:tbl>')

DOC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{body}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/></w:sectPr></w:body></w:document>"""

ARMS = []
# Q1/Q2: X present, tblpY sweep, wide (collision certain)
for ty in (0, 268, 500, 1000):
    ARMS.append((f"y{ty}_withX", [para("\u30bf\u30a4\u30c8\u30eb\u884c"), para("\u30a2\u30f3\u30ab\u30fc\u524d\u6bb5\u843d\u3042\u3042"), float_tbl(ty), para("\u5f8c\u7d9a\u6bb5\u843d\u4e00"), para("\u5f8c\u7d9a\u6bb5\u843d\u4e8c")]))
# anchor isolation: X absent
ARMS.append(("y268_noX", [para("\u30bf\u30a4\u30c8\u30eb\u884c"), float_tbl(268), para("\u5f8c\u7d9a\u6bb5\u843d\u4e00"), para("\u5f8c\u7d9a\u6bb5\u843d\u4e8c")]))
# narrow float (text fits beside)
ARMS.append(("y268_narrow", [para("\u30bf\u30a4\u30c8\u30eb\u884c"), para("\u30a2\u30f3\u30ab\u30fc\u524d\u6bb5\u843d\u3042\u3042"), float_tbl(268, wide=False, xspec="right"), para("\u5f8c\u7d9a\u6bb5\u843d\u4e00"), para("\u5f8c\u7d9a\u6bb5\u843d\u4e8c")]))
# two preceding paragraphs (which one displaces?)
ARMS.append(("y268_twoX", [para("\u30bf\u30a4\u30c8\u30eb\u884c"), para("\u524d\u6bb5\u843d\u7532"), para("\u524d\u6bb5\u843d\u4e59"), float_tbl(268), para("\u5f8c\u7d9a\u6bb5\u843d\u4e00")]))

def build(tag, blocks):
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", DOC.format(body="".join(blocks)))
    return path

def measure(path):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        try:
            for i in range(1, doc.Paragraphs.Count + 1):
                rng = doc.Paragraphs(i).Range
                cr = doc.Range(rng.Start, rng.Start)
                y = cr.Information(6)   # wdVerticalPositionRelativeToPage
                x = cr.Information(5)
                intbl = rng.Information(12)  # wdWithInTable
                txt = rng.Text.strip()[:14]
                print(f"    para{i}: y={y:7.2f} x={x:6.1f} tbl={bool(intbl)} {txt!r}")
            if doc.Tables.Count:
                t = doc.Tables(1)
                cr = doc.Range(t.Range.Start, t.Range.Start)
                print(f"    table1 start Info6 y={cr.Information(6):7.2f}")
        finally:
            doc.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    for tag, blocks in ARMS:
        p = build(tag, blocks)
        print(f"== {tag}")
        measure(p)
