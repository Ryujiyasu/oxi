# -*- coding: utf-8 -*-
"""LM2 (typed docGrid) baseline placement probe: where does Word put the glyph
BASELINE inside a grid row, as a function of (font size, line pitch)?

kyotei36spec p4 (8pt text in a 230tw=11.5pt linesAndChars grid) shows Word's
ink top = grid row top + ~2.35pt while Oxi draws +4.1pt (glyphs ~1.8pt too
low). Arms sweep fs x pitch; the PDF span origin (= baseline, 600dpi
quantized +-0.12) over 6 identical lines gives the per-row baseline; grid row
top = margin_top + n*pitch.

Usage: python _pb_gridbase_gen.py            # generate + Word render + measure
"""
import os, sys, shutil, zipfile, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = r"C:\tmp\pb_gridbase"
os.makedirs(OUT, exist_ok=True)

ARMS = [
    # (tag, linePitch_tw, sz_halfpt)
    ("p230_f16", 230, 16),   # kyotei shape: 8pt in 11.5pt grid
    ("p230_f21", 230, 21),   # 10.5pt in 11.5pt grid (tight)
    ("p320_f16", 320, 16),   # 8pt in 16pt grid (loose)
    ("p320_f21", 320, 21),   # 10.5pt in 16pt grid
    ("p360_f21", 360, 21),   # 10.5pt in 18pt grid (the classic)
    ("p360_f24", 360, 24),   # 12pt in 18pt grid
]

DOC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{paras}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="{pitch}"/></w:sectPr>
</w:body></w:document>"""

PARA = ("<w:p><w:pPr><w:rPr><w:rFonts w:ascii=\"\uff2d\uff33 \u660e\u671d\" w:eastAsia=\"\uff2d\uff33 \u660e\u671d\"/>"
        "<w:sz w:val=\"{sz}\"/></w:rPr></w:pPr>"
        "<w:r><w:rPr><w:rFonts w:ascii=\"\uff2d\uff33 \u660e\u671d\" w:eastAsia=\"\uff2d\uff33 \u660e\u671d\" w:hint=\"eastAsia\"/>"
        "<w:sz w:val=\"{sz}\"/></w:rPr><w:t>{text}</w:t></w:r></w:p>")

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
</w:styles>"""

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

def build(tag, pitch, sz):
    text = "\u69d8" * 30   # 様 x30 — full-height CJK, identical lines
    paras = "".join(PARA.format(sz=sz, text=text) for _ in range(8))
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/document.xml", DOC_XML.format(paras=paras, pitch=pitch))
    return path

def word_pdf(docx):
    import win32com.client
    pdf = docx[:-5] + ".pdf"
    if os.path.exists(pdf):
        return pdf
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(pdf), FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()
    return pdf

def measure(pdf, pitch_tw, sz_halfpt):
    import fitz
    doc = fitz.open(pdf)
    page = doc[0]
    d = page.get_text("dict")
    margin_top = 1134 / 20.0
    pitch = pitch_tw / 20.0
    fs = sz_halfpt / 2.0
    rows = []
    for b in d["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if "\u69d8" in s["text"]:
                    origin_y = s["origin"][1]
                    bbox_top = s["bbox"][1]
                    rows.append((origin_y, bbox_top, s["size"]))
    rows.sort()
    print(f"  margin_top={margin_top} pitch={pitch} fs={fs}")
    for i, (oy, bt, size) in enumerate(rows):
        row_top = margin_top + i * pitch
        print(f"  line{i}: baseline={oy:.2f} bbox_top={bt:.2f} row_top={row_top:.2f} "
              f"base-rowtop={oy-row_top:.3f} (=/fs {(oy-row_top)/fs:.4f}) inktop-rowtop={bt-row_top:.3f}")
    if rows:
        n = len(rows)
        gaps = [rows[i+1][0] - rows[i][0] for i in range(n-1)]
        print(f"  baseline gaps: {[f'{g:.2f}' for g in gaps]}")

if __name__ == "__main__":
    for tag, pitch, sz in ARMS:
        p = build(tag, pitch, sz)
        pdf = word_pdf(p)
        print(f"== {tag} (pitch={pitch}tw fs={sz/2}pt)")
        measure(pdf, pitch, sz)
