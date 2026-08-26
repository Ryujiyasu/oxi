# -*- coding: utf-8 -*-
"""rightChars unit + hanging-precedence law (the right-side S1214).

parttime body paras: ind leftChars=100 left=400 rightChars=-52 right=-125
hanging=160.  Word PDF: LEFT uses the twip (20pt — hanging blocks the left
*Chars per S1214) but the RIGHT edge = base + 0.52 x 9.0 (NOT the twip -6.25,
NOT 8pt-unit -4.16) => (a) hanging blocks only the LEFT-side *Chars,
(b) the rightChars unit at fs8/kern2 = 9.0 = fs + kern_pt?

Arms: fs {16,21} halves x kern {0,2,4} halves x rightChars {-52,-100} with
hanging, plus a no-hanging control.  Fill paragraphs with 様 so every line is
full; the max line-end x reads the effective right boundary.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_rchars"
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

def styles(kern):
    k = f'<w:kern w:val="{kern}"/>' if kern else ''
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="\uff2d\uff33 \u660e\u671d" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="\uff2d\uff33 \u660e\u671d"/>
""" + k + """<w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
</w:styles>""")

def para(sz, rchars, rtw, hanging, marker):
    h = f' w:hanging="160"' if hanging else ''
    text = marker + "\u69d8" * 90
    return (f'<w:p><w:pPr><w:ind w:leftChars="100" w:left="400" w:rightChars="{rchars}" w:right="{rtw}"{h}/>'
            f'<w:rPr><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>')

# each arm its own doc (kern is docDefaults-level)
ARMS = [
    # tag, kern, sz, rchars, rtw, hanging
    ("k2_f16_rm52_h", 2, 16, -52, -125, True),    # parttime shape
    ("k0_f16_rm52_h", 0, 16, -52, -125, True),
    ("k4_f16_rm52_h", 4, 16, -52, -125, True),
    ("k2_f21_rm52_h", 2, 21, -52, -125, True),
    ("k2_f16_rm100_h", 2, 16, -100, -240, True),
    ("k2_f16_rm52_noh", 2, 16, -52, -125, False),
    ("k2_f16_rp100_h", 2, 16, 100, 240, True),
]

def build(tag, kern, sz, rchars, rtw, hanging):
    p = os.path.join(OUT, tag + ".docx")
    body = para(sz, rchars, rtw, hanging, "\u7532")
    doc = ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>""" + body + """
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="697" w:right="697" w:bottom="697" w:left="697" w:header="45" w:footer="142" w:gutter="0"/>
<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="415"/></w:sectPr></w:body></w:document>""")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(kern))
        z.writestr("word/document.xml", doc(kern) if callable(doc) else doc)
    return p

if __name__ == "__main__":
    import win32com.client, fitz
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for tag, kern, sz, rchars, rtw, hanging in ARMS:
            p = build(tag, kern, sz, rchars, rtw, hanging)
            pdf = p[:-5] + ".pdf"
            d = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
            try:
                d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            finally:
                d.Close(False)
            doc_ = fitz.open(pdf)
            fs = sz / 2.0
            base_right = 595.35 - 34.85
            ends = []
            x0s = []
            for b in doc_[0].get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    chars = [(c["c"], c["origin"][0]) for s in l["spans"] for c in s["chars"]]
                    if len(chars) < 20:
                        continue
                    ends.append(chars[-1][1] + fs)
                    x0s.append(chars[0][1])
            mend = max(ends) if ends else 0
            ind_r = base_right - mend
            unit = (ind_r / (-rchars / 100.0)) if rchars else 0
            print(f"{tag}: max_end={mend:.2f} eff_right_ind={ind_r:+.3f} "
                  f"unit={-unit if rchars<0 else unit:.3f} x0={min(x0s):.2f}")
    finally:
        word.Quit()
