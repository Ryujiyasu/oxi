# -*- coding: utf-8 -*-
"""Does the docGrid charSpace advance apply to runs NOT at the document default
size?  parttime_000856314 sec1 (cs=-2156, docDefaults 12pt, body runs 8pt):
Word PDF advance = 8.04/7.92 (= fs, cs NOT applied), while b35123 (cs=-2714,
runs AT the 10.5pt default) applies it (S1210/S1216).  Hypothesis: the
additive charSpace advance (fs + cs/4096) fires only for runs at the DEFAULT
size; other sizes advance naturally.

Arms: cs in {-2156, +1453, none} x run sz in {24=default, 21, 16} halves.
docDefaults sz=24 (12pt), MS Mincho.  One doc per cs; three paragraphs of
40x \u69d8 each at the three sizes.  Read PDF char advances per paragraph.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_csgrid"
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
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="\uff2d\uff33 \u660e\u671d" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="\uff2d\uff33 \u660e\u671d"/>
<w:kern w:val="2"/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
</w:styles>"""

def para(sz, marker):
    text = marker + "\u69d8" * 40
    return (f'<w:p><w:pPr><w:rPr><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')

def doc(cs):
    csattr = f' w:charSpace="{cs}"' if cs is not None else ''
    body = para(24, "\u7532") + para(21, "\u4e59") + para(16, "\u4e19")
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>""" + body + f"""
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="697" w:right="697" w:bottom="697" w:left="697" w:header="45" w:footer="142" w:gutter="0"/>
<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="403"{csattr}/></w:sectPr></w:body></w:document>""")

ARMS = [("csm2156", -2156), ("csp1453", 1453), ("csnone", None)]

def build(tag, cs):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc(cs))
    return p

if __name__ == "__main__":
    import win32com.client, fitz
    from collections import Counter
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for tag, cs in ARMS:
            p = build(tag, cs)
            pdf = p[:-5] + ".pdf"
            d = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
            try:
                d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
            finally:
                d.Close(False)
            doc_ = fitz.open(pdf)
            print(f"== {tag} (cs={cs})")
            for b in doc_[0].get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    chars = [(c["c"], c["origin"][0]) for s in l["spans"] for c in s["chars"]]
                    txt = "".join(c for c, _ in chars)
                    if txt[:1] in "\u7532\u4e59\u4e19" and "\u69d8" in txt:
                        xs = [x for _, x in chars]
                        adv = [round(xs[i+1]-xs[i], 3) for i in range(len(xs)-1)]
                        cc = Counter(adv).most_common(3)
                        mean = (xs[-1]-xs[1])/(len(xs)-2)
                        fs = {round(s["size"], 2) for l2 in b.get("lines", []) for s in l2["spans"]}
                        print(f"  {txt[0]} n={len(chars)} mean_adv={mean:.4f} top={cc}")
    finally:
        word.Quit()
