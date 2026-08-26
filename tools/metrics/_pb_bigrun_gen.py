# -*- coding: utf-8 -*-
"""Mid-line 、 width vs run-size-to-grid-default relation (hypothesis H, R19).

harassmanual (no-cs linesAndChars 335, Normal 10.5): its sz-24/25 item paras
render mid-line 、 at 6.0 = fs/2 UNCONDITIONALLY (even on lines with no break
pressure), while its default-size paras and parttime's 8pt paras (< default)
keep 、 natural.  Hypothesis: in a no-charSpace linesAndChars grid with
compressPunctuation, a run LARGER than the grid default size bills standalone
、。 at HALF; runs at or below the default bill natural.

Arms (one doc, cm11 + cP + kern2 + jc=left + MS Mincho, linesAndChars 335,
docDefaults sz=21): paragraphs with runs sz {21, 24, 25, 16} halves, each a
SHORT line 「様様様様、そのた様様」 (mid 、 followed by kana, line nowhere near
full = no demand).  Read the 、 advance per paragraph.  Predictions under H:
sz21 -> 10.5 natural; sz24 -> 6.0; sz25 -> 6.25 (or 6.0 if the half is of the
、 run's own size quantized); sz16 -> 8.0 natural.
A second doc with docDefaults sz=24 shifts the boundary: there sz24 = default
-> natural 12.0, sz28 -> 7.0.
"""
import os, sys, time, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_bigrun"
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
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat>
</w:settings>"""

def styles(dsz):
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="\uff2d\uff33 \u660e\u671d" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="\uff2d\uff33 \u660e\u671d"/>
<w:kern w:val="2"/><w:sz w:val=\"""" + str(dsz) + """\"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/></w:pPr></w:style>
</w:styles>""")

MARK = {21: "\u7532", 24: "\u4e59", 25: "\u4e19", 16: "\u4e01", 28: "\u621a"}

def para(sz):
    text = MARK[sz] + "\u69d8" * 4 + "\u3001" + "\u305d\u306e\u305f" + "\u69d8" * 4
    s = f'<w:sz w:val="{sz}"/>'
    return (f'<w:p><w:pPr><w:rPr>{s}</w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/>{s}</w:rPr><w:t>{text}</w:t></w:r></w:p>')

def doc(szs):
    body = ''.join(para(s) for s in szs)
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>""" + body + """
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="697" w:right="697" w:bottom="697" w:left="697" w:header="45" w:footer="142" w:gutter="0"/>
<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="335"/></w:sectPr></w:body></w:document>""")

ARMS = [("dd21", 21, (21, 24, 25, 16)), ("dd24", 24, (24, 28, 21, 16))]

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
        for tag, dsz, szs in ARMS:
            p = os.path.join(OUT, tag + ".docx")
            with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", CT)
                z.writestr("_rels/.rels", RELS)
                z.writestr("word/_rels/document.xml.rels", DRELS)
                z.writestr("word/styles.xml", styles(dsz))
                z.writestr("word/settings.xml", SETTINGS)
                z.writestr("word/document.xml", doc(szs))
            pdf = p[:-5] + ".pdf"
            d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
            try:
                retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
            finally:
                retry(lambda: d.Close(False))
            doc_ = fitz.open(pdf)
            print(f"== {tag} (docDefaults sz={dsz})")
            for b in doc_[0].get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    chars = [(c["c"], c["origin"][0]) for s in l["spans"] for c in s["chars"]]
                    if len(chars) < 5:
                        continue
                    xs = [x for _, x in chars]
                    sz = next((k for k, v in MARK.items() if v == chars[0][0]), "?")
                    j = next(i for i, (c, _) in enumerate(chars) if c == "\u3001")
                    print(f"  run_sz={sz} halves: 、adv={xs[j+1]-xs[j]:.2f} "
                          f"様adv={xs[2]-xs[1]:.2f}")
    finally:
        word.Quit()
