# -*- coding: utf-8 -*-
"""S568 legacy oikomi discriminator: charSpace presence vs run-size regime.

parttime_000856314 (cm11, linesAndChars cs=-2156, jc=left, body 8pt vs
docDefaults 12pt): Word breaks 第25条 at NATURAL mark widths (3 lines) but
Oxi's s568_legacy_oikomi (derived on harassmanual: cm11, linesAndChars NO cs,
body 10.5 = default) packs 2 lines with cap-6.0 phantom compression.  Which
property kills the oikomi — the grid's charSpace, or the run being off the
default-size regime?

Arms (all cm11 + compressPunctuation + kern=2 + jc=left + MS Mincho,
linesAndChars linePitch=335):
  C: no cs, sz=21 (default)   — harassmanual shape, expect oikomi (packs)
  A: cs=-2156, sz=24=default  — cs-regime, layer-1 billing question
  B: no cs, runs 16 (8pt), docDefaults 24 — off-default, no cs
  D: cs=-2156, runs 16, docDefaults 24    — parttime shape, expect natural

Para = (4x様 + 、)x5 + 50x様 (5 marks amplify the width-model separation).
Margins tuned per fs so natural line1 leaves slack < 1 char but within the
5-mark oikomi capacity; line-1 char count discriminates the models.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_oikomi"
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

def styles(default_sz):
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="\uff2d\uff33 \u660e\u671d" w:eastAsia="\uff2d\uff33 \u660e\u671d" w:hAnsi="\uff2d\uff33 \u660e\u671d"/>
<w:kern w:val="2"/><w:sz w:val=\"""" + str(default_sz) + """\"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/></w:pPr></w:style>
</w:styles>""")

def doc(run_sz, cs, mar):
    csattr = f' w:charSpace="{cs}"' if cs is not None else ''
    text = ("\u69d8" * 4 + "\u3001") * 5 + "\u69d8" * 50
    szel = f'<w:sz w:val="{run_sz}"/>' if run_sz else ''
    body = (f'<w:p><w:pPr><w:rPr>{szel}</w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/>{szel}</w:rPr><w:t>{text}</w:t></w:r></w:p>')
    return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>""" + body + f"""
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="697" w:right="{mar}" w:bottom="697" w:left="{mar}" w:header="45" w:footer="142" w:gutter="0"/>
<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="335"{csattr}/></w:sectPr></w:body></w:document>""")

# tag, default_sz, run_sz(None=inherit default), cs, margin_tw
# margins: fs10.5 -> content 531.95 (mar 634); fs8 -> content 524.95 (mar 704);
# fs12 (arm A) -> content 531.95 (mar 634)
ARMS = [
    ("C_nocs_def105", 21, None, None, 634),
    ("A_cs_def12",    24, None, -2156, 634),
    ("B_nocs_run8",   24, 16,   None, 704),
    ("D_cs_run8",     24, 16,   -2156, 704),
]

def build(tag, dsz, rsz, cs, mar):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(dsz))
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc(rsz, cs, mar))
    return p

if __name__ == "__main__":
    import time
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
        for tag, dsz, rsz, cs, mar in ARMS:
            p = build(tag, dsz, rsz, cs, mar)
            pdf = p[:-5] + ".pdf"
            d = retry(lambda: word.Documents.Open(os.path.abspath(p), ReadOnly=True))
            try:
                retry(lambda: d.SaveAs2(os.path.abspath(pdf), FileFormat=17))
            finally:
                retry(lambda: d.Close(False))
            doc_ = fitz.open(pdf)
            print(f"== {tag}")
            for b in doc_[0].get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    chars = [(c["c"], c["origin"][0]) for s in l["spans"] for c in s["chars"]]
                    if len(chars) < 3:
                        continue
                    xs = [x for _, x in chars]
                    marks = [(i, round(xs[i+1]-xs[i], 2)) for i, (c, _) in enumerate(chars[:-1]) if c == "\u3001"]
                    print(f"  n={len(chars)} x0={xs[0]:.1f} 、adv={marks}")
    finally:
        word.Quit()
