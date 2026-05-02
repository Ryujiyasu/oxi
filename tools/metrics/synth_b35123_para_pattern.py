"""Synthetic minimal repro of b35123 paras 27-29 pattern.

3 paragraphs in 1 cell (form-table style) with key transitions:
  p1 (sz=21, no spacing)        ← like b35123 para 27
  p2 (sz=21, no spacing)        ← like b35123 para 28
  p3 (sz=18, afterLines=20 after=70 ind:leftChars=100 left=424 hanging=197)
                                 ← like b35123 para 29 (size+spacing transition)

Plus 2 control variants:
  V_C1: same but p3 sz=21 (no font transition)
  V_C2: same but p3 no spacing attrs (no spaceAfter)

Measure Oxi y for each para via layout dump. Compare to expected:
  Word: tight inline, p3_y = p2_y + line_h_p2 ≈ +14.5pt
  Oxi: ?
"""
import os
import sys
import zipfile
import shutil
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/b35123_synth_docs")

CTYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>
<w:sz w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"""


def gen_doc(p1_pPr, p2_pPr, p3_pPr, p3_rPr_sz):
    """Build document.xml with one table cell containing 3 paragraphs."""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
<w:tblPr><w:tblW w:w="9000" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>物理的管理措置</w:t></w:r></w:p>
</w:tc>
<w:tc>
<w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>
<w:p><w:pPr>{p1_pPr}</w:pPr><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>P1テキスト第一段落（10.5pt）</w:t></w:r></w:p>
<w:p><w:pPr>{p2_pPr}</w:pPr><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>P2テキスト第二段落（10.5pt）</w:t></w:r></w:p>
<w:p><w:pPr>{p3_pPr}</w:pPr><w:r><w:rPr><w:sz w:val="{p3_rPr_sz}"/></w:rPr><w:t>P3サブテキスト第三段落（{p3_rPr_sz}/2pt）</w:t></w:r></w:p>
</w:tc>
</w:tr>
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body></w:document>'''


def write_docx(path, p1_pPr, p2_pPr, p3_pPr, p3_sz=18):
    tmp = tempfile.mkdtemp(prefix="synth_b35_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/document.xml", gen_doc(p1_pPr, p2_pPr, p3_pPr, p3_sz)),
        ]
        for relpath, content in files:
            full = os.path.join(tmp, relpath.replace("/", os.sep))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "w", encoding="utf-8") as f:
                f.write(content)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


VARIANTS = [
    # (label, p1_pPr, p2_pPr, p3_pPr, p3_sz)
    ("V_b35_actual",
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     '<w:spacing w:afterLines="20" w:after="70"/><w:ind w:leftChars="100" w:left="424" w:hangingChars="100" w:hanging="197"/>',
     18),
    ("V_C1_p3_sz_same",  # same size as p1/p2 (no font transition)
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     '<w:spacing w:afterLines="20" w:after="70"/><w:ind w:leftChars="100" w:left="424" w:hangingChars="100" w:hanging="197"/>',
     21),
    ("V_C2_p3_no_spacing",  # no afterLines/after on p3
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     '<w:ind w:leftChars="100" w:left="424" w:hangingChars="100" w:hanging="197"/>',
     18),
    ("V_C3_p3_no_indent",  # no different indent
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     "<w:ind w:left=\"197\" w:hangingChars=\"100\" w:hanging=\"197\"/>",
     '<w:spacing w:afterLines="20" w:after="70"/>',
     18),
    ("V_C4_minimal_p3",  # only sz transition
     "",
     "",
     "",
     18),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for label, p1, p2, p3, sz in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, p1, p2, p3, sz)
        print(f"Wrote {label}.docx", flush=True)


if __name__ == "__main__":
    main()
