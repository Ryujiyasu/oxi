"""V73-V78: lh formula precision investigation across fs (font size) values.

Day 29c finding: Oxi's body lh exceeds Word's by ~+1pt/paragraph in
db9ca/b35123. Hypothesis: GDI height + CJK 83/64 inflate formula at
mod.rs:5099-5116 returns over-large lh for some fs values.

Generate minimal body-only repros at each fs value common in baseline:
- V73: fs=8 (sz=16)
- V74: fs=9 (sz=18)
- V75: fs=10 (sz=20)
- V76: fs=10.5 (sz=21)
- V77: fs=11 (sz=22)
- V78: fs=12 (sz=24)
- V79: fs=14 (sz=28)

Each: 6 paragraphs, MS Mincho, no docGrid (=LM0), no flag — to isolate
LM0 lh formula. Then a parallel set with linePitch=360 (LM1) and
adjustLineHeightInTable flag to test cell-style snap.

Output: tools/golden-test/repros/grid_snap/lh_fs_*.docx
"""
from __future__ import annotations

import os
import sys
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS_PLAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

SETTINGS_FLAG = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:adjustLineHeightInTable/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""


def styles_for(font_xml: str, sz_hp: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="{font_xml}" w:eastAsia="{font_xml}" w:hAnsi="{font_xml}"/><w:sz w:val="{sz_hp}"/><w:szCs w:val="{sz_hp}"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""


def write_docx(label: str, document_xml: str, settings_xml: str, styles_xml: str):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/settings.xml", settings_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/document.xml", document_xml)
    print(f"  wrote {out_path}")


def make_doc(body: str, grid_pitch_tw: int = 0) -> str:
    grid = (f'<w:docGrid w:type="lines" w:linePitch="{grid_pitch_tw}"/>'
            if grid_pitch_tw > 0 else '<w:docGrid w:linePitch="0"/>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
{grid}
</w:sectPr>
</w:body>
</w:document>"""


def make_para(label: str, line: int, sz_hp: int) -> str:
    return (f'<w:p><w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr>'
            f'<w:t>{label} L{line}</w:t></w:r></w:p>')


def make_paras(label: str, sz_hp: int, count: int = 6) -> str:
    return "".join(make_para(label, i, sz_hp) for i in range(1, count + 1))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # Test fs values: 8, 9, 10, 10.5, 11, 12, 14pt
    # sz_hp = fs × 2 (half-points)
    test_fs = [
        (8.0, 16, 'fs8'),
        (9.0, 18, 'fs9'),
        (10.0, 20, 'fs10'),
        (10.5, 21, 'fs10p5'),
        (11.0, 22, 'fs11'),
        (12.0, 24, 'fs12'),
        (14.0, 28, 'fs14'),
    ]

    MINCHO = "ＭＳ 明朝"
    GOTHIC = "ＭＳ ゴシック"

    # Body LM0 (no docGrid): test pure lh formula
    for fs_pt, sz_hp, tag in test_fs:
        # Mincho LM0 (no grid)
        body = make_paras(f"M{tag}", sz_hp, 6)
        write_docx(f"lh_M_{tag}_LM0",
                   make_doc(body, grid_pitch_tw=0),
                   SETTINGS_PLAIN, styles_for(MINCHO, sz_hp))

    # Body LM1 with linePitch=360 (= 18pt grid)
    for fs_pt, sz_hp, tag in test_fs:
        body = make_paras(f"M{tag}_360", sz_hp, 6)
        write_docx(f"lh_M_{tag}_LM1_360",
                   make_doc(body, grid_pitch_tw=360),
                   SETTINGS_PLAIN, styles_for(MINCHO, sz_hp))


if __name__ == "__main__":
    main()
