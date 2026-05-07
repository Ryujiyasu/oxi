"""Generate b5f706-style cell line-height variants for COM measurement.

Phase β step 3 diagnosis (Day 28): L7/L8 minimal repro shows Word does NOT
snap cell paragraphs even when snap=1 (delta_y = 13.5pt = natural for
MS Mincho 10.5pt + linePitch=330 = 16.5pt).

But b5f706 PASS->FAIL when Step 2 ships. b5f706 differs from L7 in:
  - Font: MS Gothic 10pt (sz=20) vs L7's MS Mincho 10.5pt (sz=21)
  - linePitch: 360 (18pt) vs L7's 330 (16.5pt)
  - Page orientation: landscape vs portrait

This script generates 5 variants to isolate which factor flips Word's
cell-snap behavior:

  V1: L7-as-is              MS Mincho 10.5pt + pitch=330 portrait    [control = L7]
  V2: pitch=360             MS Mincho 10.5pt + pitch=360 portrait
  V3: MS Gothic font        MS Gothic 10pt   + pitch=330 portrait
  V4: b5f706-mimic          MS Gothic 10pt   + pitch=360 portrait
  V5: b5f706-mimic landscape MS Gothic 10pt   + pitch=360 landscape

Output: tools/golden-test/repros/grid_snap/b5f706_V{1..5}.docx
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

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

# styles.xml: rPrDefault sets the EastAsia font + size for the whole doc.
def styles_for(font_name_xml_escaped: str, sz_hp: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="{font_name_xml_escaped}" w:eastAsia="{font_name_xml_escaped}" w:hAnsi="{font_name_xml_escaped}"/><w:sz w:val="{sz_hp}"/><w:szCs w:val="{sz_hp}"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""


def make_cell_doc(label: str, font_xml: str, sz_hp: int,
                  grid_pitch_tw: int, snap: bool = True,
                  landscape: bool = False) -> None:
    """Build a 1-cell table with 6 paragraphs of given font/size, snap setting."""
    paras_xml = ""
    for i in range(1, 7):
        snap_xml = "" if snap else '<w:snapToGrid w:val="0"/>'
        paras_xml += (
            f'<w:p><w:pPr>{snap_xml}</w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr>'
            f'<w:t>{label} line {i}</w:t></w:r></w:p>'
        )

    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc>
<w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{paras_xml}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""

    # Page setup
    if landscape:
        pg_sz = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    else:
        pg_sz = '<w:pgSz w:w="11906" w:h="16838"/>'
    grid = f'<w:docGrid w:type="lines" w:linePitch="{grid_pitch_tw}"/>'

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
{pg_sz}
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
{grid}
</w:sectPr>
</w:body>
</w:document>"""

    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    styles = styles_for(font_xml, sz_hp)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/settings.xml", SETTINGS)
        zf.writestr("word/styles.xml", styles)
        zf.writestr("word/document.xml", document_xml)
    print(f"  wrote {out_path}")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # Font names. Use full Japanese name for proper resolution.
    MS_MINCHO = "ＭＳ 明朝"
    MS_GOTHIC = "ＭＳ ゴシック"

    # V1: control = L7-style (MS Mincho 10.5pt sz=21, pitch=330=16.5pt, portrait, snap=1)
    make_cell_doc("b5f706_V1_control",     MS_MINCHO, 21, 330, snap=True,  landscape=False)
    # V2: pitch only differs (pitch=360=18pt)
    make_cell_doc("b5f706_V2_pitch360",    MS_MINCHO, 21, 360, snap=True,  landscape=False)
    # V3: font only differs (MS Gothic 10pt sz=20)
    make_cell_doc("b5f706_V3_gothic10",    MS_GOTHIC, 20, 330, snap=True,  landscape=False)
    # V4: b5f706 mimic (MS Gothic 10pt + pitch=360)
    make_cell_doc("b5f706_V4_mimic",       MS_GOTHIC, 20, 360, snap=True,  landscape=False)
    # V5: full mimic with landscape
    make_cell_doc("b5f706_V5_landscape",   MS_GOTHIC, 20, 360, snap=True,  landscape=True)
    # V6: V4 control with snap=0 (compare V4 vs V6 to see if snap setting matters at all)
    make_cell_doc("b5f706_V6_mimic_snap0", MS_GOTHIC, 20, 360, snap=False, landscape=False)


if __name__ == "__main__":
    main()
