"""Session 113 — parametric grid generator to COM-measure the correct
compression multiplier for '．','，','。' under the S112-isolated 4-way
AND trigger (jc=both + balanceSBDB + compressPunctuation + cs<0).

Grid dimensions:
  font      : MS Mincho, MS Gothic, Meiryo, Yu Mincho, Yu Gothic
  font size : 9, 10.5, 12, 14 pt    (sz half-points: 18, 21, 24, 28)
  cs (tw)   : -5, -9, -15, -20      (negative only; required by trigger)
  punct     : '．', '，', '。'      (compressed set per S112 phase 3)

Reduced grid (focus): MS Mincho is the 15076df font, so prioritize it.
  Full sweep: MS Mincho × all sizes × all cs × all puncts = 4×4×3 = 48
  Cross-font: each other font × fs=10.5 × cs=-9 × '．' = 4
  Total = 52 variants.

All other gates (jc=both, balanceSBDB, compressPunctuation) held constant.
Output: tools/metrics/yakumono_grid/variants/<name>.docx
"""
import os
import sys
import io
import itertools
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/yakumono_grid/variants"))
os.makedirs(OUT_DIR, exist_ok=True)


SETTINGS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    b'<w:compat>\n'
    b'<w:balanceSingleByteDoubleByteWidth/>\n'
    b'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\n'
    b'</w:compat>\n'
    b'<w:characterSpacingControl w:val="compressPunctuation"/>\n'
    b'</w:settings>\n'
)


def build_styles(font_name: str) -> bytes:
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}" w:eastAsia="{font_name}" w:cs="{font_name}"/>'
        '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
        '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault></w:docDefaults>\n'
        '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
        '<w:name w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:jc w:val="both"/></w:pPr></w:style>\n'
        '</w:styles>\n'
    )
    return s.encode('utf-8')


def build_document(font_name: str, sz_hp: int, cs_tw: int, punct: str,
                   tcw_dxa: int = 1968, body_text_tail: str = "提供を受けた匿名データの名称") -> bytes:
    """Build a document with text '１<punct><body_text_tail>' inside a cell.

    sz_hp     : font size in half-points (sz attribute)
    cs_tw     : character_spacing in twips (negative)
    """
    text = f"１{punct}{body_text_tail}"
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '<w:body>\n'
        f'<w:tbl><w:tblPr><w:tblW w:w="{tcw_dxa}" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{tcw_dxa}"/></w:tblGrid>'
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{tcw_dxa}" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr>'
        '<w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
        '<w:adjustRightInd w:val="0"/>'
        '<w:spacing w:line="240" w:lineRule="exact"/>'
        '<w:ind w:left="215" w:right="76" w:hanging="192"/>'
        '</w:pPr>'
        '<w:r><w:rPr>'
        f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}" w:eastAsia="{font_name}" w:cs="{font_name}" w:hint="eastAsia"/>'
        f'<w:sz w:val="{sz_hp}"/><w:szCs w:val="{sz_hp}"/>'
        f'<w:spacing w:val="{cs_tw}"/><w:kern w:val="0"/>'
        '</w:rPr>'
        f'<w:t>{text}</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>\n'
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="851" w:right="1134" w:bottom="567" w:left="1134"'
        ' w:header="851" w:footer="567" w:gutter="0"/>'
        '<w:docGrid w:type="lines" w:linePitch="336"/>'
        '</w:sectPr>\n'
        '</w:body></w:document>\n'
    )
    return s.encode('utf-8')


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
    '<Default Extension="xml" ContentType="application/xml"/>\n'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
    '</Types>\n'
).encode('utf-8')

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
    '</Relationships>\n'
).encode('utf-8')

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
    '</Relationships>\n'
).encode('utf-8')


def write_docx(name, font_name, sz_hp, cs_tw, punct, tcw_dxa=1968):
    out_path = os.path.join(OUT_DIR, f"{name}.docx")
    styles = build_styles(font_name)
    document = build_document(font_name, sz_hp, cs_tw, punct, tcw_dxa)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', SETTINGS_XML)


def font_slug(name):
    # Map font display name to ASCII slug for filename
    return {
        "ＭＳ 明朝": "msmincho",
        "ＭＳ ゴシック": "msgothic",
        "メイリオ": "meiryo",
        "游明朝": "yumincho",
        "游ゴシック": "yugothic",
    }.get(name, name.replace(" ", "_"))


def punct_slug(c):
    return {
        '．': "dotF",  # FULLWIDTH FULL STOP
        '，': "commaF",
        '。': "kuten",
    }[c]


def main():
    rows = []
    # MAIN GRID: MS Mincho × {9, 10.5, 12, 14} × {-5,-9,-15,-20} × {．,，,。}
    fonts_main = ["ＭＳ 明朝"]
    sizes_hp = [18, 21, 24, 28]  # 9, 10.5, 12, 14pt
    cs_values = [-5, -9, -15, -20]
    puncts = ['．', '，', '。']
    for f, sz, cs, p in itertools.product(fonts_main, sizes_hp, cs_values, puncts):
        name = f"g_{font_slug(f)}_sz{sz}_cs{cs}_{punct_slug(p)}"
        write_docx(name, f, sz, cs, p)
        rows.append((name, f, sz / 2.0, cs, p))

    # CROSS-FONT: each other font × fs=10.5 × cs=-9 × '．'
    fonts_cross = ["ＭＳ ゴシック", "メイリオ", "游明朝", "游ゴシック"]
    for f in fonts_cross:
        name = f"g_{font_slug(f)}_sz21_cs-9_dotF"
        write_docx(name, f, 21, -9, '．')
        rows.append((name, f, 10.5, -9, '．'))

    # CONTROL: jc=both removed (= S112 v28 equivalent at each fs)
    # to confirm un-triggered baseline at all sizes (no compression)
    # Already covered by v28 at 10.5pt. Add fs=9,12,14 quick controls.
    # Skip — too tangential for first measurement pass.

    print(f"Wrote {len(rows)} variants to {OUT_DIR}")
    print("\nGrid summary:")
    print(f"  MS Mincho: 4 sizes × 4 cs × 3 puncts = 48")
    print(f"  4 cross-fonts × (fs=10.5, cs=-9, '．') = 4")
    print(f"  total = {len(rows)}")


if __name__ == "__main__":
    main()
