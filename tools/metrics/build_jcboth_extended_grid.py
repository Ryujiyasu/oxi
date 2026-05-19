"""Session 120 — extended decision-boundary grid covering fs=9 and fs=12.

S116 grid was at fs=10.5 only. S117-S119 algorithm hit empirical wall:
preserving d77a (fs=12) requires kanji savings that overfits 3a4f (fs=10.5).

Hypothesis: Word's per-(fs, cs) compression ratio differs from the
single-observation 6% / 0.6% derived at fs=10.5 cs=-9. This grid measures
fs=9 and fs=12 to triangulate.

Grid:
  fs (sz)   : 18, 21, 24 (= 9, 10.5, 12pt)
  cs (tw)   : -9
  cell_dxa  : 1500, 1800, 1968, 2200, 2500
  text_len  : 5, 9, 12, 16
  yakumono_count : 1 (single '．' after '１') OR 3 (multiple ',．' patterns)
"""
import os
import sys
import io
import itertools
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/jcboth_extended_grid/variants"))
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

STYLES_XML_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>'
    '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
    '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
    '</w:rPr></w:rPrDefault></w:docDefaults>\n'
    '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
    '<w:name w:val="Normal"/><w:qFormat/>'
    '<w:pPr><w:jc w:val="both"/></w:pPr></w:style>\n'
    '</w:styles>\n'
).encode('utf-8')

KANJI_PALETTE = "提供を受けた匿名データの名称調査結果報告"


def build_text(text_len: int, yakumono_count: int) -> str:
    """Build text with N total chars, K yakumono chars after each digit.

    Example tl=9 yc=1: '１．提供を受けた匿名' (1 digit + 1 punct + 7 kanji = 9)
    Example tl=12 yc=3: '１．２．３．提供を受けた' (3 digits + 3 puncts + 6 kanji)
    """
    if yakumono_count == 1:
        n_kanji = max(0, text_len - 2)
        return "１．" + KANJI_PALETTE[:n_kanji]
    # yakumono_count = 3 → '１．２．３．' + kanji
    n_kanji = max(0, text_len - 6)
    return "１．２．３．" + KANJI_PALETTE[:n_kanji]


def build_document(font_sz: int, cs_tw: int, text: str, cell_dxa: int) -> bytes:
    s = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
        '<w:body>\n'
        f'<w:tbl><w:tblPr><w:tblW w:w="{cell_dxa}" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{cell_dxa}"/></w:tblGrid>'
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{cell_dxa}" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr>'
        '<w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
        '<w:adjustRightInd w:val="0"/>'
        '<w:spacing w:line="240" w:lineRule="exact"/>'
        '<w:ind w:left="215" w:right="76" w:hanging="192"/>'
        '</w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:cs="ＭＳ 明朝" w:hint="eastAsia"/>'
        f'<w:sz w:val="{font_sz}"/><w:szCs w:val="{font_sz}"/>'
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


def write_docx(name, font_sz, cs_tw, text, cell_dxa):
    out_path = os.path.join(OUT_DIR, f"{name}.docx")
    document = build_document(font_sz, cs_tw, text, cell_dxa)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', STYLES_XML_TEMPLATE)
        z.writestr('word/settings.xml', SETTINGS_XML)


def main():
    cell_dxas = [1500, 1800, 1968, 2200, 2500]
    text_lens = [5, 9, 12, 16]
    # Single-yakumono variants at fs=9 and fs=12 (NEW dimensions)
    fonts_new = [18, 24]  # 9pt and 12pt (10.5pt already in S116)
    count = 0
    for sz, cw, tl in itertools.product(fonts_new, cell_dxas, text_lens):
        text = build_text(tl, 1)
        name = f"eg_sz{sz}_cs-9_cw{cw}_tl{tl:02d}_yc1"
        write_docx(name, sz, -9, text, cw)
        count += 1
    # Multi-yakumono variants (3 punct) at all 3 fs values
    fonts_multi = [18, 21, 24]
    for sz, cw, tl in itertools.product(fonts_multi, cell_dxas, [9, 12, 16]):
        text = build_text(tl, 3)
        name = f"eg_sz{sz}_cs-9_cw{cw}_tl{tl:02d}_yc3"
        write_docx(name, sz, -9, text, cw)
        count += 1
    print(f"Wrote {count} variants to {OUT_DIR}")
    print(f"  fs=9 / fs=12 single-yakumono: {2*5*4} = 40")
    print(f"  fs=9/10.5/12 multi-yakumono (3 puncts): {3*5*3} = 45")


if __name__ == "__main__":
    main()
