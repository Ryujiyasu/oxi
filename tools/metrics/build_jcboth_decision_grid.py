"""Session 116 — extended decision-boundary grid for jc=both yakumono
compression.

S113 grid varied (font, fs, cs, punct) at FIXED cell width. S114/S115
attempts revealed Word's compression has a narrower trigger than
'natural overflow could be absorbed by ×0.5'. The decision rule isn't
known.

S116 grid varies cell width and text length at FIXED (fs=10.5, cs=-9,
MS Mincho) — the configuration where compression IS triggered for v8.

Grid dimensions:
  cell_dxa  : 1200, 1500, 1800, 1968, 2200, 2500, 3000, 3500
              (varies cell budget from ~58pt to ~173pt at fs=10.5)
  text_len  : 5, 7, 9, 10, 12, 14, 16, 20 chars total (incl '１' '．')
              ('１．' + N kanji where N = text_len - 2)
  punct_pos : 2 (after '１') — fixed

For each variant: COM-measure '．' advance, L1 char count, line widths.
The compression trigger pattern across (cell_w, text_len) should reveal
when Word picks compression vs natural wrap.

Total variants: 8 × 8 = 64
"""
import os
import sys
import io
import itertools
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/jcboth_decision_grid/variants"))
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

STYLES_XML = (
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

# 18-char kanji palette to fill N kanji slots
KANJI_PALETTE = "提供を受けた匿名データの名称調査結果報告"


def build_document(cell_dxa: int, text_len: int) -> bytes:
    """text = '１．' + (text_len - 2) kanji from palette."""
    n_kanji = max(0, text_len - 2)
    kanji_text = KANJI_PALETTE[:n_kanji] if n_kanji <= len(KANJI_PALETTE) else KANJI_PALETTE * (n_kanji // len(KANJI_PALETTE)) + KANJI_PALETTE[:n_kanji % len(KANJI_PALETTE)]
    text = "１．" + kanji_text
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
        '<w:sz w:val="21"/><w:szCs w:val="21"/>'
        '<w:spacing w:val="-9"/><w:kern w:val="0"/>'
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


def write_docx(name, cell_dxa, text_len):
    out_path = os.path.join(OUT_DIR, f"{name}.docx")
    document = build_document(cell_dxa, text_len)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/settings.xml', SETTINGS_XML)


def main():
    cell_dxas = [1200, 1500, 1800, 1968, 2200, 2500, 3000, 3500]
    text_lens = [5, 7, 9, 10, 12, 14, 16, 20]
    count = 0
    for cw, tl in itertools.product(cell_dxas, text_lens):
        name = f"dg_cw{cw}_tl{tl:02d}"
        write_docx(name, cw, tl)
        count += 1
    print(f"Wrote {count} variants to {OUT_DIR}")


if __name__ == "__main__":
    main()
