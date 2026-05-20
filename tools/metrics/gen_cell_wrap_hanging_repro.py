"""S139: Generate TR_V300 minimal repros to isolate why Oxi over-counts
wrap lines for hanging-indent paragraphs in a1d6's column-30 cells.

a1d6 setup (法人等 paragraph in row 21 col 30):
- Cell tcW=7654dxa (382.7pt), gridSpan=8, default cellMar L/R=108dxa
- Paragraph: pStyle=ac, wordWrap, line=280 exact (14pt), sz=20 (10pt),
  ind leftChars=150 left=533 hangingChars=100 hanging=207, jc=left
- Char spacing val=0 in run rPr (overrides style ac's val=-1)
- Text: '○　法人等であって、その役員のうちに上記のいずれかに該当する者がある者' (33 chars)
- Word: 1 line. Oxi: 2 lines. Bug.

Variants:
  V300a: NO hanging indent (sanity — should be 1 line in both)
  V300b: Exact a1d6 setup (33 chars + hanging=207 + sz=20)
         expected to reproduce the bug (Oxi 2 lines, Word 1)
  V300c: Wider cell (tcW=9000) — verify width factor
  V300d: Smaller text (sz=18, 9pt) — verify font-size factor
  V300e: No char spacing override (uses style ac's -1tw) — verify spacing factor
  V300f: jc=both instead of left
  V300g: explicit tcMar=0 (vs default 108dxa)
  V300h: Different hanging (50tw small) — verify amount factor
  V300i: Plain text (no full-width space ○) — verify the special char
  V300j: 27 chars (shorter)
  V300k: 40 chars (longer)
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cell_wrap_hanging')

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

# styles.xml mirrors a1d6's pStyle "ac": MS Mincho, line=210 exact, spacing val=-1, sz=21
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho" w:cs="MS Mincho"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac">
<w:name w:val="ac"/>
<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:cs="MS Mincho"/><w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
</w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>'''


# Text variants — full-width space + 33-char body matching a1d6 法人等 line
A1D6_TEXT_33 = '法人等であって、その役員のうちに上記のいずれかに該当する者がある者'
A1D6_TEXT_27 = '法人等であって、その役員に該当する者がある者'  # 22 chars (approx 27 with marker)
A1D6_TEXT_40 = '法人等であって、その役員のうちに上記のいずれかに該当する者がある者である者'


def make_para(text: str, hanging: int = 207, hanging_chars: int = 100,
              left: int = 533, left_chars: int = 150,
              sz: int = 20, run_spacing: int = 0, jc: str = 'left',
              use_pstyle_ac: bool = True, marker: str = '○　') -> str:
    pstyle = '<w:pStyle w:val="ac"/>' if use_pstyle_ac else ''
    indent = (f'<w:ind w:leftChars="{left_chars}" w:left="{left}" '
              f'w:hangingChars="{hanging_chars}" w:hanging="{hanging}"/>') if hanging > 0 else ''
    run_rpr = f'<w:rPr><w:spacing w:val="{run_spacing}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    return ('<w:p><w:pPr>'
            f'{pstyle}'
            '<w:wordWrap/>'
            '<w:spacing w:line="280" w:lineRule="exact"/>'
            f'{indent}'
            f'<w:jc w:val="{jc}"/>'
            f'{run_rpr}'
            '</w:pPr>'
            f'<w:r>{run_rpr}<w:t>{marker}{text}</w:t></w:r>'
            '</w:p>')


def make_cell(para_xml: str, tcw: int = 7654, gridspan: int = 8,
              tcmar_top: int | None = None, tcmar_bottom: int | None = None) -> str:
    gs = f'<w:gridSpan w:val="{gridspan}"/>' if gridspan > 1 else ''
    tcmar = ''
    if tcmar_top is not None or tcmar_bottom is not None:
        parts = []
        if tcmar_top is not None: parts.append(f'<w:top w:w="{tcmar_top}" w:type="dxa"/>')
        if tcmar_bottom is not None: parts.append(f'<w:bottom w:w="{tcmar_bottom}" w:type="dxa"/>')
        tcmar = '<w:tcMar>' + ''.join(parts) + '</w:tcMar>'
    return ('<w:tc><w:tcPr>'
            f'<w:tcW w:w="{tcw}" w:type="dxa"/>{gs}'
            '<w:vAlign w:val="center"/>'
            f'{tcmar}'
            '</w:tcPr>'
            f'{para_xml}</w:tc>')


def make_table(cell_xml: str, total_w: int = 7654, n_grid_cols: int = 8) -> str:
    # Declare n_grid_cols grid columns to match gridSpan in cell.
    col_w = total_w // n_grid_cols
    grid_cols = ''.join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(n_grid_cols))
    return ('<w:tbl>'
            '<w:tblPr>'
            f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
            '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
            '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>'
            '</w:tblBorders>'
            '</w:tblPr>'
            f'<w:tblGrid>{grid_cols}</w:tblGrid>'
            f'<w:tr>{cell_xml}</w:tr>'
            '</w:tbl>')


def doc_xml(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11904" w:h="16836" w:code="9"/>
<w:pgMar w:top="851" w:right="1134" w:bottom="851" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
<w:docGrid w:type="linesAndChars" w:linePitch="272"/>
</w:sectPr>
</w:body>
</w:document>"""


def write_docx(label: str, doc: str):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/document.xml', doc)
    return out


def main():
    # V300a: NO hanging indent (sanity)
    p = make_para(A1D6_TEXT_33, hanging=0, left=0, left_chars=0)
    write_docx('V300a_no_hanging', doc_xml(make_table(make_cell(p))))

    # V300b: Exact a1d6 setup (33 chars + hanging=207 + sz=20)
    p = make_para(A1D6_TEXT_33)
    write_docx('V300b_a1d6_exact', doc_xml(make_table(make_cell(p))))

    # V300c: Wider cell (tcW=9000)
    p = make_para(A1D6_TEXT_33)
    write_docx('V300c_wider_cell', doc_xml(make_table(make_cell(p, tcw=9000), total_w=9000)))

    # V300d: Smaller text (sz=18, 9pt)
    p = make_para(A1D6_TEXT_33, sz=18)
    write_docx('V300d_sz18', doc_xml(make_table(make_cell(p))))

    # V300e: No char spacing override (run uses style ac's -1tw)
    p = make_para(A1D6_TEXT_33, run_spacing=-1)
    write_docx('V300e_spacing_neg1', doc_xml(make_table(make_cell(p))))

    # V300f: jc=both instead of left
    p = make_para(A1D6_TEXT_33, jc='both')
    write_docx('V300f_jc_both', doc_xml(make_table(make_cell(p))))

    # V300g: explicit tcMar top=0 bottom=0 (vs default 108dxa)
    p = make_para(A1D6_TEXT_33)
    write_docx('V300g_tcmar0', doc_xml(make_table(make_cell(p, tcmar_top=0, tcmar_bottom=0))))

    # V300h: Different hanging (50tw small)
    p = make_para(A1D6_TEXT_33, hanging=50, hanging_chars=25)
    write_docx('V300h_hanging50', doc_xml(make_table(make_cell(p))))

    # V300i: No marker (no ○　 prefix)
    p = make_para(A1D6_TEXT_33, marker='')
    write_docx('V300i_no_marker', doc_xml(make_table(make_cell(p))))

    # V300j: shorter text
    p = make_para(A1D6_TEXT_27)
    write_docx('V300j_text27', doc_xml(make_table(make_cell(p))))

    # V300k: longer text
    p = make_para(A1D6_TEXT_40)
    write_docx('V300k_text40', doc_xml(make_table(make_cell(p))))

    # V300l: no pStyle ac (just direct properties)
    p = make_para(A1D6_TEXT_33, use_pstyle_ac=False)
    write_docx('V300l_no_pstyle', doc_xml(make_table(make_cell(p))))

    # V300m: multi-run mimicking a1d6 actual structure
    # 6 runs alternating with/without rFonts hint="eastAsia"
    def make_multi_run_para():
        runs = [
            ('○', True),   # eastAsia
            ('　', False),  # not
            ('法人等であって', True),
            ('、その役員のうちに', False),
            ('上記', True),
            ('のいずれかに該当する者がある者', False),
        ]
        sb_pieces = ''
        for txt, eastasia in runs:
            font_hint = '<w:rFonts w:hint="eastAsia"/>' if eastasia else ''
            sb_pieces += (
                '<w:r>'
                f'<w:rPr>{font_hint}<w:spacing w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                f'<w:t xml:space="preserve">{txt}</w:t>'
                '</w:r>'
            )
        return ('<w:p><w:pPr>'
                '<w:pStyle w:val="ac"/>'
                '<w:wordWrap/>'
                '<w:spacing w:line="280" w:lineRule="exact"/>'
                '<w:ind w:leftChars="150" w:left="533" w:hangingChars="100" w:hanging="207"/>'
                '<w:jc w:val="left"/>'
                '<w:rPr><w:spacing w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                '</w:pPr>'
                f'{sb_pieces}'
                '</w:p>')

    write_docx('V300m_multi_run', doc_xml(make_table(make_cell(make_multi_run_para()))))

    print('Done. Repros in', OUT_DIR)


if __name__ == '__main__':
    main()
