"""Day 33 part 68 (R7.22) — Full mirror of a47e6 row 6 cell 1.

Structure (extracted from a47e6 docx):
- Parent table: 1 row, 2 cells (cell 0 empty tcW=2122, cell 1 tcW=7938)
- Parent row trHeight=66tw (atLeast 3.3pt default)
- Cell 1 contents (6 paragraphs + 1 nested table):
  - p0: "（５）公表関係（統計法第34条第３項の規定によるもの）"
      pPr: pStyle=a8, spacing line=240 lineRule=exact (12pt), ind left=192 right=199 hanging=192
      rPr: sz=18 (9pt) for some runs
  - nested table (4 rows × 2 cells, gridCol 4043+3288=7331):
    - row 0 (no trHeight): "公表事項" "公表内容"
        spacing line=240 lineRule=exact (12pt), sz=14 (7pt)
    - row 1 (trHeight=340 atLeast): "① 統計の作成又は統計的研究を行うに当たって利用した調査票情報を特定するために必要な事項" / (empty)
        cell 0 spacing line=160 lineRule=exact (8pt) sz=14
        cell 1 spacing line=240 lineRule=exact (12pt) sz=14 empty
    - row 2 (trHeight=340 atLeast): "② 統計の作成又は統計的研究の方法を確認するために特に必要と認める事項" / (empty)
        both cells spacing line=160 lineRule=exact (8pt) sz=14
    - row 3 (trHeight=340 atLeast): "③ 統計又は統計的研究の成果について、掲載される学術雑誌等の名称及び掲載年月日" / (empty)
        same as row 2
  - p1: "    ※ 上記③は、（４）の公表のうち代表的なものかつ一般的に入手が困難でないものを..."
      spacing line=200 lineRule=exact (10pt), sz=14 (7pt)
  - p2: "    ※ 上記以外の公表事項の公表内容（統計若しくは統計的研究の成果又はその概要の..."
      same as p1
  - p3: "◯ 統計若しくは統計的研究又はその概要を公表するに当たって特別な事情等があれば下..."
      spacing line=240 lineRule=exact (12pt), sz=18 (9pt)
  - p4: "(                                  )"
      spacing line=240 lineRule=exact (12pt), default sz
  - p5: empty
      spacing line=240 lineRule=exact (12pt)

Settings: <w:adjustLineHeightInTable/> (matches a47e6).

Output: tools/golden-test/repros/a47e6_row6/R01_full.docx
"""

from __future__ import annotations
import os
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a47e6_row6')

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

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

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="a8"><w:name w:val="Plain Text"/></w:style>
</w:styles>"""


def para(text: str, line: int, line_rule: str = "exact", ind: dict | None = None,
         sz: int | None = None, default_run_sz: int | None = None) -> str:
    """Build a paragraph with given line-spacing + indent + run sz."""
    ind_xml = ''
    if ind:
        ind_attrs = ' '.join(f'w:{k}="{v}"' for k, v in ind.items())
        ind_xml = f'<w:ind {ind_attrs}/>'
    sz_in_rpr = f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>' if sz is not None else ''
    sz_in_run = f'<w:sz w:val="{default_run_sz}"/><w:szCs w:val="{default_run_sz}"/>' if default_run_sz is not None else ''
    if text:
        runs = f'<w:r><w:rPr>{sz_in_run}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    else:
        runs = ''
    return f'<w:p><w:pPr><w:pStyle w:val="a8"/><w:spacing w:line="{line}" w:lineRule="{line_rule}"/>{ind_xml}<w:rPr>{sz_in_rpr}</w:rPr></w:pPr>{runs}</w:p>'


def nested_table_xml() -> str:
    """Nested 4-row × 2-col table matching a47e6 row 6 cell 1's nested table."""
    # Row 0: header "公表事項" "公表内容" (line=240 exact, sz=14)
    r0_c0 = para('公表事項', line=240, ind={'left': '122', 'right': '199', 'hanging': '122'}, sz=14, default_run_sz=14)
    r0_c1 = para('公表内容', line=240, ind={'left': '122', 'right': '199', 'hanging': '122'}, sz=14, default_run_sz=14)
    row0 = f"""<w:tr><w:tc><w:tcPr><w:tcW w:w="4043" w:type="dxa"/></w:tcPr>{r0_c0}</w:tc><w:tc><w:tcPr><w:tcW w:w="3288" w:type="dxa"/></w:tcPr>{r0_c1}</w:tc></w:tr>"""

    # Row 1: ① text in c0 (line=160 exact), empty c1 (line=240 exact sz=14)
    r1_c0 = para('① 統計の作成又は統計的研究を行うに当たって利用した調査票情報を特定するために必要な事項',
                 line=160, ind={'left': '140', 'right': '199', 'hanging': '140'}, sz=14, default_run_sz=14)
    r1_c1 = para('', line=240, ind={'left': '122', 'right': '199', 'hanging': '122'}, sz=14)
    row1 = f"""<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4043" w:type="dxa"/></w:tcPr>{r1_c0}</w:tc><w:tc><w:tcPr><w:tcW w:w="3288" w:type="dxa"/></w:tcPr>{r1_c1}</w:tc></w:tr>"""

    # Row 2: ② in c0 (line=160 exact), empty c1 (line=160 exact)
    r2_c0 = para('② 統計の作成又は統計的研究の方法を確認するために特に必要と認める事項',
                 line=160, ind={'left': '140', 'right': '199', 'hanging': '140'}, sz=14, default_run_sz=14)
    r2_c1 = para('', line=160, ind={'left': '122', 'right': '199', 'hanging': '122'}, sz=14)
    row2 = f"""<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4043" w:type="dxa"/></w:tcPr>{r2_c0}</w:tc><w:tc><w:tcPr><w:tcW w:w="3288" w:type="dxa"/></w:tcPr>{r2_c1}</w:tc></w:tr>"""

    # Row 3: ③ in c0, empty c1 (both line=160 exact)
    r3_c0 = para('③ 統計又は統計的研究の成果について、掲載される学術雑誌等の名称及び掲載年月日',
                 line=160, ind={'left': '140', 'right': '199', 'hanging': '140'}, sz=14, default_run_sz=14)
    r3_c1 = para('', line=160, ind={'left': '122', 'right': '199', 'hanging': '122'}, sz=14)
    row3 = f"""<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4043" w:type="dxa"/></w:tcPr>{r3_c0}</w:tc><w:tc><w:tcPr><w:tcW w:w="3288" w:type="dxa"/></w:tcPr>{r3_c1}</w:tc></w:tr>"""

    return f"""<w:tbl>
<w:tblPr>
<w:tblW w:w="7331" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="4043"/><w:gridCol w:w="3288"/></w:tblGrid>
{row0}{row1}{row2}{row3}
</w:tbl>"""


def cell1_content() -> str:
    """Build cell 1 content: 1 label para + nested table + 5 trailing paras."""
    # p0: label "（５）公表関係..."
    p0 = para('（５）公表関係（統計法第34条第３項の規定によるもの）',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'},
              sz=18, default_run_sz=None)  # default rPr=18 → 9pt for unmarked runs

    nested = nested_table_xml()

    # p1: footnote 1 "※..."
    p1 = para('    ※ 上記③は、（４）の公表のうち代表的なものかつ一般的に入手が困難でないもの',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)

    # p2: footnote 2 "※..."
    p2 = para('    ※ 上記以外の公表事項の公表内容（統計若しくは統計的研究の成果又はその概要の公表のために必要なもの）',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)

    # p3: ◯
    p3 = para('◯　統計若しくは統計的研究又はその概要を公表するに当たって特別な事情等があれば下記に記入',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=18, default_run_sz=18)

    # p4: ( ... )
    p4 = para('　(　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　)',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)

    # p5: empty
    p5 = para('', line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)

    return f"{p0}{nested}{p1}{p2}{p3}{p4}{p5}"


def doc_xml() -> str:
    cell1 = cell1_content()
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>HEADER MARKER</w:t></w:r></w:p>
<w:tbl>
<w:tblPr>
<w:tblW w:w="0" w:type="auto"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="2122"/><w:gridCol w:w="7938"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="66"/></w:trPr>
<w:tc><w:tcPr><w:tcW w:w="2122" w:type="dxa"/></w:tcPr><w:p/></w:tc>
<w:tc><w:tcPr><w:tcW w:w="7938" w:type="dxa"/></w:tcPr>
{cell1}
</w:tc>
</w:tr>
</w:tbl>
<w:p><w:r><w:t>FOOTER MARKER</w:t></w:r></w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>"""


def build(name: str, content_fn) -> str:
    out_path = os.path.join(OUT_DIR, f'{name}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    xml = content_fn()
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', xml)
    return out_path


def cell1_no_trailing() -> str:
    """R02: label + nested table only (no trailing paragraphs)."""
    p0 = para('（５）公表関係（統計法第34条第３項の規定によるもの）',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'},
              sz=18, default_run_sz=None)
    nested = nested_table_xml()
    return f"{p0}{nested}"


def cell1_no_label() -> str:
    """R03: trailing paragraphs only (no label, no nested table)."""
    p1 = para('    ※ 上記③は、（４）の公表のうち代表的なものかつ一般的に入手が困難でないもの',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)
    p2 = para('    ※ 上記以外の公表事項の公表内容（統計若しくは統計的研究の成果又はその概要の公表のために必要なもの）',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)
    p3 = para('◯　統計若しくは統計的研究又はその概要を公表するに当たって特別な事情等があれば下記に記入',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=18, default_run_sz=18)
    p4 = para('　(　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　)',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)
    p5 = para('', line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)
    return f"{p1}{p2}{p3}{p4}{p5}"


def cell1_nested_only() -> str:
    """R04: nested table only (no other paragraphs)."""
    return nested_table_xml()


def cell1_label_plus_trailing() -> str:
    """R05: label + trailing paragraphs (no nested table)."""
    p0 = para('（５）公表関係（統計法第34条第３項の規定によるもの）',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'},
              sz=18, default_run_sz=None)
    p1 = para('    ※ 上記③は、（４）の公表のうち代表的なものかつ一般的に入手が困難でないもの',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)
    p2 = para('    ※ 上記以外の公表事項の公表内容（統計若しくは統計的研究の成果又はその概要の公表のために必要なもの）',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)
    p3 = para('◯　統計若しくは統計的研究又はその概要を公表するに当たって特別な事情等があれば下記に記入',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=18, default_run_sz=18)
    p4 = para('　(　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　)',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)
    p5 = para('', line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)
    return f"{p0}{p1}{p2}{p3}{p4}{p5}"


def cell1_nested_plus_trailing() -> str:
    """R06: nested + trailing only (no label)."""
    nested = nested_table_xml()
    p1 = para('    ※ 上記③は、（４）の公表のうち代表的なものかつ一般的に入手が困難でないもの',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)
    p2 = para('    ※ 上記以外の公表事項の公表内容（統計若しくは統計的研究の成果又はその概要の公表のために必要なもの）',
              line=200, ind={'left': '140', 'right': '199'}, sz=14, default_run_sz=14)
    p3 = para('◯　統計若しくは統計的研究又はその概要を公表するに当たって特別な事情等があれば下記に記入',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=18, default_run_sz=18)
    p4 = para('　(　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　)',
              line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)
    p5 = para('', line=240, ind={'left': '192', 'right': '199', 'hanging': '192'}, sz=None)
    return f"{nested}{p1}{p2}{p3}{p4}{p5}"


def make_doc_xml(cell1_fn) -> str:
    cell1 = cell1_fn()
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>HEADER MARKER</w:t></w:r></w:p>
<w:tbl>
<w:tblPr>
<w:tblW w:w="0" w:type="auto"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="2122"/><w:gridCol w:w="7938"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="66"/></w:trPr>
<w:tc><w:tcPr><w:tcW w:w="2122" w:type="dxa"/></w:tcPr><w:p/></w:tc>
<w:tc><w:tcPr><w:tcW w:w="7938" w:type="dxa"/></w:tcPr>
{cell1}
</w:tc>
</w:tr>
</w:tbl>
<w:p><w:r><w:t>FOOTER MARKER</w:t></w:r></w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>"""


def main() -> int:
    variants = [
        ('R01_full',                lambda: doc_xml()),
        ('R02_label_plus_nested',   lambda: make_doc_xml(cell1_no_trailing)),
        ('R03_trailing_only',       lambda: make_doc_xml(cell1_no_label)),
        ('R04_nested_only',         lambda: make_doc_xml(cell1_nested_only)),
        ('R05_label_plus_trailing', lambda: make_doc_xml(cell1_label_plus_trailing)),
        ('R06_nested_plus_trailing',lambda: make_doc_xml(cell1_nested_plus_trailing)),
    ]
    for name, fn in variants:
        path = build(name, fn)
        print(f'built: {path}')
    return 0


if __name__ == '__main__':
    import sys
    sys.exit(main())
