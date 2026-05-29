"""S445: minimal repro to isolate the +1.25pt Word adds to atLeast trHeight rows.

7ead52: rows with trHeight=860tw(43.0pt) hRule=atLeast, single 11pt line,
docGrid lines linePitch=360. Word renders 44.25pt pitch (=43.0+1.25);
Oxi renders exactly 43.0. content(~18pt) << trHeight so NOT content-driven.

Sweep what produces the +1.25:
 V0 base   : trH=860 atLeast, 11pt line, docGrid lines 360, borders sz4
 V1 nogrid : V0 but docGrid type=default (no line snap)
 V2 noborder: V0 but no table borders
 V3 exact  : V0 but hRule=exactly
 V4 trH600 : V0 but trHeight=600 (is +1.25 constant or proportional?)
 V5 sz21   : V0 but 10.5pt font
 V6 nosnap : V0 but settings has no snapToGrid / para snapToGrid=0
"""
import os, zipfile
from pathlib import Path

OUT = Path(__file__).parent / "s445_trheight_repro"
OUT.mkdir(exist_ok=True)

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS_ROOT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="21"/><w:szCs w:val="22"/></w:rPr></w:style>
</w:styles>'''

def settings(snap=True):
    snapline = '<w:doNotSnapToGridInCell/>' if not snap else ''
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{snapline}<w:compat/></w:settings>'''

def build_doc(trh=860, hrule="atLeast", grid="lines", borders=True, sz=22, nrows=5,
              valign=None, line=None, line_rule=None):
    docgrid = (f'<w:docGrid w:type="{grid}" w:linePitch="360"/>' if grid != "none"
               else '<w:docGrid w:type="default" w:linePitch="360"/>')
    if borders:
        tblb = ('<w:tblBorders>'
                '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '</w:tblBorders>')
    else:
        tblb = ''
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''
    spacing_xml = (f'<w:spacing w:line="{line}" w:lineRule="{line_rule}"/>'
                   if line is not None else '')
    rows = ""
    chars = "あいうえお"  # あいうえお
    for i in range(nrows):
        rows += (f'<w:tr><w:trPr><w:trHeight w:val="{trh}" w:hRule="{hrule}"/></w:trPr>'
                 f'<w:tc><w:tcPr><w:tcW w:w="5519" w:type="dxa"/>{valign_xml}</w:tcPr>'
                 f'<w:p><w:pPr>{spacing_xml}<w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:pPr>'
                 f'<w:r><w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
                 f'<w:t>連絡{chars[i % 5]}</w:t></w:r></w:p></w:tc></w:tr>')
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="5519" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           f'{tblb}<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="5519"/></w:tblGrid>{rows}</w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{tbl}<w:p/>'
            f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
            f'{docgrid}</w:sectPr></w:body></w:document>')

def write_docx(name, doc_xml, snap=True):
    path = OUT / name
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS_ROOT)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", settings(snap))
    print("wrote", path)

def build_faithful(nrows=5):
    """3-col matching 7ead52 rows 4-8: vMerge label col1, col2 label, col3 empty.
    col widths 621/2209/5519, trHeight=860 atLeast, 11pt, docGrid lines 360,
    all borders sz4 incl insideV."""
    tblb = ('<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders>')
    labels = ["氏名", "所属", "役職", "電話番号", "Ｅ－ｍａｉｌ"]
    rows = ""
    for i in range(nrows):
        vm = '<w:vMerge w:val="restart"/>' if i == 0 else '<w:vMerge/>'
        c1txt = '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>連絡担当窓口</w:t></w:r>' if i == 0 else ''
        rows += (f'<w:tr><w:trPr><w:trHeight w:val="860" w:hRule="atLeast"/></w:trPr>'
                 f'<w:tc><w:tcPr><w:tcW w:w="621" w:type="dxa"/>{vm}</w:tcPr>'
                 f'<w:p><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>{c1txt}</w:p></w:tc>'
                 f'<w:tc><w:tcPr><w:tcW w:w="2209" w:type="dxa"/></w:tcPr>'
                 f'<w:p><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>'
                 f'<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{labels[i]}</w:t></w:r></w:p></w:tc>'
                 f'<w:tc><w:tcPr><w:tcW w:w="5519" w:type="dxa"/></w:tcPr>'
                 f'<w:p><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr></w:p></w:tc></w:tr>')
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="8349" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           f'{tblb}<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="621"/><w:gridCol w:w="2209"/><w:gridCol w:w="5519"/></w:tblGrid>{rows}</w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{tbl}<w:p/>'
            f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')

write_docx("VF_faithful.docx", build_faithful())
write_docx("V0_base.docx", build_doc())
# real-doc attribute combos (1-col to keep pitch clean)
write_docx("VG_valign.docx", build_doc(valign="center"))
write_docx("VH_line384.docx", build_doc(line=384, line_rule="atLeast"))
write_docx("VI_valign_line.docx", build_doc(valign="center", line=384, line_rule="atLeast"))
write_docx("VJ_valign_line_exact.docx", build_doc(hrule="exact", valign="center", line=384, line_rule="atLeast"))
write_docx("V1_nogrid.docx", build_doc(grid="default"))
write_docx("V2_noborder.docx", build_doc(borders=False))
write_docx("V3_exact.docx", build_doc(hrule="exact"))
write_docx("V4_trh600.docx", build_doc(trh=600))
write_docx("V5_sz21.docx", build_doc(sz=21))
write_docx("V6_nosnap.docx", build_doc(), snap=False)
print("done")
