"""S156: Add attributes from V800k to V800l one at a time."""
from __future__ import annotations
import os, sys, zipfile, subprocess, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a1d6_row21_isolate')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')


def base_doc(cell0_extras='', cell1_extras='', tblInd='', n_fillers=18):
    fillers = '\n'.join(
        f'<w:p><w:pPr><w:pStyle w:val="ac"/><w:spacing w:line="280" w:lineRule="exact"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>FILLER_{i}_長いテキスト</w:t></w:r></w:p>'
        for i in range(n_fillers)
    )
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
            f'<w:body>\n'
            f'<w:tbl>\n'
            f'<w:tblPr><w:tblW w:w="9639" w:type="dxa"/>{tblInd}<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>\n'
            f'<w:tblGrid>\n'
            f'<w:gridCol w:w="1985"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/><w:gridCol w:w="956"/>\n'
            f'</w:tblGrid>\n'
            f'<w:tr>\n'
            f'<w:trPr><w:cantSplit/><w:trHeight w:val="4822"/></w:trPr>\n'
            f'<w:tc><w:tcPr><w:tcW w:w="1985" w:type="dxa"/><w:vMerge/>{cell0_extras}</w:tcPr><w:p><w:pPr><w:pStyle w:val="ac"/></w:pPr></w:p></w:tc>\n'
            f'<w:tc><w:tcPr><w:tcW w:w="7654" w:type="dxa"/><w:gridSpan w:val="8"/>{cell1_extras}<w:vAlign w:val="center"/></w:tcPr>\n'
            f'<w:p><w:pPr><w:pStyle w:val="ac"/><w:wordWrap/><w:spacing w:beforeLines="50" w:before="146" w:line="280" w:lineRule="exact"/><w:ind w:leftChars="50" w:left="316" w:hangingChars="100" w:hanging="207"/><w:jc w:val="left"/><w:rPr><w:spacing w:val="0"/><w:sz w:val="20"/></w:rPr></w:pPr>\n'
            f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr><w:t>※　匿名データを取り扱う者が以下のいずれにも該当しない場合</w:t></w:r>\n'
            f'</w:p>\n'
            f'{fillers}\n'
            f'</w:tc>\n'
            f'</w:tr>\n'
            f'</w:tbl>\n'
            f'<w:sectPr>\n'
            f'<w:pgSz w:w="11906" w:h="16838"/>\n'
            f'<w:pgMar w:top="1247" w:right="1077" w:bottom="1440" w:left="1077" w:header="851" w:footer="992"/>\n'
            f'<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>\n'
            f'</w:sectPr>\n'
            f'</w:body>\n'
            f'</w:document>\n')


CONTENT_TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ポス 明朝" w:hAnsi="Century" w:cs="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="ac"><w:name w:val="ac"/><w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="ポス 明朝" w:hAnsi="ポス 明朝" w:cs="ポス 明朝"/><w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style></w:styles>'
SETTINGS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:useFELayout/><w:adjustLineHeightInTable/></w:compat></w:settings>'


def write_docx(label, doc):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/document.xml', doc)
    return out


variants = [
    ('V800m_with_tcBorders_explicit',
     '<w:tcBorders><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>',
     '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>',
     ''),
    ('V800n_with_shd_clear', '', '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>', ''),
    ('V800o_both_tcBorders_shd',
     '<w:tcBorders><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>',
     '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/>',
     ''),
    ('V800p_with_tblInd', '', '', '<w:tblInd w:w="433" w:type="dxa"/>'),
    ('V800q_all_attrs',
     '<w:tcBorders><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>',
     '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="auto"/>',
     '<w:tblInd w:w="433" w:type="dxa"/>'),
]


def measure(docx_path, label):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    docx_full = os.path.abspath(docx_path)
    d = word.Documents.Open(docx_full, ReadOnly=True)
    w_y = w_pg = None
    try:
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            if '取り扱う' in (p.Range.Text or ''):
                rng = p.Range
                cr = d.Range(rng.Start, rng.Start)
                w_y = round(cr.Information(6), 2)
                w_pg = int(cr.Information(3))
                break
    finally:
        d.Close(False)
        word.Quit()
    out_layout = os.path.join(r'C:\tmp', f'{label}.json')
    r = subprocess.run([RENDERER, docx_full, os.path.join(r'C:\tmp', label), f'--dump-layout={out_layout}'], capture_output=True, text=True)
    o_y = o_pg = None
    if r.returncode == 0:
        layout = json.load(open(out_layout, encoding='utf-8'))
        for pi, page in enumerate(layout['pages']):
            for el in page['elements']:
                if el.get('type') == 'text' and '取り扱う' in el.get('text', ''):
                    o_y = round(el['y'], 2); o_pg = pi+1
                    break
            if o_y: break
    return w_y, w_pg, o_y, o_pg, r.stderr[:200] if r.returncode != 0 else None


def main():
    print(f'{"Variant":<40} | Word | Oxi | drift')
    print('-' * 75)
    for name, cell0, cell1, tblInd in variants:
        docx = write_docx(name, base_doc(cell0, cell1, tblInd))
        w_y, w_pg, o_y, o_pg, err = measure(docx, name)
        if err:
            print(f'{name:<40} | FAIL: {err}')
            continue
        if w_y is None or o_y is None:
            print(f'{name:<40} | no marker')
            continue
        drift = o_y - w_y
        print(f'{name:<40} | y={w_y:>6} | y={o_y:>6} | {drift:+.2f}')


if __name__ == '__main__':
    main()
