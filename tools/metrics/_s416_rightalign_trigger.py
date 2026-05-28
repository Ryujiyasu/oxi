"""S416: minimal repro to isolate WHY Word ignores jc=right (renders text
LEFT-positioned) for the S414/S415 fire cells.

S415 confirmed (2 docs): cells with jc=right + firstLine in multi-column
tblCellMar tables are rendered by Word at the cell LEFT origin, ignoring
jc=right / pad_l / firstLine. Text width identical to Oxi (fullwidth).

This builds ONE docx with a 3-cell-per-row table (tblCellMar=99) where
each row's cell 0 holds a test paragraph varying ONE factor from the
bug base case, then COM-measures each test paragraph's first-char x and
alignment. Compare each right/center case's char[0] x to the left-aligned
reference (T3) to decide LEFT-positioned vs honored-alignment.

Test matrix (cell width ~90pt, text width chosen << cell so L vs R is
clearly separable):
  T1  jc=right  firstLine=210(fLC=100)  "税税"     <- BUG base (expect LEFT)
  T2  jc=right  NO firstLine            "税税"     <- isolates firstLine
  T3  jc=left   firstLine=210           "税税"     <- LEFT reference
  T4  jc=right  firstLine=210           "AB"       <- ASCII (fullwidth-specific?)
  T5  jc=center firstLine=210           "税税"     <- center vs right
  T6  jc=right  firstLine=210  lead U+3000 "　税税" <- leading ideographic space
  T7  jc=right  NO firstLine            "AB"       <- right ASCII no firstLine
  T8  jc=right  firstLine=210           "税税税税税税税税" <- wide (near/over cell)

Instrumentation only. Output docx in tools/golden-test/repros/.
"""
from __future__ import annotations
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = r'c:\Users\ryuji\oxi-main'
OUT_DOCX = os.path.join(REPO, 'tools', 'golden-test', 'repros', 's416_rightalign_trigger.docx')

# Each test: (label, jc, has_firstLine, text)
TESTS = [
    ('T1', 'right',  True,  '税税'),
    ('T2', 'right',  False, '税税'),
    ('T3', 'left',   True,  '税税'),
    ('T4', 'right',  True,  'AB'),
    ('T5', 'center', True,  '税税'),
    ('T6', 'right',  True,  '　税税'),
    ('T7', 'right',  False, 'AB'),
    ('T8', 'right',  True,  '税税税税税税税税'),
]


def cell_xml(test):
    label, jc, has_fl, text = test
    ind = '<w:ind w:firstLineChars="100" w:firstLine="210"/>' if has_fl else ''
    # cell 0 holds the test paragraph; tcW ~ 90pt = 1808 twip (like ed025 col2)
    p = (f'<w:p><w:pPr>{ind}<w:jc w:val="{jc}"/></w:pPr>'
         f'<w:r><w:rPr><w:sz w:val="21"/></w:rPr>'
         f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')
    tc0 = f'<w:tc><w:tcPr><w:tcW w:w="1808" w:type="dxa"/></w:tcPr>{p}</w:tc>'
    # two filler cells to make a 3-cell row (multi-column requirement)
    filler = ('<w:tc><w:tcPr><w:tcW w:w="1600" w:type="dxa"/></w:tcPr>'
              '<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
              '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>x</w:t></w:r></w:p></w:tc>')
    return f'<w:tr>{tc0}{filler}{filler}</w:tr>'


def build():
    rows = ''.join(cell_xml(t) for t in TESTS)
    tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="0" w:type="auto"/>'
        '<w:tblInd w:w="145" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
        '<w:tblCellMar>'
        '<w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/>'
        '</w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1808"/><w:gridCol w:w="1600"/><w:gridCol w:w="1600"/></w:tblGrid>'
        f'{rows}'
        '</w:tbl>'
    )
    parts = {}
    parts['[Content_Types].xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
    parts['_rels/.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    parts['word/_rels/document.xml.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    parts['word/styles.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>'
    parts['word/document.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        f'{tbl}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    os.makedirs(os.path.dirname(OUT_DOCX), exist_ok=True)
    with zipfile.ZipFile(OUT_DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, content in parts.items():
            z.writestr(name, content)
    print(f'built {OUT_DOCX}')


def measure():
    import win32com.client as wc
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(OUT_DOCX, ReadOnly=True)
    doc.Repaginate()
    results = []
    try:
        table = doc.Tables(1)
        n_rows = table.Rows.Count
        for ri in range(1, n_rows + 1):
            cell = table.Cell(ri, 1)  # col 0 holds the test paragraph
            p = cell.Range.Paragraphs(1)
            rng = p.Range
            txt = (rng.Text or '').rstrip('\r\n\x07')
            start = rng.Start
            xs = []
            for i, ch in enumerate(txt):
                cr = doc.Range(start + i, start + i)
                try:
                    xs.append(round(float(cr.Information(5)), 2))
                except Exception:
                    xs.append(-1.0)
            label = TESTS[ri - 1][0]
            jc = TESTS[ri - 1][1]
            fl = p.Format.FirstLineIndent
            align = p.Format.Alignment
            results.append({
                'label': label, 'spec_jc': jc, 'text': txt,
                'word_align': align, 'firstLine_pt': round(float(fl), 2),
                'char0_x': xs[0] if xs else None, 'all_x': xs,
            })
    finally:
        doc.Close(SaveChanges=0)
        word.Quit()
    return results


def main():
    build()
    res = measure()
    # Left reference = T3
    left_ref = next((r['char0_x'] for r in res if r['label'] == 'T3'), None)
    print(f'\nLeft-aligned reference (T3) char0_x = {left_ref}')
    print(f"{'case':4s} {'spec_jc':7s} {'wAlign':6s} {'fl':5s} {'char0_x':8s} {'vs_left':8s} text")
    for r in res:
        d = (r['char0_x'] - left_ref) if (left_ref is not None and r['char0_x'] is not None) else None
        verdict = ''
        if d is not None:
            verdict = 'LEFT' if abs(d) < 3.0 else f'+{d:.1f}'
        print(f"{r['label']:4s} {r['spec_jc']:7s} {r['word_align']:<6} {r['firstLine_pt']:<5} {str(r['char0_x']):8s} {verdict:8s} {r['text']!r}")
    print('\nInterpretation: char0_x ~= T3 (LEFT) means Word ignored jc and left-positioned.')


if __name__ == '__main__':
    main()
