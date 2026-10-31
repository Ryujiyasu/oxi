# -*- coding: utf-8 -*-
"""Does a cell paragraph's first line wrap past the cell's inner edge when its
first-line start is negative?

ja/educational__244a3664: the 段階 column's «(10分)» labels (w:w=66, leftChars=-47
hangingChars=47, i.e. first line 0.47 char LEFT of the inner edge) sit on one
line in Word, starting 3.6pt left of the cell's text edge; Oxi clamps the
first-line start to the inner edge when it computes the wrap width and breaks
after «(10».

Arms (cell 3000tw, margins 99, 10.5pt ＭＳ 明朝, jc left / center): the first-line
start is negative through leftChars, a twip left, or a twip hanging larger than
the left. Readout from Word's PDF: x of the first and second line's first char
and how many chars line 1 holds.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cell_negfirst'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
TEXT = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも'
ARMS = {
    'none': '',
    'lc_m47_hc47': '<w:ind w:leftChars="-47" w:left="-34" w:hangingChars="47" w:hanging="65"/>',
    'lc_m100': '<w:ind w:leftChars="-100" w:left="-210"/>',
    'left_m200': '<w:ind w:left="-200"/>',
    'hang_400_left_200': '<w:ind w:left="200" w:hanging="400"/>',
    'fl_m200': '<w:ind w:firstLine="-200"/>',
}


def document(ind, jc):
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
    p = f'<w:p><w:pPr>{ind}<w:jc w:val="{jc}"/></w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'
    tc = '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + p + '</w:tc>'
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblInd w:w="1000" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders>'
           '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>{tc}</w:tr></w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p/>{tbl}<w:p/>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        print('cell left rule 106.7, inner edge 111.65, inner right 256.75')
        for name, ind in ARMS.items():
            for jc in ('left', 'center'):
                at = os.path.abspath(os.path.join(OUT, f'{name}_{jc}.docx'))
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                    z.writestr('word/document.xml', document(ind, jc))
                pdf = at[:-5] + '.pdf'
                d = app.Documents.Open(at, ReadOnly=True)
                try:
                    d.ExportAsFixedFormat(pdf, 17)
                finally:
                    d.Close(False)
                P = pymupdf.open(pdf)[0]
                lines = {}
                for b in P.get_text('rawdict')['blocks']:
                    for l in b.get('lines', []):
                        for s in l['spans']:
                            for c in s['chars']:
                                if c['c'].strip():
                                    lines.setdefault(round(s['origin'][1], 1), []).append(c['bbox'][0])
                ys = sorted(lines)
                rule = sorted(round(g['rect'].x0, 2) for g in P.get_drawings() if g['rect'].height > 5)
                print(f'{name:18} {jc:6} rule {rule[0] if rule else None}  ' + '  '.join(
                    f'L{i+1} x0 {min(lines[y]):.2f} n {len(lines[y])}' for i, y in enumerate(ys[:3])), flush=True)
    finally:
        app.Quit()
