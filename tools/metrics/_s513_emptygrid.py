# -*- coding: utf-8 -*-
"""S513 empty-para advance on docGrid type=lines: clean repro (LINE1 + N empty + LINE2),
vary linePitch L, measure LINE1->LINE2 gap to derive per-empty-para advance vs L. Also
measure the normal body pitch (= L). Derives the empty-para-on-docGrid height formula so
db9ca's 18-vs-20 deficit can be fixed WITHOUT hardcoding. cp932-safe."""
import os, zipfile, subprocess, io, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'emptypara')
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr>'


def doc_xml(L, nempty):
    body1 = '<w:p><w:r>%s<w:t>AAAA</w:t></w:r></w:p>' % RPR
    empty = ('<w:p><w:pPr>%s</w:pPr></w:p>' % RPR) * nempty
    body2 = '<w:p><w:r>%s<w:t>BBBB</w:t></w:r></w:p>' % RPR
    body3 = '<w:p><w:r>%s<w:t>CCCC</w:t></w:r></w:p>' % RPR
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>'
            '<w:docGrid w:type="lines" w:linePitch="%d"/></w:sectPr>' % L)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s%s%s</w:body></w:document>' % (NS, body1, empty, body2, body3, sect)


def build(name, L, ne):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc_xml(L, ne))
    return p


def word_bl(dx):
    import win32com.client, pythoncom, fitz
    dx = os.path.abspath(dx); pdf = os.path.splitext(dx)[0] + '_rt.pdf'
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(dx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    out = {}
    for blk in fitz.open(pdf)[0].get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                c0 = sp['chars'][0]['c'] if sp.get('chars') else ''
                if c0 in 'ABC':
                    out.setdefault(c0, sp['chars'][0]['origin'][1])
    return out


def main():
    L = ['S513 empty-para advance on docGrid lines (Lpt = L/20)']
    for lp in [300, 360, 400, 480]:
        # nempty=1 and nempty=3 to isolate per-empty advance
        b1 = word_bl(build('eg_L%d_n1.docx' % lp, lp, 1))
        b3 = word_bl(build('eg_L%d_n3.docx' % lp, lp, 3))
        # body pitch = BBBB->CCCC (no empty between) in n1 doc
        pitch = (b1['C'] - b1['B']) if ('C' in b1 and 'B' in b1) else None
        gap1 = (b1['B'] - b1['A']) if ('A' in b1 and 'B' in b1) else None  # 1 empty between
        gap3 = (b3['B'] - b3['A']) if ('A' in b3 and 'B' in b3) else None  # 3 empty between
        per_empty = ((gap3 - gap1) / 2.0) if (gap1 and gap3) else None  # marginal per-empty
        Lpt = lp / 20.0
        L.append('linePitch=%d (%.1fpt) | body_pitch=%s | gap_1empty=%s gap_3empty=%s | per_empty_advance=%s' % (
            lp, Lpt, ('%.2f' % pitch if pitch else '?'),
            ('%.2f' % gap1 if gap1 else '?'), ('%.2f' % gap3 if gap3 else '?'),
            ('%.2f' % per_empty if per_empty else '?')))
    with io.open('c:/tmp/_s513_eg_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('\n'.join(L))


if __name__ == '__main__':
    main()
