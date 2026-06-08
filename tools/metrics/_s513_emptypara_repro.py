# -*- coding: utf-8 -*-
"""S513 empty-para-after-exact-on-docGrid repro: replicate db9ca's title→body region to
derive Word's height for an EMPTY para that follows an EXACT-line title on a docGrid
type=lines. Variants vary the empty para's sz (and a no-exact control) → measure the body's
first baseline via Word PDF → infer the empty-para height. cp932-safe: UTF-8 file, ASCII out."""
import os, sys, zipfile, subprocess, tempfile, io, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'emptypara')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
RPR = lambda sz: '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="%d"/></w:rPr>' % sz


def doc_xml(empty_sz, exline):
    # exline = exact line value in twips (0 = no exact, normal title)
    title = ('<w:p><w:pPr><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr><w:r>%s<w:t>TITLE</w:t></w:r></w:p>'
             % (exline, RPR(28))) if exline else ('<w:p><w:r>%s<w:t>TITLE</w:t></w:r></w:p>' % RPR(28))
    empty = '<w:p><w:pPr>%s</w:pPr></w:p>' % RPR(empty_sz)
    body = '<w:p><w:r>%s<w:t>BODYLINE</w:t></w:r></w:p>' % RPR(21)
    body2 = '<w:p><w:r>%s<w:t>BODYLIN2</w:t></w:r></w:p>' % RPR(21)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s%s%s</w:body></w:document>' % (NS, title, empty, body, body2, sect)


def build(name, esz, exline):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc_xml(esz, exline))
    return p


def word_baselines(dx):
    import win32com.client, pythoncom
    dx = os.path.abspath(dx); pdf = os.path.splitext(dx)[0] + '_rt.pdf'
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(dx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    import fitz
    out = {}
    for blk in fitz.open(pdf)[0].get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                for ch in sp.get('chars', []):
                    if ch['c'] in ('T', 'B'):  # TITLE / BODYLINE / BODYLIN2 first chars
                        txt = ''.join(c['c'] for c in sp['chars'])
                        out.setdefault(txt[:8], ch['origin'][1])
    return out


def main():
    L = ['S513 empty-para-after-exact repro (docGrid lines 360, top margin 70.9pt)']
    # vary exact line value: 420(=21pt,db9ca), 360(=18pt=grid pitch), 357(1ec1), 0(no exact)
    variants = [('ep_ex420.docx', 24, 420), ('ep_ex360.docx', 24, 360),
                ('ep_ex357.docx', 24, 357), ('ep_ex300.docx', 24, 300), ('ep_noexact.docx', 24, 0)]
    for name, esz, exline in variants:
        dx = build(name, esz, exline)
        b = word_baselines(dx)
        ti = b.get('TITLE'); bl = b.get('BODYLINE'); bl2 = b.get('BODYLIN2')
        empty_h = (bl - ti) if (ti and bl) else None
        body_pitch = (bl2 - bl) if (bl and bl2) else None
        L.append('%-18s exline=%4d | TITLE_y=%s BODY_y=%s | title->body=%s body_pitch=%s' % (
            name, exline,
            ('%.1f' % ti if ti else '?'), ('%.1f' % bl if bl else '?'),
            ('%.2f' % empty_h if empty_h else '?'), ('%.2f' % body_pitch if body_pitch else '?')))
    with io.open('c:/tmp/_s513_repro_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('\n'.join(L))


if __name__ == '__main__':
    main()
