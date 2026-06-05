# -*- coding: utf-8 -*-
"""Extract paragraph indent/style for code-block paragraphs in a docx.
cp932-safe: writes ASCII-only summary (pStyle, ind attrs, font, ascii-prefix of text) to a file.
Identifies code paras by an ASCII-heavy first run (@prefix / void: / rdf: / http)."""
import sys, zipfile, re
import xml.etree.ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def text_of(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def main():
    docx, out = sys.argv[1], sys.argv[2]
    z = zipfile.ZipFile(docx)
    xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)
    # styles
    styles = {}
    try:
        sx = ET.fromstring(z.read('word/styles.xml').decode('utf-8'))
        for st in sx.iter(W + 'style'):
            sid = st.get(W + 'styleId')
            ind = st.find('./' + W + 'pPr/' + W + 'ind')
            rf = st.find('./' + W + 'rPr/' + W + 'rFonts')
            sz = st.find('./' + W + 'rPr/' + W + 'sz')
            styles[sid] = {
                'ind': dict(ind.attrib) if ind is not None else None,
                'rfonts': dict(rf.attrib) if rf is not None else None,
                'sz': sz.get(W + 'val') if sz is not None else None,
                'based': (st.find('./' + W + 'basedOn').get(W + 'val') if st.find('./' + W + 'basedOn') is not None else None),
            }
    except Exception as e:
        styles = {'_err': str(e)}
    lines = []
    seen_styles = {}
    for p in root.iter(W + 'p'):
        txt = text_of(p)
        ascii_part = ''.join(c for c in txt if ord(c) < 128)
        is_code = any(k in txt for k in ('@prefix', 'void:', 'rdf:', 'dcterms:', 'http://', 'foaf:'))
        ppr = p.find(W + 'pPr')
        pstyle = None; ind = None
        rfonts = None; sz = None
        if ppr is not None:
            ps = ppr.find(W + 'pStyle')
            pstyle = ps.get(W + 'val') if ps is not None else None
            indel = ppr.find(W + 'ind')
            ind = dict(indel.attrib) if indel is not None else None
            prpr = ppr.find(W + 'rPr')
            if prpr is not None:
                rf = prpr.find(W + 'rFonts'); rfonts = dict(rf.attrib) if rf is not None else None
        # first run font
        r = p.find(W + 'r')
        if r is not None:
            rp = r.find(W + 'rPr')
            if rp is not None:
                rf = rp.find(W + 'rFonts')
                if rf is not None and rfonts is None:
                    rfonts = dict(rf.attrib)
                szel = rp.find(W + 'sz')
                if szel is not None:
                    sz = szel.get(W + 'val')
        if is_code:
            def clean(d):
                return {k.split('}')[-1]: v for k, v in d.items()} if d else None
            lines.append({'style': pstyle, 'ind': clean(ind), 'rfonts': clean(rfonts),
                          'sz': sz, 'ascii': ascii_part[:30]})
            if pstyle not in seen_styles:
                seen_styles[pstyle] = styles.get(pstyle)
    with open(out, 'w', encoding='utf-8') as f:
        f.write('=== code paragraphs (first 25) ===\n')
        for L in lines[:25]:
            f.write(repr(L) + '\n')
        f.write('\n=== styles used by code paras ===\n')
        for sid, sv in seen_styles.items():
            f.write('%s: %s\n' % (sid, sv))
            # walk basedOn chain
            cur = sv
            chain = []
            while cur and cur.get('based'):
                b = cur['based']; chain.append((b, styles.get(b)))
                cur = styles.get(b)
            for b, bv in chain:
                f.write('   basedOn %s: %s\n' % (b, bv))
    print('wrote', out)


if __name__ == '__main__':
    main()
