# -*- coding: utf-8 -*-
"""For each PreformattedText paragraph: report numPr(numId/ilvl), ind, leading-whitespace
char codes, and ascii prefix. cp932-safe ASCII-only output to a file."""
import sys, zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def main():
    docx, out = sys.argv[1], sys.argv[2]
    z = zipfile.ZipFile(docx)
    root = ET.fromstring(z.read('word/document.xml').decode('utf-8'))
    rows = []
    for p in root.iter(W + 'p'):
        ppr = p.find(W + 'pPr')
        if ppr is None:
            continue
        ps = ppr.find(W + 'pStyle')
        if ps is None or ps.get(W + 'val') != 'PreformattedText':
            continue
        npr = ppr.find(W + 'numPr')
        numid = ilvl = None
        if npr is not None:
            n = npr.find(W + 'numId'); il = npr.find(W + 'ilvl')
            numid = n.get(W + 'val') if n is not None else None
            ilvl = il.get(W + 'val') if il is not None else None
        ind = ppr.find(W + 'ind')
        indd = {k.split('}')[-1]: v for k, v in ind.attrib.items()} if ind is not None else None
        txt = ''.join(t.text or '' for t in p.iter(W + 't'))
        # leading whitespace codes
        lead = []
        for c in txt:
            if c in (' ', '　', '\t'):
                lead.append('U+%04X' % ord(c))
            else:
                break
        asc = ''.join(c for c in txt if ord(c) < 128)[:24]
        rows.append((numid, ilvl, indd, lead, asc))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('PreformattedText paras: %d\n' % len(rows))
        f.write('numId ilvl  ind  | leadWS | ascii\n')
        for numid, ilvl, indd, lead, asc in rows[:40]:
            f.write('num=%s/%s ind=%s lead=%s | %s\n' % (numid, ilvl, indd, lead, asc))
    print('wrote', out)


if __name__ == '__main__':
    main()
