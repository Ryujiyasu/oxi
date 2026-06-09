# -*- coding: utf-8 -*-
"""S517: inspect b837 paragraph(s) containing circled markers (U+2460.. = (1)(2)(3)) -- dump the
run structure (per-run rPr font/sz + text) to see why the marker run gets text_y_off=0 while the
body gets 4.0. cp932-safe: UTF-8 file, results to file, ASCII out (codepoints not raw JP)."""
import os, re, io, zipfile
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx',
                    'b837808d0555_20240705_resources_data_guideline_02.docx')

CIRCLED = set(range(0x2460, 0x2474))  # circled 1..20

def run_info(run_xml):
    rpr = re.search(r'<w:rPr>.*?</w:rPr>', run_xml, re.S)
    rprs = rpr.group(0) if rpr else ''
    rfonts = re.search(r'<w:rFonts[^/]*/>', rprs)
    sz = re.search(r'<w:sz w:val="(\d+)"', rprs)
    szcs = re.search(r'<w:szCs w:val="(\d+)"', rprs)
    pos = re.search(r'<w:position w:val="(-?\d+)"', rprs)
    valign = re.search(r'<w:vertAlign w:val="(\w+)"', rprs)
    txts = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', run_xml))
    return {
        'rfonts': rfonts.group(0) if rfonts else '-',
        'sz': sz.group(1) if sz else '-',
        'szCs': szcs.group(1) if szcs else '-',
        'position': pos.group(1) if pos else '-',
        'vertAlign': valign.group(1) if valign else '-',
        'text_cps': [hex(ord(c)) for c in txts[:6]],
        'has_circled': any(ord(c) in CIRCLED for c in txts),
        'len': len(txts),
    }

def main():
    z = zipfile.ZipFile(DOCX)
    xml = z.read('word/document.xml').decode('utf-8')
    paras = re.findall(r'<w:p\b.*?</w:p>', xml, re.S)
    L = ['S517 b837 circled-marker paragraph run structure']
    found = 0
    for pi, p in enumerate(paras):
        alltxt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))
        if not any(ord(c) in CIRCLED for c in alltxt):
            continue
        found += 1
        if found > 3:
            break
        # paragraph pPr
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
        pprs = ppr.group(0) if ppr else ''
        pstyle = re.search(r'<w:pStyle w:val="([^"]+)"', pprs)
        numpr = '<w:numPr>' in pprs
        spacing = re.search(r'<w:spacing[^/]*/>', pprs)
        L.append('')
        L.append('--- para#%d  pStyle=%s numPr=%s spacing=%s  (%d chars)' % (
            pi, pstyle.group(1) if pstyle else '-', numpr,
            spacing.group(0) if spacing else '-', len(alltxt)))
        runs = re.findall(r'<w:r\b.*?</w:r>', p, re.S)
        for ri, r in enumerate(runs[:6]):
            info = run_info(r)
            L.append('  run%d circled=%s sz=%s szCs=%s pos=%s vAlign=%s rfonts=%s text=%s len=%d' % (
                ri, info['has_circled'], info['sz'], info['szCs'], info['position'],
                info['vertAlign'], info['rfonts'], info['text_cps'], info['len']))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s517_marker.txt', 'w', encoding='utf-8').write(txt + '\n')
    for line in txt.split('\n'):
        try: print(line)
        except Exception: print(line.encode('ascii', 'replace').decode())

if __name__ == '__main__':
    main()
