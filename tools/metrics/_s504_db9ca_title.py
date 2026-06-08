# -*- coding: utf-8 -*-
"""S504 db9ca title (+2.84pt) root-cause: dump the first N paragraphs' pPr (lineRule,
sz, spacing, the docDefaults) and measure Word vs Oxi baseline Y for the title + the
lines around it, to characterize the first-line leading mechanism. cp932-safe: UTF-8."""
import os, sys, io, json, zipfile, re, subprocess, tempfile
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')


def docx_for(did):
    import glob
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        g = glob.glob(os.path.join(ROOT, d, did + '*.docx'))
        if g:
            return g[0]
    return None


def main():
    dx = docx_for('db9ca18368cd')
    z = zipfile.ZipFile(dx)
    xml = z.read('word/document.xml').decode('utf-8')
    L = ['S504 db9ca title investigation', 'docx=%s' % os.path.basename(dx)]
    # sectPr top margin + docGrid
    m = re.search(r'<w:pgMar [^>]*>', xml)
    L.append('pgMar: %s' % (m.group(0) if m else '?'))
    m = re.search(r'<w:docGrid [^>]*>', xml)
    L.append('docGrid: %s' % (m.group(0) if m else '?'))
    # first 5 paragraphs pPr + first run rPr + text
    paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>', xml, re.S)
    L.append('total body paras: %d' % len(paras))
    for i, p in enumerate(paras[:5]):
        ppr = re.search(r'<w:pPr>(.*?)</w:pPr>', p, re.S)
        spacing = re.search(r'<w:spacing [^>]*>', p)
        rpr = re.search(r'<w:rPr>(.*?)</w:rPr>', p, re.S)
        sz = re.search(r'<w:sz w:val="(\d+)"', p)
        txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.S))
        in_cell = False  # body-level only for first paras
        L.append('--- para %d sz=%s spacing=%s text=%s' % (
            i, sz.group(1) if sz else '?', spacing.group(0) if spacing else 'none', txt[:24]))
    # docDefaults spacing from styles.xml
    try:
        st = z.read('word/styles.xml').decode('utf-8')
        dd = re.search(r'<w:docDefaults>.*?</w:docDefaults>', st, re.S)
        if dd:
            sp = re.search(r'<w:spacing [^>]*>', dd.group(0))
            L.append('docDefaults spacing: %s' % (sp.group(0) if sp else 'none'))
    except Exception as e:
        L.append('styles read err: %s' % e)
    # measure Word vs Oxi for first ~6 distinct baselines on p0
    wj = 'c:/tmp/db9ca183_w.json'
    if not os.path.exists(wj):
        subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), dx, wj],
                       capture_output=True, timeout=200)
    oj = 'c:/tmp/db9ca183_ox.json'
    if not os.path.exists(oj):
        subprocess.run([DW, dx, tempfile.mktemp(dir='c:/tmp'), '150', '--dump-glyphs=' + oj],
                       capture_output=True, timeout=200)
    W = json.load(io.open(wj, encoding='utf-8'))['pages'][0]['glyphs']
    O = json.load(io.open(oj, encoding='utf-8'))['pages'][0]['glyphs']

    def first_baselines(glyphs, ykey, n=6):
        ys = sorted(set(round(g[ykey], 0) for g in glyphs if g['char'].strip()))
        return ys[:n]
    wb = first_baselines(W, 'y')
    ob = first_baselines(O, 'baseline')
    L.append('\nWord first baselines: %s' % wb)
    L.append('Oxi  first baselines: %s' % ob)
    L.append('per-line dy (Oxi-Word): %s' % [round(o - w, 2) for w, o in zip(wb, ob)])
    with io.open('c:/tmp/_s504_db9ca_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('wrote c:/tmp/_s504_db9ca_out.txt')


if __name__ == '__main__':
    main()
