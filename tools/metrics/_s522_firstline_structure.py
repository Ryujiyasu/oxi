# -*- coding: utf-8 -*-
"""S522: characterize the first-line directional offset across bottom-N docs + correlate with the
first-paragraph STRUCTURE. For each doc: Oxi dump baseline (p1 top text line) vs Word PDF baseline,
delta = Oxi - Word. Extract first body paragraph's: font size, spacing before (sa/sb/beforeLines),
lineRule (exact/auto/atLeast), docGrid (type/linePitch), pStyle, top margin. Look for a rule that
predicts the sign/magnitude. cp932-safe: UTF-8 file, results to file, ASCII out.
NOTE: PDF-based (per-glyph instrument). Flag any rule for SCREENSHOT validation (S494c caveat)."""
import os, sys, json, subprocess, io, re, zipfile
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DOCX = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')

DOCS = ['1ec1091177b1_006', 'b35123fe8efc_tokumei_08_01', 'b837808d0555_20240705_resources_data_guideline_02',
        '683ffcab86e2_20230331_resources_open_data_contract_addon_00', 'd77a58485f16_20240705_resources_data_outline_08',
        '15076df085f5_tokumei_08_09', '0e7af8...', 'db9ca18368cd_20241122_resource_open_data_01',
        'a47e6c6b2ca1_order_08', '2ea81a8441cc_0025006-192', 'de6e32b5960b_tokumei_08_01-1',
        'a1d6e4efa2e7_tokumei_08_01-4', 'e3c545fac7a7_LOD_Handbook', '34140b9c5662_index-14']

def resolve(stem):
    import glob
    g = glob.glob(os.path.join(DOCX, stem.split('...')[0] + '*.docx'))
    return g[0] if g else None

def oxi_first_baseline(docx):
    pre = os.path.join('c:/tmp', 's522_' + os.path.basename(docx)[:12])
    gj = pre + '_g.json'
    subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    g = [x for x in json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs'] if x['char'].strip()]
    if not g:
        return None, None
    top = min(g, key=lambda c: c['baseline'])
    # the first line = glyphs within 3pt of the top baseline; report its leftmost glyph baseline + fs
    line = [c for c in g if abs(c['baseline'] - top['baseline']) < 3]
    lead = min(line, key=lambda c: c['x'])
    return lead['baseline'], lead['font_size']

def word_first_baseline(docx):
    import win32com.client, pythoncom, fitz
    pdf = os.path.join('c:/tmp', os.path.basename(docx)[:12] + '_w.pdf')
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    chs = []
    for blk in fitz.open(pdf)[0].get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                for c in sp.get('chars', []):
                    if c['c'].strip():
                        chs.append((c['origin'][1], c['origin'][0]))
    if not chs:
        return None
    top = min(c[0] for c in chs)
    line = [c for c in chs if abs(c[0] - top) < 3]
    return min(line, key=lambda c: c[1])[0]

def structure(docx):
    z = zipfile.ZipFile(docx)
    xml = z.read('word/document.xml').decode('utf-8', 'ignore')
    dg = re.search(r'<w:docGrid w:type="(\w+)"[^>]*w:linePitch="(\d+)"', xml)
    grid = '%s/%s' % (dg.group(1)[:4], dg.group(2)) if dg else 'none'
    tm = re.search(r'<w:pgMar[^>]*w:top="(-?\d+)"', xml)
    topm = int(tm.group(1)) / 20.0 if tm else 0
    # first paragraph WITH text
    paras = re.findall(r'<w:p\b.*?</w:p>', xml, re.S)
    first = None
    for p in paras:
        if re.search(r'<w:t[^>]*>[^<]', p):
            first = p; break
    info = {'grid': grid, 'topm': round(topm, 1)}
    if first:
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', first, re.S)
        pprs = ppr.group(0) if ppr else ''
        info['pStyle'] = (re.search(r'<w:pStyle w:val="([^"]+)"', pprs) or [None, '-'])[1]
        info['exact'] = (re.search(r'<w:lineRule w:val="(\w+)"', pprs) or re.search(r'w:lineRule="(\w+)"', pprs) or [None, 'auto'])[1]
        sp = re.search(r'<w:spacing([^/]*)/>', pprs)
        info['spacing'] = sp.group(1).strip() if sp else '-'
        sz = re.search(r'<w:sz w:val="(\d+)"', first)
        info['sz'] = sz.group(1) if sz else '?'
    return info

def main():
    L = ['S522 first-line offset (Oxi dump - Word PDF) + first-para structure  (PDF-based; SCREENSHOT-validate any rule)']
    L.append('%-26s %7s %5s | %-12s %4s %-7s %-8s %s' % ('doc', 'dOxiW', 'fs', 'grid', 'topm', 'exact', 'pStyle', 'spacing'))
    for stem in DOCS:
        dx = resolve(stem)
        if not dx:
            L.append('%-26s (not found)' % stem[:26]); continue
        ob, ofs = oxi_first_baseline(dx)
        wb = word_first_baseline(dx)
        st = structure(dx)
        d = (ob - wb) if (ob and wb) else None
        L.append('%-26s %+7.2f %5s | %-12s %4s %-7s %-8s %s' % (
            os.path.basename(dx)[:26], d if d is not None else 0, str(ofs),
            st.get('grid'), st.get('topm'), st.get('exact'), str(st.get('pStyle'))[:8], st.get('spacing', '-')[:30]))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s522_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    for line in txt.split('\n'):
        try: print(line)
        except Exception: print(line.encode('ascii', 'replace').decode())

if __name__ == '__main__':
    main()
