# -*- coding: utf-8 -*-
"""S516: per-line baseline offset for the FULL db9ca doc (the SSIM-gate doc, renders MS Mincho/
Century — NOT the per-page extract's Yu Mincho), page 1 only. Confirms whether the real gate-doc
vertical residual exists independent of the per-page Yu Mincho extraction artifact.
cp932-safe: UTF-8 file, results to file, ASCII verdict."""
import os, json, subprocess, io, collections
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DOCX = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx', 'db9ca18368cd_20241122_resource_open_data_01.docx')

def word_lines_p1(docx):
    import win32com.client, pythoncom, fitz
    docx = os.path.abspath(docx); pdf = os.path.join('c:/tmp', 'db9ca_FULL_word.pdf')
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    pg = fitz.open(pdf)[0]  # page 1
    lines = []
    for blk in pg.get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            chs = [c for sp in ln.get('spans', []) for c in sp.get('chars', [])]
            txt = ''.join(c['c'] for c in chs).strip()
            if txt:
                lines.append((txt, chs[0]['origin'][1]))
    return lines

def oxi_lines_p1(docx):
    pre = os.path.join('c:/tmp', 'db9ca_FULL_oxi')
    gj = pre + '_glyphs.json'
    subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    g = sorted([x for x in json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs'] if x['char'].strip()],
               key=lambda c: (c['baseline'], c['x']))
    lines = []; cur = []
    for x in g:
        if cur and abs(x['baseline'] - cur[0]['baseline']) > 8:
            lines.append(cur); cur = []
        cur.append(x)
    if cur: lines.append(cur)
    res = []
    for ln in lines:
        bl = collections.Counter(round(c['baseline'] * 2) / 2 for c in ln).most_common(1)[0][0]
        res.append((''.join(c['char'] for c in sorted(ln, key=lambda c: c['x'])).strip(), bl))
    return res

def main():
    wl = word_lines_p1(DOCX); ol = oxi_lines_p1(DOCX)
    L = ['S516 FULL db9ca p1 per-line offset (Oxi dwrite dpi72 - Word PDF), gate doc (MS Mincho)']
    L.append('n_word=%d n_oxi=%d' % (len(wl), len(ol)))
    n = min(len(wl), len(ol))
    for i in range(n):
        wt, wy = wl[i]; ot, oy = ol[i]
        flag = '' if wt[:2] == ot[:2] else ' MISMATCH'
        L.append('%2d w=%7.2f o=%7.2f d=%+6.2f%s' % (i, wy, oy, oy - wy, flag))
    io.open('c:/tmp/_s516_full_db9ca.txt', 'w', encoding='utf-8').write('\n'.join(L) + '\n')
    print('n_word=%d n_oxi=%d -> c:/tmp/_s516_full_db9ca.txt' % (len(wl), len(ol)))

if __name__ == '__main__':
    main()
