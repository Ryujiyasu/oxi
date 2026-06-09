# -*- coding: utf-8 -*-
"""S513 GROUND TRUTH: per-text-line vertical offset (Oxi dwrite baseline - Word PDF baseline)
for the REAL db9ca p1. Exports db9ca p1 to PDF via Word, extracts each text line's first-char
baseline, runs Oxi --dump-glyphs at dpi=72, groups Oxi glyphs into lines by baseline, aligns
in reading order with a text-prefix sanity check. cp932-safe: UTF-8 file, results to file, ASCII verdicts."""
import os, json, subprocess, io
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DOCX = os.path.join(ROOT, 'pipeline_data', 'golden_per_page', 'db9ca18368cd_20241122_resource_open_data_01_p1.docx')

def word_lines(docx):
    import win32com.client, pythoncom, fitz
    docx = os.path.abspath(docx); pdf = os.path.join('c:/tmp', 'db9ca_p1_word.pdf')
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    lines = []
    pg = fitz.open(pdf)[0]
    for blk in pg.get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            chars = []
            for sp in ln.get('spans', []):
                chars.extend(sp.get('chars', []))
            if not chars:
                continue
            txt = ''.join(c['c'] for c in chars).strip()
            if txt:
                lines.append((txt, chars[0]['origin'][1]))
    return lines

def oxi_lines(docx):
    out_prefix = os.path.join('c:/tmp', 'db9ca_p1_oxi')
    gj = out_prefix + '_glyphs.json'
    subprocess.run([EXE, os.path.abspath(docx), out_prefix, '72', '--dump-glyphs=' + gj],
                   capture_output=True, text=True)
    data = json.load(open(gj, encoding='utf-8'))
    glyphs = data['pages'][0]['glyphs']
    # group by baseline (rounded to 0.5pt)
    groups = {}
    for g in glyphs:
        key = round(g['baseline'] * 2) / 2
        groups.setdefault(key, []).append(g)
    lines = []
    for bl in sorted(groups):
        gs = sorted(groups[bl], key=lambda g: g['x'])
        txt = ''.join(g['char'] for g in gs).strip()
        if txt:
            lines.append((txt, bl))
    return lines

def main():
    wl = word_lines(DOCX)
    ol = oxi_lines(DOCX)
    L = ['S513 db9ca p1 per-line baseline offset (Oxi dwrite dpi72 - Word PDF), px==pt']
    L.append('n_word=%d n_oxi=%d' % (len(wl), len(ol)))
    L.append('%3s %8s %8s %7s  %-20s' % ('#', 'word_y', 'oxi_y', 'oxi-wd', 'text[:14]'))
    n = min(len(wl), len(ol))
    for i in range(n):
        wt, wy = wl[i]; ot, oy = ol[i]
        # sanity: do first 2 chars match?
        match = '' if (wt[:2] == ot[:2]) else ' <MISMATCH>'
        L.append('%3d %8.2f %8.2f %+7.2f  %-20s%s' % (i, wy, oy, oy - wy, wt[:14], match))
    if len(wl) != len(ol):
        L.append('LINE COUNT DIFF: word=%d oxi=%d (alignment past min may be invalid)' % (len(wl), len(ol)))
    txt = '\n'.join(L)
    with io.open('c:/tmp/_s513_db9ca_lines.txt', 'w', encoding='utf-8') as f:
        f.write(txt + '\n')
    # print ASCII-safe summary only
    print('n_word=%d n_oxi=%d' % (len(wl), len(ol)))
    print('wrote c:/tmp/_s513_db9ca_lines.txt')

if __name__ == '__main__':
    main()
