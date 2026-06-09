# -*- coding: utf-8 -*-
"""S513: run Oxi dwrite dump-glyphs on the emptypara repros at dpi=72 (px==pt),
extract TITLE / BODYLINE / BODYLIN2 baselines, compare to Word PDF numbers.
ASCII-only output."""
import os, sys, json, subprocess, io
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'emptypara')
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')

WORD = {  # from c:/tmp/_s513_repro_out.txt  (exline -> (TITLE_y, BODY_y, gap, pitch))
    'ep_ex420.docx': (87.7, 122.8),
    'ep_ex360.docx': (85.3, 119.8),
    'ep_ex357.docx': (85.2, 119.7),
    'ep_ex300.docx': (82.9, 116.8),
    'ep_noexact.docx': (85.0, 119.8),
}

def oxi_glyphs(docx):
    out_prefix = os.path.join('c:/tmp', os.path.splitext(os.path.basename(docx))[0] + '_oxi')
    gj = out_prefix + '_glyphs.json'
    r = subprocess.run([EXE, docx, out_prefix, '72', '--dump-glyphs=' + gj],
                       capture_output=True, text=True)
    if not os.path.exists(gj):
        return None, r.stderr[-300:]
    data = json.load(open(gj, encoding='utf-8'))
    return data, None

def first_baseline(glyphs, ch):
    for g in glyphs:
        if g['char'] == ch:
            return g['baseline']
    return None

def main():
    L = ['S513 Oxi(dwrite dpi72) vs Word baselines  (px==pt at 72dpi)']
    L.append('%-18s | TITLE oxi/word d | BODY oxi/word d | t->body oxi/word | bpitch' % 'doc')
    for name, (wt, wb) in WORD.items():
        dx = os.path.join(REPRO, name)
        data, err = oxi_glyphs(dx)
        if data is None:
            L.append('%-18s ERROR %s' % (name, err)); continue
        glyphs = data['pages'][0]['glyphs']
        ot = first_baseline(glyphs, 'T')  # TITLE
        ob = first_baseline(glyphs, 'B')  # BODYLINE
        # BODYLIN2 also starts with B; find second distinct B-line baseline
        bls = sorted({round(g['baseline'], 2) for g in glyphs if g['char'] == 'B'})
        ob = bls[0] if bls else None
        ob2 = bls[1] if len(bls) > 1 else None
        bpitch = (ob2 - ob) if (ob and ob2) else None
        dt = (ot - wt) if ot else None
        db = (ob - wb) if ob else None
        L.append('%-18s | %6.2f/%6.2f %+5.2f | %6.2f/%6.2f %+5.2f | %5.2f/%5.2f | %s' % (
            name,
            ot or 0, wt, dt or 0,
            ob or 0, wb, db or 0,
            (ob - ot) if (ob and ot) else 0, (wb - wt),
            ('%.2f' % bpitch) if bpitch else '?'))
    txt = '\n'.join(L)
    with io.open('c:/tmp/_s513_oxi_out.txt', 'w', encoding='utf-8') as f:
        f.write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
