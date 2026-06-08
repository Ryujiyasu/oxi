# -*- coding: utf-8 -*-
"""S502 full-corpus ss3 gate: compare current (post-S502) ss3 render to the ss3
ssim_baseline.json for ALL docs. Reports mean delta, regressions>0.005, improvements,
and bottom-10 sum. Non-destructive. cp932-safe."""
import os, glob, subprocess, tempfile, io, json
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
BASE = json.load(io.open(os.path.join(ROOT, 'pipeline_data', 'ssim_baseline.json'), encoding='utf-8'))


def docx_for(did):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT, d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did:
                return p
        g = glob.glob(os.path.join(ROOT, d, did + '*.docx'))
        if g:
            return g[0]
    g = glob.glob(os.path.join(ROOT, 'pipeline_data', '**', did + '*.docx'), recursive=True)
    return g[0] if g else None


def rgb(wp, op):
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))


def ss3_mean(dx, did):
    wdir = os.path.join(ROOT, 'pipeline_data', 'word_png', did)
    wp = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    if not wp:
        return None
    with tempfile.TemporaryDirectory() as td:
        base = os.path.join(td, 'p.png')
        subprocess.run([DW, os.path.abspath(dx), base, '150', '--supersample=3'], capture_output=True, timeout=600)
        oxi = {int(os.path.basename(f).split('_p')[-1].split('.')[0]): f
               for f in glob.glob(os.path.join(td, 'p.png_p*.png'))}
        v = [rgb(w, oxi[int(os.path.basename(w)[5:9])]) for w in wp
             if int(os.path.basename(w)[5:9]) in oxi]
    return sum(v) / len(v) if v else None


def main():
    keys = sorted(BASE.keys())
    rows = []
    n = 0
    for k in keys:
        dx = docx_for(k)
        if not dx:
            continue
        m_base = sum(BASE[k].values()) / len(BASE[k])
        m_new = ss3_mean(dx, k)
        if m_new is None:
            continue
        rows.append((k, m_base, m_new))
        n += 1
        if n % 20 == 0:
            print('...%d/%d' % (n, len(keys)))
    base_sum = sum(b for _, b, _ in rows)
    new_sum = sum(n for _, _, n in rows)
    regress = sorted([(nn - bb, k) for k, bb, nn in rows if nn < bb - 0.005])
    improve = sorted([(nn - bb, k) for k, bb, nn in rows if nn > bb + 0.005], reverse=True)
    bot = sorted(rows, key=lambda r: r[1])[:10]
    out = []
    out.append('=== S502 FULL GATE (ss3 doc-mean, %d docs) ===' % len(rows))
    out.append('mean base=%.4f -> new=%.4f  delta=%+.5f' % (
        base_sum / len(rows), new_sum / len(rows), (new_sum - base_sum) / len(rows)))
    out.append('regress>0.005: %d   improve>0.005: %d' % (len(regress), len(improve)))
    out.append('-- regressions --')
    for d, k in regress[:15]:
        out.append('  %+.4f  %s' % (d, k))
    out.append('-- improvements --')
    for d, k in improve[:15]:
        out.append('  %+.4f  %s' % (d, k))
    bb = sum(b for _, b, _ in bot); bn = sum(n for _, _, n in bot)
    out.append('bottom-10 (by base) sum: base=%.4f -> new=%.4f  delta=%+.4f' % (bb, bn, bn - bb))
    for k, b, nn in bot:
        out.append('  %-40s %.4f -> %.4f  %+.4f' % (k[:40], b, nn, nn - b))
    txt = '\n'.join(out)
    print(txt)
    with io.open('c:/tmp/_s502_fullgate_out.txt', 'w', encoding='utf-8') as f:
        f.write(txt + '\n')


if __name__ == '__main__':
    main()
