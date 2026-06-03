# -*- coding: utf-8 -*-
"""Refresh ssim_baseline.json at the new dwrite default (ss2). Renders each baseline doc
full-doc at 150dpi (NO --supersample flag => uses the new ss2 default), computes per-PAGE
canonical RGB SSIM vs word_png (matches pipeline/ssim_calculator.py). Backs up the old baseline,
writes the new one, prints the Phase-3 gate comparison (mean, bottom-N sum, regress/improve).
Docs that fail to render keep their OLD baseline values. cp932-safe."""
import json, os, glob, subprocess, tempfile, shutil
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

DW = os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
DOCXDIRS = ['pipeline_data/docx', 'tools/golden-test/documents/docx']
BASE = 'pipeline_data/ssim_baseline.json'
old = json.load(open(BASE, encoding='utf-8'))


def docx_for(docid):
    for d in DOCXDIRS:
        f = glob.glob(os.path.join(d, docid + '*.docx'))
        # prefer exact stem match
        for p in f:
            if os.path.splitext(os.path.basename(p))[0] == docid:
                return p
        if f:
            return f[0]
    return None


def rgb_ssim(wp, op):
    a = np.array(Image.open(wp).convert('RGB'))
    b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]):
        b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))


new = {}
fail = []
for i, docid in enumerate(sorted(old)):
    wdir = f'pipeline_data/word_png/{docid}'
    dx = docx_for(docid)
    wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png'))) if os.path.isdir(wdir) else []
    if not dx or not wpages:
        new[docid] = old[docid]  # keep old
        continue
    with tempfile.TemporaryDirectory() as td:
        pref = os.path.join(td, 'o')
        try:
            subprocess.run([DW, os.path.abspath(dx), pref, '150'],
                           capture_output=True, timeout=400)
        except Exception:
            new[docid] = old[docid]; fail.append(docid); continue
        pages = {}
        for wp in wpages:
            pg = int(os.path.basename(wp)[5:9])
            op = os.path.join(td, f'o_p{pg}.png')
            if not os.path.exists(op):
                # keep old value for missing page
                if str(pg) in old[docid]:
                    pages[str(pg)] = old[docid][str(pg)]
                continue
            pages[str(pg)] = rgb_ssim(wp, op)
        new[docid] = pages if pages else old[docid]
    if (i + 1) % 20 == 0:
        print('...%d/%d' % (i + 1, len(old)), flush=True)

# gate comparison (per-doc mean + per-page)
def docmean(d):
    return sum(d.values()) / len(d) if d else 0
om = {k: docmean(old[k]) for k in old}
nm = {k: docmean(new[k]) for k in new}
common = [k for k in old if k in new and old[k] and new[k]]
mean_o = sum(om[k] for k in common) / len(common)
mean_n = sum(nm[k] for k in common) / len(common)
deltas = sorted((nm[k] - om[k], k) for k in common)
reg = [(d, k) for d, k in deltas if d < -0.005]
up = [(d, k) for d, k in deltas if d > 0.005]
bot = sorted(common, key=lambda k: om[k])[:10]
bo = sum(om[k] for k in bot); bn = sum(nm[k] for k in bot)
print('\n=== SS2 BASELINE REFRESH GATE (doc-mean) ===')
print('docs=%d  mean ss1=%.4f -> ss2=%.4f  delta=%+.4f' % (len(common), mean_o, mean_n, mean_n - mean_o))
print('regress>0.005: %d   improve>0.005: %d' % (len(reg), len(up)))
print('worst reg:', [(round(d, 4), k[:28]) for d, k in reg[:8]])
print('bottom-10 (by old) sum: ss1=%.4f -> ss2=%.4f  delta=%+.4f' % (bo, bn, bn - bo))
for k in bot:
    print('  %-40s %.4f -> %.4f  %+.4f' % (k[:40], om[k], nm[k], nm[k] - om[k]))
if fail:
    print('render-failed (kept old):', len(fail), fail[:5])

shutil.copy(BASE, BASE + '.ss1_bak')
json.dump(new, open(BASE, 'w', encoding='utf-8'), indent=2, ensure_ascii=False)
allv = [v for d in new.values() for v in d.values()]
print('\nwrote %s (backup .ss1_bak). new all-page mean=%.4f' % (BASE, sum(allv) / len(allv)))
print('DONE')
