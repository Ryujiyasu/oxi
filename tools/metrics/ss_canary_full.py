# -*- coding: utf-8 -*-
"""Full-corpus supersample canary: render every baseline doc at ss=1 and ss=2 (dwrite OUTLINE),
SSIM vs word_png. Decides if global ss2 is a net Phase-3 win (bottom-N strictly up, mean not
down >0.005, no doc regress >0.005). Writes c:/tmp/ss_full_result.json incrementally. cp932-safe."""
import json, os, glob, subprocess, tempfile, sys
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

DW = os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
DOCXDIR = 'tools/golden-test/documents/docx'
base = json.load(open('pipeline_data/ssim_baseline.json', encoding='utf-8'))
OUT = 'c:/tmp/ss_full_result.json'

docids = [k for k in base if os.path.isdir(f'pipeline_data/word_png/{k}')]
results = {}


def docx_for(docid):
    f = glob.glob(os.path.join(DOCXDIR, docid + '*.docx'))
    return f[0] if f else None


def doc_means(docid):
    dx = docx_for(docid)
    if not dx:
        return None
    wdir = f'pipeline_data/word_png/{docid}'
    wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    if not wpages:
        return None
    out = {}
    with tempfile.TemporaryDirectory() as td:
        for ss in (1, 2):
            pref = os.path.join(td, f's{ss}_')
            try:
                subprocess.run([DW, dx, pref, '150', f'--supersample={ss}'],
                               capture_output=True, timeout=300)
            except Exception:
                return None
        for ss in (1, 2):
            scs = []
            for wp in wpages:
                pg = int(os.path.basename(wp)[5:9])
                op = os.path.join(td, f's{ss}__p{pg}.png')
                if not os.path.exists(op):
                    continue
                a = np.asarray(Image.open(wp).convert('L'))
                b = Image.open(op).convert('L')
                if b.size != a.shape[::-1]:
                    b = b.resize(a.shape[::-1], Image.LANCZOS)
                scs.append(float(ssim(a, np.asarray(b), data_range=255)))
            out[ss] = scs
    return out


for i, d in enumerate(docids):
    r = doc_means(d)
    if not r or not r.get(1):
        continue
    m1 = sum(r[1]) / len(r[1])
    m2 = sum(r[2]) / len(r[2]) if r.get(2) and len(r[2]) == len(r[1]) else None
    results[d] = {'ss1': m1, 'ss2': m2, 'n': len(r[1])}
    if (i + 1) % 20 == 0:
        json.dump(results, open(OUT, 'w'), indent=0)
        print("...%d/%d done" % (i + 1, len(docids)), flush=True)

json.dump(results, open(OUT, 'w'), indent=0)
# summarize
vals = [(d, v['ss1'], v['ss2']) for d, v in results.items() if v['ss2'] is not None]
n = len(vals)
mean1 = sum(v[1] for v in vals) / n
mean2 = sum(v[2] for v in vals) / n
deltas = sorted((v[2] - v[1], v[0]) for v in vals)
reg = [(dl, d) for dl, d in deltas if dl < -0.005]
up = [(dl, d) for dl, d in deltas if dl > 0.005]
b1 = sorted(v[1] for v in vals)[:10]
# bottom-10 by ss1, their ss2
bottomdocs = sorted(vals, key=lambda v: v[1])[:10]
print("\n=== FULL SS CANARY (ss1 vs ss2), n=%d docs ===" % n)
print("mean ss1=%.4f  ss2=%.4f  delta=%+.4f" % (mean1, mean2, mean2 - mean1))
print("docs regress >0.005: %d   improve >0.005: %d" % (len(reg), len(up)))
print("worst regressions:", [(round(dl, 4), d[:30]) for dl, d in reg[:8]])
print("bottom-10 docs (by ss1) ss1->ss2:")
for d, s1, s2 in bottomdocs:
    print("  %-40s %.4f -> %.4f  %+.4f" % (d[:40], s1, s2, s2 - s1))
b1s = sum(s for d, s, _ in bottomdocs); b2s = sum(s2 for d, _, s2 in bottomdocs)
print("bottom-10 sum: ss1=%.4f ss2=%.4f  delta=%+.4f" % (b1s, b2s, b2s - b1s))
print("DONE")
