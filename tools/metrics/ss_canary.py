# -*- coding: utf-8 -*-
"""Supersample canary: render each doc at ss=1/2/4 (dwrite OUTLINE), SSIM vs word_png.
Tests whether supersampling the S460 OUTLINE AA broadly improves Word-match (weight/AA cap)
or regresses already-soft docs. cp932-safe (paths via glob in Python; ASCII out)."""
import json, os, glob, subprocess, tempfile
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

DW = os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
DOCXDIR = 'tools/golden-test/documents/docx'
base = json.load(open('pipeline_data/ssim_baseline.json', encoding='utf-8'))

LO = ['b35123fe8efc_tokumei_08_01', '15076df085f5_tokumei_08_09',
      '1ec1091177b1_006', 'db9ca18368cd_20241122_resource_open_data_01']
CTRL = ['gen2_009', 'gen_japanese', 'test_exact_spacing', 'gen2_034']


def docx_for(docid):
    f = glob.glob(os.path.join(DOCXDIR, docid + '*.docx'))
    return f[0] if f else None


def score_doc(docid):
    dx = docx_for(docid)
    if not dx:
        return None
    wdir = f'pipeline_data/word_png/{docid}'
    if not os.path.isdir(wdir):
        return None
    res = {}
    with tempfile.TemporaryDirectory() as td:
        for ss in (1, 2, 4):
            pref = os.path.join(td, f's{ss}_')
            subprocess.run([DW, dx, pref, '150', f'--supersample={ss}'],
                           capture_output=True, timeout=300)
        # score each page that exists in word_png
        wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
        for ss in (1, 2, 4):
            tot = 0.0; n = 0
            for wp in wpages:
                pgnum = int(os.path.basename(wp)[5:9])
                op = os.path.join(td, f's{ss}__p{pgnum}.png')
                if not os.path.exists(op):
                    continue
                a = np.asarray(Image.open(wp).convert('L'))
                b = Image.open(op).convert('L')
                if b.size != a.shape[::-1]:
                    b = b.resize(a.shape[::-1], Image.LANCZOS)
                tot += ssim(a, np.asarray(b), data_range=255); n += 1
            res[ss] = (tot / n if n else 0, n)
    return res


print("%-44s %7s %7s %7s   %8s %8s" % ('doc', 'ss1', 'ss2', 'ss4', 'd2-1', 'd4-1'))
print('-' * 90)
agg = {'lo': [0, 0, 0], 'ctrl': [0, 0, 0]}
cnt = {'lo': 0, 'ctrl': 0}
for grp, docs in (('lo', LO), ('ctrl', CTRL)):
    for d in docs:
        r = score_doc(d)
        if not r:
            print("%-44s  (no docx/word_png)" % d[:44]); continue
        s1, s2, s4 = r[1][0], r[2][0], r[4][0]
        print("%-44s %7.4f %7.4f %7.4f   %+8.4f %+8.4f" % (d[:44], s1, s2, s4, s2 - s1, s4 - s1))
        agg[grp][0] += s1; agg[grp][1] += s2; agg[grp][2] += s4; cnt[grp] += 1
print('-' * 90)
for grp in ('lo', 'ctrl'):
    if cnt[grp]:
        a = [x / cnt[grp] for x in agg[grp]]
        print("%-44s %7.4f %7.4f %7.4f   %+8.4f %+8.4f" % (
            grp.upper() + ' mean', a[0], a[1], a[2], a[1] - a[0], a[2] - a[0]))
