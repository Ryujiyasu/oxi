# -*- coding: utf-8 -*-
"""S492q — build the bottom-N lowest-SSIM page target set for a fresh no-ROI lever
scout. From ssim_baseline.json pick the lowest-SSIM pages; for each, locate the docx +
the cached Word PNG and render the Oxi (DWrite) page PNG. Output targets.json for the
Workflow. cp932-safe (UTF-8 file)."""
import os, glob, json, subprocess, shutil
from pathlib import Path

ROOT = Path(os.path.abspath('.'))
RENDERER = ROOT / 'tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe'
DOCX_DIR = ROOT / 'tools/golden-test/documents/docx'
WORD_PNG = ROOT / 'pipeline_data/word_png'
STAGE = Path('c:/tmp/bottomN'); STAGE.mkdir(parents=True, exist_ok=True)

base = json.load(open(ROOT / 'pipeline_data/ssim_baseline.json', encoding='utf-8'))
pages = []
for doc, pm in base.items():
    if not isinstance(pm, dict):
        continue
    for pg, s in pm.items():
        try:
            pages.append((doc, int(pg), float(s)))
        except Exception:
            pass
pages.sort(key=lambda t: t[2])
bottom = pages[:16]

# map doc stem -> docx path
def find_docx(doc):
    stem = doc.split('_')[0]
    g = glob.glob(str(DOCX_DIR / (stem + '*.docx')))
    return g[0] if g else None

def find_word_png(doc, page):
    # word_png/<doc>/page_NNNN.png
    d = WORD_PNG / doc
    cands = [d / ('page_%04d.png' % page), d / ('page_%d.png' % page)]
    for c in cands:
        if c.exists():
            return str(c)
    # try any dir starting with stem
    stem = doc.split('_')[0]
    for dd in glob.glob(str(WORD_PNG / (stem + '*'))):
        for nm in ('page_%04d.png' % page, 'page_%d.png' % page):
            p = Path(dd) / nm
            if p.exists():
                return str(p)
    return None

# render each unique doc once (DWrite), then pick the page PNG
targets = []
rendered = {}
for doc, page, ssim in bottom:
    docx = find_docx(doc)
    if not docx:
        continue
    if docx not in rendered:
        prefix = str(STAGE / ('oxi_' + doc.split('_')[0]))
        subprocess.run([str(RENDERER), docx, prefix, '150'],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        rendered[docx] = prefix
    prefix = rendered[docx]
    oxi_png = prefix + ('_p%d.png' % page)
    if not os.path.exists(oxi_png):
        continue
    wpng = find_word_png(doc, page)
    targets.append({'doc': doc, 'doc_id': doc.split('_')[0], 'page': page,
                    'ssim': round(ssim, 4), 'docx': os.path.abspath(docx),
                    'oxi_png': os.path.abspath(oxi_png),
                    'word_png': os.path.abspath(wpng) if wpng else None})

json.dump(targets, open(str(STAGE / 'targets.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print("bottom-N targets:", len(targets))
for t in targets:
    print("  %-14s p%-2d ssim=%.4f  oxi=%s word=%s" %
          (t['doc_id'], t['page'], t['ssim'], 'Y' if os.path.exists(t['oxi_png']) else 'N',
           'Y' if t['word_png'] and os.path.exists(t['word_png']) else 'N'))
print("targets ->", STAGE / 'targets.json')
