# -*- coding: utf-8 -*-
"""Find ALL baseline docs whose glyph layout changes with OXI_S497_CELL_HANG.
Dumps glyphs ON vs OFF per doc, reports docs that differ (+ page count delta).
cp932-safe. Fast (glyph dump only, no PNG)."""
import os, json, glob, subprocess, tempfile, io
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
base = json.load(io.open(os.path.join(ROOT, 'pipeline_data', 'ssim_baseline.json'), encoding='utf-8'))

def docx_for(did):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT, d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did:
                return p
        g = glob.glob(os.path.join(ROOT, d, did + '*.docx'))
        if g: return g[0]
    g = glob.glob(os.path.join(ROOT, 'pipeline_data', '**', did + '*.docx'), recursive=True)
    return g[0] if g else None

def dump(dx, on):
    env = dict(os.environ)
    if on: env['OXI_S497_CELL_HANG'] = '1'
    else: env.pop('OXI_S497_CELL_HANG', None)
    gj = tempfile.mktemp(suffix='.json')
    try:
        subprocess.run([DW, os.path.abspath(dx), tempfile.mktemp(), '150', '--dump-glyphs=' + gj],
                       capture_output=True, timeout=300, env=env)
        return json.load(open(gj, encoding='utf-8'))
    except Exception:
        return None
    finally:
        if os.path.exists(gj): os.unlink(gj)

def sig(d):
    """compact signature per page: (nglyphs, sum of x+top rounded)."""
    if not d: return None
    return [(len(p['glyphs']), round(sum(g['x'] + g['top'] for g in p['glyphs']), 1)) for p in d['pages']]

affected = []
ids = sorted(base)
for i, did in enumerate(ids):
    dx = docx_for(did)
    if not dx: continue
    off = sig(dump(dx, False)); on = sig(dump(dx, True))
    if off != on:
        # page count delta
        pcd = (len(on) - len(off)) if (on and off) else '?'
        affected.append((did, pcd))
    if (i + 1) % 40 == 0:
        print('...%d/%d  affected so far %d' % (i + 1, len(ids), len(affected)), flush=True)
print('\n=== S497 affected docs: %d ===' % len(affected))
for did, pcd in affected:
    print('  %-44s page_count_delta=%s' % (did[:44], pcd))
print('DONE')
