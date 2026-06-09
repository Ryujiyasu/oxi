# -*- coding: utf-8 -*-
"""S517 targeted gate: re-render + SSIM the 10 numPr-list-marker docs vs baseline (the only docs
the S517 marker-text_y_off fix can change; all others are byte-identical). Reuses the pipeline
machinery. Clears stale oxi_png for the affected docs first (CLAUDE.md hygiene). Reports
per-page improved/unchanged/regressed vs ssim_baseline. Does NOT mutate the baseline."""
import os, sys, glob, shutil, json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi, OXI_PNG_DIR
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline

AFFECTED = [
    '3a4f9fbe1a83_001620506',
    '459f05f1e877_kyodokenkyuyoushiki01',
    'b35123fe8efc_tokumei_08_01',
    'b5f706e9f6ad_kyodokenkyuyoushiki_bessi',
    'b837808d0555_20240705_resources_data_guideline_02',
    'd77a58485f16_20240705_resources_data_outline_08',
    'd7_v4_bullet',
    'd7_v5_no_indent_bullet',
    'e3c545fac7a7_LOD_Handbook',
    'ed025cbecffb_index-23',
]

def main():
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    paths = []
    for stem in AFFECTED:
        p = docx_dir / (stem + '.docx')
        if p.exists():
            paths.append(str(p))
        # clear stale oxi PNGs so the new binary re-renders
        cache = Path(OXI_PNG_DIR) / stem
        if cache.exists():
            shutil.rmtree(cache)
    print('rendering %d docs (cleared oxi_png)...' % len(paths))
    baseline = load_baseline()
    word = render_with_word(paths)
    oxi = render_with_oxi(paths)
    scores = calculate_ssim(word, oxi, skip_heatmap=True)
    imp = []; reg = []; unch = 0
    for s in scores:
        did = s['doc_id']; page = str(s['page']); new = s['ssim_score']
        pk = page if (did in baseline and page in baseline.get(did, {})) else f"{int(page):04d}"
        if did in baseline and pk in baseline[did]:
            old = baseline[did][pk]; diff = new - old
            if diff < -0.001: reg.append((did, int(page), old, new, diff))
            elif diff > 0.001: imp.append((did, int(page), old, new, diff))
            else: unch += 1
    out = []
    out.append('S517 numPr-marker fix gate (%d docs, dwrite default ss)' % len(paths))
    out.append('improved=%d unchanged=%d regressed=%d' % (len(imp), unch, len(reg)))
    out.append('net gain=%.4f loss=%.4f net=%+.4f' % (
        sum(x[4] for x in imp), sum(abs(x[4]) for x in reg),
        sum(x[4] for x in imp) - sum(abs(x[4]) for x in reg)))
    out.append('--- improvements ---')
    for x in sorted(imp, key=lambda r: r[4]):
        out.append('  %-40s p%d %.4f -> %.4f (%+.4f)' % (x[0][:40], x[1], x[2], x[3], x[4]))
    out.append('--- REGRESSIONS ---')
    for x in sorted(reg, key=lambda r: r[4]):
        out.append('  %-40s p%d %.4f -> %.4f (%+.4f)' % (x[0][:40], x[1], x[2], x[3], x[4]))
    txt = '\n'.join(out)
    (Path('c:/tmp') / '_s517_gate.txt').write_text(txt + '\n', encoding='utf-8')
    print(txt)

if __name__ == '__main__':
    main()
