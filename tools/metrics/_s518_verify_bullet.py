# -*- coding: utf-8 -*-
"""S518 bullet-fix gate: re-render + SSIM all gen2 baseline docs (69 render a Symbol bullet, the
rest are byte-identical sentinels) + the numPr sentinels (b837/d7 = confirm S517 holds) vs the
ssim_baseline. Clears stale oxi_png first. Zero-regression rule. Does NOT mutate the baseline."""
import os, sys, shutil, json, io
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi, OXI_PNG_DIR
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline

def main():
    baseline = load_baseline()
    gen2 = [k for k in baseline if k.startswith('gen2_')]
    sentinels = ['b837808d0555_20240705_resources_data_guideline_02', 'd7_v4_bullet', 'd7_v5_no_indent_bullet']
    stems = sorted(set(gen2 + [s for s in sentinels if s in baseline]))
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    paths = []
    for stem in stems:
        p = docx_dir / (stem + '.docx')
        if p.exists():
            paths.append(str(p))
            cache = Path(OXI_PNG_DIR) / stem
            if cache.exists():
                shutil.rmtree(cache)
    print('gating %d docs (cleared oxi_png)...' % len(paths))
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
    out = ['S518 bullet-fix gate (%d docs)' % len(paths),
           'improved=%d unchanged=%d regressed=%d' % (len(imp), unch, len(reg)),
           'net gain=%.4f loss=%.4f net=%+.4f' % (
               sum(x[4] for x in imp), sum(abs(x[4]) for x in reg),
               sum(x[4] for x in imp) - sum(abs(x[4]) for x in reg)),
           '--- top improvements ---']
    for x in sorted(imp, key=lambda r: r[4])[:20]:
        out.append('  %-34s p%d %.4f -> %.4f (%+.4f)' % (x[0][:34], x[1], x[2], x[3], x[4]))
    out.append('--- REGRESSIONS ---')
    for x in sorted(reg, key=lambda r: r[4]):
        out.append('  %-34s p%d %.4f -> %.4f (%+.4f)' % (x[0][:34], x[1], x[2], x[3], x[4]))
    txt = '\n'.join(out)
    io.open('c:/tmp/_s518_bullet_gate.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
