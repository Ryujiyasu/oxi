# -*- coding: utf-8 -*-
"""S521 SSIM gate: render the 33 ALH+docGrid+table docs with OXI_S521_CELL_NATURAL=1 (cell uses
natural line height, no grid-snap) and compare to ssim_baseline. Zero-regression check; reports
per-page improved/regressed. Special attention to b35 (S492y winner) and e3c545 (non-ALH control,
should be ~unchanged). Clears stale oxi_png. Does NOT mutate the baseline."""
import os, sys, shutil, io
from pathlib import Path
os.environ['OXI_S521_CELL_NATURAL'] = '1'   # MUST be set before importing/rendering
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi, OXI_PNG_DIR
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline

def main():
    baseline = load_baseline()
    stems = [l.strip() for l in io.open('c:/tmp/_alh_docs.txt', encoding='utf-8') if l.strip()]
    stems = [s for s in stems if s in baseline]
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    paths = []
    for s in stems:
        p = docx_dir / (s + '.docx')
        if p.exists():
            paths.append(str(p))
            c = Path(OXI_PNG_DIR) / s
            if c.exists():
                shutil.rmtree(c)
    print('gating %d ALH docs with OXI_S521_CELL_NATURAL=1 ...' % len(paths))
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
    out = ['S521 cell-natural gate (%d ALH docs)' % len(paths),
           'improved=%d unchanged=%d regressed=%d' % (len(imp), unch, len(reg)),
           'net gain=%.4f loss=%.4f net=%+.4f' % (
               sum(x[4] for x in imp), sum(abs(x[4]) for x in reg),
               sum(x[4] for x in imp) - sum(abs(x[4]) for x in reg)),
           '--- improvements (top 20) ---']
    for x in sorted(imp, key=lambda r: r[4])[:20]:
        out.append('  %-34s p%d %.4f -> %.4f (%+.4f)' % (x[0][:34], x[1], x[2], x[3], x[4]))
    out.append('--- REGRESSIONS (top 25) ---')
    for x in sorted(reg, key=lambda r: r[4])[:25]:
        out.append('  %-34s p%d %.4f -> %.4f (%+.4f)' % (x[0][:34], x[1], x[2], x[3], x[4]))
    txt = '\n'.join(out)
    io.open('c:/tmp/_s521_gate_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
