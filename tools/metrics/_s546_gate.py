# -*- coding: utf-8 -*-
"""S546 FULL-corpus SSIM gate: fs/2-exact halfwidth + fs/4 autospace touches any
doc with UPM256 halfwidth chars or DE/DN boundaries -> gate the ENTIRE baseline.
Clears ALL oxi_png caches first (CLAUDE.md hygiene; word_png is cached/reused).
Reports per-page improved/unchanged/regressed vs ssim_baseline. Does NOT mutate
the baseline."""
import os, sys, glob, shutil, json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi, OXI_PNG_DIR
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline


def main():
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    baseline = load_baseline()
    paths = []
    for p in sorted(docx_dir.glob('*.docx')):
        if p.stem in baseline:
            paths.append(str(p))
            cache = Path(OXI_PNG_DIR) / p.stem
            if cache.exists():
                shutil.rmtree(cache)
    print('rendering %d baseline docs (cleared ALL oxi_png)...' % len(paths))
    word = render_with_word(paths)
    oxi = render_with_oxi(paths)
    scores = calculate_ssim(word, oxi, skip_heatmap=True)
    imp = []; reg = []; unch = 0
    old_sum = 0.0; new_sum = 0.0; n = 0
    for s in scores:
        did = s['doc_id']; page = str(s['page']); new = s['ssim_score']
        pk = page if (did in baseline and page in baseline.get(did, {})) else f"{int(page):04d}"
        if did in baseline and pk in baseline[did]:
            old = baseline[did][pk]; diff = new - old
            old_sum += old; new_sum += new; n += 1
            if diff < -0.001: reg.append((did, int(page), old, new, diff))
            elif diff > 0.001: imp.append((did, int(page), old, new, diff))
            else: unch += 1
    out = []
    out.append('S546 fs/2-halfwidth + fs/4-autospace FULL gate (%d docs)' % len(paths))
    out.append('improved=%d unchanged=%d regressed=%d (pages=%d)' % (len(imp), unch, len(reg), n))
    out.append('net gain=%.4f loss=%.4f net=%+.4f' % (
        sum(x[4] for x in imp), sum(abs(x[4]) for x in reg),
        sum(x[4] for x in imp) - sum(abs(x[4]) for x in reg)))
    out.append('mean: %.6f -> %.6f (%+.6f)' % (old_sum / n, new_sum / n, (new_sum - old_sum) / n))
    out.append('--- improvements ---')
    for x in sorted(imp, key=lambda r: r[4]):
        out.append('  %-44s p%d %.4f -> %.4f (%+.4f)' % (x[0][:44], x[1], x[2], x[3], x[4]))
    out.append('--- REGRESSIONS ---')
    for x in sorted(reg, key=lambda r: r[4]):
        out.append('  %-44s p%d %.4f -> %.4f (%+.4f)' % (x[0][:44], x[1], x[2], x[3], x[4]))
    txt = '\n'.join(out)
    (Path('c:/tmp') / '_s546_gate.txt').write_text(txt + '\n', encoding='utf-8')
    print(txt)


if __name__ == '__main__':
    main()
