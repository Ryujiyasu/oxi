# -*- coding: utf-8 -*-
"""S529 w:w gate: re-render + SSIM the 5 docs using <w:w> character scaling (the only docs the
text_scale render fix can change; all others byte-identical) vs ssim_baseline. Zero-regression
check. Render-only (element positions unchanged) -> Phase-1 unaffected. Does NOT mutate baseline."""
import os, sys, shutil, io
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi, OXI_PNG_DIR
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline

WW = ['3a4f9fbe1a83_001620506', 'b35123fe8efc_tokumei_08_01',
      'b837808d0555_20240705_resources_data_guideline_02',
      'cb8be715d839_kyodokenkyuyoushiki03', 'd4d126dfe1d9_tokumei_08_01-3']

def main():
    baseline = load_baseline()
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    paths = []
    for s in WW:
        p = docx_dir / (s + '.docx')
        if p.exists() and s in baseline:
            paths.append(str(p))
            c = Path(OXI_PNG_DIR) / s
            if c.exists():
                shutil.rmtree(c)
    print('gating %d w:w docs ...' % len(paths))
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
    out = ['S529 w:w text-scale gate (%d docs)' % len(paths),
           'improved=%d unchanged=%d regressed=%d' % (len(imp), unch, len(reg)),
           'net gain=%.4f loss=%.4f net=%+.4f' % (sum(x[4] for x in imp), sum(abs(x[4]) for x in reg),
                                                   sum(x[4] for x in imp) - sum(abs(x[4]) for x in reg)),
           '--- improvements ---']
    for x in sorted(imp, key=lambda r: r[4]):
        out.append('  %-34s p%d %.4f -> %.4f (%+.4f)' % (x[0][:34], x[1], x[2], x[3], x[4]))
    out.append('--- REGRESSIONS ---')
    for x in sorted(reg, key=lambda r: r[4]):
        out.append('  %-34s p%d %.4f -> %.4f (%+.4f)' % (x[0][:34], x[1], x[2], x[3], x[4]))
    txt = '\n'.join(out)
    io.open('c:/tmp/_s529_gate.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
