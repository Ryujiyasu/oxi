# -*- coding: utf-8 -*-
"""S531 full-baseline SSIM gate: re-render ALL baseline docs with the current (post-fix) Oxi
binary and compare per-page SSIM vs ssim_baseline.json. The fix (table-style cellMar inheritance
no longer gated on borders + single-cell cellMar wrap-budget subtraction) touches a shared parser/
layout path, so a full gate is required. Does NOT mutate the baseline. cp932-safe."""
import os, sys, shutil, io
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi, OXI_PNG_DIR
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline


def main():
    only = sys.argv[1:]  # optional stem prefixes to restrict
    baseline = load_baseline()
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    paths = []
    for did in sorted(baseline.keys()):
        if only and not any(did.startswith(o) for o in only):
            continue
        p = docx_dir / (did + '.docx')
        if p.exists():
            paths.append(str(p))
            c = Path(OXI_PNG_DIR) / did
            if c.exists():
                shutil.rmtree(c)
    print('gating %d docs (clearing oxi_png cache) ...' % len(paths))
    word = render_with_word(paths)
    oxi = render_with_oxi(paths)
    scores = calculate_ssim(word, oxi, skip_heatmap=True)
    imp = []; reg = []; unch = 0
    new_by_doc = {}
    for s in scores:
        did = s['doc_id']; page = str(s['page']); new = s['ssim_score']
        new_by_doc.setdefault(did, {})[page] = new
        pk = page if (did in baseline and page in baseline.get(did, {})) else f"{int(page):04d}"
        if did in baseline and pk in baseline[did]:
            old = baseline[did][pk]; diff = new - old
            if diff < -0.001: reg.append((did, int(page), old, new, diff))
            elif diff > 0.001: imp.append((did, int(page), old, new, diff))
            else: unch += 1
    # per-doc mean delta
    docdelta = []
    for did, pages in new_by_doc.items():
        if did not in baseline:
            continue
        olds = []; news = []
        for page, new in pages.items():
            pk = page if page in baseline[did] else f"{int(page):04d}"
            if pk in baseline[did]:
                olds.append(baseline[did][pk]); news.append(new)
        if olds:
            docdelta.append((did, sum(news)/len(news) - sum(olds)/len(olds)))
    out = ['S531 full-baseline cellMar gate (%d docs, %d pages scored)' % (len(paths), len(scores)),
           'improved=%d unchanged=%d regressed=%d' % (len(imp), unch, len(reg)),
           'gain=%+.4f loss=%-.4f net=%+.4f' % (sum(x[4] for x in imp), sum(abs(x[4]) for x in reg),
                                                sum(x[4] for x in imp) - sum(abs(x[4]) for x in reg)),
           '--- improvements (page) ---']
    for x in sorted(imp, key=lambda r: r[4]):
        out.append('  %-36s p%d %.4f -> %.4f (%+.4f)' % (x[0][:36], x[1], x[2], x[3], x[4]))
    out.append('--- REGRESSIONS (page) ---')
    for x in sorted(reg, key=lambda r: r[4]):
        out.append('  %-36s p%d %.4f -> %.4f (%+.4f)' % (x[0][:36], x[1], x[2], x[3], x[4]))
    out.append('--- per-DOC mean delta (|>0.0005|) ---')
    for did, dd in sorted(docdelta, key=lambda r: r[1]):
        if abs(dd) > 0.0005:
            out.append('  %-36s %+.4f' % (did[:36], dd))
    txt = '\n'.join(out)
    io.open('c:/tmp/_s531_gate.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)


if __name__ == '__main__':
    main()
