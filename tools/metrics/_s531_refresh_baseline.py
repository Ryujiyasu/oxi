# -*- coding: utf-8 -*-
"""S531 baseline refresh: recompute SSIM for all baseline docs from the CURRENT caches
(word_png cached; oxi_png fresh post-S531 from the gate run) and write the updated
ssim_baseline.json. Reports every entry that changes. Run ONLY after the gate passed."""
import json, sys, io
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline

BASELINE_PATH = ROOT / 'pipeline_data' / 'ssim_baseline.json'


def main():
    baseline = load_baseline()
    docx_dir = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
    paths = [str(docx_dir / (d + '.docx')) for d in sorted(baseline) if (docx_dir / (d + '.docx')).exists()]
    print('recomputing %d docs from caches ...' % len(paths))
    word = render_with_word(paths)
    oxi = render_with_oxi(paths)
    scores = calculate_ssim(word, oxi, skip_heatmap=True)
    changed = []
    new_baseline = {d: dict(pages) for d, pages in baseline.items()}
    for s in scores:
        did = s['doc_id']; page = str(s['page']); new = s['ssim_score']
        if did not in new_baseline:
            continue
        pk = page if page in new_baseline[did] else ('%04d' % int(page))
        if pk not in new_baseline[did]:
            pk = page
        old = new_baseline[did].get(pk)
        if old is None or abs(new - old) > 1e-9:
            changed.append((did, pk, old, new))
            new_baseline[did][pk] = new
    with io.open(BASELINE_PATH, 'w', encoding='utf-8') as f:
        json.dump(new_baseline, f, indent=2, ensure_ascii=False)
    olds = [v for p in baseline.values() for v in p.values()]
    news = [v for p in new_baseline.values() for v in p.values()]
    out = ['S531 baseline refresh: %d entries changed' % len(changed),
           'mean %.6f -> %.6f (%+.6f)' % (sum(olds)/len(olds), sum(news)/len(news),
                                          sum(news)/len(news)-sum(olds)/len(olds))]
    for did, pk, old, new in sorted(changed, key=lambda x: (x[3]-(x[2] or 0))):
        out.append('  %-40s p%-3s %.4f -> %.4f (%+.4f)' % (did[:40], pk, old or 0, new, new-(old or 0)))
    txt = '\n'.join(out)
    io.open('c:/tmp/_s531_refresh.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)


if __name__ == '__main__':
    main()
