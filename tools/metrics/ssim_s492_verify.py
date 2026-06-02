"""S492 SSIM sentinel (Phase-3 primary gate) on the docs F1 touches, using the
canonical pipeline building blocks (Word EMF render + SSIM calc). OFF vs ON, no
baseline mutation. Clears OXI_PNG_DIR per doc between renders (render_with_oxi caches).
"""
import os, sys, glob, shutil
from pathlib import Path

ROOT = Path(os.path.abspath('.'))
sys.path.insert(0, str(ROOT))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi
from pipeline.ssim_calculator import calculate_ssim
from pipeline.config import OXI_PNG_DIR

DOCX_DIR = ROOT / 'tools' / 'golden-test' / 'documents' / 'docx'
STEMS = ['683ffcab86e2', '0e7af1ae8f21', 'd77a58485f16']
docx_paths = []
for s in STEMS:
    docx_paths += sorted(glob.glob(str(DOCX_DIR / (s + '*.docx'))))
print('affected docx:', [Path(p).name for p in docx_paths])


def clear_cache():
    for p in docx_paths:
        d = Path(OXI_PNG_DIR) / Path(p).stem
        if d.exists():
            shutil.rmtree(d)


def to_map(scores):
    return {(s['doc_id'], int(s['page'])): s['ssim_score'] for s in scores}


word = render_with_word(docx_paths)  # canonical Word EMF (ground truth)

clear_cache()
os.environ.pop('OXI_S492_JCNATURAL', None)
off = to_map(calculate_ssim(word, render_with_oxi(docx_paths), skip_heatmap=True))

clear_cache()
os.environ['OXI_S492_JCNATURAL'] = '1'
on = to_map(calculate_ssim(word, render_with_oxi(docx_paths), skip_heatmap=True))

clear_cache()  # leave cache clean

print('\n=== per-page SSIM: OFF -> ON (delta) ===')
keys = sorted(set(off) | set(on))
tot_off = tot_on = 0.0
n = 0
bottom_off = []
bottom_on = []
for k in keys:
    o = off.get(k); nn = on.get(k)
    if o is None or nn is None:
        print('  %s p%d  missing (OFF=%s ON=%s)' % (k[0][:20], k[1], o, nn)); continue
    d = nn - o
    flag = '' if abs(d) < 0.001 else ('  UP' if d > 0 else '  DOWN <<<')
    print('  %-22s p%-2d  %.4f -> %.4f  %+.4f%s' % (k[0][:22], k[1], o, nn, d, flag))
    tot_off += o; tot_on += nn; n += 1
    bottom_off.append(o); bottom_on.append(nn)

if n:
    print('\nmean over %d pages: OFF %.4f -> ON %.4f  (%+.4f)' % (n, tot_off / n, tot_on / n, (tot_on - tot_off) / n))
    bottom_off.sort(); bottom_on.sort()
    for N in (3, 5):
        print('bottom-%d sum: OFF %.4f -> ON %.4f  (%+.4f)' %
              (N, sum(bottom_off[:N]), sum(bottom_on[:N]), sum(bottom_on[:N]) - sum(bottom_off[:N])))
    regs = [(k, off[k], on[k]) for k in keys if k in off and k in on and on[k] - off[k] < -0.001]
    print('\nREGRESSIONS (>0.001):', len(regs))
    for k, o, nn in sorted(regs, key=lambda x: x[2] - x[1])[:10]:
        print('  %-22s p%-2d  %.4f -> %.4f  %+.4f' % (k[0][:22], k[1], o, nn, nn - o))
