"""S417d: x-fidelity diagnostic — Word TRUE rendered x (word_true_x cache)
vs Oxi rendered x, per matched paragraph. STANDALONE diagnostic; does NOT
touch the live Phase 2 Y-only gate (element_iou_diff).

Motivation: S416 showed Phase 2 position_iou is Y-only and the Word ref x
(Information(5)) is a left-flow artifact, so horizontal fixes (S412 cellMar
gate) are invisible to the current metric. This tool measures horizontal
fidelity against the GetPoint TRUE-x reference.

Inputs:
  - pipeline_data/word_true_x/<doc>.json  (Word true x, S417 cache)
  - Oxi gdi --dump-layout (rendered x per fragment), env-controlled so the
    same binary can be run OFF (default) and ON (OXI_S412_ENABLE=1).

Per matched paragraph (matched by page + text-prefix + nearest y):
  dx        = oxi_x_left - word_x_true        (signed pt error)
  x_pos_iou = max(0, 1 - |dx| / max(word_w, oxi_w))   (parallel to Y position_iou)
Focus reporting on non-left cells (align 1 center / 2 right) where the
artifact and S412 matter; left/justify (0/3) reported separately.

Usage:
  python tools/metrics/x_fidelity_diff.py <doc_id> [--on]
    --on sets OXI_S412_ENABLE=1 for the Oxi render.
"""
from __future__ import annotations
import os, sys, json, subprocess, tempfile
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = r'c:\Users\ryuji\oxi-main'
DOCS_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TRUEX_DIR = os.path.join(REPO, 'pipeline_data', 'word_true_x')


def docx_for(doc_id):
    import glob
    for p in glob.glob(os.path.join(DOCS_DIR, '*.docx')):
        if os.path.basename(p).startswith(doc_id):
            return p
    return None


def oxi_render(docx, enable_s412):
    env = dict(os.environ)
    if enable_s412:
        env['OXI_S412_ENABLE'] = '1'
    else:
        env.pop('OXI_S412_ENABLE', None)
    with tempfile.TemporaryDirectory(prefix='xfid_') as tmp:
        pre = os.path.join(tmp, 'p_')
        dump = os.path.join(tmp, 'l.json')
        subprocess.run([RENDERER, docx, pre, '--dump-layout=' + dump],
                       capture_output=True, text=True, timeout=180, env=env)
        with open(dump, encoding='utf-8') as f:
            return json.load(f)


def oxi_paragraphs(dump):
    """Aggregate fragments per (page, para_idx, cell_para_idx, cell_col_idx)
    into a paragraph record: x_left=min(x), x_right=max(x+w), y, text."""
    out = []
    for page in dump.get('pages', []):
        pg = page.get('page')
        groups = defaultdict(lambda: {'xl': 1e9, 'xr': -1e9, 'y': 1e9, 'parts': []})
        for el in page.get('elements', []):
            if 'text' not in el or not el.get('text'):
                continue
            key = (el.get('para_idx'), el.get('cell_para_idx'), el.get('cell_col_idx'))
            g = groups[key]
            x = el['x']; w = el.get('w', 0.0); y = el['y']
            g['xl'] = min(g['xl'], x)
            g['xr'] = max(g['xr'], x + w)
            g['y'] = min(g['y'], y)
            g['parts'].append((y, x, el.get('text', '')))
        for key, g in groups.items():
            g['parts'].sort()
            text = ''.join(t for _, _, t in g['parts'])
            out.append({'page': pg, 'x_left': round(g['xl'], 2),
                        'x_right': round(g['xr'], 2), 'w': round(g['xr'] - g['xl'], 2),
                        'y': round(g['y'], 2), 'text': text})
    return out


def norm(s):
    # Keep ideographic spaces (U+3000) — they are part of the cell content
    # and help disambiguate short cells like "　　　　税". Strip only ASCII
    # spaces and line markers.
    return (s or '').replace(' ', '').strip('\r\n\x07')


def match_and_score(word_recs, oxi_recs):
    """Match each Word paragraph to the nearest-y Oxi record on the same page
    with matching normalized text prefix. Return scored pairs."""
    oxi_by_page = defaultdict(list)
    for o in oxi_recs:
        oxi_by_page[o['page']].append(o)
    pairs = []
    used = set()
    for w in word_recs:
        if w.get('x_true') is None:
            continue
        wt = norm(w['text'])
        if len(wt) < 1:
            continue
        cands = oxi_by_page.get(w['page'], [])
        best = None; bestdy = 1e9
        for i, o in enumerate(cands):
            oid = id(o)
            ot = norm(o['text'])
            if not ot:
                continue
            k = min(6, len(wt), len(ot))
            if not (ot[:k] == wt[:k]):
                continue
            dy = abs(o['y'] - w['y'])
            if dy < bestdy and oid not in used:
                bestdy = dy; best = o
        if best is None or bestdy > 12.0:
            continue
        used.add(id(best))
        dx = best['x_left'] - w['x_true']
        ww = w.get('w_true') or 0.0
        ow = best.get('w') or 0.0
        denom = max(ww, ow, 1.0)
        x_pos_iou = max(0.0, 1.0 - abs(dx) / denom)
        pairs.append({'page': w['page'], 'align': w['align'], 'text': w['text'][:16],
                      'word_x': w['x_true'], 'oxi_x': best['x_left'],
                      'dx': round(dx, 2), 'x_pos_iou': round(x_pos_iou, 4)})
    return pairs


def run(doc_id, enable_s412):
    truex = os.path.join(TRUEX_DIR, f'{doc_id}.json')
    if not os.path.exists(truex):
        print(f'no true-x cache for {doc_id} (run build_word_true_x_cache.py)', file=sys.stderr)
        return None
    word_recs = json.load(open(truex, encoding='utf-8'))['paragraphs']
    docx = docx_for(doc_id)
    dump = oxi_render(docx, enable_s412)
    oxi_recs = oxi_paragraphs(dump)
    pairs = match_and_score(word_recs, oxi_recs)
    return pairs


def summarize(pairs, label):
    nonleft = [p for p in pairs if p['align'] in (1, 2)]
    def stats(ps):
        if not ps:
            return (0, None, None)
        mean_iou = sum(p['x_pos_iou'] for p in ps) / len(ps)
        mean_adx = sum(abs(p['dx']) for p in ps) / len(ps)
        return (len(ps), round(mean_iou, 4), round(mean_adx, 2))
    n_all, iou_all, adx_all = stats(pairs)
    n_nl, iou_nl, adx_nl = stats(nonleft)
    print(f'  [{label}] matched={n_all} mean_x_iou={iou_all} mean|dx|={adx_all}pt | '
          f'non-left: n={n_nl} mean_x_iou={iou_nl} mean|dx|={adx_nl}pt')
    return {'n': n_all, 'x_iou': iou_all, 'adx': adx_all,
            'nl_n': n_nl, 'nl_x_iou': iou_nl, 'nl_adx': adx_nl}


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    on = '--on' in sys.argv
    doc_id = args[0]
    print(f'=== x-fidelity {doc_id} (S412 {"ON" if on else "OFF"}) ===')
    pairs = run(doc_id, on)
    if pairs is not None:
        summarize(pairs, 'ON' if on else 'OFF')
        # show worst non-left cells
        nl = sorted((p for p in pairs if p['align'] in (1, 2)), key=lambda p: -abs(p['dx']))
        for p in nl[:8]:
            print(f"    dx={p['dx']:+7.2f} x_iou={p['x_pos_iou']} align={p['align']} "
                  f"word_x={p['word_x']} oxi_x={p['oxi_x']} {p['text']!r}")
