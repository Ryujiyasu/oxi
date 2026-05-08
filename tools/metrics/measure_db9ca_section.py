"""Measure Word vs Oxi paragraph y-positions on DS_V116-V120 repros.

For each variant, capture per-paragraph (start_y, page) for both Word
COM and Oxi --dump-layout. Compute drift trajectory.

Goal: identify whether the +18pt drift jump in db9ca emerges in V116
(paras 1-22, full context including page break) or only in larger
variants. Compare to V119 (paras 1-15, no page break = control) and
V120 (paras 14-22, drift event without preceding context).
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'db9ca_section')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
TMP = r'C:\tmp'

VARIANTS = [
    'DS_V116_paras_1_to_22',
    'DS_V117_paras_1_to_25',
    'DS_V118_paras_1_to_30',
    'DS_V119_paras_1_to_15',
    'DS_V120_paras_14_to_22',
    'DS_V122_paras_40_to_46',
    'DS_V123_paras_1_to_45',
    'DS_V124_paras_30_to_46',
    'DS_V125_para11_only',
    'DS_V126_para15_only',
    'DS_V127_para25_only',
]

# Full db9ca for comparison
FULL_DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx',
                         'db9ca18368cd_20241122_resource_open_data_01.docx')

PAGE_HEIGHT = 841.95  # A4


def measure_word(docx_path: str) -> list[dict]:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    paras = []
    try:
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            text = (r.Text or '').strip()
            paras.append({
                'i': i,
                'text': text[:50],
                'page': int(cr.Information(3)),
                'y': round(cr.Information(6), 2),
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def measure_oxi(docx_path: str) -> list[dict]:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    out_prefix = os.path.join(TMP, f'{label}')
    out_layout = os.path.join(TMP, f'{label}_layout.json')
    cmd = [RENDERER, docx_path, out_prefix, f'--dump-layout={out_layout}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        return []
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    pages = layout.get('pages', [])
    by_para = {}
    for page_idx, page in enumerate(pages):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            pi = el.get('para_idx')
            if pi is None:
                continue
            if pi not in by_para or el['y'] < by_para[pi]['y']:
                by_para[pi] = {
                    'para_idx': pi,
                    'page': page.get('page', page_idx + 1),
                    'y': round(el['y'], 2),
                    'text': el.get('text', '')[:50],
                }
    return [by_para[k] for k in sorted(by_para.keys())]


def cross_match(word_paras: list[dict], oxi_paras: list[dict]) -> list[dict]:
    """Match by paragraph index. Word's i is 1-indexed, Oxi's para_idx is 0-indexed.
    If counts match, pair word_paras[i] with oxi_paras[i-1].
    """
    matches = []
    # Build oxi by-idx map
    oxi_by_idx = {p['para_idx']: p for p in oxi_paras}
    for wp in word_paras:
        # Word i is 1-indexed; Oxi para_idx may be 0- or 1-indexed
        # Try i-1 first (most common 0-indexed)
        oxi_p = oxi_by_idx.get(wp['i'] - 1)
        if oxi_p is None:
            oxi_p = oxi_by_idx.get(wp['i'])
        if oxi_p is None:
            continue
        w_y_abs = (wp['page'] - 1) * PAGE_HEIGHT + wp['y']
        o_y_abs = (oxi_p['page'] - 1) * PAGE_HEIGHT + oxi_p['y']
        matches.append({
            'word_i': wp['i'], 'oxi_idx': oxi_p['para_idx'],
            'word_page': wp['page'], 'oxi_page': oxi_p['page'],
            'word_y': wp['y'], 'oxi_y': oxi_p['y'],
            'dy_abs': round(o_y_abs - w_y_abs, 2),
            'text': wp['text'][:30],
        })
    return matches


def main():
    results = []
    # Include full db9ca as DS_V121
    targets = [(label, os.path.join(REPRO_DIR, f'{label}.docx')) for label in VARIANTS]
    targets.append(('DS_V121_full_db9ca', FULL_DOCX))
    for label, docx in targets:
        if not os.path.exists(docx):
            print(f'SKIP {label} (not found)')
            continue
        print(f'=== {label} ===')
        try:
            wp = measure_word(docx)
        except Exception as e:
            print(f'  Word ERROR: {e}')
            continue
        op = measure_oxi(docx)
        m = cross_match(wp, op)
        # Compute trajectory metrics
        if m:
            cum = m[-1]['dy_abs'] - m[0]['dy_abs']
            max_abs = max(abs(x['dy_abs']) for x in m)
        else:
            cum = max_abs = 0
        print(f'  word_n={len(wp)} oxi_n={len(op)} matched={len(m)} cum_drift={cum:+.2f} max_abs={max_abs:+.2f}')
        # Show key transitions
        for i in (0, 1, 2, len(m)//4, len(m)//2, 3*len(m)//4, len(m)-1):
            if 0 <= i < len(m):
                p = m[i]
                print(f'    word_i={p["word_i"]:>2} pg w/o={p["word_page"]}/{p["oxi_page"]} y={p["word_y"]:>6.1f}/{p["oxi_y"]:>6.1f} dy={p["dy_abs"]:>+7.2f} | {p["text"]!r}')
        results.append({'label': label, 'matches': m, 'cum_drift': cum, 'max_abs': max_abs,
                        'word_paras': len(wp), 'oxi_paras': len(op)})
        print()

    out = os.path.join(REPO, 'pipeline_data', 'db9ca_section_results.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'Wrote {out}')


if __name__ == '__main__':
    main()
