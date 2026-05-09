"""Per-paragraph Word-vs-Oxi y drift trace for 4 Class A docs.

Day 32 Week 1 part 2: identify drift accumulation source.
For each Class A doc:
1. Render Oxi at SOFT=0pt, capture per-paragraph y per page
2. Use Word COM to capture Word per-paragraph y
3. Compute per-paragraph dy = oxi_y - word_y
4. Find earliest paragraph where dy starts accumulating

Output: pipeline_data/class_a_drift_trace_{doc_id}.json + cross-doc summary
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
PAGE_HEIGHT = 841.95


def find_docx(doc_id: str) -> str | None:
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def render_oxi(docx: str) -> dict:
    """Return {para_idx: (page, first_y)}."""
    label = os.path.splitext(os.path.basename(docx))[0]
    out_layout = os.path.join(r'C:\tmp', f'{label}_drift_layout.json')
    cmd = [RENDERER, docx, os.path.join(r'C:\tmp', f'{label}_drift'), f'--dump-layout={out_layout}']
    subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    with open(out_layout, encoding='utf-8') as f:
        layout = json.load(f)
    by_para = {}
    for page in layout.get('pages', []):
        pg = page.get('page')
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            pi = el.get('para_idx')
            if pi is None:
                continue
            y = el.get('y', 0)
            if pi not in by_para or (pg, y) < (by_para[pi]['page'], by_para[pi]['y']):
                by_para[pi] = {'page': pg, 'y': y, 'sample': el.get('text', '')[:30]}
    return by_para


def measure_word(docx: str) -> list[dict]:
    """Return [{i, page, start_y, text}]."""
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    paras = []
    try:
        n = d.Paragraphs.Count
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr_start = d.Range(r.Start, r.Start)
            text = (r.Text or '').strip()
            paras.append({
                'i': i,
                'text': text[:30],
                'page': int(cr_start.Information(3)),
                'start_y': round(cr_start.Information(6), 2),
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def trace_doc(doc_id: str) -> dict:
    docx = find_docx(doc_id)
    if not docx:
        return {'error': 'docx not found'}
    print(f'\n=== {doc_id} ===')
    print('Rendering Oxi...')
    oxi = render_oxi(docx)
    print(f'  {len(oxi)} paragraphs in Oxi')
    print('Measuring Word...')
    word_paras = measure_word(docx)
    print(f'  {len(word_paras)} paragraphs in Word')

    # Match by index (Oxi para_idx 0-based, Word i 1-based)
    drifts = []
    for w in word_paras:
        oxi_pi = w['i'] - 1
        if oxi_pi not in oxi:
            continue
        o = oxi[oxi_pi]
        w_abs = (w['page'] - 1) * PAGE_HEIGHT + w['start_y']
        o_abs = (o['page'] - 1) * PAGE_HEIGHT + o['y']
        drifts.append({
            'word_i': w['i'],
            'oxi_pi': oxi_pi,
            'word_pg': w['page'],
            'word_y': w['start_y'],
            'oxi_pg': o['page'],
            'oxi_y': round(o['y'], 2),
            'dy_abs': round(o_abs - w_abs, 2),
            'sample': w['text'],
        })

    # Find first paragraph with dy > 0.5pt
    first_drift = None
    for d in drifts:
        if abs(d['dy_abs']) > 0.5:
            first_drift = d
            break
    print(f'  First non-zero drift: {first_drift}')
    return {
        'doc_id': doc_id,
        'n_paras': len(drifts),
        'first_drift': first_drift,
        'drifts': drifts[:60],  # first 60 for trace
    }


def main():
    docs = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']
    results = {}
    for doc_id in docs:
        results[doc_id] = trace_doc(doc_id)

    print('\n\n=== Cross-doc summary ===')
    print(f'{"doc_id":<32} {"#paras":>7} {"first_drift_pi":>15} {"first_dy":>10} {"sample"}')
    for doc_id, r in results.items():
        if r.get('error'):
            print(f'{doc_id:<32} ERROR: {r["error"]}')
            continue
        fd = r.get('first_drift')
        if fd:
            print(f'{doc_id:<32} {r["n_paras"]:>7} {fd["oxi_pi"]:>15} {fd["dy_abs"]:>+10.2f} {fd["sample"]!r}')
        else:
            print(f'{doc_id:<32} {r["n_paras"]:>7} {"(none)":>15}')

    out_path = os.path.join(REPO, 'pipeline_data', 'class_a_drift_trace.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nWrote {out_path}')


if __name__ == '__main__':
    main()
