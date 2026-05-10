"""Measure PB_DRIFT_EMPTY — body→empty(s)→body cursor advance."""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..', '..'))
GDI_EXE = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    "DR_EMP_01", "DR_EMP_02", "DR_EMP_03", "DR_EMP_04", "DR_EMP_05",
    "DR_EMP_06", "DR_EMP_07", "DR_EMP_08", "DR_EMP_09", "DR_EMP_10",
]


def measure_word(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        out = []
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            t = (p.Range.Text or '').strip()
            rng = p.Range
            cr = d.Range(rng.Start, rng.Start)
            out.append({
                'idx': i, 'text': t[:6],
                'page': int(cr.Information(3)),
                'y': round(cr.Information(6), 3),
            })
        return out
    finally:
        d.Close(False); word.Quit()


def measure_oxi(docx_path):
    layout_path = os.path.join(HERE, '_tmp_layout.json')
    out_prefix = os.path.join(HERE, '_tmp_oxi')
    cmd = [GDI_EXE, docx_path, out_prefix, '150', f'--dump-layout={layout_path}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0: return []
    with open(layout_path, encoding='utf-8') as f:
        layout = json.load(f)
    by_pi = {}
    for page_idx, page in enumerate(layout.get('pages', [])):
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            y = el.get('y', 0)
            if pi not in by_pi or y < by_pi[pi]['y']:
                by_pi[pi] = {'pi': pi, 'page': page_idx + 1, 'y': y, 'text': el.get('text', '')[:6]}
    return [by_pi[k] for k in sorted(by_pi)]


def main():
    results = []
    for vid in VARIANTS:
        path = os.path.join(HERE, f'{vid}.docx')
        if not os.path.exists(path): continue
        wd = measure_word(path); ox = measure_oxi(path)
        # First word para is A (text), last is C
        wA = wd[0]; wC = wd[-1]
        # Oxi may have empty paragraphs missing from layout (no text element)
        # Find A and C by text match
        oA = next((o for o in ox if 'A' in o['text']), None)
        oC = next((o for o in ox if 'C' in o['text']), None)
        if oA is None or oC is None:
            print(f'{vid}: oxi A or C not found'); continue
        word_through = round(wC['y'] - wA['y'], 3)
        oxi_through = round(oC['y'] - oA['y'], 3)
        through_drift = round(oxi_through - word_through, 3)
        n_empty = len(wd) - 2
        print(f'{vid}: n_empty={n_empty} Word={word_through:6.2f} Oxi={oxi_through:6.2f} Δ={through_drift:+.2f}')
        results.append({
            'variant': vid, 'n_empty_word': n_empty, 'n_empty_oxi': len(ox) - 2,
            'word_y_A': wA['y'], 'word_y_C': wC['y'],
            'oxi_y_A': oA['y'], 'oxi_y_C': oC['y'],
            'word_through': word_through, 'oxi_through': oxi_through, 'through_drift': through_drift,
            'word_paras': wd, 'oxi_paras': ox,
        })

    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f'\n{"variant":<11} {"n_empty":>7} {"Word":>6} {"Oxi":>6} {"Δ":>6}')
    for r in results:
        print(f'  {r["variant"]:<11} {r["n_empty_word"]:>7} {r["word_through"]:>6.2f} {r["oxi_through"]:>6.2f} {r["through_drift"]:>+6.2f}')


if __name__ == '__main__':
    main()
