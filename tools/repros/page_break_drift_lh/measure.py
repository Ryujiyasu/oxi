"""Measure PB_DRIFT_LH — Word vs Oxi cursor advance under lh transitions."""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..', '..'))
GDI_EXE = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    ("DR_LH_01", "auto", "auto", "auto"),
    ("DR_LH_02", "auto", "exact14", "auto"),
    ("DR_LH_03", "auto", "exact16", "auto"),
    ("DR_LH_04", "auto", "exact18", "auto"),
    ("DR_LH_05", "auto", "exact21", "auto"),
    ("DR_LH_06", "auto", "mult1.15", "auto"),
    ("DR_LH_07", "auto", "mult1.5", "auto"),
    ("DR_LH_08", "exact16", "exact16", "exact16"),
    ("DR_LH_09", "exact14", "exact16", "exact14"),
    ("DR_LH_10", "auto", "exact14", "exact14"),
]


def measure_word(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False; word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        ys = []
        for i in range(1, d.Paragraphs.Count + 1):
            rng = d.Paragraphs(i).Range
            cr = d.Range(rng.Start, rng.Start)
            ys.append(round(cr.Information(6), 3))
        return ys
    finally:
        d.Close(False); word.Quit()


def measure_oxi(docx_path):
    layout_path = os.path.join(HERE, '_tmp_layout.json')
    out_prefix = os.path.join(HERE, '_tmp_oxi')
    cmd = [GDI_EXE, docx_path, out_prefix, '150', f'--dump-layout={layout_path}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0: return []
    try:
        with open(layout_path, encoding='utf-8') as f:
            layout = json.load(f)
    except Exception: return []
    by_pi = {}
    for page in layout.get('pages', []):
        for el in page.get('elements', []):
            if el.get('type') != 'text': continue
            pi = el.get('para_idx')
            if pi is None: continue
            y = el.get('y', 0)
            if pi not in by_pi or y < by_pi[pi]:
                by_pi[pi] = y
    return [by_pi.get(pi, None) for pi in sorted(by_pi)]


def main():
    results = []
    for vid, A, B, C in VARIANTS:
        path = os.path.join(HERE, f'{vid}.docx')
        if not os.path.exists(path): continue
        wys = measure_word(path); oys = measure_oxi(path)
        if len(wys) >= 3 and len(oys) >= 3:
            r = {
                'variant': vid, 'A': A, 'B': B, 'C': C,
                'word_y': wys[:3], 'oxi_y': oys[:3],
                'word_AB': round(wys[1]-wys[0], 3),
                'word_BC': round(wys[2]-wys[1], 3),
                'oxi_AB': round(oys[1]-oys[0], 3),
                'oxi_BC': round(oys[2]-oys[1], 3),
            }
            r['delta_AB'] = round(r['oxi_AB']-r['word_AB'], 3)
            r['delta_BC'] = round(r['oxi_BC']-r['word_BC'], 3)
            r['init_dy'] = round(oys[0]-wys[0], 3)
            results.append(r)

    print(f'\n{"variant":<10} {"A":<10} {"B":<10} {"C":<10} {"init_dy":>8} {"W_AB":>6} {"O_AB":>6} {"Δ_AB":>6} {"W_BC":>6} {"O_BC":>6} {"Δ_BC":>6}')
    for r in results:
        print(f'  {r["variant"]:<10} {r["A"]:<10} {r["B"]:<10} {r["C"]:<10} {r["init_dy"]:>+8.2f} '
              f'{r["word_AB"]:>6.2f} {r["oxi_AB"]:>6.2f} {r["delta_AB"]:>+6.2f} '
              f'{r["word_BC"]:>6.2f} {r["oxi_BC"]:>6.2f} {r["delta_BC"]:>+6.2f}')

    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nSaved: {out}')

    print('\n=== Drift events |Δ| > 0.5pt ===')
    for r in results:
        if abs(r['delta_AB']) > 0.5:
            print(f'  {r["variant"]}: A→B ({r["A"]} → {r["B"]}): Δ={r["delta_AB"]:+.2f}pt')
        if abs(r['delta_BC']) > 0.5:
            print(f'  {r["variant"]}: B→C ({r["B"]} → {r["C"]}): Δ={r["delta_BC"]:+.2f}pt')

    print('\n=== Initial position offset ===')
    for r in results:
        if abs(r['init_dy']) > 0.5:
            print(f'  {r["variant"]}: y_A initial offset Δ={r["init_dy"]:+.2f}pt')


if __name__ == '__main__':
    main()
