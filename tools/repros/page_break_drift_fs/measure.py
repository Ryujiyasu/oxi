"""Measure PB_DRIFT_FS — Word COM advance + Oxi layout JSON advance per variant."""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..', '..'))
GDI_EXE = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

VARIANTS = [
    ("DR_FS_01", 10.5, 10.5, 10.5),
    ("DR_FS_02", 10.5, 11.0, 10.5),
    ("DR_FS_03", 10.5, 12.0, 10.5),
    ("DR_FS_04", 10.5, 14.0, 10.5),
    ("DR_FS_05", 14.0, 10.5, 14.0),
    ("DR_FS_06", 11.5, 10.5, 11.5),
    ("DR_FS_07", 10.5,  8.0, 10.5),
    ("DR_FS_08",  8.0, 10.5,  8.0),
    ("DR_FS_09", 10.5, 16.0, 10.5),
    ("DR_FS_10",  9.0, 10.5,  9.0),
]


def measure_word(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
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
    """Run GDI renderer with --dump-layout, return per-paragraph y."""
    layout_path = os.path.join(HERE, '_tmp_layout.json')
    out_prefix = os.path.join(HERE, '_tmp_oxi')
    cmd = [GDI_EXE, docx_path, out_prefix, '150', f'--dump-layout={layout_path}']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(f'  GDI render failed: {r.stderr[:200]}')
        return []
    try:
        with open(layout_path, encoding='utf-8') as f:
            layout = json.load(f)
    except Exception as e:
        print(f'  layout parse failed: {e}')
        return []
    # First text y per paragraph index
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
    print(f'GDI renderer: {GDI_EXE}')
    print(f'Exists: {os.path.exists(GDI_EXE)}')
    results = []
    for vid, A, B, C in VARIANTS:
        path = os.path.join(HERE, f'{vid}.docx')
        if not os.path.exists(path):
            continue
        print(f'\n--- {vid}: A={A} / B={B} / C={C} ---')
        wys = measure_word(path)
        oys = measure_oxi(path)
        print(f'  Word ys: {wys[:5]}')
        print(f'  Oxi ys:  {oys[:5]}')
        if len(wys) >= 3 and len(oys) >= 3:
            w_AB = round(wys[1] - wys[0], 3)
            w_BC = round(wys[2] - wys[1], 3)
            o_AB = round(oys[1] - oys[0], 3)
            o_BC = round(oys[2] - oys[1], 3)
            d_AB = round(o_AB - w_AB, 3)
            d_BC = round(o_BC - w_BC, 3)
            print(f'  advance A→B: Word={w_AB:6.2f}  Oxi={o_AB:6.2f}  Δ={d_AB:+.2f}')
            print(f'  advance B→C: Word={w_BC:6.2f}  Oxi={o_BC:6.2f}  Δ={d_BC:+.2f}')
            results.append({
                'variant': vid, 'A_fs': A, 'B_fs': B, 'C_fs': C,
                'word_y_A': wys[0], 'word_y_B': wys[1], 'word_y_C': wys[2],
                'oxi_y_A': oys[0], 'oxi_y_B': oys[1], 'oxi_y_C': oys[2],
                'word_advance_AB': w_AB, 'word_advance_BC': w_BC,
                'oxi_advance_AB': o_AB, 'oxi_advance_BC': o_BC,
                'delta_AB': d_AB, 'delta_BC': d_BC,
            })

    # Summary
    print(f'\n=== Cursor advance comparison summary ===')
    print(f'{"variant":<10} {"A":>5} {"B":>5} {"C":>5} {"W_AB":>6} {"O_AB":>6} {"Δ_AB":>6} {"W_BC":>6} {"O_BC":>6} {"Δ_BC":>6}')
    for r in results:
        print(f'  {r["variant"]:<10} {r["A_fs"]:>5.1f} {r["B_fs"]:>5.1f} {r["C_fs"]:>5.1f} '
              f'{r["word_advance_AB"]:>6.2f} {r["oxi_advance_AB"]:>6.2f} {r["delta_AB"]:>+6.2f} '
              f'{r["word_advance_BC"]:>6.2f} {r["oxi_advance_BC"]:>6.2f} {r["delta_BC"]:>+6.2f}')

    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nSaved: {out}')

    # Identify drift sources (|Δ| > 0.5pt)
    print('\n=== Drift events (|Δ| > 0.5pt) ===')
    drifts = [(r, 'AB', r['delta_AB']) for r in results if abs(r['delta_AB']) > 0.5]
    drifts += [(r, 'BC', r['delta_BC']) for r in results if abs(r['delta_BC']) > 0.5]
    if drifts:
        for r, where, d in sorted(drifts, key=lambda x: -abs(x[2])):
            print(f'  {r["variant"]:<10} {where} ({r["A_fs"] if where=="AB" else r["B_fs"]:>4.1f}→{r["B_fs"] if where=="AB" else r["C_fs"]:>4.1f}): Δ={d:+.2f}pt')
    else:
        print('  No drift events found (all |Δ| ≤ 0.5pt)')


if __name__ == '__main__':
    main()
