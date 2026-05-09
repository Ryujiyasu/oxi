"""Measure PB_DG variants with Word COM, characterize tolerance threshold."""
from __future__ import annotations
import os, sys, json, glob
sys.stdout.reconfigure(encoding='utf-8')

HERE = os.path.dirname(os.path.abspath(__file__))


def measure(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        test_idx = None
        for i in range(1, n + 1):
            t = (d.Paragraphs(i).Range.Text or '').strip()
            if 'TEST' in t:
                test_idx = i
                break
        if test_idx is None:
            return {'error': 'TEST not found'}
        rng = d.Paragraphs(test_idx).Range
        cr = d.Range(rng.Start, rng.Start)
        return {
            'test_idx': test_idx,
            'test_page': int(cr.Information(3)),
            'test_y': round(cr.Information(6), 3),
            'n_pages': int(d.ComputeStatistics(2)),
        }
    finally:
        d.Close(False); word.Quit()


D_VALUES = [-3, -1, 0, 1, 3, 5, 7, 10]


def main():
    results = []
    series_label = {'A': 'lines', 'B': 'linesAndChars'}
    for series in ('A', 'B'):
        for i, D in enumerate(D_VALUES, 1):
            vid = f'PB_DG_{series}_{i:02d}'
            path = os.path.join(HERE, f'{vid}.docx')
            if not os.path.exists(path):
                continue
            print(f'  {vid}...')
            m = measure(path)
            m.update({'variant': vid, 'series': series, 'D': D, 'grid': series_label[series]})
            results.append(m)

    print(f'\n{"variant":<12} {"grid":<14} {"D":>3} {"test_pg":>7} {"test_y":>8}')
    for r in results:
        if 'error' in r:
            print(f'  {r["variant"]:<12} ERROR'); continue
        print(f'  {r["variant"]:<12} {r["grid"]:<14} {r["D"]:>+3d} '
              f'{r["test_page"]:>7} {r["test_y"]:>8.2f}')

    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nSaved: {out}')

    # Tolerance per series
    print('\n=== Tolerance per series ===')
    for series in ('A', 'B'):
        srs = [r for r in results if r.get('series') == series and 'test_page' in r]
        fits = [r['D'] for r in srs if r['test_page'] == 1]
        breaks = [r['D'] for r in srs if r['test_page'] > 1]
        label = series_label[series]
        if fits and breaks:
            print(f'  Series {series} ({label:<14}): fits up to D={max(fits):+d}, breaks at D={min(breaks):+d}')
        else:
            print(f'  Series {series} ({label}): all-fit or all-break')


if __name__ == '__main__':
    main()
