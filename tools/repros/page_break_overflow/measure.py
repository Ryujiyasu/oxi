"""Series PB_OF — measure Word's actual page placement for each variant.

For each PB_OF_NN.docx, open in Word COM, locate the "TEST" paragraph
(last paragraph), and record:
- Word's page index for "TEST"
- Word's actual y for "TEST" (Information(6) at collapsed start range,
  per R30 fix)
- Computed: cursor_y just before TEST paragraph (= y_TEST - space_before)

Output: measurements.json with per-variant decision.

For each variant:
  - design_D = target overflow at TEST paragraph bottom
  - measured_page = 1 (fits) or 2 (broke)
  - actual_y = Word's y for TEST
  - actual_overflow = actual_y + lh - page_bottom (only if measured_page=1)

Tolerance characterization:
  - Smallest D at which Word breaks: that is Word's overflow tolerance.
  - If Word breaks at D=+1pt: Word matches Oxi's strict > rule (no soft margin).
  - If Word breaks at D=+5pt: Word allows ~5pt soft overflow.
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')


VARIANTS = [
    ("PB_OF_01", -15),
    ("PB_OF_02", -10),
    ("PB_OF_03",  -5),
    ("PB_OF_04",  -2),
    ("PB_OF_05",  -1),
    ("PB_OF_06",   0),
    ("PB_OF_07",   1),
    ("PB_OF_08",   2),
    ("PB_OF_09",   3),
    ("PB_OF_10",   5),
    ("PB_OF_11",   7),
    ("PB_OF_12",  10),
]

HERE = os.path.dirname(os.path.abspath(__file__))


def measure(docx_path):
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    try:
        n = d.Paragraphs.Count
        # Last paragraph = TEST
        # Or find by text
        test_para = None
        test_idx = None
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            t = (p.Range.Text or '').strip()
            if 'TEST' in t:
                test_para = p
                test_idx = i
                break
        if test_para is None:
            return {'error': 'TEST paragraph not found', 'n_paragraphs': n}

        # Information(3) = page index, (6) = y
        rng = test_para.Range
        cr = d.Range(rng.Start, rng.Start)
        page = int(cr.Information(3))
        y = round(cr.Information(6), 3)

        # Last fill paragraph (test_idx - 1) for cursor reference
        if test_idx >= 2:
            prev = d.Paragraphs(test_idx - 1)
            prev_rng = prev.Range
            prev_cr = d.Range(prev_rng.Start, prev_rng.Start)
            prev_page = int(prev_cr.Information(3))
            prev_y = round(prev_cr.Information(6), 3)
        else:
            prev_page = None
            prev_y = None

        # Total pages
        n_pages = int(d.ComputeStatistics(2))  # wdStatisticPages = 2

        return {
            'n_paragraphs': n,
            'test_idx': test_idx,
            'test_page': page,
            'test_y': y,
            'prev_page': prev_page,
            'prev_y': prev_y,
            'n_pages': n_pages,
        }
    finally:
        d.Close(False)
        word.Quit()


def main():
    results = []
    for vid, D in VARIANTS:
        path = os.path.join(HERE, f'{vid}.docx')
        if not os.path.exists(path):
            print(f'  {vid}: missing'); continue
        print(f'  measuring {vid}...')
        m = measure(path)
        m['variant'] = vid
        m['design_D'] = D
        results.append(m)

    # Print table
    print(f'\n{"variant":<12} {"D":>4} {"n_pages":>7} {"test_pg":>7} {"test_y":>8} {"prev_pg":>7} {"prev_y":>8}')
    for r in results:
        if 'error' in r:
            print(f'  {r["variant"]:<12} ERROR: {r["error"]}')
            continue
        print(f'  {r["variant"]:<12} {r["design_D"]:>+4d} '
              f'{r["n_pages"]:>7} {r["test_page"]:>7} {r["test_y"]:>8.2f} '
              f'{r.get("prev_page", "?"):>7} {r.get("prev_y", 0):>8.2f}')

    # Save JSON
    out = os.path.join(HERE, 'measurements.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\nSaved: {out}')

    # Tolerance characterization
    print('\n=== Tolerance characterization ===')
    fits = [(r['design_D'], r) for r in results if r.get('test_page') == 1]
    breaks = [(r['design_D'], r) for r in results if r.get('test_page', 0) > 1]
    if fits and breaks:
        max_fit_D = max(d for d, _ in fits)
        min_break_D = min(d for d, _ in breaks)
        print(f'  Max D where Word fits:  D = {max_fit_D:+d}pt')
        print(f'  Min D where Word breaks: D = {min_break_D:+d}pt')
        print(f'  Tolerance threshold: between {max_fit_D}pt and {min_break_D}pt')
    elif fits:
        print(f'  All variants fit on page 1 (max D tested: {max(d for d,_ in fits)})')
    elif breaks:
        print(f'  All variants break (min D tested: {min(d for d,_ in breaks)})')


if __name__ == '__main__':
    main()
