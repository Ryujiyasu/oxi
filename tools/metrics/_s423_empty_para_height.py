"""S423: COM-measure Word's empty-paragraph height determinants in 3a4f.

S422 found 3a4f's IoU debt is dominated by empty-paragraph height mismatch:
Oxi clusters at 72pt (4×18) while Word varies 14-144 (1-8 lines). This
measures, for every EMPTY paragraph, Word's rendered height (= Information(6)
y-gap to the next paragraph on the same page) alongside its spacing / line /
font properties, to find what drives Word's value.

Output: pipeline_data/ra_manual_measurements/s423_3a4f_empty_heights.json
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = r'c:\Users\ryuji\oxi-main'
DOC = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', '3a4f9fbe1a83_001620506.docx')
OUT = os.path.join(REPO, 'pipeline_data', 'ra_manual_measurements', 's423_3a4f_empty_heights.json')

wdVerticalPositionRelativeToPage = 6
wdActiveEndPageNumber = 3


def main():
    import win32com.client as wc
    import shutil, tempfile
    # 3a4f's original file fails Documents.Open ("command could not complete")
    # — likely protected-view / zone-identifier / lock. Opening a temp copy
    # works (S423). Copy first.
    tmp = os.path.join(tempfile.gettempdir(), 's423_3a4f.docx')
    shutil.copy(DOC, tmp)
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(tmp, ReadOnly=True)
    doc.Repaginate()
    recs = []
    try:
        n = doc.Paragraphs.Count
        # Pre-fetch (page, y) for every paragraph start (collapsed; R30).
        py = []
        for pi in range(1, n + 1):
            rng = doc.Paragraphs(pi).Range
            s = doc.Range(rng.Start, rng.Start)
            try:
                pg = int(s.Information(wdActiveEndPageNumber))
                y = float(s.Information(wdVerticalPositionRelativeToPage))
            except Exception:
                pg, y = -1, -1.0
            py.append((pg, y))
        for pi in range(1, n + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            txt = (rng.Text or '').rstrip('\r\n\x07')
            if txt.strip():
                continue  # only empty paragraphs
            pg, y = py[pi - 1]
            # rendered height = gap to next paragraph if same page
            h = None
            if pi < n:
                npg, ny = py[pi]
                if npg == pg and ny > y:
                    h = round(ny - y, 1)
            fmt = p.Format
            font = rng.Font
            # has field/bookmark?
            has_field = rng.Fields.Count > 0
            has_bookmark = rng.Bookmarks.Count > 0
            recs.append({
                'pi': pi,
                'page': pg,
                'y': round(y, 1),
                'h': h,
                'space_before': round(float(fmt.SpaceBefore), 1),
                'space_after': round(float(fmt.SpaceAfter), 1),
                'line_spacing': round(float(fmt.LineSpacing), 1),
                'line_rule': int(fmt.LineSpacingRule),  # 0 single,1 1.5,2 dbl,3 atLeast,4 exactly,5 multiple
                'font_size': round(float(font.Size), 1) if font.Size else None,
                'space_before_auto': int(fmt.SpaceBeforeAuto),
                'space_after_auto': int(fmt.SpaceAfterAuto),
                'has_field': int(has_field),
                'has_bookmark': int(has_bookmark),
            })
    finally:
        doc.Close(SaveChanges=0)
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(recs, f, ensure_ascii=False, indent=1)

    # Summarize: group by rendered height, show the property profile.
    from collections import defaultdict
    by_h = defaultdict(list)
    for r in recs:
        if r['h'] is not None:
            by_h[r['h']].append(r)
    print(f'{len(recs)} empty paras; {sum(1 for r in recs if r["h"] is not None)} with same-page height')
    print(f"{'h':>6} {'n':>4} {'spBef':>6} {'spAft':>6} {'lnsp':>6} {'rule':>4} {'fs':>5} {'fld':>3} {'bkm':>3}")
    for h in sorted(by_h, reverse=True):
        g = by_h[h]
        import statistics as st
        def md(k):
            vs = [r[k] for r in g if r[k] is not None]
            return st.median(vs) if vs else None
        print(f"{h:>6} {len(g):>4} {md('space_before'):>6} {md('space_after'):>6} {md('line_spacing'):>6} {md('line_rule'):>4} {str(md('font_size')):>5} {sum(r['has_field'] for r in g):>3} {sum(r['has_bookmark'] for r in g):>3}")
    print(f'\nsaved -> {OUT}')


if __name__ == '__main__':
    main()
