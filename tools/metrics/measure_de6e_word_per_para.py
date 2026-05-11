"""Day 33 part 43 — Per-paragraph Word measurement for de6e.

Walks d.Paragraphs (Word's flat paragraph list) rather than Cells. Avoids
gridSpan errors. For each paragraph:
- collapsed-start Information(6) for y
- spacing.before, spacing.after, line, lineRule
- font size, snap
- in_table flag
- predicted_advance = sb + line + sa
- actual_advance to NEXT paragraph
- compaction = predicted - actual

Output: pipeline_data/de6e_per_para_advance.csv

Compaction sites = where actual < predicted. Categorize by attributes
to find the discriminator (first-in-cell? last? row-context?).
"""
from __future__ import annotations
import os, sys, csv
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCX = 'tools/golden-test/documents/docx/de6e32b5960b_tokumei_08_01-1.docx'
PAGE_HEIGHT = 841.95
WD_VPOS = 6
WD_PAGE = 3
WD_IN_TABLE = 12


def measure(docx_path):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    paras = []
    try:
        n = d.Paragraphs.Count
        print(f'  {n} paragraphs total')
        for i in range(1, n + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try: y = round(cr.Information(WD_VPOS), 2)
            except: y = -1
            try: pg = int(cr.Information(WD_PAGE))
            except: pg = -1
            try: text = (r.Text or '').replace('\r', ' ').replace('\x07', '').strip()
            except: text = ''
            try: in_table = bool(r.Information(WD_IN_TABLE))
            except: in_table = False
            try: fs = float(r.Font.Size)
            except: fs = -1
            try: sb = round(float(p.Format.SpaceBefore), 2)
            except: sb = -1
            try: sa = round(float(p.Format.SpaceAfter), 2)
            except: sa = -1
            try: lh_rule = int(p.Format.LineSpacingRule)
            except: lh_rule = -1
            try: lh_val = round(float(p.Format.LineSpacing), 2)
            except: lh_val = -1
            try: snap = int(p.Format.SnapToGrid)
            except: snap = -1
            try: style = str(p.Style.NameLocal)
            except: style = '?'
            paras.append({
                'i': i, 'pg': pg, 'y': y, 'in_table': in_table,
                'fs': fs, 'sb': sb, 'sa': sa, 'lh_rule': lh_rule, 'lh_val': lh_val,
                'snap': snap, 'style': style, 'text': text[:30],
            })
    finally:
        d.Close(False)
        word.Quit()
    return paras


def main():
    print(f'Measuring de6e per-paragraph...')
    paras = measure(DOCX)

    # Compute advances
    out_rows = []
    for idx, p in enumerate(paras):
        nxt = paras[idx + 1] if idx + 1 < len(paras) else None
        abs_y = (p['pg'] - 1) * PAGE_HEIGHT + p['y'] if p['pg'] > 0 else None
        nxt_abs_y = ((nxt['pg'] - 1) * PAGE_HEIGHT + nxt['y']) if nxt and nxt['pg'] > 0 else None
        actual_adv = (nxt_abs_y - abs_y) if (abs_y is not None and nxt_abs_y is not None) else None
        # Predicted: sb + line + sa  (line interpreted as lh_val for rule=exact, otherwise nat lh from fs)
        lh = p['lh_val'] if p['lh_rule'] in (4, 3) else max(p['lh_val'], p['fs'] * 1.2)
        predicted = (p['sb'] if p['sb'] >= 0 else 0) + lh + (p['sa'] if p['sa'] >= 0 else 0)
        compaction = (predicted - actual_adv) if actual_adv is not None else None
        out_rows.append({
            **p,
            'actual_adv': round(actual_adv, 2) if actual_adv is not None else '',
            'predicted_adv': round(predicted, 2),
            'compaction': round(compaction, 2) if compaction is not None else '',
        })

    # Save CSV
    out_path = 'pipeline_data/de6e_per_para_advance.csv'
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
        writer.writeheader()
        writer.writerows(out_rows)
    print(f'Wrote {out_path}')

    # Aggregate: in-table paragraphs only, where compaction > 1pt
    in_tbl = [r for r in out_rows if r['in_table'] and r['compaction'] != '' and float(r['compaction']) > 1.0]
    print(f'\nIn-table paragraphs with compaction > 1pt: {len(in_tbl)}')
    # Top 20 compaction sites
    in_tbl.sort(key=lambda r: -float(r['compaction']))
    print(f'\nTop 20 compaction sites (in-table):')
    print(f'{"i":>4} {"pg":>3} {"y":>7} {"sb":>5} {"lh":>5} {"sa":>4} {"actual":>7} {"pred":>6} {"comp":>6}  text')
    for r in in_tbl[:20]:
        print(f'{r["i"]:>4} {r["pg"]:>3} {r["y"]:>7} {r["sb"]:>5} {r["lh_val"]:>5} {r["sa"]:>4} '
              f'{r["actual_adv"]:>7} {r["predicted_adv"]:>6} {r["compaction"]:>6}  {r["text"][:25]!r}')

    # Aggregate: cumulative compaction
    total_compact = sum(float(r['compaction']) for r in out_rows if r['compaction'] != '' and float(r['compaction']) > 0)
    total_expand = sum(float(r['compaction']) for r in out_rows if r['compaction'] != '' and float(r['compaction']) < 0)
    n_compact = sum(1 for r in out_rows if r['compaction'] != '' and float(r['compaction']) > 0)
    n_expand = sum(1 for r in out_rows if r['compaction'] != '' and float(r['compaction']) < 0)
    print(f'\nTotal Word compaction (pred > actual): {total_compact:.0f}pt across {n_compact} paragraphs')
    print(f'Total Word expansion (pred < actual): {total_expand:.0f}pt across {n_expand} paragraphs')


if __name__ == '__main__':
    main()
