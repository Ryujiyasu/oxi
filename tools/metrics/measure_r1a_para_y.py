"""S136: Re-measure R1A_spacing_lineRule and R1A_all4 to check
whether Day 33 part 17's "Word suppresses spacing.before for first
paragraph in cell" is true at the PARAGRAPH-Y level (not just Row.Height).

V200 shows Word DOES apply sb to row extent (16.55pt = 12 line + 4.35 sb).
Day 33 part 17 claimed Word renders row at 12.5pt. Discrepancy: maybe
they measured Row.Height (excludes sb) vs visual extent (includes sb).

Here we measure both:
- Row.Height via Word COM API
- Cell para y via Information(6)
- Following body para y (= cell bottom + table border + sb_body if any)

Net per-cell-extent = body_para_y - cell_para_y. If = 12.5pt → sb suppressed.
If = 16.85pt → sb applied.
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'row1_attr_isolation')

VARIANTS = [
    'R1A_baseline',
    'R1A_spacing_before',
    'R1A_spacing_lineRule',
    'R1A_lineRule_exact',
    'R1A_vAlign_center',
    'R1A_vAlign_lineRule',
    'R1A_pStyle_ac',
    'R1A_all4',
]


def measure(docx_path: str) -> dict:
    import win32com.client as wc
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(docx_path, ReadOnly=True)
    out = {}
    try:
        # Page geometry
        ps = d.PageSetup
        out['top_margin'] = round(ps.TopMargin, 3)

        # All paragraphs
        paras = []
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            txt = (p.Range.Text or '').strip()
            rng = p.Range
            cr = d.Range(rng.Start, rng.Start)
            paras.append({
                'i': i,
                'text': txt[:40],
                'page': int(cr.Information(3)),
                'y': round(cr.Information(6), 3),  # vertical pos rel to page
                'x': round(cr.Information(5), 3),
                'in_table': bool(cr.Information(12)),
            })
        out['paragraphs'] = paras

        # Table 1 row 1 metrics
        if d.Tables.Count >= 1:
            t = d.Tables(1)
            row = t.Rows(1)
            cell = row.Cells(1)
            out['table1'] = {
                'row_height': round(row.Height, 3),  # Word's reported row height
                'row_height_rule': int(row.HeightRule),  # 0=auto, 1=at least, 2=exact
                'cell_height': round(cell.Height, 3),
                'cell_top_padding': round(cell.TopPadding, 3),
                'cell_bottom_padding': round(cell.BottomPadding, 3),
            }
    finally:
        d.Close(False)
        word.Quit()
    return out


def main():
    results = []
    for label in VARIANTS:
        docx = os.path.join(REPRO_DIR, f'{label}.docx')
        if not os.path.exists(docx):
            print(f'SKIP {label}')
            continue
        print(f'=== {label} ===')
        try:
            r = measure(docx)
        except Exception as e:
            print(f'  ERROR: {e}')
            continue
        r['label'] = label
        # Derive: cell extent = body_para_y - cell_para_y
        paras = r.get('paragraphs', [])
        cell_para = next((p for p in paras if p['in_table']), None)
        body_para = next((p for p in paras if not p['in_table']), None)
        if cell_para and body_para:
            extent = round(body_para['y'] - cell_para['y'], 3)
            r['cell_extent_via_paras'] = extent
            print(f'  cell_para y={cell_para["y"]} text={cell_para["text"][:25]!r}')
            print(f'  body_para y={body_para["y"]} text={body_para["text"][:25]!r}')
            print(f'  cell_extent (body-cell) = {extent}pt')
        if 'table1' in r:
            t1 = r['table1']
            print(f'  Word Row.Height={t1["row_height"]}pt (rule={t1["row_height_rule"]}) '
                  f'Cell.Height={t1["cell_height"]}pt '
                  f'pad_t={t1["cell_top_padding"]}pt pad_b={t1["cell_bottom_padding"]}pt')
        results.append(r)
        print()
    out_path = os.path.join(REPO, 'pipeline_data', 'r1a_para_y_results.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'Wrote {out_path}')


if __name__ == '__main__':
    main()
