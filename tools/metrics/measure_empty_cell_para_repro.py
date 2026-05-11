"""Day 33 part 15 — Measure Word's row height for empty_cell_para repros.

For each variant, get Tables(1).Rows(1).Range collapsed-start y (= row top)
and compare to anchor paragraph y (next paragraph after table).

row_height = anchor_y - row_y - row_y_to_anchor_padding
But simpler: just compute anchor_y - row_y = total_row_height + post-table-spacing.
"""
from __future__ import annotations
import os, sys, json, glob
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
REPRO_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'empty_cell_para')
OUT = os.path.join(REPO, 'pipeline_data', 'empty_cell_para_repro.json')


def measure(word, docx_path: str) -> dict:
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    out = {}
    try:
        # Row 1 of table 1 top
        row = d.Tables(1).Rows(1)
        cr_row = d.Range(row.Range.Start, row.Range.Start)
        out['row_y_pt'] = cr_row.Information(6)
        # Each cell: paragraph y values
        cells_data = []
        for ci in range(1, row.Cells.Count + 1):
            cell = row.Cells(ci)
            paras = cell.Range.Paragraphs
            ys = []
            for pi in range(1, paras.Count + 1):
                p = paras(pi)
                cr = d.Range(p.Range.Start, p.Range.Start)
                txt = (p.Range.Text or '').rstrip('\r\x07')
                ys.append({
                    'pi': pi,
                    'y_pt': cr.Information(6),
                    'is_empty': not txt.strip(),
                    'text': txt[:20],
                    'fs': p.Range.Font.Size,
                })
            cells_data.append({'ci': ci, 'paras': ys})
        out['cells'] = cells_data
        # Anchor (paragraph after table)
        n_paras = d.Paragraphs.Count
        # Find first paragraph after table — typically right after row's range end
        # Or simpler: find "下のアンカー"
        for i in range(1, n_paras + 1):
            p = d.Paragraphs(i)
            if '下のアンカー' in (p.Range.Text or ''):
                cr = d.Range(p.Range.Start, p.Range.Start)
                out['anchor_y_pt'] = cr.Information(6)
                out['anchor_i'] = i
                break
        # Row height advance
        if 'anchor_y_pt' in out:
            out['anchor_minus_row_y_pt'] = out['anchor_y_pt'] - out['row_y_pt']
    finally:
        d.Close(SaveChanges=False)
    return out


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    docs = sorted(glob.glob(os.path.join(REPRO_DIR, '*.docx')))
    results = []
    try:
        for fp in docs:
            label = os.path.basename(fp).replace('.docx', '')
            m = measure(word, fp)
            m['label'] = label
            results.append(m)
            print(f'{label}: row_y={m.get("row_y_pt")} anchor_y={m.get("anchor_y_pt")} delta={m.get("anchor_minus_row_y_pt")}')
            for cell in m.get('cells', []):
                for p in cell['paras']:
                    print(f'    cell{cell["ci"]} p{p["pi"]} y={p["y_pt"]:.2f} empty={p["is_empty"]} fs={p["fs"]} text={p["text"]!r}')
    finally:
        word.Quit()
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\nSaved: {OUT}')

if __name__ == '__main__':
    main()
