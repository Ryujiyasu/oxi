"""Day 33 part 15 — Direct table row y measurement via Word COM.

For each table in the doc, walk Tables(t).Rows(r) and report:
- row.Range.Information(6) → row top y (collapsed to start)
- row.Cells.Count → number of cells
- per-cell first-paragraph y (collapsed start)
- row.Height (twips)
- row.HeightRule (auto/atLeast/exact)

Output: pipeline_data/table_row_y_<doc_id>.json
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))


def measure(docx_path: str, label: str) -> list:
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    out = []
    try:
        n_tables = d.Tables.Count
        print(f'  Tables: {n_tables}')
        for ti in range(1, min(n_tables, 3) + 1):
            tbl = d.Tables(ti)
            n_rows = tbl.Rows.Count
            print(f'  Table {ti}: {n_rows} rows')
            for ri in range(1, n_rows + 1):
                try:
                    row = tbl.Rows(ri)
                    rng = row.Range
                    cr_start = d.Range(rng.Start, rng.Start)
                    row_y = cr_start.Information(6)
                    row_pg = cr_start.Information(3)
                    height_twips = row.Height  # twips
                    height_rule = row.HeightRule  # 0=auto, 1=atLeast, 2=exact
                    n_cells = row.Cells.Count
                    cell_data = []
                    for ci in range(1, n_cells + 1):
                        try:
                            cell = row.Cells(ci)
                            crng = cell.Range
                            ccr = d.Range(crng.Start, crng.Start)
                            cell_y = ccr.Information(6)
                            cell_x = ccr.Information(5)  # X
                            cell_w_twips = cell.Width  # pt actually
                            n_paras = cell.Range.Paragraphs.Count
                            valign = None
                            try:
                                valign = cell.VerticalAlignment
                            except:
                                pass
                            cell_data.append({
                                'ci': ci,
                                'y_pt': cell_y,
                                'x_pt': cell_x,
                                'width_pt': cell_w_twips,
                                'n_paras': n_paras,
                                'valign': valign,
                            })
                        except Exception as e:
                            cell_data.append({'ci': ci, 'error': str(e)[:60]})
                    entry = {
                        'table': ti,
                        'row': ri,
                        'pg': row_pg,
                        'row_y_pt': row_y,
                        'row_height_twips': height_twips,
                        'row_height_rule': height_rule,
                        'n_cells': n_cells,
                        'cells': cell_data,
                    }
                    out.append(entry)
                    print(f'    row {ri}: y={row_y:>7.2f} h={height_twips/20:>6.2f}pt rule={height_rule} cells={n_cells}')
                except Exception as e:
                    print(f'    row {ri}: ERROR {e!r}')
    finally:
        d.Close(SaveChanges=False)
        word.Quit()
    out_path = os.path.join(REPO, 'pipeline_data', f'table_row_y_{label}.json')
    json.dump(out, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'\n  Saved: {out_path}\n')
    return out


def main():
    docs = [
        ('d4d126dfe1d9', 'tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx'),
        ('664c38001b40', 'tools/golden-test/documents/docx/664c38001b40_order_12.docx'),
        ('de6e32b5960b', 'tools/golden-test/documents/docx/de6e32b5960b_tokumei_08_01-1.docx'),
    ]
    for label, path in docs:
        print(f'=== {label} ===')
        measure(path, label)


if __name__ == '__main__':
    main()
