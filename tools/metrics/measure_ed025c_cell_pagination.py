"""Measure Word's per-page-per-cell paragraph count in ed025c 損益計算書 table.

Day 37 multi-session investigation. Per [[session58_day37_per_cell_full_refactor_recursive_lrpb_consumed]],
ed025c's 1 remaining outlier (17 cell-2 × × × delta=-1) is caused by Word's
cell-content-height calc being smaller than Oxi's available area. Word fits
~21 × × × per page in cell 2 while Oxi fits ~39. Source of the discrepancy is
non-LRPB (LRPB consumed by first split; recursive uses row-wide split).

This script COM-measures, for each cell of the 損益計算書 table:
  - cell column index
  - per-paragraph: text prefix, page, y
  - per-page-per-cell: count, first y, last y, content_h span

Output: pipeline_data/ra_manual_measurements/ed025c_cell_pagination.json

Cross-compare with pipeline_data/pagination_oxi/ed025cbecffb.json to find:
  - Word's per-page paragraph count per cell
  - Where Word's break point differs from Oxi's
  - Whether the offset is uniform per page or accumulating

Pre-req: ed025c must be openable in Word; no special setup.
"""
from __future__ import annotations
import os, sys, json, traceback
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\ed025cbecffb_index-23.docx'
OUT = r'c:\Users\ryuji\oxi-main\pipeline_data\ra_manual_measurements\ed025c_cell_pagination.json'


def main():
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()

        # Find 損益計算書 table — look for "Ⅰ営業損益" paragraph, then its containing table
        target_text = '営業損益'
        target_para_idx = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            txt = (doc.Paragraphs(pi).Range.Text or '').strip()
            if 'Ⅰ' in txt and target_text in txt:
                target_para_idx = pi
                break
        if not target_para_idx:
            print('ERROR: target paragraph not found')
            return
        para = doc.Paragraphs(target_para_idx)
        # Find enclosing table via Range.Tables
        tables = para.Range.Tables
        if tables.Count == 0:
            print(f'ERROR: paragraph {target_para_idx} not in a table')
            return
        table = tables(1)
        print(f'Target table found. Rows={table.Rows.Count}, Columns={table.Columns.Count}')

        # Iterate cells flat (vMerge-safe; same pattern as a1d6 survey)
        all_cells_data = []
        for ci in range(1, table.Range.Cells.Count + 1):
            try:
                cell = table.Range.Cells(ci)
                row_idx = cell.RowIndex
                col_idx = cell.ColumnIndex
                # Get cell paragraphs
                paras = cell.Range.Paragraphs
                cell_paras = []
                for pi_in_cell in range(1, paras.Count + 1):
                    p = paras(pi_in_cell)
                    txt = (p.Range.Text or '').rstrip('\r\n\x07')[:40]
                    rng = p.Range
                    start_rng = doc.Range(rng.Start, rng.Start)
                    try:
                        page = int(start_rng.Information(1))  # wdActiveEndPageNumber=1
                        y = float(start_rng.Information(6))   # wdVerticalPositionRelativeToPage=6
                    except Exception as e:
                        page = -1
                        y = -1.0
                    cell_paras.append({
                        'cpi': pi_in_cell - 1,  # 0-based to match Oxi
                        'text': txt,
                        'page': page,
                        'y': round(y, 2),
                    })
                all_cells_data.append({
                    'row_idx': row_idx,
                    'col_idx': col_idx,
                    'n_paragraphs': len(cell_paras),
                    'paragraphs': cell_paras,
                })
                print(f'Cell row={row_idx} col={col_idx}: {len(cell_paras)} paragraphs')
            except Exception as e:
                print(f'Cell {ci}: error {e}')

        # Compute per-page-per-cell summaries
        # key: (row_idx, col_idx) -> page -> [(cpi, text, y), ...]
        from collections import defaultdict
        per_cell_per_page = defaultdict(lambda: defaultdict(list))
        for cell in all_cells_data:
            key = (cell['row_idx'], cell['col_idx'])
            for p in cell['paragraphs']:
                per_cell_per_page[key][p['page']].append((p['cpi'], p['text'], p['y']))

        summaries = []
        for (ri, ci), pages_map in sorted(per_cell_per_page.items()):
            for pg in sorted(pages_map.keys()):
                paras_on_page = pages_map[pg]
                ys = [y for _, _, y in paras_on_page if y > 0]
                if ys:
                    first_y = min(ys)
                    last_y = max(ys)
                    span = last_y - first_y
                else:
                    first_y = last_y = span = -1
                summaries.append({
                    'cell_row': ri,
                    'cell_col': ci,
                    'page': pg,
                    'n_paragraphs': len(paras_on_page),
                    'first_y': first_y,
                    'last_y': last_y,
                    'span_pt': round(span, 2) if span >= 0 else -1,
                    'first_text': paras_on_page[0][1] if paras_on_page else '',
                    'last_text': paras_on_page[-1][1] if paras_on_page else '',
                })

        # Page setup for context
        ps = doc.Sections(1).PageSetup
        page_info = {
            'page_width': ps.PageWidth,
            'page_height': ps.PageHeight,
            'top_margin': ps.TopMargin,
            'bottom_margin': ps.BottomMargin,
            'footer_distance': ps.FooterDistance,
        }

        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        result = {
            'doc': 'ed025cbecffb_index-23.docx',
            'target_paragraph_idx': target_para_idx,
            'table_rows': table.Rows.Count,
            'table_columns': table.Columns.Count,
            'page_info': page_info,
            'per_page_per_cell_summary': summaries,
            'cells_detail': all_cells_data,
        }
        with open(OUT, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f'\nSaved to {OUT}')

        # Print summary table
        print('\n=== Per-page-per-cell summary ===')
        print(f'{"row":>3} {"col":>3} {"page":>4} {"n_paras":>7} {"first_y":>7} {"last_y":>7} {"span":>6}  first_text')
        for s in summaries:
            print(f'{s["cell_row"]:>3} {s["cell_col"]:>3} {s["page"]:>4} {s["n_paragraphs"]:>7} {s["first_y"]:>7} {s["last_y"]:>7} {s["span_pt"]:>6}  {s["first_text"][:30]!r}')

    except Exception:
        traceback.print_exc()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()


if __name__ == '__main__':
    main()
