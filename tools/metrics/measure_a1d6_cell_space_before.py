"""COM survey: a1d6 first-paragraph-in-cell space_before applied vs suppressed.

Day 37 autonomous loop, Session 58.

Goal: discover Word's hidden differentiator for "suppress space_before of first
paragraph in cell" rule. Day 33 part 17 broad rule (suppress for ALL first-in-cell)
caused 1 a1d6 outlier (wi=141 over-suppressed). Day 37 narrowings (row_idx==0 or
cell.blocks.len()==1) caused 21 a1d6 outliers (Word DOES suppress in those
mid-row / multi-paragraph cells too). Need to find what differentiates the cells
where Word applies vs suppresses.

Method:
- For each Table -> flat Cells collection (avoids vMerge row enumeration error):
  - For each cell -> first paragraph in cell:
  - Measure Y of first line (Information(6) on collapsed start)
  - Measure Y of cell start (cell.Range.Start collapsed Information(6))
  - Measure Y of all paragraphs in cell -> compute applied gaps
  - Capture cell/row/table context: index, count, height-rule, height, vMerge,
    cantSplit, vertical-alignment, top-padding, pStyle, space_before setting
- Classify each as applied/suppressed/ambiguous based on:
  - If 2+ paragraphs in cell:
      gap1 = first_y - cell_start_y
      gap_inter = second_y - first_y (full line_height + p2.space_before)
      If gap1 - gap_inter < -1pt and p1.sb > 0 -> likely suppressed
      If |gap1 - gap_inter| < 1pt -> likely applied
- If 1 paragraph in cell: less reliable, attempt cell_top_y -> first_y inference

Output: pipeline_data/a1d6_first_para_in_cell_survey.json
"""
from __future__ import annotations
import os, sys, json, traceback
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = os.path.abspath('tools/golden-test/documents/docx/a1d6e4efa2e7_tokumei_08_01-4.docx')
OUT = os.path.abspath('pipeline_data/a1d6_first_para_in_cell_survey.json')


def info5_pt(d, pos):
    try: return float(d.Range(pos, pos).Information(5))
    except Exception: return None

def info6_pt(d, pos):
    try: return float(d.Range(pos, pos).Information(6))
    except Exception: return None

def info1_page(d, pos):
    try: return int(d.Range(pos, pos).Information(1))
    except Exception: return None

def safe_get(obj, attr, default=None):
    try: return getattr(obj, attr)
    except Exception: return default

def short(s, n=40):
    s = (s or '').rstrip('\r\n\x07')
    return s if len(s) <= n else s[:n] + '...'


def measure_cell(doc, cell, table_index, table_n_rows):
    try:
        cell_paras = cell.Range.Paragraphs
        n_paras = cell_paras.Count
    except Exception:
        return None
    if n_paras < 1:
        return None
    try:
        p1 = cell_paras(1)
    except Exception:
        return None
    p1_rng = p1.Range
    p1_start = p1_rng.Start
    p1_y = info6_pt(doc, p1_start)
    p1_x = info5_pt(doc, p1_start)
    p1_page = info1_page(doc, p1_start)
    p1_text = short(p1_rng.Text, 40)
    fmt = p1.Format
    sb_pt = safe_get(fmt, 'SpaceBefore', None)
    sa_pt = safe_get(fmt, 'SpaceAfter', None)
    ls_pt = safe_get(fmt, 'LineSpacing', None)
    ls_rule = safe_get(fmt, 'LineSpacingRule', None)
    p_style_obj = safe_get(p1, 'Style', None)
    p_style = safe_get(p_style_obj, 'NameLocal', None) if p_style_obj else None
    # Para index in whole doc
    p1_index_in_doc = None
    try:
        for pi in range(1, doc.Paragraphs.Count + 1):
            if doc.Paragraphs(pi).Range.Start == p1_start:
                p1_index_in_doc = pi
                break
    except Exception:
        pass
    # Capture up to 4 paragraphs in cell with their Y
    para_y_list = []
    for i in range(1, min(n_paras, 4) + 1):
        try:
            pi_obj = cell_paras(i)
            y_i = info6_pt(doc, pi_obj.Range.Start)
            sb_i = safe_get(pi_obj.Format, 'SpaceBefore', None)
            txt_i = short(pi_obj.Range.Text, 25)
            para_y_list.append({'i': i, 'y': round(y_i, 3) if y_i else None,
                                'sb': sb_i, 'text': txt_i})
        except Exception:
            pass
    # Cell start Y
    cell_start_y = info6_pt(doc, cell.Range.Start)
    # Row geometry via cell properties
    row_idx = safe_get(cell, 'RowIndex', None)
    col_idx = safe_get(cell, 'ColumnIndex', None)
    top_pad = safe_get(cell, 'TopPadding', None)
    bot_pad = safe_get(cell, 'BottomPadding', None)
    vert_align = safe_get(cell, 'VerticalAlignment', None)
    # Try row-level lookup
    row_height_rule = None
    row_height_setting = None
    cant_split = None
    try:
        row = cell.Row
        row_height_rule = safe_get(row, 'HeightRule', None)
        row_height_setting = safe_get(row, 'Height', None)
        cant_split = safe_get(row, 'AllowBreakAcrossPages', None)
    except Exception:
        pass
    rec = {
        'table_index': table_index,
        'n_rows_in_table': table_n_rows,
        'cell_row_index_word': row_idx,
        'cell_col_index_word': col_idx,
        'n_paras_in_cell': n_paras,
        'row_height_rule': row_height_rule,
        'row_height_setting': row_height_setting,
        'cant_split': cant_split,
        'top_padding': top_pad,
        'bottom_padding': bot_pad,
        'vertical_alignment': vert_align,
        'p1_index_in_doc': p1_index_in_doc,
        'p1_text': p1_text,
        'p1_style': p_style,
        'p1_space_before_pt': sb_pt,
        'p1_space_after_pt': sa_pt,
        'p1_line_spacing_pt': ls_pt,
        'p1_line_rule': ls_rule,
        'p1_y': round(p1_y, 3) if p1_y else None,
        'p1_x': round(p1_x, 3) if p1_x else None,
        'p1_page': p1_page,
        'cell_start_y': round(cell_start_y, 3) if cell_start_y else None,
        'paragraphs': para_y_list,
    }
    return rec


def main():
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()
        result = {'doc': DOC, 'tables': []}
        n_tables = doc.Tables.Count
        print(f'[+] {n_tables} tables')
        for ti in range(1, n_tables + 1):
            table = doc.Tables(ti)
            t_entry = {
                'table_index': ti,
                'n_rows': table.Rows.Count,
                'n_columns': safe_get(table.Columns, 'Count', None) if safe_get(table, 'Columns', None) else None,
                'cells': [],
            }
            # Use flat table.Range.Cells iteration to bypass vMerge row error
            try:
                flat_cells = table.Range.Cells
                n_flat = flat_cells.Count
            except Exception:
                print(f'  table {ti}: flat enumeration failed')
                result['tables'].append(t_entry)
                continue
            for fi in range(1, n_flat + 1):
                try:
                    cell = flat_cells(fi)
                except Exception:
                    continue
                rec = measure_cell(doc, cell, ti, table.Rows.Count)
                if rec is not None:
                    t_entry['cells'].append(rec)
            result['tables'].append(t_entry)
            print(f'  table {ti}: {len(t_entry["cells"])} cell-first-paras (flat n={n_flat})')
        with open(OUT, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f'[+] wrote {OUT}')
        n_total = sum(len(t['cells']) for t in result['tables'])
        n_multi = sum(1 for t in result['tables'] for c in t['cells'] if c['n_paras_in_cell'] >= 2)
        print(f'Total: {n_total}, multi-paragraph cells: {n_multi}')
    except Exception:
        traceback.print_exc()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()


if __name__ == '__main__':
    main()
