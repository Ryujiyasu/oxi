"""d4d126 全テーブルの first-paragraph-in-cell sb-applied-vs-suppressed 計測.

a1d6 と同じ方法 (cell_start_y からの前段落 delta) で d4d126 のテーブル先頭段落を判定。
これにより memory note の「d4d126 では Word が sb 適用」主張が正しいか確認する。
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = os.path.abspath('tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx')
OUT = os.path.abspath('pipeline_data/d4d126_first_para_in_cell_survey.json')

def info5(d, pos):
    try: return float(d.Range(pos, pos).Information(5))
    except: return None
def info6(d, pos):
    try: return float(d.Range(pos, pos).Information(6))
    except: return None
def info1(d, pos):
    try: return int(d.Range(pos, pos).Information(1))
    except: return None
def safe(o, a, dft=None):
    try: return getattr(o, a)
    except: return dft
def short(s, n=40):
    s = (s or '').rstrip('\r\n\x07')
    return s if len(s) <= n else s[:n] + '...'

def measure_cell(doc, cell, ti, tn_rows):
    try:
        cps = cell.Range.Paragraphs
        n = cps.Count
    except: return None
    if n < 1: return None
    try: p1 = cps(1)
    except: return None
    p1r = p1.Range
    p1_start = p1r.Start
    fmt = p1.Format
    cell_start = cell.Range.Start
    cell_start_y = info6(doc, cell_start)
    return dict(
        table_index=ti,
        n_rows_in_table=tn_rows,
        cell_row_index_word=safe(cell, 'RowIndex'),
        cell_col_index_word=safe(cell, 'ColumnIndex'),
        n_paras_in_cell=n,
        top_padding=safe(cell, 'TopPadding', 0.0),
        p1_index_in_doc=safe(p1, 'Range') and None,  # we'll fill below
        p1_text=short(p1r.Text, 40),
        p1_space_before_pt=safe(fmt, 'SpaceBefore'),
        p1_line_spacing_pt=safe(fmt, 'LineSpacing'),
        p1_line_rule=safe(fmt, 'LineSpacingRule'),
        p1_y=info6(doc, p1_start),
        p1_x=info5(doc, p1_start),
        p1_page=info1(doc, p1_start),
        cell_start_y=cell_start_y,
        paragraphs=[
            dict(i=i, y=info6(doc, cps(i).Range.Start), sb=float(cps(i).Format.SpaceBefore or 0),
                 text=short(cps(i).Range.Text, 30))
            for i in range(1, min(n, 4) + 1)
        ],
    )

def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(DOC, ReadOnly=True)
    try:
        tables = doc.Tables
        n_tables = tables.Count
        print(f'doc: {os.path.basename(DOC)} n_tables: {n_tables}')
        data = dict(doc=os.path.basename(DOC), tables=[])
        for ti in range(1, n_tables + 1):
            t = tables(ti)
            tn_rows = t.Rows.Count
            cells_t = t.Range.Cells
            n_cells = cells_t.Count
            print(f'\nTable {ti}: rows={tn_rows} cells={n_cells}')
            cells_data = []
            for ci in range(1, n_cells + 1):
                try:
                    c = cells_t(ci)
                except Exception as e:
                    continue
                cd = measure_cell(doc, c, ti, tn_rows)
                if cd:
                    cells_data.append(cd)
            data['tables'].append(dict(
                table_index=ti, n_rows=tn_rows, n_cells=len(cells_data), cells=cells_data
            ))
        with open(OUT, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f'\nWrote {OUT}')
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

if __name__ == '__main__':
    main()
