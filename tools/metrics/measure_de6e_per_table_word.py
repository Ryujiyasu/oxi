"""Day 33 part 46 — Word per-table vertical span measurement for de6e.

For each table in de6e, compute Word's table-top y to first-paragraph-after-table y.
Span = (page_break-aware absolute y of after) - (absolute y of first cell).

Compare to Oxi sum_row_h from Day 33 part 45:
  t1=161.3, t2=144.5, t3=85.6, t4=85.6, t5=3028.5, t6=608.5
  total Oxi tables: 4113.5pt

Goal: identify which table holds Oxi's +728pt over-pump relative to Word.
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc
import glob

DOCX = glob.glob('tools/golden-test/documents/docx/de6e32b5960b*')[0]
PAGE_H = 841.95
WD_VPOS = 6
WD_PAGE = 3
WD_IN_TABLE = 12


def abs_y(pg, y):
    if pg is None or pg < 1 or y is None or y < 0:
        return None
    return (pg - 1) * PAGE_H + y


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        n_tables = d.Tables.Count
        print(f'de6e: {n_tables} tables, {d.Paragraphs.Count} paragraphs')
        for t_idx in range(1, n_tables + 1):
            t = d.Tables(t_idx)
            try: n_rows = t.Rows.Count
            except: n_rows = -1
            try: n_cols = t.Columns.Count
            except: n_cols = -1
            # First-char y of table (Range.Start)
            try:
                first = d.Range(t.Range.Start, t.Range.Start)
                top_y = round(first.Information(WD_VPOS), 2)
                top_pg = int(first.Information(WD_PAGE))
            except Exception as e:
                top_y, top_pg = -1, -1
            # First paragraph AFTER the table
            after_start = t.Range.End
            after_y, after_pg, after_text = -1, -1, ''
            try:
                # Step through document.paragraphs to find first paragraph with Range.Start >= after_start
                # AND in_table=False
                n_para = d.Paragraphs.Count
                for pi in range(1, n_para + 1):
                    p = d.Paragraphs(pi)
                    pr = p.Range
                    if pr.Start < after_start:
                        continue
                    cr = d.Range(pr.Start, pr.Start)
                    try: in_t = bool(cr.Information(WD_IN_TABLE))
                    except: in_t = False
                    if in_t:
                        continue
                    after_y = round(cr.Information(WD_VPOS), 2)
                    after_pg = int(cr.Information(WD_PAGE))
                    after_text = (pr.Text or '').replace('\r', ' ').replace('\x07', '').strip()[:25]
                    break
            except Exception as e:
                pass
            # Compute span
            ay_top = abs_y(top_pg, top_y)
            ay_after = abs_y(after_pg, after_y)
            span = (ay_after - ay_top) if (ay_top is not None and ay_after is not None) else None
            span_str = f'{span:.1f}pt' if span is not None else 'n/a'
            print(f'\nt{t_idx}: {n_rows} rows × {n_cols} cols')
            print(f'  Word top: pg={top_pg} y={top_y}')
            print(f'  Word after: pg={after_pg} y={after_y} text={after_text!r}')
            print(f'  Word span: {span_str}')

        # Also measure body-only advance (between tables / before t1 / after last t)
        # Approximation: total doc body = last_para.y - first_para.y
        # Body advance = total - sum(table spans)
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
