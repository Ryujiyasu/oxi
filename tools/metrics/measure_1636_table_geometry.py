"""Day 33 part 34 (2026-05-11) — Measure 1636 first-table geometry in Word.

Goal: identify the +8.6pt initial-table offset source.
- Word's table top y (where row 0 begins)
- First cell's content top y
- Border thickness
- Cell padding / margin
- vAlign offset
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCX = 'tools/golden-test/documents/docx/1636d28e2c46_tokumei_08_04.docx'


def main():
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(DOCX)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        # First few paragraphs
        n = d.Paragraphs.Count
        print('First 8 paragraphs:')
        for i in range(1, min(8, n) + 1):
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            try: pg = int(cr.Information(3))
            except: pg = -1
            try: y = round(cr.Information(6), 2)
            except: y = -1
            try: text = (r.Text or '').strip()[:25]
            except: text = ''
            try: in_table = bool(r.Information(12))
            except: in_table = False
            print(f'  wi={i} pg={pg} y={y:>7} in_table={in_table} text={text!r}')

        # First table
        tables = d.Tables
        print(f'\nDocument has {tables.Count} tables')
        if tables.Count >= 1:
            t = tables(1)
            try: n_rows = t.Rows.Count
            except: n_rows = -1
            try: n_cols = t.Columns.Count
            except: n_cols = -1
            print(f'\nTable 1: rows={n_rows} cols={n_cols}')

            # Row 0 first cell
            try:
                c = t.Cell(1, 1)
                r = c.Range
                first_char_range = d.Range(r.Start, r.Start)
                cell_y = round(first_char_range.Information(6), 2)
                cell_pg = int(first_char_range.Information(3))
                try: cell_height = round(c.Height, 2)
                except: cell_height = -1
                try: cell_height_rule = c.HeightRule
                except: cell_height_rule = -1
                try: top_pad = c.TopPadding
                except: top_pad = -1
                try: bottom_pad = c.BottomPadding
                except: bottom_pad = -1
                try: left_pad = c.LeftPadding
                except: left_pad = -1
                try: right_pad = c.RightPadding
                except: right_pad = -1
                try: v_align = c.VerticalAlignment
                except: v_align = -1
                print(f'  Cell(1,1): first_char_y={cell_y} pg={cell_pg}')
                print(f'    height={cell_height} rule={cell_height_rule}')
                print(f'    padding: top={top_pad} bottom={bottom_pad} left={left_pad} right={right_pad}')
                print(f'    v_align={v_align}')

                # Borders
                for bdr_label, bdr in [('top', t.Borders(-1)), ('inside_h', t.Borders(-7))]:
                    try: line_width = bdr.LineWidth
                    except: line_width = -1
                    try: line_style = bdr.LineStyle
                    except: line_style = -1
                    print(f'  Table border {bdr_label}: width={line_width} style={line_style}')
            except Exception as e:
                print(f'  Cell access error: {e}')

            # Row 0 dimensions
            try:
                r0 = t.Rows(1)
                try: r0_height = round(r0.Height, 2)
                except: r0_height = -1
                try: r0_height_rule = r0.HeightRule
                except: r0_height_rule = -1
                try: r0_topPad = r0.TopPadding
                except: r0_topPad = -1
                print(f'  Row 1: height={r0_height} rule={r0_height_rule} topPadding={r0_topPad}')
            except Exception as e:
                print(f'  Row access error: {e}')
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
