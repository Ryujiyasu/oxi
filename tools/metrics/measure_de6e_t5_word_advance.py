"""Day 33 part 42 — Measure Word's per-row advance in de6e t5.

Goal: identify Word's "hidden compaction" mechanism by measuring per-row
y advance against per-row content (spacing.before + line + spacing.after).
Rows where Word advances LESS than the sum = compaction sites.

For each row in de6e t5 (32 rows):
- Cell.Range first-char collapsed-y (Information(6))
- Row height (Rows(i).Height)
- HeightRule
- vAlign (Cell.VerticalAlignment)
- pPr spacing (before, after, line, lineRule) — via first cell first paragraph
- Predicted advance = before + line + after (without compaction)
- Actual advance = (next row word_y) - (current row word_y)
- Compaction = predicted - actual
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCX = 'tools/golden-test/documents/docx/de6e32b5960b_tokumei_08_01-1.docx'

# Word measurement constants
WD_INFORMATION_VERTICAL_POS = 6


def measure_table_rows(docx_path, t_idx):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    rows_data = []
    try:
        if d.Tables.Count < t_idx:
            print(f'  table {t_idx} not in doc (only {d.Tables.Count} tables)')
            return rows_data
        t = d.Tables(t_idx)
        n_rows = t.Rows.Count
        n_cols = t.Columns.Count
        print(f'  table {t_idx}: rows={n_rows} cols={n_cols}')

        for r in range(1, n_rows + 1):
            # Cell (r, 1) — first column
            try:
                cell = t.Cell(r, 1)
                rng = cell.Range
                first = d.Range(rng.Start, rng.Start)
                y = round(first.Information(WD_INFORMATION_VERTICAL_POS), 2)
                pg = int(first.Information(3))
                text = (rng.Text or '').replace('\r', ' ').replace('\x07', '').strip()[:20]
            except Exception as e:
                y, pg, text = -1, -1, f'err:{e}'
            try: row_h = round(float(t.Rows(r).Height), 2)
            except: row_h = -1
            try: row_h_rule = int(t.Rows(r).HeightRule)
            except: row_h_rule = -1
            try: v_align = int(cell.VerticalAlignment)
            except: v_align = -1

            # First paragraph pPr attrs
            try: p1 = cell.Range.Paragraphs(1)
            except: p1 = None
            fs, lh_rule, lh_val, sb, sa, snap, style = -1, -1, -1, -1, -1, -1, '?'
            if p1 is not None:
                try: fs = float(p1.Range.Font.Size)
                except: pass
                try: lh_rule = int(p1.Format.LineSpacingRule)
                except: pass
                try: lh_val = round(float(p1.Format.LineSpacing), 2)
                except: pass
                try: sb = round(float(p1.Format.SpaceBefore), 2)
                except: pass
                try: sa = round(float(p1.Format.SpaceAfter), 2)
                except: pass
                try: snap = int(p1.Format.SnapToGrid)
                except: pass
                try: style = str(p1.Style.NameLocal)
                except: pass
            rows_data.append({
                'r': r, 'pg': pg, 'y': y, 'text': text,
                'row_h': row_h, 'row_h_rule': row_h_rule,
                'v_align': v_align,
                'fs': fs, 'lh_rule': lh_rule, 'lh_val': lh_val,
                'sb': sb, 'sa': sa, 'snap': snap, 'style': style,
            })
    finally:
        d.Close(False)
        word.Quit()
    return rows_data


def main():
    print(f'=== de6e t5 per-row measurement ===\n')
    rows = measure_table_rows(DOCX, t_idx=5)
    print()
    print(f'{"r":>3} {"pg":>2} {"y":>7} {"row_h":>7} {"rule":>4} {"v_a":>3} '
          f'{"fs":>5} {"lhR":>3} {"lhV":>6} {"sb":>5} {"sa":>5} {"snap":>4}  '
          f'{"adv_y":>7} {"adv_pg":>4}  text  /  style')
    prev_y = None
    prev_pg = None
    abs_y = lambda r: r['y'] + (r['pg'] - 1) * 841.95  # A4 height in pt
    for row in rows:
        if prev_y is not None and row['pg'] == prev_pg:
            adv = row['y'] - prev_y
            adv_str = f'{adv:+6.2f}'
        elif prev_y is not None and row['pg'] != prev_pg:
            adv = abs_y(row) - (prev_y + (prev_pg - 1) * 841.95)
            adv_str = f'{adv:+6.2f}*'  # crossed page
        else:
            adv_str = '   ---'
        print(f'{row["r"]:>3} {row["pg"]:>2} {row["y"]:>7} {row["row_h"]:>7} '
              f'{row["row_h_rule"]:>4} {row["v_align"]:>3} '
              f'{row["fs"]:>5} {row["lh_rule"]:>3} {row["lh_val"]:>6} '
              f'{row["sb"]:>5} {row["sa"]:>5} {row["snap"]:>4}  '
              f'{adv_str:>7} {"" if row["pg"]==prev_pg or prev_pg is None else "PG":>4}  '
              f'{row["text"]!r}  /  {row["style"]!r}')
        prev_y = row['y']
        prev_pg = row['pg']


if __name__ == '__main__':
    main()
