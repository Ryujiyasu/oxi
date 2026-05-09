"""Day 32 part 6 continued — bd90b00 trigger paragraph deep dive.

Day 32 part 6 found 4 drift triggers in bd90b00:
- pi=2 (Bug 2 initial)
- pi=11 (out-of-table, +3.0pt jump)
- pi=33 (table first-row first-para, +2.3pt jump)
- pi=42 (※ comment, +1.3pt jump)

This tool inspects each trigger paragraph's structure (font, size,
line_height, snap, tabs, runs) to identify what makes them differ.
"""
from __future__ import annotations
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding='utf-8')


def main():
    import win32com.client as wc
    docx = r'C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\bd90b00ab7a7_order_05.docx'
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
    try:
        targets = [1, 2, 9, 10, 11, 12, 13, 26, 27, 32, 33, 34, 41, 42, 43, 44]
        print(f'{"i":>3} {"pg":>2} {"y":>7} {"lh_set":>10} {"sz":>5} {"snap":>4} {"in_t":>4} text')
        for i in targets:
            p = d.Paragraphs(i)
            r = p.Range
            cr = d.Range(r.Start, r.Start)
            text = (r.Text or '').strip()
            try:
                lh_rule = p.Format.LineSpacingRule  # 0=single,1=1.5,2=double,3=at_least,4=exactly,5=multiple
                lh_val = p.Format.LineSpacing
            except Exception:
                lh_rule, lh_val = -1, -1
            try:
                snap = p.Format.SnapToGrid
            except Exception:
                snap = -1
            try:
                fs = r.Font.Size
            except Exception:
                fs = -1
            in_t = bool(r.Information(12))
            print(f'{i:>3} {int(cr.Information(3)):>2} {round(cr.Information(6), 2):>7} {f"{lh_rule}/{lh_val}":>10} {fs:>5} {snap:>4} {"T" if in_t else "-":>4} {text[:60]!r}')

        # For pi=11 vs pi=10/12: get all run details
        print('\n--- pi=11 detailed runs ---')
        p11 = d.Paragraphs(11)
        runs = []
        for run_idx in range(1, p11.Range.Runs.Count + 1):
            try:
                run = p11.Range.Runs(run_idx)
                runs.append({
                    'i': run_idx,
                    'text': (run.Text or '')[:30],
                    'font_name': run.Font.Name,
                    'font_size': run.Font.Size,
                    'bold': run.Font.Bold,
                    'underline': run.Font.Underline,
                })
            except Exception as e:
                pass
        for r in runs:
            print(f'  run {r["i"]:>2}: {r["font_name"]:<20} sz={r["font_size"]:>5} {r["text"]!r}')

        # For pi=33 cell first paragraph
        print('\n--- pi=33 (table first-row first-para) cell info ---')
        p33 = d.Paragraphs(33)
        try:
            cell = p33.Range.Cells(1)
            print(f'  cell row: {cell.RowIndex}, column: {cell.ColumnIndex}')
            print(f'  cell vAlign: {cell.VerticalAlignment}')
            print(f'  cell height_rule: {cell.HeightRule}, height: {cell.Height}')
        except Exception as e:
            print(f'  not in table or error: {e}')

        # For pi=42 (※ comment after table)
        print('\n--- pi=42 (※ comment after table) ---')
        p42 = d.Paragraphs(42)
        text = (p42.Range.Text or '').strip()
        print(f'  text: {text[:80]!r}')
        try:
            print(f'  lh_rule: {p42.Format.LineSpacingRule}, lh_val: {p42.Format.LineSpacing}')
            print(f'  fs first: {p42.Range.Runs(1).Font.Size}')
            print(f'  space_before: {p42.Format.SpaceBefore}')
            print(f'  space_after: {p42.Format.SpaceAfter}')
        except Exception as e:
            print(f'  err: {e}')
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    main()
