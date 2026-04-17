"""Measure Word's row height for all multi-line cells in b35 and
record (n_lines, font_size, pitch, row_height) tuples for formula research.
"""
import sys, os, json, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
results = []
try:
    wdoc.Repaginate()
    for ti in range(1, wdoc.Tables.Count + 1):
        tbl = wdoc.Tables(ti)
        n_rows = tbl.Rows.Count
        n_cols = tbl.Columns.Count
        # Collect per-row first cell Y (use col with smallest index that has a cell)
        row_ys = {}
        for r in range(1, n_rows + 1):
            for c in range(1, n_cols + 1):
                try:
                    cell = tbl.Cell(r, c)
                    y = cell.Range.Information(6)
                    row_ys[r] = (c, y, cell)
                    break
                except:
                    continue
        # Pair consecutive rows
        for r in sorted(row_ys.keys())[:-1]:
            next_r = r + 1
            if next_r not in row_ys: continue
            c1, y1, cell1 = row_ys[r]
            c2, y2, cell2 = row_ys[next_r]
            row_height = y2 - y1
            # Count lines across all cells of this row
            max_lines_this_row = 0
            for c in range(1, n_cols + 1):
                try:
                    cell = tbl.Cell(r, c)
                    rng = cell.Range
                    # Count unique Y positions of chars
                    ys = set()
                    for i in range(1, rng.Characters.Count + 1):
                        try:
                            ch = rng.Characters(i)
                            cy = ch.Information(6)
                            ys.add(round(cy, 1))
                            if ch.Text in ('\r', '\x07', '\n'): continue
                        except: break
                    if len(ys) > max_lines_this_row:
                        max_lines_this_row = len(ys)
                except: pass
            first_char = cell1.Range.Characters(1)
            font = first_char.Font.NameFarEast or first_char.Font.Name
            size = first_char.Font.Size
            entry = {'table': ti, 'row': r, 'n_lines': max_lines_this_row,
                     'font': font, 'size': size, 'row_height': round(row_height, 3)}
            print(entry)
            results.append(entry)
finally:
    wdoc.Close(False)
    word.Quit()

out = r"C:\Users\ryuji\oxi-1\pipeline_data\b35_multiline_rows.json"
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f'Wrote {len(results)} entries to {out}')
