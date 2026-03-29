"""
Ra: テーブルセル内のexact spacing + character spacing 累積挙動
- テーブルセル内でexact line spacingは通常と同じか？
- cs累積がセル幅を超える場合の折り返し
- セル内パディングとの相互作用
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_table_exact():
    """Table cell with exact line spacing."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72
        sec.PageSetup.RightMargin = 72

        wdoc.Content.Text = ""

        # Create 3-row table
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 3, 2)

        # Set table borders
        tbl.Borders.Enable = True

        for ri in range(1, 4):
            for ci in range(1, 3):
                cell = tbl.Cell(ri, ci)
                cell.Range.Text = f"R{ri}C{ci} exact line spacing test text"
                cell.Range.Font.Name = "Calibri"
                cell.Range.Font.Size = 9

                # Set exact line spacing
                for pi in range(1, cell.Range.Paragraphs.Count + 1):
                    para = cell.Range.Paragraphs(pi)
                    para.Format.LineSpacingRule = 2  # wdLineSpaceExactly
                    para.Format.LineSpacing = 10  # 10pt exact
                    para.Format.SpaceBefore = 0
                    para.Format.SpaceAfter = 0

        wdoc.Repaginate()

        data = {"scenario": "table_exact_ls", "rows": []}
        for ri in range(1, 4):
            row_data = {"row": ri, "cells": []}
            for ci in range(1, 3):
                cell = tbl.Cell(ri, ci)
                prng = cell.Range.Paragraphs(1).Range
                cell_data = {
                    "col": ci,
                    "y_pt": round(prng.Information(6), 4),
                    "x_pt": round(prng.Information(5), 4),
                    "line_spacing": round(cell.Range.Paragraphs(1).Format.LineSpacing, 4),
                    "ls_rule": cell.Range.Paragraphs(1).Format.LineSpacingRule,
                    "row_height": round(tbl.Rows(ri).Height, 4),
                }
                row_data["cells"].append(cell_data)
            data["rows"].append(row_data)

        return data
    finally:
        wdoc.Close(False)


def test_table_cs_accumulation():
    """Character spacing accumulation in table cells."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 4, 1)

        # Set narrow column width
        tbl.Columns(1).Width = 150  # ~2 inches

        cs_values = [0, -6, -12, 6]  # twips
        for ri in range(1, 5):
            cell = tbl.Cell(ri, 1)
            cell.Range.Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            cell.Range.Font.Name = "Calibri"
            cell.Range.Font.Size = 9

            # Set character spacing
            cs_tw = cs_values[ri - 1]
            if cs_tw != 0:
                cell.Range.Font.Spacing = cs_tw / 20.0  # convert twips to pt

            cell.Range.Paragraphs(1).Format.SpaceBefore = 0
            cell.Range.Paragraphs(1).Format.SpaceAfter = 0

        wdoc.Repaginate()

        data = {"scenario": "table_cs_accumulation", "rows": []}
        for ri in range(1, 5):
            cell = tbl.Cell(ri, 1)
            para = cell.Range.Paragraphs(1)
            nlines = para.Range.ComputeStatistics(1)  # wdStatisticLines

            # Measure char positions
            chars = []
            prng = para.Range
            for ci_idx in range(prng.Start, min(prng.End, prng.Start + 20)):
                cr = wdoc.Range(ci_idx, ci_idx + 1)
                ch = cr.Text
                if ord(ch) not in (13, 7):
                    chars.append({"ch": ch, "x": round(cr.Information(5), 4)})

            data["rows"].append({
                "row": ri,
                "cs_twips": cs_values[ri - 1],
                "cs_pt": round(cell.Range.Font.Spacing, 4),
                "line_count": nlines,
                "first_chars": chars[:10],
            })

        return data
    finally:
        wdoc.Close(False)


def test_table_cell_padding():
    """Measure actual cell padding values."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 2, 2)
        tbl.Borders.Enable = True

        for ri in range(1, 3):
            for ci in range(1, 3):
                cell = tbl.Cell(ri, ci)
                cell.Range.Text = f"R{ri}C{ci}"
                cell.Range.Font.Name = "Calibri"
                cell.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "cell_padding", "cells": []}
        for ri in range(1, 3):
            for ci in range(1, 3):
                cell = tbl.Cell(ri, ci)
                prng = cell.Range.Paragraphs(1).Range
                data["cells"].append({
                    "row": ri, "col": ci,
                    "text_x": round(prng.Information(5), 4),
                    "text_y": round(prng.Information(6), 4),
                    "top_padding": round(cell.TopPadding, 4),
                    "bottom_padding": round(cell.BottomPadding, 4),
                    "left_padding": round(cell.LeftPadding, 4),
                    "right_padding": round(cell.RightPadding, 4),
                })

        # Table position
        data["table_x"] = round(tbl.Rows(1).Cells(1).Range.Information(5), 4)

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_table_exact()
    results.append(d1)
    print("=== table_exact_ls ===")
    for row in d1["rows"]:
        for cell in row["cells"]:
            print(f"  R{row['row']}C{cell['col']}: y={cell['y_pt']}, x={cell['x_pt']}, "
                  f"ls={cell['line_spacing']}(rule={cell['ls_rule']}), row_h={cell['row_height']}")

    d2 = test_table_cs_accumulation()
    results.append(d2)
    print(f"\n=== table_cs_accumulation (col_width=150pt) ===")
    for row in d2["rows"]:
        char_gaps = []
        for i in range(1, len(row["first_chars"])):
            gap = row["first_chars"][i]["x"] - row["first_chars"][i-1]["x"]
            char_gaps.append(round(gap, 2))
        print(f"  R{row['row']}: cs={row['cs_twips']}tw({row['cs_pt']}pt), "
              f"lines={row['line_count']}, char_gaps={char_gaps[:5]}")

    d3 = test_table_cell_padding()
    results.append(d3)
    print(f"\n=== cell_padding ===")
    for cell in d3["cells"]:
        print(f"  R{cell['row']}C{cell['col']}: x={cell['text_x']}, y={cell['text_y']}, "
              f"pad(T={cell['top_padding']}, B={cell['bottom_padding']}, "
              f"L={cell['left_padding']}, R={cell['right_padding']})")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_table_exact_cs.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
