"""
Ra: テーブル行高さ(trHeight)の計算精度
- trHeight atLeast vs exact
- autoHeight (trHeight指定なし) の場合のセル高さ
- セル内コンテンツが行高さに与える影響
- テーブル行のY位置計算
"""
import win32com.client, json, os

word = win32com.client.DispatchEx('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_row_height_auto():
    """Auto row height: height determined by content."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72; ps.TopMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 4, 2)
        tbl.Borders.Enable = True

        # Row 1: single line
        tbl.Cell(1,1).Range.Text = "Single line"
        tbl.Cell(1,2).Range.Text = "Single"

        # Row 2: two lines in col 1
        tbl.Cell(2,1).Range.Text = "Line 1\rLine 2"
        tbl.Cell(2,2).Range.Text = "One line"

        # Row 3: three lines
        tbl.Cell(3,1).Range.Text = "L1\rL2\rL3"
        tbl.Cell(3,2).Range.Text = "Short"

        # Row 4: different font size
        tbl.Cell(4,1).Range.Text = "Big font"
        tbl.Cell(4,1).Range.Font.Size = 18
        tbl.Cell(4,2).Range.Text = "Normal"

        # Set all to Calibri
        for r in range(1, 5):
            for c in range(1, 3):
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                if r != 4 or c != 1:
                    tbl.Cell(r, c).Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "row_height_auto", "rows": []}
        for r in range(1, 5):
            row = tbl.Rows(r)
            c1 = tbl.Cell(r, 1)
            c2 = tbl.Cell(r, 2)
            y1 = c1.Range.Paragraphs(1).Range.Information(6)
            y2 = c2.Range.Paragraphs(1).Range.Information(6)
            data["rows"].append({
                "row": r,
                "height": round(row.Height, 4),
                "height_rule": row.HeightRule,  # 0=auto, 1=atLeast, 2=exact
                "cell1_y": round(y1, 4),
                "cell2_y": round(y2, 4),
            })

        # Compute row Y gaps
        for i in range(1, len(data["rows"])):
            data["rows"][i]["gap_from_prev"] = round(
                data["rows"][i]["cell1_y"] - data["rows"][i-1]["cell1_y"], 4)

        return data
    finally:
        wdoc.Close(False)


def test_row_height_exact():
    """Exact row height."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72; ps.TopMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 4, 2)
        tbl.Borders.Enable = True

        heights = [20, 30, 40, 50]  # pt
        for r in range(1, 5):
            tbl.Rows(r).Height = heights[r-1]
            tbl.Rows(r).HeightRule = 2  # wdRowHeightExactly
            for c in range(1, 3):
                tbl.Cell(r, c).Range.Text = f"R{r} h={heights[r-1]}"
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                tbl.Cell(r, c).Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "row_height_exact", "rows": []}
        for r in range(1, 5):
            row = tbl.Rows(r)
            c1 = tbl.Cell(r, 1)
            y1 = c1.Range.Paragraphs(1).Range.Information(6)
            data["rows"].append({
                "row": r,
                "set_height": heights[r-1],
                "actual_height": round(row.Height, 4),
                "cell1_y": round(y1, 4),
            })

        for i in range(1, len(data["rows"])):
            gap = data["rows"][i]["cell1_y"] - data["rows"][i-1]["cell1_y"]
            data["rows"][i]["y_gap"] = round(gap, 4)

        return data
    finally:
        wdoc.Close(False)


def test_row_height_atleast():
    """AtLeast row height with varying content."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72; ps.TopMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 3, 2)
        tbl.Borders.Enable = True

        # All rows: atLeast=25pt
        for r in range(1, 4):
            tbl.Rows(r).Height = 25
            tbl.Rows(r).HeightRule = 1  # wdRowHeightAtLeast

        # Row 1: small content (should use 25pt)
        tbl.Cell(1,1).Range.Text = "Small"
        tbl.Cell(1,2).Range.Text = "S"

        # Row 2: content taller than 25pt
        tbl.Cell(2,1).Range.Text = "Line1\rLine2\rLine3"
        tbl.Cell(2,2).Range.Text = "S"

        # Row 3: exact match
        tbl.Cell(3,1).Range.Text = "Fit"
        tbl.Cell(3,2).Range.Text = "F"

        for r in range(1, 4):
            for c in range(1, 3):
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                tbl.Cell(r, c).Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "row_height_atleast", "rows": []}
        for r in range(1, 4):
            row = tbl.Rows(r)
            y = tbl.Cell(r, 1).Range.Paragraphs(1).Range.Information(6)
            data["rows"].append({
                "row": r,
                "set_height": 25,
                "actual_height": round(row.Height, 4),
                "cell1_y": round(y, 4),
            })

        for i in range(1, len(data["rows"])):
            data["rows"][i]["y_gap"] = round(
                data["rows"][i]["cell1_y"] - data["rows"][i-1]["cell1_y"], 4)

        return data
    finally:
        wdoc.Close(False)


def test_table_row_y_positions():
    """Precise Y positions of table rows (for SSIM matching)."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72; ps.TopMargin = 72

        # Add a paragraph before table
        wdoc.Content.Text = "Before table paragraph."
        wdoc.Paragraphs(1).Range.Font.Name = "Calibri"
        wdoc.Paragraphs(1).Range.Font.Size = 11
        wdoc.Paragraphs(1).Format.SpaceAfter = 0

        # Add table
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertParagraphAfter()
        rng2 = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        tbl = wdoc.Tables.Add(rng2, 5, 3)
        tbl.Borders.Enable = True

        for r in range(1, 6):
            for c in range(1, 4):
                tbl.Cell(r, c).Range.Text = f"R{r}C{c} text"
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                tbl.Cell(r, c).Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "table_row_y", "margin_top": round(ps.TopMargin, 4)}

        # Before-table paragraph
        p_before = wdoc.Paragraphs(1).Range
        data["before_para_y"] = round(p_before.Information(6), 4)

        data["rows"] = []
        for r in range(1, 6):
            y = tbl.Cell(r, 1).Range.Paragraphs(1).Range.Information(6)
            data["rows"].append({
                "row": r,
                "text_y": round(y, 4),
                "height": round(tbl.Rows(r).Height, 4),
            })

        for i in range(1, len(data["rows"])):
            data["rows"][i]["y_gap"] = round(
                data["rows"][i]["text_y"] - data["rows"][i-1]["text_y"], 4)

        # First row gap from before_para
        data["rows"][0]["gap_from_para"] = round(
            data["rows"][0]["text_y"] - data["before_para_y"], 4)

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_row_height_auto()
    results.append(d1)
    print("=== row_height_auto ===")
    for r in d1["rows"]:
        gap = r.get('gap_from_prev', '-')
        print(f"  R{r['row']}: h={r['height']}, rule={r['height_rule']}, "
              f"y1={r['cell1_y']}, gap={gap}")

    d2 = test_row_height_exact()
    results.append(d2)
    print(f"\n=== row_height_exact ===")
    for r in d2["rows"]:
        gap = r.get('y_gap', '-')
        print(f"  R{r['row']}: set={r['set_height']}, actual={r['actual_height']}, "
              f"y={r['cell1_y']}, y_gap={gap}")

    d3 = test_row_height_atleast()
    results.append(d3)
    print(f"\n=== row_height_atleast ===")
    for r in d3["rows"]:
        gap = r.get('y_gap', '-')
        print(f"  R{r['row']}: set=25, actual={r['actual_height']}, y={r['cell1_y']}, gap={gap}")

    d4 = test_table_row_y_positions()
    results.append(d4)
    print(f"\n=== table_row_y ===")
    print(f"  Before para y: {d4['before_para_y']}")
    for r in d4["rows"]:
        gap = r.get('y_gap', r.get('gap_from_para', '-'))
        print(f"  R{r['row']}: y={r['text_y']}, h={r['height']}, gap={gap}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_trheight.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
