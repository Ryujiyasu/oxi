"""
Ra: テーブル入れ子 + テーブル内TextBox の位置精度
- ネストテーブルの幅計算 (親セル幅 - 親セルパディング)
- ネストテーブルのX/Y位置
- テーブルセル内のTextBox位置
- TextBox内のテーブル
- 3段ネスト
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_nested_table_2level():
    """2-level nested table: outer 2x2, inner 2x2 in cell(1,1)."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        outer = wdoc.Tables.Add(rng, 2, 2)
        outer.Borders.Enable = True

        # Set explicit widths
        outer.Columns(1).Width = 250
        outer.Columns(2).Width = 200

        for r in range(1, 3):
            for c in range(1, 3):
                outer.Cell(r, c).Range.Text = f"Outer R{r}C{c}"
                outer.Cell(r, c).Range.Font.Name = "Calibri"
                outer.Cell(r, c).Range.Font.Size = 10

        # Insert nested table in outer(1,1)
        cell11 = outer.Cell(1, 1)
        # Move to end of cell content
        inner_rng = wdoc.Range(cell11.Range.End - 1, cell11.Range.End - 1)
        inner = wdoc.Tables.Add(inner_rng, 2, 2)
        inner.Borders.Enable = True

        for r in range(1, 3):
            for c in range(1, 3):
                inner.Cell(r, c).Range.Text = f"Inner {r}{c}"
                inner.Cell(r, c).Range.Font.Name = "Calibri"
                inner.Cell(r, c).Range.Font.Size = 9

        wdoc.Repaginate()

        data = {"scenario": "nested_2level"}

        # Outer table
        data["outer"] = {
            "col1_width": round(outer.Columns(1).Width, 4),
            "col2_width": round(outer.Columns(2).Width, 4),
            "cell11_pad_l": round(outer.Cell(1,1).LeftPadding, 4),
            "cell11_pad_r": round(outer.Cell(1,1).RightPadding, 4),
        }

        # Measure all outer cells
        data["outer_cells"] = []
        for r in range(1, 3):
            for c in range(1, 3):
                cell = outer.Cell(r, c)
                p1 = cell.Range.Paragraphs(1).Range
                data["outer_cells"].append({
                    "row": r, "col": c,
                    "text_x": round(p1.Information(5), 4),
                    "text_y": round(p1.Information(6), 4),
                    "width": round(cell.Width, 4),
                })

        # Inner table
        data["inner"] = {"cells": []}
        for r in range(1, 3):
            for c in range(1, 3):
                cell = inner.Cell(r, c)
                p1 = cell.Range.Paragraphs(1).Range
                data["inner"]["cells"].append({
                    "row": r, "col": c,
                    "text_x": round(p1.Information(5), 4),
                    "text_y": round(p1.Information(6), 4),
                    "width": round(cell.Width, 4),
                })

        # Key question: inner table width vs outer cell content area
        inner_total_w = sum(inner.Columns(c).Width for c in range(1, 3))
        outer_cell_content_w = outer.Columns(1).Width - outer.Cell(1,1).LeftPadding - outer.Cell(1,1).RightPadding
        data["inner_total_width"] = round(inner_total_w, 4)
        data["outer_cell_content_width"] = round(outer_cell_content_w, 4)

        return data
    finally:
        wdoc.Close(False)


def test_textbox_in_table():
    """TextBox positioned inside a table cell."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 2, 2)
        tbl.Borders.Enable = True

        for r in range(1, 3):
            for c in range(1, 3):
                tbl.Cell(r, c).Range.Text = f"Cell R{r}C{c}"
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                tbl.Cell(r, c).Range.Font.Size = 11

        # Add textbox anchored in cell(1,1)
        cell11 = tbl.Cell(1, 1)
        tb = wdoc.Shapes.AddTextbox(1, 10, 10, 100, 40, cell11.Range)
        tf = tb.TextFrame
        tf.TextRange.Text = "TB in cell"
        tf.TextRange.Font.Name = "Calibri"
        tf.TextRange.Font.Size = 9

        wdoc.Repaginate()

        data = {"scenario": "textbox_in_table"}
        data["textbox"] = {
            "left": round(tb.Left, 4),
            "top": round(tb.Top, 4),
            "width": round(tb.Width, 4),
            "height": round(tb.Height, 4),
            "margin_left": round(tf.MarginLeft, 4),
            "margin_top": round(tf.MarginTop, 4),
        }

        # Cell position
        cell_para = cell11.Range.Paragraphs(1).Range
        data["cell11"] = {
            "text_x": round(cell_para.Information(5), 4),
            "text_y": round(cell_para.Information(6), 4),
            "width": round(cell11.Width, 4),
        }

        # TB text position
        tb_para = tf.TextRange.Paragraphs(1).Range
        data["textbox"]["text_x"] = round(tb_para.Information(5), 4)
        data["textbox"]["text_y"] = round(tb_para.Information(6), 4)

        return data
    finally:
        wdoc.Close(False)


def test_table_in_textbox():
    """Table inside a TextBox."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72

        wdoc.Content.Text = "Body text."

        # Create textbox
        tb = wdoc.Shapes.AddTextbox(1, 72, 200, 300, 150, wdoc.Range(0, 0))
        tf = tb.TextFrame

        # Add table inside textbox
        tf.TextRange.Text = ""
        inner_rng = tf.TextRange
        tbl = wdoc.Tables.Add(inner_rng, 2, 2)
        tbl.Borders.Enable = True

        for r in range(1, 3):
            for c in range(1, 3):
                tbl.Cell(r, c).Range.Text = f"TB-T R{r}C{c}"
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                tbl.Cell(r, c).Range.Font.Size = 9

        wdoc.Repaginate()

        data = {"scenario": "table_in_textbox"}
        data["textbox"] = {
            "left": round(tb.Left, 4),
            "top": round(tb.Top, 4),
            "width": round(tb.Width, 4),
            "margin_left": round(tf.MarginLeft, 4),
            "margin_top": round(tf.MarginTop, 4),
        }

        # Table cells
        data["table_cells"] = []
        for r in range(1, 3):
            for c in range(1, 3):
                cell = tbl.Cell(r, c)
                p = cell.Range.Paragraphs(1).Range
                data["table_cells"].append({
                    "row": r, "col": c,
                    "text_x": round(p.Information(5), 4),
                    "text_y": round(p.Information(6), 4),
                    "width": round(cell.Width, 4),
                })

        # Expected: table should be within textbox bounds
        tb_content_x = tb.Left + tf.MarginLeft
        tb_content_w = tb.Width - tf.MarginLeft - tf.MarginRight
        data["expected_content_x"] = round(tb_content_x, 4)
        data["expected_content_width"] = round(tb_content_w, 4)

        return data
    finally:
        wdoc.Close(False)


def test_3level_nest():
    """3-level nesting: outer table > inner table > innermost table."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        outer = wdoc.Tables.Add(rng, 1, 1)
        outer.Borders.Enable = True
        outer.Columns(1).Width = 400

        outer.Cell(1,1).Range.Text = "Outer"
        outer.Cell(1,1).Range.Font.Name = "Calibri"
        outer.Cell(1,1).Range.Font.Size = 10

        # Level 2
        mid_rng = wdoc.Range(outer.Cell(1,1).Range.End - 1, outer.Cell(1,1).Range.End - 1)
        mid = wdoc.Tables.Add(mid_rng, 1, 1)
        mid.Borders.Enable = True
        mid.Cell(1,1).Range.Text = "Mid"
        mid.Cell(1,1).Range.Font.Size = 9

        # Level 3
        inner_rng = wdoc.Range(mid.Cell(1,1).Range.End - 1, mid.Cell(1,1).Range.End - 1)
        inner = wdoc.Tables.Add(inner_rng, 1, 1)
        inner.Borders.Enable = True
        inner.Cell(1,1).Range.Text = "Inner"
        inner.Cell(1,1).Range.Font.Size = 8

        wdoc.Repaginate()

        data = {"scenario": "3level_nest", "levels": []}

        for label, tbl in [("outer", outer), ("mid", mid), ("inner", inner)]:
            cell = tbl.Cell(1, 1)
            p = cell.Range.Paragraphs(1).Range
            data["levels"].append({
                "label": label,
                "cell_width": round(cell.Width, 4),
                "text_x": round(p.Information(5), 4),
                "text_y": round(p.Information(6), 4),
                "left_pad": round(cell.LeftPadding, 4),
            })

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_nested_table_2level()
    results.append(d1)
    ml = 72
    print("=== nested_2level ===")
    print(f"  Outer col widths: {d1['outer']['col1_width']}, {d1['outer']['col2_width']}")
    print(f"  Outer cell padding: L={d1['outer']['cell11_pad_l']}, R={d1['outer']['cell11_pad_r']}")
    for c in d1["outer_cells"]:
        print(f"  Outer R{c['row']}C{c['col']}: x={c['text_x']}, y={c['text_y']}, w={c['width']}")
    for c in d1["inner"]["cells"]:
        print(f"  Inner R{c['row']}C{c['col']}: x={c['text_x']}, y={c['text_y']}, w={c['width']}")
    print(f"  Inner total width: {d1['inner_total_width']}pt")
    print(f"  Outer cell content area: {d1['outer_cell_content_width']}pt")

    d2 = test_textbox_in_table()
    results.append(d2)
    print(f"\n=== textbox_in_table ===")
    print(f"  Cell(1,1): x={d2['cell11']['text_x']}, y={d2['cell11']['text_y']}, w={d2['cell11']['width']}")
    print(f"  TextBox: left={d2['textbox']['left']}, top={d2['textbox']['top']}, "
          f"w={d2['textbox']['width']}, h={d2['textbox']['height']}")
    print(f"  TB text: x={d2['textbox']['text_x']}, y={d2['textbox']['text_y']}")

    d3 = test_table_in_textbox()
    results.append(d3)
    print(f"\n=== table_in_textbox ===")
    print(f"  TextBox: left={d3['textbox']['left']}, top={d3['textbox']['top']}, w={d3['textbox']['width']}")
    print(f"  Expected content x: {d3['expected_content_x']}, w: {d3['expected_content_width']}")
    for c in d3["table_cells"]:
        print(f"  Table R{c['row']}C{c['col']}: x={c['text_x']}, y={c['text_y']}, w={c['width']}")

    d4 = test_3level_nest()
    results.append(d4)
    print(f"\n=== 3level_nest ===")
    for lv in d4["levels"]:
        print(f"  {lv['label']}: w={lv['cell_width']}, x={lv['text_x']}, y={lv['text_y']}, pad={lv['left_pad']}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_nested_table_textbox.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
