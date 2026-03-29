"""
Ra: TextBox (txbxContent) の内部パディングをCOM計測
- デフォルト内部マージン (inset)
- 明示的設定時の値
- パディングがテキスト位置に与える影響
- AutoFit (shrink/grow) の挙動
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_textbox_default_padding():
    """Default textbox internal margins."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = "Body text."

        # Add textbox
        tb = wdoc.Shapes.AddTextbox(1, 100, 100, 200, 100, wdoc.Range(0, 0))
        tf = tb.TextFrame

        # Measure default margins
        data = {
            "scenario": "textbox_default_padding",
            "margin_left": round(tf.MarginLeft, 4),
            "margin_right": round(tf.MarginRight, 4),
            "margin_top": round(tf.MarginTop, 4),
            "margin_bottom": round(tf.MarginBottom, 4),
            "box_left": round(tb.Left, 4),
            "box_top": round(tb.Top, 4),
            "box_width": round(tb.Width, 4),
            "box_height": round(tb.Height, 4),
        }

        # Add text and measure position
        tf.TextRange.Text = "Padding test"
        tf.TextRange.Font.Name = "Calibri"
        tf.TextRange.Font.Size = 11

        wdoc.Repaginate()

        para = tf.TextRange.Paragraphs(1)
        data["text_x"] = round(para.Range.Information(5), 4)
        data["text_y"] = round(para.Range.Information(6), 4)

        # Expected: text_x = box_left + margin_left
        data["expected_text_x"] = round(data["box_left"] + data["margin_left"], 4)
        data["expected_text_y"] = round(data["box_top"] + data["margin_top"], 4)

        return data
    finally:
        wdoc.Close(False)


def test_textbox_custom_padding():
    """TextBox with various internal margin settings."""
    wdoc = word.Documents.Add()
    try:
        data = {"scenario": "textbox_custom_padding", "tests": []}

        configs = [
            ("default", None, None, None, None),
            ("zero", 0, 0, 0, 0),
            ("large", 20, 20, 15, 15),
            ("asymmetric", 10, 5, 8, 3),
        ]

        for label, l, r, t, b in configs:
            wdoc.Content.Text = "Body."
            tb = wdoc.Shapes.AddTextbox(1, 100, 100, 200, 100, wdoc.Range(0, 0))
            tf = tb.TextFrame

            if l is not None:
                tf.MarginLeft = l
                tf.MarginRight = r
                tf.MarginTop = t
                tf.MarginBottom = b

            tf.TextRange.Text = "Padding test text"
            tf.TextRange.Font.Name = "Calibri"
            tf.TextRange.Font.Size = 11

            wdoc.Repaginate()

            para = tf.TextRange.Paragraphs(1)
            entry = {
                "label": label,
                "margin_left": round(tf.MarginLeft, 4),
                "margin_right": round(tf.MarginRight, 4),
                "margin_top": round(tf.MarginTop, 4),
                "margin_bottom": round(tf.MarginBottom, 4),
                "text_x": round(para.Range.Information(5), 4),
                "text_y": round(para.Range.Information(6), 4),
                "box_left": round(tb.Left, 4),
                "box_top": round(tb.Top, 4),
            }
            entry["text_offset_x"] = round(entry["text_x"] - entry["box_left"], 2)
            entry["text_offset_y"] = round(entry["text_y"] - entry["box_top"], 2)
            data["tests"].append(entry)

            tb.Delete()

        return data
    finally:
        wdoc.Close(False)


def test_textbox_autofit():
    """TextBox AutoFit behavior."""
    wdoc = word.Documents.Add()
    try:
        data = {"scenario": "textbox_autofit", "tests": []}

        # AutoSize modes: 0=None, 1=FitShape, 2=ShrinkText
        for mode, label in [(0, "NoAutoFit"), (1, "ResizeShape"), (2, "ShrinkText")]:
            wdoc.Content.Text = "Body."
            tb = wdoc.Shapes.AddTextbox(1, 100, 200, 150, 50, wdoc.Range(0, 0))
            tf = tb.TextFrame

            try:
                tf.AutoSize = mode
            except:
                pass  # Some modes may not be available

            # Add text that might overflow
            tf.TextRange.Text = "This is a long text that might overflow the textbox boundaries."
            tf.TextRange.Font.Name = "Calibri"
            tf.TextRange.Font.Size = 11

            wdoc.Repaginate()

            entry = {
                "label": label,
                "autosize_mode": mode,
                "box_width": round(tb.Width, 4),
                "box_height": round(tb.Height, 4),
                "para_count": tf.TextRange.Paragraphs.Count,
            }

            # Check line count
            try:
                entry["line_count"] = tf.TextRange.ComputeStatistics(1)
            except:
                entry["line_count"] = -1

            data["tests"].append(entry)
            tb.Delete()

        return data
    finally:
        wdoc.Close(False)


def test_textbox_vs_table_padding():
    """Compare TextBox padding with Table cell padding side by side."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72

        wdoc.Content.Text = ""

        # Create table
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 1, 1)
        tbl.Borders.Enable = True
        cell = tbl.Cell(1, 1)
        cell.Range.Text = "Table cell text"
        cell.Range.Font.Name = "Calibri"
        cell.Range.Font.Size = 11

        # Create textbox
        tb = wdoc.Shapes.AddTextbox(1, 100, 200, 200, 100, wdoc.Range(0, 0))
        tf = tb.TextFrame
        tf.TextRange.Text = "TextBox text"
        tf.TextRange.Font.Name = "Calibri"
        tf.TextRange.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "textbox_vs_table_padding"}

        # Table cell
        cell_para = cell.Range.Paragraphs(1).Range
        data["table_cell"] = {
            "text_x": round(cell_para.Information(5), 4),
            "text_y": round(cell_para.Information(6), 4),
            "left_pad": round(cell.LeftPadding, 4),
            "top_pad": round(cell.TopPadding, 4),
            "right_pad": round(cell.RightPadding, 4),
            "bottom_pad": round(cell.BottomPadding, 4),
        }

        # TextBox
        tb_para = tf.TextRange.Paragraphs(1).Range
        data["textbox"] = {
            "text_x": round(tb_para.Information(5), 4),
            "text_y": round(tb_para.Information(6), 4),
            "margin_left": round(tf.MarginLeft, 4),
            "margin_top": round(tf.MarginTop, 4),
            "margin_right": round(tf.MarginRight, 4),
            "margin_bottom": round(tf.MarginBottom, 4),
            "box_left": round(tb.Left, 4),
            "box_top": round(tb.Top, 4),
        }

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_textbox_default_padding()
    results.append(d1)
    print("=== textbox_default_padding ===")
    print(f"  Margins: L={d1['margin_left']}, R={d1['margin_right']}, T={d1['margin_top']}, B={d1['margin_bottom']}")
    print(f"  Box pos: ({d1['box_left']}, {d1['box_top']})")
    print(f"  Text pos: ({d1['text_x']}, {d1['text_y']})")
    print(f"  Expected: ({d1['expected_text_x']}, {d1['expected_text_y']})")
    print(f"  Diff: x={round(d1['text_x']-d1['expected_text_x'],2)}, y={round(d1['text_y']-d1['expected_text_y'],2)}")

    d2 = test_textbox_custom_padding()
    results.append(d2)
    print(f"\n=== textbox_custom_padding ===")
    for t in d2["tests"]:
        print(f"  {t['label']}: margins(L={t['margin_left']},T={t['margin_top']}), "
              f"text_offset=({t['text_offset_x']}, {t['text_offset_y']})")

    d3 = test_textbox_autofit()
    results.append(d3)
    print(f"\n=== textbox_autofit ===")
    for t in d3["tests"]:
        print(f"  {t['label']}: w={t['box_width']}, h={t['box_height']}, lines={t['line_count']}")

    d4 = test_textbox_vs_table_padding()
    results.append(d4)
    print(f"\n=== textbox_vs_table ===")
    tc = d4["table_cell"]
    tb = d4["textbox"]
    print(f"  Table cell: pad(L={tc['left_pad']}, T={tc['top_pad']}, R={tc['right_pad']}, B={tc['bottom_pad']})")
    print(f"  TextBox:    pad(L={tb['margin_left']}, T={tb['margin_top']}, R={tb['margin_right']}, B={tb['margin_bottom']})")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_textbox_padding.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
