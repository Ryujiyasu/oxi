"""
Ra: マルチカラムレイアウトの位置計算をCOM計測で確定
- 2カラム/3カラムのX位置
- カラム幅とギャップ
- カラム間のテキスト折り返し
- 均等幅 vs カスタム幅
- カラムブレーク（次カラムへの移動）
"""
import win32com.client, json, os, tempfile

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def create_and_measure(scenario, num_cols=2, equal_width=True, col_widths=None,
                       col_spacing=None, num_paragraphs=6):
    """Create multi-column doc via COM and measure paragraph positions."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72  # 1 inch
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72

        # Set columns
        tc = ps.TextColumns
        tc.SetCount(num_cols)

        if equal_width:
            tc.EvenlySpaced = True
            if col_spacing is not None:
                tc.Spacing = col_spacing
        else:
            tc.EvenlySpaced = False
            if col_widths:
                for i, (w, sp) in enumerate(col_widths):
                    col = tc.Item(i + 1)
                    col.Width = w
                    if i < len(col_widths) - 1:
                        col.SpaceAfter = sp

        # Add paragraphs
        wdoc.Content.Text = ""
        for i in range(num_paragraphs):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            rng = wdoc.Paragraphs(i + 1 if i < wdoc.Paragraphs.Count else wdoc.Paragraphs.Count).Range
            rng.Text = f"Paragraph {i+1}: Lorem ipsum dolor sit amet consectetur."
            rng.Font.Name = "Calibri"
            rng.Font.Size = 11
            rng.ParagraphFormat.SpaceBefore = 0
            rng.ParagraphFormat.SpaceAfter = 0

        # Measure
        data = {
            "scenario": scenario,
            "num_cols": num_cols,
            "page_width": round(ps.PageWidth, 4),
            "margin_left": round(ps.LeftMargin, 4),
            "margin_right": round(ps.RightMargin, 4),
            "text_width": round(ps.PageWidth - ps.LeftMargin - ps.RightMargin, 4),
            "columns": [],
            "paragraphs": []
        }

        # Get column info
        for i in range(1, tc.Count + 1):
            col = tc.Item(i)
            col_data = {"index": i, "width": round(col.Width, 4)}
            if i < tc.Count:
                col_data["space_after"] = round(col.SpaceAfter, 4)
            data["columns"].append(col_data)

        # Measure paragraph positions
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            x = rng.Information(5)  # wdHorizontalPositionRelativeToPage
            y = rng.Information(6)  # wdVerticalPositionRelativeToPage
            col_num = rng.Information(16)  # wdFirstCharacterColumnNumber
            page_num = rng.Information(3)  # wdActiveEndPageNumber
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(x, 4),
                "y_pt": round(y, 4),
                "column": col_num,
                "page": page_num,
                "text": rng.Text.strip()[:50]
            })

        return data
    finally:
        wdoc.Close(False)


def create_column_break_test():
    """Test column break behavior."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72

        ps.TextColumns.SetCount(2)
        ps.TextColumns.EvenlySpaced = True

        wdoc.Content.Text = ""

        # Add paragraphs with a column break after P2
        for i in range(4):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Para {i+1}: text content here."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

            if i == 1:  # Insert column break after P2
                rng = wdoc.Range(para.Range.End - 1, para.Range.End - 1)
                rng.InsertBreak(8)  # wdColumnBreak = 8

        data = {"scenario": "column_break", "paragraphs": []}
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "column": rng.Information(16),
                "text": rng.Text.strip()[:50]
            })

        return data
    finally:
        wdoc.Close(False)


try:
    # Test 1: 2 columns, equal width
    data = create_and_measure("2col_equal", num_cols=2, equal_width=True, num_paragraphs=8)
    results.append(data)
    print(f"\n=== 2col_equal ===")
    print(f"  text_width={data['text_width']}pt")
    for c in data["columns"]:
        print(f"  Column {c['index']}: width={c['width']}pt" +
              (f", space_after={c.get('space_after', '-')}" if 'space_after' in c else ""))
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, col={p['column']}, pg={p['page']}")

    # Test 2: 3 columns, equal width
    data = create_and_measure("3col_equal", num_cols=3, equal_width=True, num_paragraphs=12)
    results.append(data)
    print(f"\n=== 3col_equal ===")
    for c in data["columns"]:
        print(f"  Column {c['index']}: width={c['width']}pt" +
              (f", space_after={c.get('space_after', '-')}" if 'space_after' in c else ""))
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, col={p['column']}, pg={p['page']}")

    # Test 3: 2 columns with custom spacing
    data = create_and_measure("2col_wide_gap", num_cols=2, equal_width=True,
                               col_spacing=36, num_paragraphs=6)
    results.append(data)
    print(f"\n=== 2col_wide_gap (spacing=36pt) ===")
    for c in data["columns"]:
        print(f"  Column {c['index']}: width={c['width']}pt" +
              (f", space_after={c.get('space_after', '-')}" if 'space_after' in c else ""))
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, col={p['column']}")

    # Test 4: 2 columns, unequal width
    data = create_and_measure("2col_unequal", num_cols=2, equal_width=False,
                               col_widths=[(200, 36), (215.3, 0)], num_paragraphs=6)
    results.append(data)
    print(f"\n=== 2col_unequal ===")
    for c in data["columns"]:
        print(f"  Column {c['index']}: width={c['width']}pt" +
              (f", space_after={c.get('space_after', '-')}" if 'space_after' in c else ""))
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, col={p['column']}")

    # Test 5: Column break
    data = create_column_break_test()
    results.append(data)
    print(f"\n=== column_break ===")
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, col={p['column']}  [{p['text']}]")

finally:
    word.Quit()

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_column_layout.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Analysis
print("\n\n========== COLUMN LAYOUT ANALYSIS ==========")
for data in results:
    if "columns" not in data:
        continue
    sc = data["scenario"]
    ml = data["margin_left"]
    tw = data["text_width"]
    print(f"\n{sc}:")
    print(f"  Page: margin_left={ml}pt, text_width={tw}pt")

    # Determine column X starts from paragraph positions
    col_x = {}
    for p in data["paragraphs"]:
        col = p["column"]
        if col not in col_x:
            col_x[col] = p["x_pt"]

    print(f"  Column X positions: {dict(sorted(col_x.items()))}")
    print(f"  Column X (margin-rel): {dict((k, round(v-ml, 2)) for k, v in sorted(col_x.items()))}")

    # Verify: col1 starts at margin_left, col2 at margin_left + col1_width + spacing
    if len(data["columns"]) >= 2:
        c1 = data["columns"][0]
        expected_col2_x = ml + c1["width"] + c1.get("space_after", 0)
        actual_col2_x = col_x.get(2, "N/A")
        print(f"  Expected col2 x: margin({ml}) + width({c1['width']}) + space({c1.get('space_after', 0)}) = {round(expected_col2_x, 2)}")
        print(f"  Actual col2 x: {actual_col2_x}")
