"""
Ra: マルチカラムレイアウトv2 — 十分なテキストでカラムオーバーフローを確認
- Repaginate() で確実にレイアウト計算
- 長いテキストでカラム間フロー確認
- カラム幅・ギャップの正確な位置
"""
import win32com.client, json, os, tempfile

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []

LONG_TEXT = ("Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
             "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
             "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. ")


def create_col_test(scenario, num_cols=2, col_spacing_pt=None, num_paras=30,
                    text=None, equal=True, col1_width=None, col2_width=None):
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72

    tc = ps.TextColumns
    tc.SetCount(num_cols)

    if equal:
        tc.EvenlySpaced = True
        if col_spacing_pt is not None:
            tc.Spacing = col_spacing_pt
    else:
        tc.EvenlySpaced = False
        if col1_width and col2_width:
            tc.Item(1).Width = col1_width
            tc.Item(1).SpaceAfter = col_spacing_pt or 21.25
            tc.Item(2).Width = col2_width

    wdoc.Content.Text = ""

    # Add many paragraphs to fill multiple columns
    para_text = text or LONG_TEXT
    for i in range(num_paras):
        if i > 0:
            rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
            rng.InsertParagraphAfter()
        para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
        para.Range.Text = f"[P{i+1}] {para_text}"
        para.Range.Font.Name = "Calibri"
        para.Range.Font.Size = 11
        para.Format.SpaceBefore = 0
        para.Format.SpaceAfter = 0

    # Force repagination
    wdoc.Repaginate()

    # Measure
    data = {
        "scenario": scenario,
        "num_cols": num_cols,
        "page_width": round(ps.PageWidth, 4),
        "margin_left": round(ps.LeftMargin, 4),
        "text_width": round(ps.PageWidth - ps.LeftMargin - ps.RightMargin, 4),
        "columns": [],
        "paragraphs": []
    }

    for i in range(1, tc.Count + 1):
        col = tc.Item(i)
        cd = {"index": i, "width": round(col.Width, 4)}
        if i < tc.Count:
            cd["space_after"] = round(col.SpaceAfter, 4)
        data["columns"].append(cd)

    for i in range(1, wdoc.Paragraphs.Count + 1):
        para = wdoc.Paragraphs(i)
        rng = para.Range
        x = rng.Information(5)
        y = rng.Information(6)
        pg = rng.Information(3)
        data["paragraphs"].append({
            "index": i,
            "x_pt": round(x, 4),
            "y_pt": round(y, 4),
            "page": pg,
        })

    wdoc.Close(False)
    return data


def create_col_break_multi(num_cols=2):
    """Column break tests with explicit breaks."""
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72

    tc = ps.TextColumns
    tc.SetCount(num_cols)
    tc.EvenlySpaced = True

    wdoc.Content.Text = ""

    # P1 in col1
    wdoc.Content.InsertAfter("Col1-Para1 short text\r")
    wdoc.Content.InsertAfter("Col1-Para2 short text\r")

    # Column break
    rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
    rng.InsertBreak(8)  # wdColumnBreak

    wdoc.Content.InsertAfter("Col2-Para1 short text\r")
    wdoc.Content.InsertAfter("Col2-Para2 short text\r")

    if num_cols == 3:
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertBreak(8)
        wdoc.Content.InsertAfter("Col3-Para1 short text\r")

    wdoc.Repaginate()

    data = {"scenario": f"{num_cols}col_breaks", "paragraphs": []}
    for i in range(1, wdoc.Paragraphs.Count + 1):
        para = wdoc.Paragraphs(i)
        rng = para.Range
        data["paragraphs"].append({
            "index": i,
            "x_pt": round(rng.Information(5), 4),
            "y_pt": round(rng.Information(6), 4),
            "page": rng.Information(3),
            "text": rng.Text.strip()[:40]
        })

    # Get column info
    data["columns"] = []
    for i in range(1, tc.Count + 1):
        col = tc.Item(i)
        cd = {"index": i, "width": round(col.Width, 4)}
        if i < tc.Count:
            cd["space_after"] = round(col.SpaceAfter, 4)
        data["columns"].append(cd)

    data["margin_left"] = round(ps.LeftMargin, 4)

    wdoc.Close(False)
    return data


try:
    # Test 1: 2 columns, many paragraphs (should overflow to col2)
    data = create_col_test("2col_overflow", num_cols=2, num_paras=40)
    results.append(data)

    # Analyze column positions
    x_values = sorted(set(p["x_pt"] for p in data["paragraphs"]))
    print(f"=== 2col_overflow ===")
    print(f"  Unique X positions: {x_values}")
    print(f"  Columns: {[(c['width'], c.get('space_after', '-')) for c in data['columns']]}")
    for xv in x_values:
        paras_at_x = [p for p in data["paragraphs"] if p["x_pt"] == xv]
        print(f"  x={xv}: {len(paras_at_x)} paragraphs (P{paras_at_x[0]['index']}-P{paras_at_x[-1]['index']})")

    # Test 2: 3 columns overflow
    data = create_col_test("3col_overflow", num_cols=3, num_paras=60)
    results.append(data)
    x_values = sorted(set(p["x_pt"] for p in data["paragraphs"]))
    print(f"\n=== 3col_overflow ===")
    print(f"  Unique X positions: {x_values}")
    print(f"  Columns: {[(c['width'], c.get('space_after', '-')) for c in data['columns']]}")
    for xv in x_values:
        paras_at_x = [p for p in data["paragraphs"] if p["x_pt"] == xv]
        print(f"  x={xv}: {len(paras_at_x)} paragraphs")

    # Test 3: 2 columns with wide gap
    data = create_col_test("2col_gap36", num_cols=2, col_spacing_pt=36, num_paras=40)
    results.append(data)
    x_values = sorted(set(p["x_pt"] for p in data["paragraphs"]))
    print(f"\n=== 2col_gap36 ===")
    print(f"  Unique X positions: {x_values}")
    print(f"  Columns: {[(c['width'], c.get('space_after', '-')) for c in data['columns']]}")

    # Test 4: 2 columns unequal
    data = create_col_test("2col_unequal", num_cols=2, equal=False,
                            col1_width=150, col2_width=265.3, col_spacing_pt=36, num_paras=40)
    results.append(data)
    x_values = sorted(set(p["x_pt"] for p in data["paragraphs"]))
    print(f"\n=== 2col_unequal ===")
    print(f"  Unique X positions: {x_values}")
    print(f"  Columns: {[(c['width'], c.get('space_after', '-')) for c in data['columns']]}")

    # Test 5: Explicit column breaks - 2col
    data = create_col_break_multi(2)
    results.append(data)
    print(f"\n=== 2col_breaks ===")
    ml = data["margin_left"]
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']} (margin+{round(p['x_pt']-ml, 2)}), y={p['y_pt']}  [{p['text']}]")
    print(f"  Columns: {[(c['width'], c.get('space_after', '-')) for c in data['columns']]}")

    # Test 6: Explicit column breaks - 3col
    data = create_col_break_multi(3)
    results.append(data)
    print(f"\n=== 3col_breaks ===")
    ml = data["margin_left"]
    for p in data["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']} (margin+{round(p['x_pt']-ml, 2)}), y={p['y_pt']}  [{p['text']}]")
    print(f"  Columns: {[(c['width'], c.get('space_after', '-')) for c in data['columns']]}")

finally:
    word.Quit()

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_column_layout_v2.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Final analysis
print("\n\n========== COLUMN POSITION FORMULA ==========")
for data in results:
    if "columns" not in data or len(data.get("columns", [])) < 2:
        continue
    sc = data["scenario"]
    ml = data["margin_left"]
    cols = data["columns"]
    x_vals = sorted(set(p["x_pt"] for p in data["paragraphs"]))

    print(f"\n{sc}:")
    print(f"  margin_left = {ml}")

    expected_x = ml
    for i, col in enumerate(cols):
        print(f"  Col{i+1}: expected_x={round(expected_x, 2)}, width={col['width']}, "
              f"space={col.get('space_after', 0)}")
        # Find actual x closest to expected
        actual = min(x_vals, key=lambda x: abs(x - expected_x)) if x_vals else None
        if actual is not None:
            print(f"    actual_x={actual}, diff={round(actual - expected_x, 2)}")
        expected_x += col["width"] + col.get("space_after", 0)
