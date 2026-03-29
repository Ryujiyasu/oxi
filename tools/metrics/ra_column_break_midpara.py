"""
Ra: 段落がカラムを跨ぐ時の行分割ルールをCOM計測で確定
- 長い段落がカラム境界で分割される位置
- widow/orphan制御はカラム間でも有効か
- 段落途中でカラムが変わる時のY位置
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []

LONG_PARA = ("This is a long paragraph that should span across column boundaries. " * 8).strip()
SHORT_PARA = "Short paragraph."


def test_midpara_column_break():
    """Long paragraph that overflows column 1 into column 2."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.TextColumns.SetCount(2)
        ps.TextColumns.EvenlySpaced = True

        wdoc.Content.Text = ""
        # Fill column 1 partially, then add a long paragraph
        for i in range(15):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = SHORT_PARA if i < 14 else LONG_PARA
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0

        wdoc.Repaginate()

        data = {"scenario": "midpara_column_break", "paragraphs": []}
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
                "text": rng.Text.strip()[:30],
                "line_count": para.Range.ComputeStatistics(1),  # wdStatisticLines
            })

        # Also measure individual lines of the long paragraph
        long_para = wdoc.Paragraphs(15)
        long_rng = long_para.Range
        data["long_para_lines"] = []

        # Walk through the range character by character to find line boundaries
        prev_y = None
        line_start_x = None
        line_y = None
        for ci in range(long_rng.Start, min(long_rng.End, long_rng.Start + 500)):
            cr = wdoc.Range(ci, ci + 1)
            y = cr.Information(6)
            x = cr.Information(5)
            if prev_y is None or abs(y - prev_y) > 1:
                if line_y is not None:
                    data["long_para_lines"].append({
                        "x_pt": round(line_start_x, 4),
                        "y_pt": round(line_y, 4),
                    })
                line_start_x = x
                line_y = y
            prev_y = y

        if line_y is not None:
            data["long_para_lines"].append({
                "x_pt": round(line_start_x, 4),
                "y_pt": round(line_y, 4),
            })

        return data
    finally:
        wdoc.Close(False)


def test_widow_orphan_column():
    """Test widow/orphan control at column boundaries."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.TextColumns.SetCount(2)

        wdoc.Content.Text = ""
        # Fill column 1 almost completely, then add 3-line paragraph
        for i in range(25):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            if i < 24:
                para.Range.Text = SHORT_PARA
            else:
                # 3-line paragraph at the end
                para.Range.Text = "Line one of orphan test. " * 6
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            para.Format.WidowControl = True

        wdoc.Repaginate()

        data = {"scenario": "widow_orphan_column", "paragraphs": []}
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
            })

        return data
    finally:
        wdoc.Close(False)


def test_column_break_char():
    """Test explicit column break character (ctrl+shift+enter)."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.TextColumns.SetCount(2)

        wdoc.Content.Text = "Col1 text before break"
        # Insert column break
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertBreak(8)  # wdColumnBreak
        rng2 = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng2.InsertAfter("Col2 text after break")

        wdoc.Repaginate()

        data = {"scenario": "explicit_column_break", "paragraphs": []}
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
                "text": rng.Text.strip()[:40],
            })

        # Check column info
        data["columns"] = []
        tc = ps.TextColumns
        for ci in range(1, tc.Count + 1):
            col = tc.Item(ci)
            cd = {"index": ci, "width": round(col.Width, 4)}
            if ci < tc.Count:
                cd["space_after"] = round(col.SpaceAfter, 4)
            data["columns"].append(cd)
        data["margin_left"] = round(ps.LeftMargin, 4)

        return data
    finally:
        wdoc.Close(False)


def test_keeplines_column():
    """Test keepTogether/keepLines with columns."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.TextColumns.SetCount(2)

        wdoc.Content.Text = ""
        # Fill col 1 almost full, then add keepTogether paragraph
        for i in range(23):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            if i < 22:
                para.Range.Text = SHORT_PARA
            else:
                para.Range.Text = "KeepTogether paragraph. " * 8
                para.Format.KeepTogether = True
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0

        wdoc.Repaginate()

        data = {"scenario": "keeplines_column", "paragraphs": []}
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
            })

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_midpara_column_break()
    results.append(d1)
    print("=== midpara_column_break ===")
    for p in d1["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, pg={p['page']}, lines={p.get('line_count','-')}  [{p['text']}]")
    print(f"\n  Long para (P15) line positions:")
    for ll in d1["long_para_lines"]:
        print(f"    x={ll['x_pt']}, y={ll['y_pt']}")

    d2 = test_widow_orphan_column()
    results.append(d2)
    print(f"\n=== widow_orphan_column ===")
    # Show last 5 paragraphs
    for p in d2["paragraphs"][-6:]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, pg={p['page']}")

    d3 = test_column_break_char()
    results.append(d3)
    print(f"\n=== explicit_column_break ===")
    ml = d3["margin_left"]
    for p in d3["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}(margin+{round(p['x_pt']-ml,2)}), y={p['y_pt']}  [{p['text']}]")
    print(f"  Columns: {d3['columns']}")

    d4 = test_keeplines_column()
    results.append(d4)
    print(f"\n=== keeplines_column ===")
    for p in d4["paragraphs"][-5:]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, pg={p['page']}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_column_break_midpara.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
