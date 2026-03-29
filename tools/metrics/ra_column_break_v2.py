"""
Ra: カラム分割v2 — 十分なテキストでカラムオーバーフロー + mid-para break確認
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []

LONG = "Long paragraph text for overflow testing purposes here. " * 12


def test(scenario, num_short=35, long_text=None, keep_together=False, widow=True):
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
        total = num_short + (1 if long_text else 0)
        for i in range(total):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            if long_text and i == num_short:
                para.Range.Text = long_text
                if keep_together:
                    para.Format.KeepTogether = True
            else:
                para.Range.Text = f"Short line {i+1}."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            para.Format.WidowControl = widow

        wdoc.Repaginate()

        data = {"scenario": scenario, "paragraphs": [], "margin_left": round(ps.LeftMargin, 4)}
        col_info = ps.TextColumns
        data["columns"] = []
        for ci in range(1, col_info.Count + 1):
            col = col_info.Item(ci)
            cd = {"width": round(col.Width, 4)}
            if ci < col_info.Count:
                cd["space_after"] = round(col.SpaceAfter, 4)
            data["columns"].append(cd)

        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
            })

        # For long paragraph: measure individual line positions
        if long_text:
            long_para = wdoc.Paragraphs(num_short + 1)
            lr = long_para.Range
            data["long_lines"] = []
            prev_y = None
            for ci in range(lr.Start, min(lr.End, lr.Start + 800)):
                cr = wdoc.Range(ci, ci + 1)
                y = cr.Information(6)
                x = cr.Information(5)
                if prev_y is None or abs(y - prev_y) > 1:
                    data["long_lines"].append({"x": round(x, 4), "y": round(y, 4)})
                prev_y = y

        return data
    finally:
        wdoc.Close(False)


try:
    # Test 1: Enough short paras to overflow col1 → col2
    d1 = test("overflow_40short", num_short=40)
    results.append(d1)
    ml = d1["margin_left"]
    x_vals = sorted(set(p["x_pt"] for p in d1["paragraphs"]))
    print(f"=== overflow_40short ===")
    print(f"  Unique X: {x_vals}")
    print(f"  Columns: {d1['columns']}")
    for xv in x_vals:
        ps = [p for p in d1["paragraphs"] if p["x_pt"] == xv]
        print(f"  x={xv}(margin+{round(xv-ml,1)}): {len(ps)} paras (P{ps[0]['index']}-P{ps[-1]['index']})")

    # Test 2: Long paragraph that spans columns
    d2 = test("midpara_break", num_short=35, long_text=LONG)
    results.append(d2)
    print(f"\n=== midpara_break ===")
    x_vals = sorted(set(p["x_pt"] for p in d2["paragraphs"]))
    print(f"  Unique X: {x_vals}")
    if "long_lines" in d2:
        print(f"  Long para lines:")
        for ll in d2["long_lines"]:
            print(f"    x={ll['x']}(margin+{round(ll['x']-ml,1)}), y={ll['y']}")

    # Test 3: KeepTogether that forces column break
    d3 = test("keeptogether_colbreak", num_short=35, long_text=LONG, keep_together=True)
    results.append(d3)
    print(f"\n=== keeptogether_colbreak ===")
    x_vals = sorted(set(p["x_pt"] for p in d3["paragraphs"]))
    print(f"  Unique X: {x_vals}")
    last_short = d3["paragraphs"][34] if len(d3["paragraphs"]) > 35 else None
    long_p = d3["paragraphs"][35] if len(d3["paragraphs"]) > 35 else None
    if last_short and long_p:
        print(f"  Last short: x={last_short['x_pt']}, y={last_short['y_pt']}")
        print(f"  Long(keepTogether): x={long_p['x_pt']}, y={long_p['y_pt']}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_column_break_v2.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
