"""
Ra: タブストップ + indent v2 — Repaginate() + python-docxではなくCOM直接生成
前回 Information(5) がindent反映していなかった問題を修正
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test(scenario, indent=0, hanging=0, first_line=0, tab_positions=None, tab_types=None,
         text="AAA\tBBB\tCCC"):
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72
        sec.PageSetup.RightMargin = 72

        wdoc.Content.Text = text
        para = wdoc.Paragraphs(1)
        para.Range.Font.Name = "Calibri"
        para.Range.Font.Size = 11
        para.Format.SpaceBefore = 0
        para.Format.SpaceAfter = 0
        para.Format.LeftIndent = indent
        para.Format.FirstLineIndent = first_line
        if hanging > 0:
            para.Format.FirstLineIndent = -hanging

        para.Format.TabStops.ClearAll()
        if tab_positions:
            for i, pos in enumerate(tab_positions):
                align = 0  # left
                if tab_types and i < len(tab_types):
                    align = {"left": 0, "center": 1, "right": 2}[tab_types[i]]
                para.Format.TabStops.Add(pos, align, 0)

        # Add a second paragraph for comparison (same settings, second line of text)
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertParagraphAfter()
        p2 = wdoc.Paragraphs(2)
        p2.Range.Text = "XXX\tYYY\tZZZ"
        p2.Range.Font.Name = "Calibri"
        p2.Range.Font.Size = 11
        p2.Format.SpaceBefore = 0
        p2.Format.SpaceAfter = 0
        p2.Format.LeftIndent = indent
        p2.Format.FirstLineIndent = 0  # No first-line indent for P2 (continuation)
        p2.Format.TabStops.ClearAll()
        if tab_positions:
            for i, pos in enumerate(tab_positions):
                align = 0
                if tab_types and i < len(tab_types):
                    align = {"left": 0, "center": 1, "right": 2}[tab_types[i]]
                p2.Format.TabStops.Add(pos, align, 0)

        wdoc.Repaginate()

        # Read back actual values
        actual_indent = round(para.Format.LeftIndent, 4)
        actual_first = round(para.Format.FirstLineIndent, 4)

        data = {
            "scenario": scenario,
            "set_indent": indent, "set_hanging": hanging, "set_first_line": first_line,
            "actual_indent": actual_indent, "actual_first_line_indent": actual_first,
            "tab_positions": tab_positions or [],
            "margin_left": round(sec.PageSetup.LeftMargin, 4),
            "paragraphs": []
        }

        for pi in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(pi)
            prng = p.Range
            segments = []
            cur = {"start_x": None, "chars": ""}

            for ci in range(prng.Start, prng.End):
                cr = wdoc.Range(ci, ci + 1)
                ch = cr.Text
                x = cr.Information(5)

                if ord(ch) == 9:
                    if cur["start_x"] is not None:
                        segments.append(cur)
                    cur = {"start_x": None, "chars": ""}
                elif ord(ch) == 13:
                    pass
                else:
                    if cur["start_x"] is None:
                        cur["start_x"] = round(x, 4)
                    cur["chars"] += ch

            if cur["start_x"] is not None:
                segments.append(cur)

            data["paragraphs"].append({
                "index": pi,
                "y_pt": round(prng.Information(6), 4),
                "first_line_indent": round(p.Format.FirstLineIndent, 4),
                "left_indent": round(p.Format.LeftIndent, 4),
                "segments": segments
            })

        return data
    finally:
        wdoc.Close(False)


try:
    tests = [
        ("baseline", {}),
        ("tabs_144_288", {"tab_positions": [144, 288]}),
        ("indent36_tabs144_288", {"indent": 36, "tab_positions": [144, 288]}),
        ("indent72_tabs144_288", {"indent": 72, "tab_positions": [144, 288]}),
        ("indent180_tabs144_288", {"indent": 180, "tab_positions": [144, 288]}),
        ("hanging36_indent72_tabs144", {"indent": 72, "hanging": 36, "tab_positions": [144, 288]}),
        ("firstline36_tabs144", {"first_line": 36, "tab_positions": [144, 288]}),
        ("indent36_no_tabs", {"indent": 36}),
        ("indent72_no_tabs", {"indent": 72}),
    ]

    for name, kwargs in tests:
        data = test(name, **kwargs)
        results.append(data)
        ml = data["margin_left"]

        print(f"\n=== {name} ===")
        print(f"  actual_indent={data['actual_indent']}, actual_firstLine={data['actual_first_line_indent']}")
        print(f"  tabs={data['tab_positions']}")

        for pd in data["paragraphs"]:
            prefix = "P1(firstLine)" if pd["index"] == 1 else "P2(continue)"
            print(f"  {prefix}: leftIndent={pd['left_indent']}, firstLineIndent={pd['first_line_indent']}")
            for si, seg in enumerate(pd["segments"]):
                mr = round(seg["start_x"] - ml, 2)
                print(f"    Seg{si}: x={seg['start_x']}pt (margin+{mr}pt) \"{seg['chars']}\"")

finally:
    word.Quit()

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_tab_indent_v2.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Analysis
print("\n\n========== TAB + INDENT ANALYSIS ==========")
for data in results:
    ml = data["margin_left"]
    tabs = data["tab_positions"]
    print(f"\n{data['scenario']}:")
    for pd in data["paragraphs"]:
        segs_mr = [round(s["start_x"] - ml, 2) for s in pd["segments"]]
        li = pd["left_indent"]
        fli = pd["first_line_indent"]
        effective_start = li + fli
        print(f"  P{pd['index']}: effectiveStart={effective_start}pt, segments(margin-rel)={segs_mr}")
        if tabs and len(pd["segments"]) > 1:
            for ti, tab in enumerate(tabs):
                if ti + 1 < len(pd["segments"]):
                    actual = segs_mr[ti + 1]
                    print(f"    tab[{ti}]={tab}pt, actual_seg={actual}pt, diff={round(actual - tab, 2)}")
