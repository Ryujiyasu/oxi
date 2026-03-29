"""
Ra: タブストップとindentの相互作用をCOM計測で確定
- left_indent がタブ位置に影響するか？
- hanging indent + タブの挙動
- firstLineIndent + タブの挙動
- タブ位置 < indent の場合の挙動
"""
import win32com.client, json, os, tempfile

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def measure(doc, scenario):
    """Save doc, open in Word, measure character x positions."""
    tmp = os.path.join(tempfile.gettempdir(), f"ra_tabind_{scenario}.docx")
    doc.Close(False) if hasattr(doc, 'Close') else None

    # Create via COM directly for precise control
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72  # 1 inch = 72pt
        sec.PageSetup.RightMargin = 72

        return wdoc, sec
    except:
        wdoc.Close(False)
        raise


def create_and_measure(scenario, indent_pt=0, hanging_pt=0, first_line_pt=0,
                       tab_positions=None, tab_types=None, text="A\tB\tC"):
    """Create doc via COM with precise indent/tab control and measure."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        sec.PageSetup.LeftMargin = 72  # 1 inch

        # Clear default content
        wdoc.Content.Text = ""

        # Add paragraph with text
        rng = wdoc.Content
        rng.Text = text
        para = wdoc.Paragraphs(1)

        # Set font
        para.Range.Font.Name = "Calibri"
        para.Range.Font.Size = 11

        # Set spacing
        para.Format.SpaceBefore = 0
        para.Format.SpaceAfter = 0

        # Set indents
        para.Format.LeftIndent = indent_pt
        para.Format.FirstLineIndent = first_line_pt
        if hanging_pt > 0:
            para.Format.FirstLineIndent = -hanging_pt  # negative = hanging

        # Clear existing tabs and add custom ones
        para.Format.TabStops.ClearAll()
        if tab_positions:
            for i, pos in enumerate(tab_positions):
                align = 0  # wdAlignTabLeft
                if tab_types and i < len(tab_types):
                    align = {"left": 0, "center": 1, "right": 2, "decimal": 3}[tab_types[i]]
                para.Format.TabStops.Add(pos, align, 0)  # 0 = no leader

        # Measure
        data = {
            "scenario": scenario,
            "indent_pt": indent_pt,
            "hanging_pt": hanging_pt,
            "first_line_pt": first_line_pt,
            "tab_positions": tab_positions or [],
            "margin_left": round(sec.PageSetup.LeftMargin, 4),
            "segments": []
        }

        para_rng = para.Range
        current_seg = {"start_x": None, "chars": ""}

        for ci in range(para_rng.Start, para_rng.End):
            char_rng = wdoc.Range(ci, ci + 1)
            ch = char_rng.Text
            x = char_rng.Information(5)  # wdHorizontalPositionRelativeToPage

            if ord(ch) == 9:  # tab
                if current_seg["start_x"] is not None:
                    data["segments"].append(current_seg)
                current_seg = {"start_x": None, "chars": ""}
            elif ord(ch) == 13:  # paragraph mark
                pass
            else:
                if current_seg["start_x"] is None:
                    current_seg["start_x"] = round(x, 4)
                current_seg["chars"] += ch

        if current_seg["start_x"] is not None:
            data["segments"].append(current_seg)

        return data
    finally:
        wdoc.Close(False)


try:
    tests = [
        # Baseline: no indent, custom tabs
        ("no_indent_tabs", {"indent_pt": 0, "tab_positions": [144, 288], "text": "A\tB\tC"}),

        # Left indent + tabs (tab positions are absolute from margin)
        ("indent36_tabs144", {"indent_pt": 36, "tab_positions": [144, 288], "text": "A\tB\tC"}),
        ("indent72_tabs144", {"indent_pt": 72, "tab_positions": [144, 288], "text": "A\tB\tC"}),

        # Indent > first tab position (what happens?)
        ("indent180_tabs144", {"indent_pt": 180, "tab_positions": [144, 288], "text": "A\tB\tC"}),

        # Hanging indent
        ("hanging36_indent72_tabs144", {"indent_pt": 72, "hanging_pt": 36, "tab_positions": [144, 288],
                                         "text": "A\tB\tC"}),

        # FirstLineIndent + tabs
        ("firstline36_tabs144", {"first_line_pt": 36, "tab_positions": [144, 288], "text": "A\tB\tC"}),

        # No custom tabs, with indent (default tabs)
        ("indent36_default_tabs", {"indent_pt": 36, "text": "A\tB\tC\tD"}),

        # Indent + default tabs: does default tab start from indent or margin?
        ("indent72_default_tabs", {"indent_pt": 72, "text": "A\tB\tC\tD"}),

        # Tab position = indent position exactly
        ("indent144_tabs144", {"indent_pt": 144, "tab_positions": [144, 288], "text": "A\tB\tC"}),

        # Center/right tabs with indent
        ("indent36_center_tab", {"indent_pt": 36, "tab_positions": [216], "tab_types": ["center"],
                                  "text": "Left\tCenter"}),

        # Multiple tab stops, some before indent
        ("indent100_tabs72_144_216", {"indent_pt": 100, "tab_positions": [72, 144, 216],
                                       "text": "A\tB\tC\tD"}),
    ]

    for name, kwargs in tests:
        data = create_and_measure(name, **kwargs)
        results.append(data)

        ml = data["margin_left"]
        print(f"\n=== {name} ===")
        print(f"  indent={data['indent_pt']}pt, hanging={data['hanging_pt']}pt, "
              f"firstLine={data['first_line_pt']}pt")
        print(f"  tabs={data['tab_positions']}")
        for si, seg in enumerate(data["segments"]):
            margin_rel = round(seg["start_x"] - ml, 2)
            print(f"  Seg{si}: x={seg['start_x']}pt (margin+{margin_rel}pt) \"{seg['chars']}\"")

finally:
    word.Quit()

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_tab_indent_interaction.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)

print(f"\nResults saved to {out_path}")

# Analysis
print("\n\n========== ANALYSIS ==========")
baseline = next(r for r in results if r["scenario"] == "no_indent_tabs")
bl_segs = [s["start_x"] for s in baseline["segments"]]
ml = baseline["margin_left"]
print(f"\nBaseline (no indent, tabs@144,288): segments at margin+{[round(x-ml,1) for x in bl_segs]}")

for data in results:
    if data["scenario"] == "no_indent_tabs":
        continue
    segs = [s["start_x"] for s in data["segments"]]
    margin_rels = [round(x - ml, 1) for x in segs]
    print(f"\n{data['scenario']}:")
    print(f"  segments at margin+{margin_rels}")

    # Check: does indent affect tab positions?
    if data["tab_positions"]:
        for si, seg in enumerate(data["segments"][1:], 1):  # skip first segment
            expected_tab = data["tab_positions"][si-1] if si-1 < len(data["tab_positions"]) else None
            if expected_tab:
                actual_margin_rel = round(seg["start_x"] - ml, 1)
                if abs(actual_margin_rel - expected_tab) < 1:
                    print(f"  Seg{si}: at tab position {expected_tab}pt (absolute from margin)")
                else:
                    print(f"  Seg{si}: at {actual_margin_rel}pt, NOT at tab {expected_tab}pt "
                          f"(diff={round(actual_margin_rel - expected_tab, 1)}pt)")
