"""
Ra: タブストップ + indent — PDF座標で計測（Information(5)のindent問題を回避）
- Word COM で文書作成 → PDF保存 → PDF座標を読み取り
- 代替: Selection.MoveRight + Information(5) で各文字位置を確認
"""
import win32com.client, json, os, tempfile

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def create_and_measure(scenario, indent=0, hanging=0, first_line=0,
                       tab_positions=None, text="AAA\tBBB\tCCC"):
    """Create doc via COM, print to PDF, extract text positions."""
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
            for pos in tab_positions:
                para.Format.TabStops.Add(pos, 0, 0)  # left tab, no leader

        # Add P2 (continuation line, no firstLineIndent)
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertParagraphAfter()
        p2 = wdoc.Paragraphs(2)
        p2.Range.Text = "XXX\tYYY\tZZZ"
        p2.Range.Font.Name = "Calibri"
        p2.Range.Font.Size = 11
        p2.Format.LeftIndent = indent
        p2.Format.FirstLineIndent = 0
        p2.Format.TabStops.ClearAll()
        if tab_positions:
            for pos in tab_positions:
                p2.Format.TabStops.Add(pos, 0, 0)

        wdoc.Repaginate()

        data = {
            "scenario": scenario,
            "indent": indent, "hanging": hanging, "first_line": first_line,
            "tab_positions": tab_positions or [],
            "margin_left": round(sec.PageSetup.LeftMargin, 4),
            "paragraphs": []
        }

        # Method: use Selection to move character by character and get precise positions
        # Selection.Information gives more reliable positions than Range.Information
        for pi in range(1, 3):
            para = wdoc.Paragraphs(pi)
            prng = para.Range

            # Select start of paragraph
            prng.Select()
            sel = wdoc.Application.Selection
            sel.HomeKey(5)  # wdLine - go to start of line

            segments = []
            cur_seg = {"start_x": None, "chars": ""}

            # Move through each character
            for ci in range(prng.Start, prng.End):
                cr = wdoc.Range(ci, ci + 1)
                ch = cr.Text

                # Get position by selecting the character
                cr.Select()
                x = wdoc.Application.Selection.Information(5)  # wdHorizontalPositionRelativeToPage

                if ord(ch) == 9:  # tab
                    if cur_seg["start_x"] is not None:
                        segments.append(cur_seg)
                    cur_seg = {"start_x": None, "chars": ""}
                elif ord(ch) == 13:  # para mark
                    pass
                else:
                    if cur_seg["start_x"] is None:
                        cur_seg["start_x"] = round(x, 4)
                    cur_seg["chars"] += ch

            if cur_seg["start_x"] is not None:
                segments.append(cur_seg)

            data["paragraphs"].append({
                "index": pi,
                "left_indent": round(para.Format.LeftIndent, 4),
                "first_line_indent": round(para.Format.FirstLineIndent, 4),
                "segments": segments
            })

        # Also save as PDF and note the path for external verification
        pdf_path = os.path.join(tempfile.gettempdir(), f"ra_tabind_{scenario}.pdf")
        wdoc.SaveAs2(pdf_path, FileFormat=17)  # wdFormatPDF
        data["pdf_path"] = pdf_path

        return data
    finally:
        wdoc.Close(False)


try:
    tests = [
        ("baseline", {}),
        ("tabs144", {"tab_positions": [144, 288]}),
        ("indent36_tabs144", {"indent": 36, "tab_positions": [144, 288]}),
        ("indent72_tabs144", {"indent": 72, "tab_positions": [144, 288]}),
        ("indent180_tabs144", {"indent": 180, "tab_positions": [144, 288]}),
        ("hanging36_indent72", {"indent": 72, "hanging": 36, "tab_positions": [144, 288]}),
        ("firstline36", {"first_line": 36, "tab_positions": [144, 288]}),
        ("indent36_default", {"indent": 36}),
        ("indent72_default", {"indent": 72}),
    ]

    for name, kwargs in tests:
        data = create_and_measure(name, **kwargs)
        results.append(data)
        ml = data["margin_left"]
        print(f"\n=== {name} ===")
        print(f"  indent={data['indent']}, hanging={data['hanging']}, firstLine={data['first_line']}")
        for pd in data["paragraphs"]:
            label = "P1" if pd["index"] == 1 else "P2"
            segs = [(round(s["start_x"] - ml, 2), s["chars"]) for s in pd["segments"]]
            print(f"  {label} (li={pd['left_indent']}, fli={pd['first_line_indent']}): {segs}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_tab_indent_pdf.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Final analysis
print("\n\n========== TAB + INDENT DEFINITIVE ANALYSIS ==========")
for data in results:
    ml = data["margin_left"]
    tabs = data["tab_positions"]
    indent = data["indent"]
    hanging = data["hanging"]
    fl = data["first_line"]
    print(f"\n{data['scenario']}:")
    for pd in data["paragraphs"]:
        effective = pd["left_indent"] + pd["first_line_indent"]
        segs_abs = [round(s["start_x"], 2) for s in pd["segments"]]
        segs_mr = [round(s["start_x"] - ml, 2) for s in pd["segments"]]
        print(f"  P{pd['index']}: effective_indent={effective}pt, segs(margin-rel)={segs_mr}")
        # Check if first segment starts at effective indent
        if pd["segments"]:
            first_x = pd["segments"][0]["start_x"] - ml
            print(f"    first_char at margin+{round(first_x, 2)}pt (expected indent={effective}pt, diff={round(first_x - effective, 2)})")
