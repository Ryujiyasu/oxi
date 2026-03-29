"""
Ra: 脚注(footnote)の位置・レイアウトをCOM計測で確定
- 脚注テキストのY位置
- 脚注セパレータの位置
- 脚注がページ本文領域を圧迫するか
- 複数脚注の並び
- 脚注がページ跨ぎする場合
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []

LOREM = ("Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
         "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ")


def test_single_footnote():
    """Single footnote - measure position."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72

        # Add body text with footnote
        wdoc.Content.Text = ""
        wdoc.Content.InsertAfter("This is body text with a footnote reference")

        # Insert footnote at end of text
        rng = wdoc.Range(wdoc.Content.End - 2, wdoc.Content.End - 1)
        fn = wdoc.Footnotes.Add(rng, Text="This is the footnote text.")

        # Add more body text
        body_end = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        body_end.InsertParagraphAfter()
        p2 = wdoc.Paragraphs(wdoc.Paragraphs.Count)
        p2.Range.Text = "Second body paragraph after footnote."
        p2.Range.Font.Name = "Calibri"
        p2.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "single_footnote"}

        # Measure body paragraphs
        data["body_paragraphs"] = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng_p = para.Range
            data["body_paragraphs"].append({
                "index": i,
                "y_pt": round(rng_p.Information(6), 4),
                "text": rng_p.Text.strip()[:50],
            })

        # Measure footnote
        if wdoc.Footnotes.Count > 0:
            fn_rng = wdoc.Footnotes(1).Range
            data["footnote"] = {
                "y_pt": round(fn_rng.Information(6), 4),
                "x_pt": round(fn_rng.Information(5), 4),
                "text": fn_rng.Text.strip()[:50],
                "page": fn_rng.Information(3),
            }

        data["page_setup"] = {
            "page_height": round(ps.PageHeight, 4),
            "top_margin": round(ps.TopMargin, 4),
            "bottom_margin": round(ps.BottomMargin, 4),
        }

        return data
    finally:
        wdoc.Close(False)


def test_multiple_footnotes():
    """Multiple footnotes on same page."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72

        wdoc.Content.Text = ""

        # Add 3 paragraphs with footnotes
        for i in range(3):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Paragraph {i+1} with footnote reference here."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

            # Add footnote
            fn_rng = wdoc.Range(para.Range.End - 2, para.Range.End - 1)
            wdoc.Footnotes.Add(fn_rng, Text=f"Footnote {i+1} text content.")

        wdoc.Repaginate()

        data = {"scenario": "multiple_footnotes"}

        # Measure footnotes
        data["footnotes"] = []
        for i in range(1, wdoc.Footnotes.Count + 1):
            fn = wdoc.Footnotes(i)
            fn_rng = fn.Range
            data["footnotes"].append({
                "index": i,
                "y_pt": round(fn_rng.Information(6), 4),
                "x_pt": round(fn_rng.Information(5), 4),
                "page": fn_rng.Information(3),
                "text": fn_rng.Text.strip()[:40],
            })

        data["page_setup"] = {
            "page_height": round(ps.PageHeight, 4),
            "bottom_margin": round(ps.BottomMargin, 4),
        }

        return data
    finally:
        wdoc.Close(False)


def test_footnote_font_size():
    """Check default footnote font size and line height."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        wdoc.Content.Text = "Body text."
        rng = wdoc.Range(0, 9)
        fn = wdoc.Footnotes.Add(rng, Text="Footnote text for measuring.")

        wdoc.Repaginate()

        fn_rng = fn.Range
        data = {
            "scenario": "footnote_font",
            "font_name": fn_rng.Font.Name,
            "font_size": round(fn_rng.Font.Size, 4),
            "line_spacing": round(fn_rng.ParagraphFormat.LineSpacing, 4),
            "ls_rule": fn_rng.ParagraphFormat.LineSpacingRule,
            "space_after": round(fn_rng.ParagraphFormat.SpaceAfter, 4),
            "space_before": round(fn_rng.ParagraphFormat.SpaceBefore, 4),
        }

        return data
    finally:
        wdoc.Close(False)


def test_footnote_with_long_body():
    """Footnote position when body text is long (footnote at page bottom)."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72

        wdoc.Content.Text = ""

        # Fill most of the page with body text
        for i in range(30):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Line {i+1}: {LOREM[:60]}"
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0

        # Add footnote to first paragraph
        p1_rng = wdoc.Range(0, 5)
        wdoc.Footnotes.Add(p1_rng, Text="Long page footnote.")

        wdoc.Repaginate()

        data = {"scenario": "footnote_long_body"}

        # Last body paragraph on page 1
        data["body_paragraphs"] = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            pg = para.Range.Information(3)
            if pg == 1:
                data["body_paragraphs"].append({
                    "index": i,
                    "y_pt": round(para.Range.Information(6), 4),
                })

        data["last_body_on_p1"] = data["body_paragraphs"][-1] if data["body_paragraphs"] else None
        data["body_count_p1"] = len(data["body_paragraphs"])

        fn_rng = wdoc.Footnotes(1).Range
        data["footnote"] = {
            "y_pt": round(fn_rng.Information(6), 4),
            "page": fn_rng.Information(3),
        }

        data["page_setup"] = {
            "page_height": round(ps.PageHeight, 4),
            "bottom_margin": round(ps.BottomMargin, 4),
        }

        # Footnote distance from page bottom
        data["footnote_from_bottom"] = round(ps.PageHeight - fn_rng.Information(6), 4)

        return data
    finally:
        wdoc.Close(False)


try:
    # Test 1
    d1 = test_single_footnote()
    results.append(d1)
    print("=== single_footnote ===")
    for bp in d1["body_paragraphs"]:
        print(f"  Body P{bp['index']}: y={bp['y_pt']}  [{bp['text']}]")
    if "footnote" in d1:
        fn = d1["footnote"]
        print(f"  Footnote: y={fn['y_pt']}, page={fn['page']}  [{fn['text']}]")
    ps = d1["page_setup"]
    print(f"  Page: height={ps['page_height']}, botMargin={ps['bottom_margin']}")

    # Test 2
    d2 = test_multiple_footnotes()
    results.append(d2)
    print(f"\n=== multiple_footnotes ===")
    for fn in d2["footnotes"]:
        print(f"  FN{fn['index']}: y={fn['y_pt']}, page={fn['page']}  [{fn['text']}]")

    # Test 3
    d3 = test_footnote_font_size()
    results.append(d3)
    print(f"\n=== footnote_font ===")
    print(f"  Font: {d3['font_name']} {d3['font_size']}pt")
    print(f"  LineSpacing: {d3['line_spacing']}pt (rule={d3['ls_rule']})")
    print(f"  SpaceBefore: {d3['space_before']}, SpaceAfter: {d3['space_after']}")

    # Test 4
    d4 = test_footnote_with_long_body()
    results.append(d4)
    print(f"\n=== footnote_long_body ===")
    print(f"  Body paragraphs on page 1: {d4['body_count_p1']}")
    if d4["last_body_on_p1"]:
        print(f"  Last body Y: {d4['last_body_on_p1']['y_pt']}")
    print(f"  Footnote: y={d4['footnote']['y_pt']}, page={d4['footnote']['page']}")
    print(f"  Footnote from page bottom: {d4['footnote_from_bottom']}pt")
    print(f"  Page height: {d4['page_setup']['page_height']}, botMargin={d4['page_setup']['bottom_margin']}")

finally:
    word.Quit()

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_footnote.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Analysis
print("\n\n========== FOOTNOTE POSITION ANALYSIS ==========")
if "footnote" in d4:
    fn_y = d4["footnote"]["y_pt"]
    page_h = d4["page_setup"]["page_height"]
    bot_m = d4["page_setup"]["bottom_margin"]
    body_area_bottom = page_h - bot_m
    print(f"Body area bottom: {body_area_bottom}pt")
    print(f"Footnote Y: {fn_y}pt")
    print(f"Footnote is {'above' if fn_y < body_area_bottom else 'below'} body area bottom")
    if d4["last_body_on_p1"]:
        last_y = d4["last_body_on_p1"]["y_pt"]
        print(f"Last body Y: {last_y}pt")
        print(f"Gap between last body and footnote: {round(fn_y - last_y, 2)}pt")
