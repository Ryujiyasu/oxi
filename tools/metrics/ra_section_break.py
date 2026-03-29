"""
Ra: セクション区切り(section break)のレイアウト影響をCOM計測
- continuous section break: 同一ページ内でのセクション変更
- nextPage section break: 強制改ページ
- セクション間でのマージン・カラム変更
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_continuous_section():
    """Continuous section break - same page, different formatting."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72

        wdoc.Content.Text = ""
        # Add text in section 1
        for i in range(3):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Section 1, paragraph {i+1}."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        # Insert continuous section break
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertBreak(3)  # wdSectionBreakContinuous

        # Add text in section 2 (will be same page)
        for i in range(3):
            rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
            rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Section 2, paragraph {i+1}."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "continuous_section", "sections": [], "paragraphs": []}

        # Section info
        for si in range(1, wdoc.Sections.Count + 1):
            sec = wdoc.Sections(si)
            data["sections"].append({
                "index": si,
                "start_page": sec.Range.Information(3),
                "left_margin": round(sec.PageSetup.LeftMargin, 4),
                "cols": sec.PageSetup.TextColumns.Count,
            })

        # Paragraph positions
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
                "section": rng.Sections(1).Index if rng.Sections.Count > 0 else -1,
                "text": rng.Text.strip()[:40],
            })

        return data
    finally:
        wdoc.Close(False)


def test_continuous_with_columns():
    """Continuous section: switch from 1 column to 2 columns mid-page."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72

        wdoc.Content.Text = ""
        for i in range(3):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Single column section, paragraph {i+1}."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        # Continuous break
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertBreak(3)

        # Set section 2 to 2 columns
        wdoc.Sections(2).PageSetup.TextColumns.SetCount(2)

        # Add enough text to fill both columns
        for i in range(20):
            rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
            rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Two column text line {i+1}."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "continuous_columns", "sections": [], "paragraphs": []}

        for si in range(1, wdoc.Sections.Count + 1):
            sec = wdoc.Sections(si)
            data["sections"].append({
                "index": si,
                "cols": sec.PageSetup.TextColumns.Count,
            })

        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
                "page": rng.Information(3),
                "text": rng.Text.strip()[:30],
            })

        return data
    finally:
        wdoc.Close(False)


def test_nextpage_section():
    """Next page section break."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72

        wdoc.Content.Text = "Section 1 text."

        # Next page break
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng.InsertBreak(2)  # wdSectionBreakNextPage

        rng2 = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng2.InsertAfter("Section 2 text on next page.")

        # Change section 2 margins
        wdoc.Sections(2).PageSetup.LeftMargin = 108  # 1.5 inch

        wdoc.Repaginate()

        data = {"scenario": "nextpage_section", "paragraphs": []}
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

        data["sections"] = []
        for si in range(1, wdoc.Sections.Count + 1):
            sec = wdoc.Sections(si)
            data["sections"].append({
                "index": si,
                "left_margin": round(sec.PageSetup.LeftMargin, 4),
                "page": sec.Range.Information(3),
            })

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_continuous_section()
    results.append(d1)
    print("=== continuous_section ===")
    print(f"  Sections: {d1['sections']}")
    for p in d1["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, pg={p['page']}  [{p['text']}]")

    d2 = test_continuous_with_columns()
    results.append(d2)
    print(f"\n=== continuous_columns ===")
    print(f"  Sections: {d2['sections']}")
    x_vals = sorted(set(p["x_pt"] for p in d2["paragraphs"]))
    print(f"  Unique X: {x_vals}")
    for p in d2["paragraphs"][:5]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}  [{p['text']}]")
    print("  ...")
    for p in d2["paragraphs"][-3:]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}  [{p['text']}]")

    d3 = test_nextpage_section()
    results.append(d3)
    print(f"\n=== nextpage_section ===")
    print(f"  Sections: {d3['sections']}")
    for p in d3["paragraphs"]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}, pg={p['page']}  [{p['text']}]")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_section_break.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
