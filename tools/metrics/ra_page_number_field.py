"""
Ra: ページ番号フィールド(PAGE/NUMPAGES)のレンダリング仕様
- フィールドテキストの位置
- フィールドの幅計算
- 右寄せフッター内のページ番号位置
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_page_number_field():
    """Page number field in footer."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72

        # Add enough text for 2 pages
        wdoc.Content.Text = ("Body text paragraph. " * 20 + "\n") * 3

        # Add page number to footer
        sec = wdoc.Sections(1)
        ftr = sec.Footers(1)  # Primary footer
        ftr.Range.Text = ""

        # Add "Page X of Y" format
        ftr.Range.InsertAfter("Page ")
        rng = ftr.Range
        rng.Collapse(0)
        wdoc.Fields.Add(rng, 33)  # wdFieldPage
        rng2 = ftr.Range
        rng2.Collapse(0)
        rng2.InsertAfter(" of ")
        rng3 = ftr.Range
        rng3.Collapse(0)
        wdoc.Fields.Add(rng3, 26)  # wdFieldNumPages

        ftr.Range.Font.Name = "Calibri"
        ftr.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "page_number_field"}

        # Measure footer text
        ftr_rng = ftr.Range
        data["footer_text"] = ftr_rng.Text.strip()
        data["footer_y"] = round(ftr_rng.Information(6), 4)
        data["footer_x"] = round(ftr_rng.Information(5), 4)

        # Fields info
        data["fields"] = []
        for i in range(1, wdoc.Fields.Count + 1):
            field = wdoc.Fields(i)
            data["fields"].append({
                "index": i,
                "type": field.Type,
                "result": field.Result.Text,
            })

        # Total pages
        data["total_pages"] = wdoc.ComputeStatistics(2)  # wdStatisticPages

        return data
    finally:
        wdoc.Close(False)


def test_right_aligned_page_number():
    """Right-aligned page number in footer."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72

        wdoc.Content.Text = ("Text. " * 50 + "\n") * 3

        sec = wdoc.Sections(1)
        ftr = sec.Footers(1)

        # Right-align the footer
        ftr.Range.ParagraphFormat.Alignment = 2  # wdAlignParagraphRight

        ftr.Range.Text = ""
        rng = ftr.Range
        rng.Collapse(0)
        wdoc.Fields.Add(rng, 33)  # wdFieldPage

        ftr.Range.Font.Name = "Calibri"
        ftr.Range.Font.Size = 11

        wdoc.Repaginate()

        ftr_rng = ftr.Range
        data = {
            "scenario": "right_aligned_page",
            "footer_text": ftr_rng.Text.strip(),
            "footer_x": round(ftr_rng.Information(5), 4),
            "footer_y": round(ftr_rng.Information(6), 4),
            "page_width": round(ps.PageWidth, 4),
            "right_margin": round(ps.RightMargin, 4),
        }

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_page_number_field()
    results.append(d1)
    print("=== page_number_field ===")
    print(f"  Footer: \"{d1['footer_text']}\"")
    print(f"  Footer pos: ({d1['footer_x']}, {d1['footer_y']})")
    print(f"  Total pages: {d1['total_pages']}")
    for f in d1["fields"]:
        print(f"  Field {f['index']}: type={f['type']}, result=\"{f['result']}\"")

    d2 = test_right_aligned_page_number()
    results.append(d2)
    print(f"\n=== right_aligned_page ===")
    print(f"  Footer: \"{d2['footer_text']}\"")
    print(f"  Footer x: {d2['footer_x']} (right edge: {d2['page_width'] - d2['right_margin']})")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_page_number.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
