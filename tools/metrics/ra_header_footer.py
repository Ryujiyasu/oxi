"""
Ra: ヘッダー/フッターの位置・サイズをCOM計測で確定
- ヘッダーのY位置（ページ上端からの距離）
- フッターのY位置（ページ下端からの距離）
- ヘッダー/フッターの高さがページ本文領域に影響するか
- 奇偶ページ別ヘッダー
- first page ヘッダー
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_basic_header_footer():
    """Basic header/footer positioning."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.HeaderDistance = 36  # 0.5 inch from page top
        ps.FooterDistance = 36  # 0.5 inch from page bottom

        # Add header text
        hdr = sec.Headers(1)  # wdHeaderFooterPrimary
        hdr.Range.Text = "Header Line 1"
        hdr.Range.Font.Name = "Calibri"
        hdr.Range.Font.Size = 11

        # Add footer text
        ftr = sec.Footers(1)
        ftr.Range.Text = "Footer Line 1"
        ftr.Range.Font.Name = "Calibri"
        ftr.Range.Font.Size = 11

        # Add body text
        wdoc.Content.Text = ""
        for i in range(5):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Body paragraph {i+1}"
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0

        wdoc.Repaginate()

        data = {"scenario": "basic_header_footer"}

        # Measure header
        hdr_rng = hdr.Range
        data["header"] = {
            "y_pt": round(hdr_rng.Information(6), 4),
            "x_pt": round(hdr_rng.Information(5), 4),
            "text": hdr_rng.Text.strip(),
        }

        # Measure footer
        ftr_rng = ftr.Range
        data["footer"] = {
            "y_pt": round(ftr_rng.Information(6), 4),
            "x_pt": round(ftr_rng.Information(5), 4),
            "text": ftr_rng.Text.strip(),
        }

        # Measure body
        data["body_paragraphs"] = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["body_paragraphs"].append({
                "index": i,
                "y_pt": round(rng.Information(6), 4),
                "text": rng.Text.strip()[:40],
            })

        # Page setup values
        data["page_setup"] = {
            "top_margin": round(ps.TopMargin, 4),
            "bottom_margin": round(ps.BottomMargin, 4),
            "header_distance": round(ps.HeaderDistance, 4),
            "footer_distance": round(ps.FooterDistance, 4),
            "page_height": round(ps.PageHeight, 4),
        }

        return data
    finally:
        wdoc.Close(False)


def test_tall_header():
    """Header that's taller than the gap between headerDistance and topMargin."""
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72  # 1 inch
        ps.BottomMargin = 72
        ps.HeaderDistance = 36
        ps.FooterDistance = 36

        # Add tall header (3 lines)
        hdr = sec.Headers(1)
        hdr.Range.Text = "Header Line 1\rHeader Line 2\rHeader Line 3"
        hdr.Range.Font.Name = "Calibri"
        hdr.Range.Font.Size = 14  # larger font

        # Add body text
        wdoc.Content.Text = "Body paragraph 1"
        wdoc.Paragraphs(1).Range.Font.Name = "Calibri"
        wdoc.Paragraphs(1).Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "tall_header"}

        hdr_rng = hdr.Range
        data["header_y"] = round(hdr_rng.Information(6), 4)

        # Measure each header paragraph
        data["header_paragraphs"] = []
        for i in range(1, hdr.Range.Paragraphs.Count + 1):
            p = hdr.Range.Paragraphs(i)
            data["header_paragraphs"].append({
                "index": i,
                "y_pt": round(p.Range.Information(6), 4),
                "text": p.Range.Text.strip()[:30],
            })

        # Measure body paragraph
        body_p = wdoc.Paragraphs(1)
        data["body_y"] = round(body_p.Range.Information(6), 4)
        data["body_text"] = body_p.Range.Text.strip()[:30]

        data["page_setup"] = {
            "top_margin": round(ps.TopMargin, 4),
            "header_distance": round(ps.HeaderDistance, 4),
        }

        return data
    finally:
        wdoc.Close(False)


def test_header_footer_margins():
    """Test different headerDistance/footerDistance values."""
    configs = [
        ("hdr18_ftr18", 18, 18, 72, 72),
        ("hdr36_ftr36", 36, 36, 72, 72),
        ("hdr54_ftr54", 54, 54, 72, 72),
        ("margin36_hdr18", 18, 18, 36, 36),
        ("margin108_hdr36", 36, 36, 108, 108),
    ]

    all_data = []
    for name, hdr_dist, ftr_dist, top_m, bot_m in configs:
        wdoc = word.Documents.Add()
        try:
            sec = wdoc.Sections(1)
            ps = sec.PageSetup
            ps.LeftMargin = 72
            ps.RightMargin = 72
            ps.TopMargin = top_m
            ps.BottomMargin = bot_m
            ps.HeaderDistance = hdr_dist
            ps.FooterDistance = ftr_dist

            hdr = sec.Headers(1)
            hdr.Range.Text = "Header"
            hdr.Range.Font.Name = "Calibri"
            hdr.Range.Font.Size = 11

            ftr = sec.Footers(1)
            ftr.Range.Text = "Footer"
            ftr.Range.Font.Name = "Calibri"
            ftr.Range.Font.Size = 11

            wdoc.Content.Text = "Body text"
            wdoc.Paragraphs(1).Range.Font.Name = "Calibri"
            wdoc.Paragraphs(1).Range.Font.Size = 11

            wdoc.Repaginate()

            data = {
                "name": name,
                "header_distance": hdr_dist,
                "footer_distance": ftr_dist,
                "top_margin": top_m,
                "bottom_margin": bot_m,
                "header_y": round(hdr.Range.Information(6), 4),
                "footer_y": round(ftr.Range.Information(6), 4),
                "body_y": round(wdoc.Paragraphs(1).Range.Information(6), 4),
                "page_height": round(ps.PageHeight, 4),
            }
            all_data.append(data)
        finally:
            wdoc.Close(False)

    return all_data


try:
    # Test 1: Basic header/footer
    data1 = test_basic_header_footer()
    results.append(data1)
    print("=== basic_header_footer ===")
    print(f"  Header: y={data1['header']['y_pt']}pt")
    print(f"  Footer: y={data1['footer']['y_pt']}pt")
    print(f"  Body P1: y={data1['body_paragraphs'][0]['y_pt']}pt")
    ps = data1['page_setup']
    print(f"  PageSetup: topMargin={ps['top_margin']}, headerDist={ps['header_distance']}, "
          f"pageHeight={ps['page_height']}")

    # Test 2: Tall header
    data2 = test_tall_header()
    results.append(data2)
    print(f"\n=== tall_header ===")
    for hp in data2["header_paragraphs"]:
        print(f"  Header P{hp['index']}: y={hp['y_pt']}pt [{hp['text']}]")
    print(f"  Body: y={data2['body_y']}pt [{data2['body_text']}]")
    print(f"  topMargin={data2['page_setup']['top_margin']}, headerDist={data2['page_setup']['header_distance']}")

    # Test 3: Various margin/distance combos
    data3_list = test_header_footer_margins()
    results.extend(data3_list)
    print(f"\n=== margin/distance variations ===")
    for d in data3_list:
        print(f"  {d['name']}: hdr_y={d['header_y']}, body_y={d['body_y']}, ftr_y={d['footer_y']}")
        print(f"    topMargin={d['top_margin']}, hdrDist={d['header_distance']}, "
              f"botMargin={d['bottom_margin']}, ftrDist={d['footer_distance']}")

finally:
    word.Quit()

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_header_footer.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Analysis
print("\n\n========== HEADER/FOOTER ANALYSIS ==========")
print("\nHeader Y position = headerDistance from page top?")
for d in data3_list:
    expected_hdr_y = d["header_distance"]
    print(f"  {d['name']}: expected={expected_hdr_y}, actual={d['header_y']}, "
          f"match={'YES' if abs(d['header_y'] - expected_hdr_y) < 1 else 'NO'}")

print("\nBody Y position = topMargin?")
for d in data3_list:
    expected_body_y = d["top_margin"]
    print(f"  {d['name']}: expected={expected_body_y}, actual={d['body_y']}, "
          f"match={'YES' if abs(d['body_y'] - expected_body_y) < 1 else 'NO'}")

print("\nFooter Y position analysis:")
for d in data3_list:
    page_h = d["page_height"]
    expected_ftr_y = page_h - d["bottom_margin"]
    from_bottom = page_h - d["footer_y"]
    print(f"  {d['name']}: footer_y={d['footer_y']}, page_h={page_h}, "
          f"from_bottom={round(from_bottom, 2)}, footerDist={d['footer_distance']}, "
          f"botMargin={d['bottom_margin']}")
