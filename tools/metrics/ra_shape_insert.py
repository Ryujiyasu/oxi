"""
Ra: 図形(Shape)挿入の位置精度
- インライン図形 vs アンカー図形
- 図形の位置基準 (page/margin/paragraph/column/character)
- wrapSquare/wrapTight/wrapTopAndBottom の影響
- 図形がテキストフローに与える影響
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_shape_position_refs():
    """Shape positioning relative to different anchors."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72

        # Add body text
        wdoc.Content.Text = ""
        for i in range(10):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Body paragraph {i+1} with some text content."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11

        data = {"scenario": "shape_position_refs", "shapes": []}

        # Shape 1: position relative to page
        s1 = wdoc.Shapes.AddShape(1, 100, 200, 80, 60)  # msoShapeRectangle
        s1.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionPage
        s1.RelativeVerticalPosition = 0    # wdRelativeVerticalPositionPage
        s1.Left = 100
        s1.Top = 200

        data["shapes"].append({
            "label": "page_relative",
            "left": round(s1.Left, 4),
            "top": round(s1.Top, 4),
            "width": round(s1.Width, 4),
            "height": round(s1.Height, 4),
            "h_rel": s1.RelativeHorizontalPosition,
            "v_rel": s1.RelativeVerticalPosition,
        })

        # Shape 2: position relative to margin
        s2 = wdoc.Shapes.AddShape(1, 50, 100, 60, 40)
        s2.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionMargin
        s2.RelativeVerticalPosition = 1    # wdRelativeVerticalPositionMargin
        s2.Left = 50
        s2.Top = 100

        data["shapes"].append({
            "label": "margin_relative",
            "left": round(s2.Left, 4),
            "top": round(s2.Top, 4),
            "width": round(s2.Width, 4),
            "height": round(s2.Height, 4),
            "h_rel": s2.RelativeHorizontalPosition,
            "v_rel": s2.RelativeVerticalPosition,
            "expected_abs_x": round(72 + 50, 4),  # margin + offset
            "expected_abs_y": round(72 + 100, 4),
        })

        # Shape 3: position relative to paragraph
        s3 = wdoc.Shapes.AddShape(1, 0, 0, 70, 50, wdoc.Paragraphs(5).Range)
        s3.RelativeHorizontalPosition = 2  # wdRelativeHorizontalPositionColumn
        s3.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph
        s3.Left = 200
        s3.Top = 0

        p5_y = wdoc.Paragraphs(5).Range.Information(6)
        data["shapes"].append({
            "label": "paragraph_relative",
            "left": round(s3.Left, 4),
            "top": round(s3.Top, 4),
            "width": round(s3.Width, 4),
            "height": round(s3.Height, 4),
            "h_rel": s3.RelativeHorizontalPosition,
            "v_rel": s3.RelativeVerticalPosition,
            "anchor_para_y": round(p5_y, 4),
        })

        return data
    finally:
        wdoc.Close(False)


def test_wrap_types():
    """Different text wrapping modes and their impact on text flow."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72

        wdoc.Content.Text = ""
        for i in range(15):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            para = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            para.Range.Text = f"Paragraph {i+1} text that flows around shapes."
            para.Range.Font.Name = "Calibri"
            para.Range.Font.Size = 11
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0

        # Add shape with wrapSquare between P3 and P4
        s = wdoc.Shapes.AddShape(1, 72, 0, 100, 80, wdoc.Paragraphs(4).Range)
        s.WrapFormat.Type = 0  # wdWrapSquare
        s.WrapFormat.Side = 0  # wdWrapBoth
        s.RelativeHorizontalPosition = 1  # margin
        s.RelativeVerticalPosition = 2    # paragraph
        s.Left = 0
        s.Top = 0

        wdoc.Repaginate()

        data = {"scenario": "wrap_square", "paragraphs": []}
        for i in range(1, 16):
            para = wdoc.Paragraphs(i)
            rng = para.Range
            data["paragraphs"].append({
                "index": i,
                "x_pt": round(rng.Information(5), 4),
                "y_pt": round(rng.Information(6), 4),
            })

        data["shape"] = {
            "left": round(s.Left, 4),
            "top": round(s.Top, 4),
            "width": round(s.Width, 4),
            "height": round(s.Height, 4),
            "wrap_type": s.WrapFormat.Type,
        }

        return data
    finally:
        wdoc.Close(False)


def test_inline_shape():
    """Inline shape (part of text flow)."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72

        wdoc.Content.Text = "Before shape "
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)

        # Add inline shape (rectangle)
        ishape = wdoc.InlineShapes.AddOLEObject(Range=rng)
        # This might fail, try AddPicture or simple shape
        # Actually, let's just measure inline image behavior

        # Simpler: add text with different sizes to simulate
        wdoc.Content.Text = "Before "
        # Can't easily create inline shapes via COM without a file
        # Skip this test

        data = {"scenario": "inline_shape", "note": "skipped - COM limitation"}
        return data
    except:
        return {"scenario": "inline_shape", "note": "skipped"}
    finally:
        wdoc.Close(False)


try:
    d1 = test_shape_position_refs()
    results.append(d1)
    print("=== shape_position_refs ===")
    for s in d1["shapes"]:
        print(f"  {s['label']}: left={s['left']}, top={s['top']}, "
              f"w={s['width']}, h={s['height']}, h_rel={s['h_rel']}, v_rel={s['v_rel']}")
        if "expected_abs_x" in s:
            print(f"    expected_abs: ({s['expected_abs_x']}, {s['expected_abs_y']})")
        if "anchor_para_y" in s:
            print(f"    anchor_para_y: {s['anchor_para_y']}")

    d2 = test_wrap_types()
    results.append(d2)
    print(f"\n=== wrap_square ===")
    print(f"  Shape: left={d2['shape']['left']}, top={d2['shape']['top']}, "
          f"w={d2['shape']['width']}, h={d2['shape']['height']}")
    for p in d2["paragraphs"][:8]:
        print(f"  P{p['index']}: x={p['x_pt']}, y={p['y_pt']}")

    d3 = test_inline_shape()
    results.append(d3)

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_shape_insert.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
