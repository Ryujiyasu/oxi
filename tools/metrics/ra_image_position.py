"""
Ra: 画像の位置・サイズ精度
- インライン画像のY位置 (テキストと同じ行)
- フローティング画像の絶対位置
- 画像サイズ (原寸 vs スケーリング)
- クリッピング (srcRect)
"""
import win32com.client, json, os

word = win32com.client.DispatchEx('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_inline_image():
    """Inline image positioning within text flow."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72; ps.TopMargin = 72

        wdoc.Content.Text = ""
        # P1: text before image
        wdoc.Content.InsertAfter("Before image ")

        # Try to add a simple inline shape (rectangle as OLE)
        # Use a built-in clip art or create a simple image
        rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)

        # Add an inline shape (horizontal line as substitute)
        ishape = wdoc.InlineShapes.AddHorizontalLineStandard(rng)
        ishape.Width = 100
        ishape.Height = 50

        rng2 = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng2.InsertAfter(" After image.")

        wdoc.Repaginate()

        data = {"scenario": "inline_image"}
        data["inline_shape"] = {
            "width": round(ishape.Width, 4),
            "height": round(ishape.Height, 4),
        }

        # Paragraph Y
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            y = p.Range.Information(6)
            data[f"para{i}_y"] = round(y, 4)

        return data
    except Exception as e:
        return {"scenario": "inline_image", "error": str(e)}
    finally:
        wdoc.Close(False)


def test_floating_image_position():
    """Floating image with different position references."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72; ps.TopMargin = 72

        wdoc.Content.Text = ""
        for i in range(5):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            p = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            p.Range.Text = f"Paragraph {i+1} body text."
            p.Range.Font.Name = "Calibri"
            p.Range.Font.Size = 11

        # Add floating shape (rectangle)
        s = wdoc.Shapes.AddShape(1, 200, 150, 120, 80)
        s.WrapFormat.Type = 3  # wdWrapNone (behind/in front)

        # Test different position references
        data = {"scenario": "floating_image_position", "tests": []}

        # Page relative
        s.RelativeHorizontalPosition = 0  # page
        s.RelativeVerticalPosition = 0
        s.Left = 200; s.Top = 300
        wdoc.Repaginate()
        data["tests"].append({
            "ref": "page",
            "set_left": 200, "set_top": 300,
            "actual_left": round(s.Left, 4),
            "actual_top": round(s.Top, 4),
        })

        # Margin relative
        s.RelativeHorizontalPosition = 1
        s.RelativeVerticalPosition = 1
        s.Left = 100; s.Top = 50
        wdoc.Repaginate()
        data["tests"].append({
            "ref": "margin",
            "set_left": 100, "set_top": 50,
            "actual_left": round(s.Left, 4),
            "actual_top": round(s.Top, 4),
            "expected_abs_x": 72 + 100,
            "expected_abs_y": 72 + 50,
        })

        # Paragraph relative
        s.RelativeVerticalPosition = 2  # paragraph
        s.Top = 10
        wdoc.Repaginate()
        p3_y = wdoc.Paragraphs(3).Range.Information(6)
        data["tests"].append({
            "ref": "paragraph",
            "set_top": 10,
            "actual_top": round(s.Top, 4),
            "anchor_para_y": round(p3_y, 4),
        })

        return data
    finally:
        wdoc.Close(False)


def test_image_scaling():
    """Image size and scaling behavior."""
    wdoc = word.Documents.Add()
    try:
        wdoc.Content.Text = "Body."

        # Add shape with specific dimensions
        s = wdoc.Shapes.AddShape(1, 72, 72, 144, 108)  # 2in x 1.5in

        data = {"scenario": "image_scaling"}
        data["original"] = {
            "width": round(s.Width, 4),
            "height": round(s.Height, 4),
        }

        # Scale to 50%
        s.ScaleWidth(0.5, 0)  # relative to original
        s.ScaleHeight(0.5, 0)
        data["scaled_50pct"] = {
            "width": round(s.Width, 4),
            "height": round(s.Height, 4),
        }

        # Lock aspect ratio and change width
        s.LockAspectRatio = True
        s.Width = 200
        data["locked_aspect_w200"] = {
            "width": round(s.Width, 4),
            "height": round(s.Height, 4),
            "ratio": round(s.Width / s.Height, 4),
        }

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_inline_image()
    results.append(d1)
    print("=== inline_image ===")
    if "error" in d1:
        print(f"  Error: {d1['error']}")
    else:
        print(f"  Shape: w={d1['inline_shape']['width']}, h={d1['inline_shape']['height']}")
        for k, v in d1.items():
            if k.startswith("para"):
                print(f"  {k}: {v}")

    d2 = test_floating_image_position()
    results.append(d2)
    print(f"\n=== floating_image_position ===")
    for t in d2["tests"]:
        print(f"  {t['ref']}: set=({t.get('set_left','-')},{t.get('set_top','-')}), "
              f"actual=({t.get('actual_left','-')},{t.get('actual_top','-')})")
        if "expected_abs_x" in t:
            print(f"    expected_abs=({t['expected_abs_x']},{t['expected_abs_y']})")

    d3 = test_image_scaling()
    results.append(d3)
    print(f"\n=== image_scaling ===")
    print(f"  Original: {d3['original']}")
    print(f"  50%: {d3['scaled_50pct']}")
    print(f"  Locked w=200: {d3['locked_aspect_w200']}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_image_position.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
