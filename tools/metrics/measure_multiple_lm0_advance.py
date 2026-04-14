"""Create LM=0 minimal repro documents to measure pure Multiple spacing advance.
Sets LayoutMode=0 explicitly, no spacing, consecutive same-font paragraphs."""
import win32com.client

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    test_cases = [
        ("ＭＳ 明朝", 11.0),
        ("ＭＳ 明朝", 10.5),
        ("ＭＳ ゴシック", 14.0),
        ("ＭＳ ゴシック", 13.0),
        ("ＭＳ ゴシック", 11.0),
        ("Calibri", 11.0),
    ]

    for font_name, font_size in test_cases:
        doc = word.Documents.Add()
        ps = doc.Sections(1).PageSetup
        ps.LayoutMode = 0  # LM=0, no grid
        ps.TopMargin = 72
        ps.BottomMargin = 72

        doc.Content.Delete()

        # Create 10 consecutive paragraphs with same font, Multiple 1.15x, no spacing
        for i in range(10):
            if i > 0:
                sel = word.Selection
                sel.EndKey(6)
                sel.TypeParagraph()

            p = doc.Paragraphs(i + 1)
            p.Range.Text = f"Line {i+1} test text"
            p.Range.Font.Name = font_name
            p.Range.Font.Size = font_size
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            # Multiple 1.15x
            p.Format.LineSpacingRule = 5
            p.Format.LineSpacing = 15.87  # placeholder, will be set by Word

        # Set spacing via XML-like method: 276/240 = 1.15x
        # Actually for rule=5, LineSpacing = rendered_height.
        # The multiplier is stored as w:line="276" w:lineRule="auto" in XML.
        # COM setting: use LinesToPoints for multiple
        for i in range(10):
            p = doc.Paragraphs(i + 1)
            # wdLineSpaceMultiple=5, LineSpacing = multiplier * 12
            # 1.15 * 12 = 13.8
            p.Format.LineSpacing = 13.8
            p.Format.LineSpacingRule = 5

        # Verify and measure
        print(f"=== {font_name} {font_size}pt, LM=0, Multiple 1.15x ===")
        print(f"  LayoutMode: {ps.LayoutMode}")

        prev_y = None
        gaps = []
        for i in range(1, 11):
            p = doc.Paragraphs(i)
            y = p.Range.Information(6)
            gap = y - prev_y if prev_y else 0
            gaps.append(gap)
            prev_y = y

        # Show advances (skip gap[0]=0)
        advances = gaps[1:]
        print(f"  Advances: {[round(a, 1) for a in advances]}")
        unique = sorted(set(round(a, 1) for a in advances))
        print(f"  Unique: {unique}")
        if len(unique) <= 2:
            print(f"  Average: {sum(advances)/len(advances):.3f}pt")

        doc.Close(False)

finally:
    word.Quit()
