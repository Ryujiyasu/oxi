"""Measure exact line height for Multiple spacing paragraphs with different fonts.
Create a minimal repro document with consecutive same-font paragraphs to isolate
the per-line advance without spacing interference."""
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    # Create a test document with controlled paragraphs
    doc = word.Documents.Add()

    # Set page to A4, margins 72pt
    ps = doc.Sections(1).PageSetup
    ps.PageHeight = 841.9
    ps.PageWidth = 595.3
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LeftMargin = 72
    ps.RightMargin = 72

    # Clear default content
    doc.Content.Delete()

    # Test cases: font + size + Multiple 1.15x, 5 consecutive paragraphs each
    test_cases = [
        ("MS 明朝", 11.0),
        ("MS 明朝", 10.5),
        ("MS ゴシック", 14.0),
        ("MS ゴシック", 13.0),
        ("MS ゴシック", 11.0),
        ("游ゴシック", 11.0),
    ]

    para_idx = 1
    for font, size in test_cases:
        # Add 5 paragraphs with same font/size, Multiple 1.15x
        for j in range(5):
            if para_idx > 1:
                rng = doc.Content
                rng.InsertAfter("\r")

            p = doc.Paragraphs(para_idx)
            p.Range.Text = f"Test {font} {size}pt line {j+1}"
            p.Range.Font.Name = font
            p.Range.Font.Size = size
            p.Format.LineSpacingRule = 5  # wdLineSpaceMultiple
            p.Format.LineSpacing = size * 1.15 * 12 / size  # 1.15x
            # Actually LineSpacing for Multiple is the value in points
            # For 1.15x: LineSpacing = single_height * 1.15
            # But the rule=5 interprets LineSpacing as the actual pt value
            # Use: p.Format.SpaceMultiple = 1.15? No, use raw XML approach
            para_idx += 1

        # Separator
        if para_idx > 1:
            rng = doc.Content
            rng.InsertAfter("\r")
        p = doc.Paragraphs(para_idx)
        p.Range.Text = "---"
        p.Range.Font.Size = 8
        p.Format.LineSpacingRule = 0  # Single
        para_idx += 1

    # Actually setting Multiple spacing properly via COM:
    # Reset and use proper method
    doc.Content.Delete()

    # Simple approach: just set line spacing to auto 276 (1.15x)
    rng = doc.Content
    rng.Text = ""

    fonts_and_sizes = [
        ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 11),
    ]

    # Just measure existing gen2_001 document instead
    doc.Close(False)

    import glob
    docx = glob.glob(os.path.join(os.path.abspath("tools/golden-test/documents/docx"), "gen2_001*"))[0]
    doc = word.Documents.Open(docx, ReadOnly=True)

    # Measure consecutive same-font paragraphs
    print("=== Measuring consecutive body paragraphs (should be pure advance) ===")
    print("Looking for P6-P7 (both body 11pt, both Normal style, sa suppressed by different style)...")
    print()

    # Actually let's find true consecutive pairs:
    # P5→P6: Normal→Normal, sa=10, gap=26.5
    # P6→P7: Normal→ListBullet, sa=10, gap=26.0 (style change)
    # P12→P13: Normal→Normal, sa=10, gap=26.5
    # P13→P14: Normal→Normal, sa=10, gap=26.5

    # To get pure advance without sa: look at P7→P8 (ListBullet→ListBullet, contextual, sa suppressed)
    # P7→P8: gap=16.5 = pure advance
    # P19→P20: gap=16.5 = pure advance (ListBullet→ListBullet)

    # These are j=0 each time (heading resets before them).
    # Need consecutive body paras with no sa to see j>0 pattern.

    # Alternative: create minimal repro
    doc.Close(False)

    # Create minimal repro with 10 consecutive body paragraphs, no spacing
    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.PageHeight = 841.9
    ps.PageWidth = 595.3
    ps.TopMargin = 72
    ps.BottomMargin = 72

    doc.Content.Delete()

    for i in range(10):
        if i > 0:
            sel = word.Selection
            sel.EndKey(6)  # wdStory
            sel.TypeParagraph()

        p = doc.Paragraphs(i + 1)
        p.Range.Text = f"Line {i+1} test text MS Mincho 11pt multiple 1.15x spacing"
        p.Range.Font.Name = "ＭＳ 明朝"
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 5  # Multiple
        p.Format.LineSpacing = 12.65  # 11pt * 1.15 ≈ 12.65

    # Wait, LineSpacing for Multiple is the RENDERED value, not multiplier.
    # For Multiple 1.15x with 11pt: ls = base_height * 1.15
    # COM returns ls=13.8 for gen2 body. Let me use that.
    for i in range(10):
        p = doc.Paragraphs(i + 1)
        p.Format.LineSpacing = 13.8

    # Measure
    print("=== Minimal repro: 10 paragraphs ＭＳ 明朝 11pt, Multiple 1.15x, no spacing ===")
    prev_y = None
    for i in range(1, 11):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        ls = p.Format.LineSpacing
        gap = y - prev_y if prev_y else 0
        print(f"P{i:2d} y={y:7.1f} gap={gap:5.1f} ls={ls:.1f}")
        prev_y = y

    print()

    # Now try with MS Gothic 14pt
    for i in range(10):
        p = doc.Paragraphs(i + 1)
        p.Range.Font.Name = "ＭＳ ゴシック"
        p.Range.Font.Size = 14

    print("=== MS Gothic 14pt, Multiple 1.15x (same ls=13.8 setting) ===")
    prev_y = None
    for i in range(1, 11):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        ls = p.Format.LineSpacing
        gap = y - prev_y if prev_y else 0
        print(f"P{i:2d} y={y:7.1f} gap={gap:5.1f} ls={ls:.1f}")
        prev_y = y

    # Try with proper line spacing for 14pt
    for i in range(10):
        p = doc.Paragraphs(i + 1)
        # For 14pt 1.15x, the ls should be different
        # CJK 83/64: 18.125 * 1.15 = 20.84375
        p.Format.LineSpacing = 20.8

    print()
    print("=== MS Gothic 14pt, ls=20.8 (approx 1.15x) ===")
    prev_y = None
    for i in range(1, 11):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        ls = p.Format.LineSpacing
        gap = y - prev_y if prev_y else 0
        print(f"P{i:2d} y={y:7.1f} gap={gap:5.1f} ls={ls:.1f}")
        prev_y = y

    doc.Close(False)

finally:
    word.Quit()
