"""Systematic COM measurement of spacing collapse and contextual spacing.
Tests heading→body transitions to determine exact spacing formula."""
import win32com.client
import math

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    # ===== TEST 1: Heading → Body spacing with Multiple 1.15x =====
    print("=" * 60)
    print("TEST 1: Heading→Body spacing (Multiple 1.15x, LM=0)")
    print("=" * 60)

    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.LayoutMode = 0
    ps.TopMargin = 72

    doc.Content.Delete()

    # Create: H1(sb=24,sa=0) → Body(sb=0,sa=10) → Body → H1 → Body → Body
    configs = [
        ("Heading", 14, 0, 24, 0),   # H1: fs=14, sa=0, sb=24
        ("Body", 11, 0, 0, 10),       # Body: fs=11, sa=10, sb=0
        ("Body", 11, 0, 0, 10),
        ("Body", 11, 0, 0, 10),
        ("Heading", 14, 0, 24, 0),
        ("Body", 11, 0, 0, 10),
        ("Body", 11, 0, 0, 10),
        ("Body", 11, 0, 0, 10),
        ("Body", 11, 0, 0, 10),
        ("Heading", 13, 0, 10, 0),   # H2: fs=13, sa=0, sb=10
        ("Body", 11, 0, 0, 10),
        ("Body", 11, 0, 0, 10),
    ]

    for i, (kind, fs, sb_override, sb, sa) in enumerate(configs):
        if i > 0:
            sel = word.Selection
            sel.EndKey(6)
            sel.TypeParagraph()

        p = doc.Paragraphs(i + 1)
        p.Range.Text = f"{kind} {i+1} text content here"
        p.Range.Font.Name = "ＭＳ 明朝" if kind == "Body" else "ＭＳ ゴシック"
        p.Range.Font.Size = fs
        p.Format.SpaceBefore = sb
        p.Format.SpaceAfter = sa
        p.Format.LineSpacingRule = 5  # Multiple
        p.Format.LineSpacing = 13.8  # 1.15x

    # Measure all Y positions and gaps
    prev_y = None
    prev_sa = 0
    for i in range(1, len(configs) + 1):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        fs = p.Range.Font.Size
        sa = p.Format.SpaceAfter
        sb = p.Format.SpaceBefore
        gap = y - prev_y if prev_y else 0
        collapsed = max(prev_sa, sb) if prev_y else 0
        advance = gap - collapsed if collapsed > 0 and gap > collapsed else gap

        kind = configs[i-1][0]
        print(f"  P{i:2d} {kind:7s} fs={fs:4.0f} y={y:7.1f} gap={gap:5.1f} "
              f"sa={sa:.0f} sb={sb:.0f} collapsed={collapsed:.0f} advance={advance:.1f}")
        prev_y = y
        prev_sa = sa

    doc.Close(False)

    # ===== TEST 2: Same test with contextualSpacing =====
    print()
    print("=" * 60)
    print("TEST 2: Same layout WITH contextualSpacing on Body")
    print("=" * 60)

    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.LayoutMode = 0
    ps.TopMargin = 72
    doc.Content.Delete()

    for i, (kind, fs, sb_override, sb, sa) in enumerate(configs):
        if i > 0:
            sel = word.Selection
            sel.EndKey(6)
            sel.TypeParagraph()

        p = doc.Paragraphs(i + 1)
        p.Range.Text = f"{kind} {i+1} text content here"
        p.Range.Font.Name = "ＭＳ 明朝" if kind == "Body" else "ＭＳ ゴシック"
        p.Range.Font.Size = fs
        p.Format.SpaceBefore = sb
        p.Format.SpaceAfter = sa
        p.Format.LineSpacingRule = 5
        p.Format.LineSpacing = 13.8

    # Can't set contextualSpacing via COM ParagraphFormat.
    # Skip this test for now.
    doc.Close(False)

    # ===== TEST 3: Heading advance = per-line or cumul? =====
    print()
    print("=" * 60)
    print("TEST 3: Heading line height isolation (no spacing)")
    print("=" * 60)

    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.LayoutMode = 0
    ps.TopMargin = 72
    doc.Content.Delete()

    # All paragraphs: sa=0, sb=0, Multiple 1.15x
    # Pattern: body, body, heading, body, body
    sizes = [11, 11, 14, 11, 11, 11, 13, 11, 11]
    for i, fs in enumerate(sizes):
        if i > 0:
            sel = word.Selection
            sel.EndKey(6)
            sel.TypeParagraph()
        p = doc.Paragraphs(i + 1)
        p.Range.Text = f"Line {i+1} fs={fs}"
        p.Range.Font.Name = "ＭＳ 明朝"
        p.Range.Font.Size = fs
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 5
        p.Format.LineSpacing = 13.8

    prev_y = None
    for i in range(1, len(sizes) + 1):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        fs = p.Range.Font.Size
        gap = y - prev_y if prev_y else 0
        print(f"  P{i} fs={fs:4.0f} y={y:7.1f} gap={gap:5.1f}")
        prev_y = y

    doc.Close(False)

    # ===== TEST 4: Spacing collapse formula =====
    print()
    print("=" * 60)
    print("TEST 4: Spacing collapse (sa/sb combinations)")
    print("=" * 60)

    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.LayoutMode = 0
    ps.TopMargin = 72
    doc.Content.Delete()

    # All body 11pt, Multiple 1.15x, varying sa/sb
    spacing_tests = [
        (0, 0),   # no spacing
        (10, 0),  # sa=10
        (0, 10),  # sb=10
        (10, 24), # sa=10, next sb=24 → collapsed=24
        (15, 24), # sa=15, next sb=24 → collapsed=24
        (0, 0),   # back to normal
    ]

    for i, (sa, sb) in enumerate(spacing_tests):
        if i > 0:
            sel = word.Selection
            sel.EndKey(6)
            sel.TypeParagraph()
        p = doc.Paragraphs(i + 1)
        p.Range.Text = f"Para {i+1} sa={sa} sb={sb}"
        p.Range.Font.Name = "ＭＳ 明朝"
        p.Range.Font.Size = 11
        p.Format.SpaceAfter = sa
        p.Format.SpaceBefore = sb
        p.Format.LineSpacingRule = 5
        p.Format.LineSpacing = 13.8

    prev_y = None
    prev_sa = 0
    for i in range(1, len(spacing_tests) + 1):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        sa = p.Format.SpaceAfter
        sb = p.Format.SpaceBefore
        gap = y - prev_y if prev_y else 0
        expected_collapsed = max(prev_sa, sb)
        print(f"  P{i} y={y:7.1f} gap={gap:5.1f} sa={sa:.0f} sb={sb:.0f} "
              f"max(prev_sa={prev_sa:.0f},sb={sb:.0f})={expected_collapsed:.0f}")
        prev_y = y
        prev_sa = sa

    doc.Close(False)

finally:
    word.Quit()
