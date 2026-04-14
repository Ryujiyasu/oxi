"""Systematic COM measurement of charGrid line break spec.
Creates minimal repro documents to isolate each variable."""
import win32com.client
import math

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

def create_chargrid_doc(word, font_name, font_size, char_space=0, indent_left=0, first_line_indent=0):
    """Create a minimal charGrid document with controlled parameters."""
    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.PageHeight = 841.9  # A4
    ps.PageWidth = 595.3
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LeftMargin = 85
    ps.RightMargin = 85
    # Content width = 595.3 - 170 = 425.3pt

    # Set docGrid type=linesAndChars via XML manipulation is complex.
    # Instead set LayoutMode and CharacterUnitFirstLineIndent.
    # Actually we need to set the grid via section properties.
    # Use a workaround: modify the section via COM.

    # For now, let's measure against the existing d1e8ac document
    # and create simple test cases with Add.
    doc.Close(False)
    return None

try:
    # ===== TEST 1: charGrid with fontSize > pitch =====
    # Use d1e8ac directly and measure all affected paragraphs
    import os, glob
    docx = os.path.abspath('tools/golden-test/documents/docx/d1e8ac8fd1cc_kyodokenkyuyoushiki06.docx')
    doc = word.Documents.Open(docx, ReadOnly=True)

    ps = doc.Sections(1).PageSetup
    content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    chars_line = ps.CharsLine
    pitch = content_w / chars_line

    print("=" * 60)
    print(f"TEST 1: charGrid line break (d1e8ac)")
    print(f"Content width: {content_w:.2f}pt")
    print(f"CharsLine: {chars_line:.0f}")
    print(f"Char pitch: {pitch:.4f}pt")
    print("=" * 60)

    # Measure ALL paragraphs: chars per line, indent, fontSize
    for i in range(1, min(doc.Paragraphs.Count + 1, 30)):
        p = doc.Paragraphs(i)
        rng = p.Range
        page = rng.Information(3)
        if page > 1:
            break

        y = rng.Information(6)
        fs = rng.Font.Size
        li = p.Format.LeftIndent
        fi = p.Format.FirstLineIndent

        # Count chars on first line
        start = rng.Start
        start_y = rng.Information(6)
        line1_chars = 0
        for offset in range(0, 80):
            r = doc.Range(start + offset, start + offset + 1)
            if r.Text == '\r' or r.Text == '\x07':
                line1_chars = offset
                break
            cy = r.Information(6)
            if offset > 0 and cy != start_y:
                line1_chars = offset
                break
        else:
            line1_chars = offset + 1

        effective_w = content_w - li
        max_cells = math.floor(effective_w / pitch)
        max_cells_raw = math.floor(content_w / pitch)
        indent_cells = round(li / pitch) if li > 0.5 else 0

        if line1_chars > 0 and fs > 0:
            print(f"P{i:2d} y={y:7.1f} fs={fs:4.1f} li={li:5.1f} fi={fi:6.1f} L1={line1_chars:2d}ch "
                  f"max_cells={max_cells_raw:.0f}-{indent_cells}={max_cells_raw-indent_cells:.0f} "
                  f"{'OK' if line1_chars == max_cells_raw - indent_cells else 'DIFF'}")

    doc.Close(False)

    # ===== TEST 2: charGrid with different fontSize/pitch combinations =====
    print()
    print("=" * 60)
    print("TEST 2: charGrid minimal repro (different fs/pitch combos)")
    print("=" * 60)

    # Create document with charGrid via modifying existing template
    # Actually, let's use COM to create a charGrid document
    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.PageHeight = 841.9
    ps.PageWidth = 595.3
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LeftMargin = 85
    ps.RightMargin = 85

    # Set charGrid
    ps.LayoutMode = 2  # wdLayoutModeLineGrid=1, wdLayoutModeGenko=2
    # LayoutMode 2 enables linesAndChars grid

    content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    print(f"Content width: {content_w:.2f}pt")
    print(f"CharsLine: {ps.CharsLine:.0f}")
    print(f"Char pitch: {content_w/ps.CharsLine:.4f}pt")

    doc.Content.Delete()

    # Test cases: different font sizes
    test_sizes = [9, 10, 10.5, 11, 12, 14]

    for fs in test_sizes:
        # Add paragraph
        sel = word.Selection
        sel.EndKey(6)  # wdStory
        if doc.Paragraphs.Count > 1 or doc.Content.Text.strip():
            sel.TypeParagraph()

        p_idx = doc.Paragraphs.Count
        p = doc.Paragraphs(p_idx)
        # Fill with 80 'あ' characters
        p.Range.Text = chr(0x3042) * 80
        p.Range.Font.Name = "ＭＳ 明朝"
        p.Range.Font.Size = fs
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LeftIndent = 0
        p.Format.FirstLineIndent = 0

    # Measure
    pitch = content_w / ps.CharsLine
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        fs = rng.Font.Size
        start = rng.Start
        start_y = rng.Information(6)

        # Count chars on first line
        line1_chars = 0
        for offset in range(0, 81):
            r = doc.Range(start + offset, start + offset + 1)
            if r.Text == '\r':
                line1_chars = offset
                break
            cy = r.Information(6)
            if offset > 0 and cy != start_y:
                line1_chars = offset
                break

        expected = math.floor(content_w / pitch)
        print(f"  fs={fs:5.1f} pitch={pitch:.2f} L1={line1_chars:2d}ch "
              f"expected={expected:.0f} "
              f"{'OK' if line1_chars == expected else f'DIFF(got {line1_chars})'}")

    doc.Close(False)

    # ===== TEST 3: charGrid with indent =====
    print()
    print("=" * 60)
    print("TEST 3: charGrid + indent combinations")
    print("=" * 60)

    doc = word.Documents.Add()
    ps = doc.Sections(1).PageSetup
    ps.PageHeight = 841.9
    ps.PageWidth = 595.3
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LeftMargin = 85
    ps.RightMargin = 85
    ps.LayoutMode = 2

    content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    pitch = content_w / ps.CharsLine
    chars_line = ps.CharsLine
    print(f"Content width: {content_w:.2f}pt, CharsLine={chars_line:.0f}, pitch={pitch:.4f}pt")

    doc.Content.Delete()

    # Test indent combinations with fs=11 (> pitch)
    indent_tests = [
        (0, 0, "no indent"),
        (11, 0, "left=11 (1cell)"),
        (22, 0, "left=22 (2cells)"),
        (11, -11, "hanging (left=11,fi=-11)"),
        (22, -11, "left=22,fi=-11"),
        (0, 11, "firstLine=11"),
    ]

    for li, fi, label in indent_tests:
        sel = word.Selection
        sel.EndKey(6)
        if doc.Paragraphs.Count > 1 or doc.Content.Text.strip():
            sel.TypeParagraph()

        p_idx = doc.Paragraphs.Count
        p = doc.Paragraphs(p_idx)
        p.Range.Text = chr(0x3042) * 80
        p.Range.Font.Name = "ＭＳ 明朝"
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LeftIndent = li
        p.Format.FirstLineIndent = fi

    # Measure
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        li = p.Format.LeftIndent
        fi = p.Format.FirstLineIndent
        start = rng.Start
        start_y = rng.Information(6)

        # Count chars on first line
        line1_chars = 0
        for offset in range(0, 81):
            r = doc.Range(start + offset, start + offset + 1)
            if r.Text == '\r':
                line1_chars = offset
                break
            cy = r.Information(6)
            if offset > 0 and cy != start_y:
                line1_chars = offset
                break

        # Count chars on second line
        line2_chars = 0
        if line1_chars < 80:
            for offset in range(line1_chars, 81):
                r = doc.Range(start + offset, start + offset + 1)
                if r.Text == '\r':
                    line2_chars = offset - line1_chars
                    break
                cy = r.Information(6)
                if offset > line1_chars and cy != doc.Range(start + line1_chars, start + line1_chars + 1).Information(6):
                    line2_chars = offset - line1_chars
                    break

        indent_cells = round(li / pitch) if li > 0.5 else 0
        fi_cells = round(fi / pitch) if abs(fi) > 0.5 else 0
        net_first = chars_line - indent_cells + fi_cells
        net_subseq = chars_line - indent_cells

        label = indent_tests[i-1][2] if i <= len(indent_tests) else "?"
        print(f"  {label:25s} li={li:5.1f} fi={fi:6.1f} L1={line1_chars:2d}ch L2={line2_chars:2d}ch "
              f"expected_L1={net_first:.0f} expected_L2={net_subseq:.0f}")

    doc.Close(False)

finally:
    word.Quit()
