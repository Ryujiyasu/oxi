"""Measure empty paragraph line heights in 1ec1 document.
Focus on blocks 2-25 (empty paragraphs used as TextBox anchors).
Check if wrapNone TextBox anchors affect paragraph height."""
import win32com.client
import os
import time

docx_path = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    ps = doc.Sections(1).PageSetup
    print(f"Page: {ps.PageWidth:.1f}pt x {ps.PageHeight:.1f}pt")
    print(f"Margins: T={ps.TopMargin:.1f}, B={ps.BottomMargin:.1f}")

    # Build a set of paragraphs that are TextBox anchors
    anchor_paras = set()
    for i in range(1, doc.Shapes.Count + 1):
        shp = doc.Shapes(i)
        anchor_range = shp.Anchor
        # Find which paragraph this anchor belongs to
        for j in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(j)
            if p.Range.Start <= anchor_range.Start <= p.Range.End:
                anchor_paras.add(j)
                break

    print(f"\nTextBox anchor paragraphs: {sorted(anchor_paras)}")

    print(f"\n{'Para#':<6} {'Page':<5} {'Y(pt)':<10} {'LineSpacing':<12} {'SpaceBefore':<12} {'SpaceAfter':<11} {'IsAnchor':<9} {'LineRule':<10} {'Text'}")
    print("-" * 120)

    prev_y = None
    for i in range(1, min(doc.Paragraphs.Count + 1, 48)):
        para = doc.Paragraphs(i)
        rng = para.Range

        # Page and Y position
        word.Selection.SetRange(rng.Start, rng.Start)
        y_pos = float(word.Selection.Information(6))  # wdVerticalPositionRelativeToPage
        page = rng.Information(3)

        # Paragraph formatting
        pf = para.Format
        line_spacing = pf.LineSpacing
        space_before = pf.SpaceBefore
        space_after = pf.SpaceAfter
        line_rule = pf.LineSpacingRule
        # 0=wdLineSpaceSingle, 1=wdLineSpace1pt5, 2=wdLineSpaceDouble,
        # 3=wdLineSpaceAtLeast, 4=wdLineSpaceExactly, 5=wdLineSpaceMultiple

        rule_names = {0: "Single", 1: "1.5", 2: "Double", 3: "AtLeast", 4: "Exactly", 5: "Multiple"}
        rule_str = rule_names.get(line_rule, str(line_rule))

        is_anchor = "ANCHOR" if i in anchor_paras else ""

        text = rng.Text[:40].replace('\r', '\\r').replace('\n', '\\n')
        is_empty = rng.Text.strip() in ('', '\r', '\r\n')

        # Delta from previous paragraph
        delta = ""
        if prev_y is not None:
            delta = f" (delta={y_pos - prev_y:.2f})"

        print(f"{i:<6} {page:<5} {y_pos:<10.2f} {line_spacing:<12.2f} {space_before:<12.2f} {space_after:<11.2f} {is_anchor:<9} {rule_str:<10} {'[empty]' if is_empty else text}{delta}")

        prev_y = y_pos

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
