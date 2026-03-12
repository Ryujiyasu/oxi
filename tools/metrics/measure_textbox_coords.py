"""Measure TextBox positions via Word COM to determine coordinate system.
Checks whether Shape.Left/Top are page-relative or document-global.
Also reports the anchor paragraph's page and position."""
import win32com.client
import os
import time

docx_path = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx")

print(f"Opening: {docx_path}")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    # Page setup
    ps = doc.Sections(1).PageSetup
    print(f"\nPage: {ps.PageWidth:.1f}pt x {ps.PageHeight:.1f}pt")
    print(f"Margins: T={ps.TopMargin:.1f}, B={ps.BottomMargin:.1f}, L={ps.LeftMargin:.1f}, R={ps.RightMargin:.1f}")

    print(f"\nTotal Shapes (InlineShapes): {doc.InlineShapes.Count}")
    print(f"Total Shapes (Shapes): {doc.Shapes.Count}")

    print(f"\n{'='*90}")
    print(f"{'#':<4} {'Name':<20} {'Type':<6} {'Left':>8} {'Top':>8} {'Width':>8} {'Height':>8} {'AnchorPage':>10} {'AnchorY':>10}")
    print(f"{'-'*90}")

    for i in range(1, doc.Shapes.Count + 1):
        shp = doc.Shapes(i)
        name = shp.Name
        shp_type = shp.Type  # 17=TextBox(msoTextBox), 1=AutoShape, etc.

        left = shp.Left
        top = shp.Top
        width = shp.Width
        height = shp.Height

        # Get anchor paragraph info
        anchor_range = shp.Anchor
        anchor_page = anchor_range.Information(3)  # wdActiveEndPageNumber

        # Get anchor paragraph's vertical position on page
        word.Selection.SetRange(anchor_range.Start, anchor_range.Start)
        anchor_y = word.Selection.Information(6)  # wdVerticalPositionRelativeToPage

        # Wrap type
        wrap_type = shp.WrapFormat.Type  # 3=wdWrapNone, 0=wdWrapInline, etc.
        wrap_names = {0: "Inline", 1: "TopBottom", 2: "Around", 3: "None", 4: "Tight", 5: "Through", 6: "Left", 7: "Right"}
        wrap_str = wrap_names.get(wrap_type, str(wrap_type))

        # RelativeVerticalPosition
        # 0=wdRelativeVerticalPositionMargin, 1=wdRelativeVerticalPositionPage, 2=wdRelativeVerticalPositionParagraph
        rel_vert = shp.RelativeVerticalPosition
        rel_names = {0: "margin", 1: "page", 2: "paragraph", 3: "line", 4: "topMarginArea", 5: "bottomMarginArea", 6: "insideMargin", 7: "outsideMargin"}
        rel_str = rel_names.get(rel_vert, str(rel_vert))

        rel_horiz = shp.RelativeHorizontalPosition
        rel_h_names = {0: "margin", 1: "page", 2: "column", 3: "character", 4: "leftMarginArea", 5: "rightMarginArea", 6: "insideMargin", 7: "outsideMargin"}
        rel_h_str = rel_h_names.get(rel_horiz, str(rel_horiz))

        print(f"{i:<4} {name:<20} {shp_type:<6} {left:>8.1f} {top:>8.1f} {width:>8.1f} {height:>8.1f} {anchor_page:>10} {anchor_y:>10.1f}")
        print(f"     Wrap={wrap_str}, RelVert={rel_str}, RelHoriz={rel_h_str}")

        # Text content preview
        if shp.TextFrame.HasText:
            text = shp.TextFrame.TextRange.Text[:60].replace('\r', '\\r')
            print(f"     Text: \"{text}\"")

    # Also measure paragraph positions for cross-reference
    print(f"\n{'='*90}")
    print(f"Paragraph positions (first 20):")
    print(f"{'Para#':<6} {'Page':<5} {'Y(pt)':<10} {'Text preview'}")
    print(f"{'-'*60}")

    for i in range(1, min(doc.Paragraphs.Count + 1, 25)):
        para = doc.Paragraphs(i)
        rng = para.Range
        page_num = rng.Information(3)
        word.Selection.SetRange(rng.Start, rng.Start)
        y_pos = float(word.Selection.Information(6))
        text = rng.Text[:50].replace('\r', '\\r').replace('\n', '\\n')
        print(f"{i:<6} {page_num:<5} {y_pos:<10.2f} {text}")

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
