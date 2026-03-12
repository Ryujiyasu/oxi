"""Check total pages and shapes that overflow page boundary."""
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
    page_h = ps.PageHeight
    margin_top = ps.TopMargin

    # Total pages
    total_pages = doc.ComputeStatistics(2)  # wdStatisticPages
    print(f"Total pages: {total_pages}")
    print(f"Page height: {page_h:.1f}pt, Margin top: {margin_top:.1f}pt")
    print(f"Total paragraphs: {doc.Paragraphs.Count}")

    print(f"\nShape overflow analysis (anchor_y + top vs page_height):")
    for i in range(1, doc.Shapes.Count + 1):
        shp = doc.Shapes(i)
        anchor_range = shp.Anchor
        word.Selection.SetRange(anchor_range.Start, anchor_range.Start)
        anchor_y = float(word.Selection.Information(6))
        anchor_page = anchor_range.Information(3)

        top = shp.Top
        left = shp.Left
        abs_y = anchor_y + top
        overflow = abs_y > page_h

        # Get the actual rendered page of the shape by checking a point inside it
        # Use Selection approach: select text inside shape, check page
        shape_page = anchor_page
        try:
            if shp.TextFrame.HasText:
                tr = shp.TextFrame.TextRange
                word.Selection.SetRange(tr.Start, tr.Start)
                shape_page = word.Selection.Information(3)
        except:
            pass

        print(f"  Shape {i}: anchor_y={anchor_y:.1f} + top={top:.1f} = abs_y={abs_y:.1f} {'OVERFLOW' if overflow else 'ok'} | anchor_page={anchor_page} shape_text_page={shape_page}")
        if left < -999000:
            print(f"    Left={left:.0f} (alignment sentinel)")

    # Also dump paragraphs around the anchor area (near Y=748)
    print(f"\nParagraphs near bottom of page (Y > 600):")
    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(i)
        rng = para.Range
        word.Selection.SetRange(rng.Start, rng.Start)
        y_pos = float(word.Selection.Information(6))
        page = rng.Information(3)
        if y_pos > 600 or page > 1:
            text = rng.Text[:40].replace('\r', '\\r')
            print(f"  Para {i}: page={page} y={y_pos:.1f} \"{text}\"")

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
