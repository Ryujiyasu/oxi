"""Measure paragraph Y positions in 1ec1 document via Word COM automation."""
import win32com.client
import os
import time

docx_path = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\golden-test\golden-test\documents\docx\1ec1091177b1_006.docx")

print(f"Opening: {docx_path}")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    # Get page setup
    ps = doc.Sections(1).PageSetup
    print(f"\nPage: {ps.PageWidth}pt x {ps.PageHeight}pt")
    print(f"Margins: T={ps.TopMargin}, B={ps.BottomMargin}, L={ps.LeftMargin}, R={ps.RightMargin}")

    # Measure each paragraph's position
    print(f"\nTotal paragraphs: {doc.Paragraphs.Count}")
    print(f"{'Para#':<6} {'Page':<5} {'Top(pt)':<10} {'Left(pt)':<10} {'Height(pt)':<11} {'Text preview'}")
    print("-" * 80)

    for i in range(1, min(doc.Paragraphs.Count + 1, 35)):
        para = doc.Paragraphs(i)
        rng = para.Range

        # Get position info via Range
        # wdBoundOfCurrentPage = 1, wdActiveEndPageNumber
        page_num = rng.Information(3)  # wdActiveEndPageNumber

        # Use Window.GetPoint to get screen coordinates, but that's unreliable
        # Instead, measure the paragraph's vertical position via PDF coordinates
        # or use the Range.Information approach

        # Get line position
        top = -1
        left = -1
        try:
            # PageSetup-relative position via Selection
            word.Selection.SetRange(rng.Start, rng.Start)
            top_info = word.Selection.Information(10)  # wdVerticalPositionRelativeToPage
            left_info = word.Selection.Information(5)  # wdHorizontalPositionRelativeToPage
            top = float(top_info)
            left = float(left_info)
        except:
            pass

        text = rng.Text[:40].replace('\r', '\r').replace('\n', '\n')
        if len(rng.Text) > 40:
            text += "..."

        print(f"{i:<6} {page_num:<5} {top:<10.2f} {left:<10.2f} {'':11} {text}")

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
