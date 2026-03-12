"""Measure paragraph Y positions in 1ec1 document via Word COM automation."""
import win32com.client
import os
import time

docx_path = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx"
print(f"Opening: {docx_path}")
assert os.path.exists(docx_path), f"File not found: {docx_path}"

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    ps = doc.Sections(1).PageSetup
    print(f"\nPage: {ps.PageWidth}pt x {ps.PageHeight}pt")
    print(f"Margins: T={ps.TopMargin}, B={ps.BottomMargin}, L={ps.LeftMargin}, R={ps.RightMargin}")

    print(f"\nTotal paragraphs: {doc.Paragraphs.Count}")
    print(f"{'Para#':<6} {'Page':<5} {'Top(pt)':<10} {'Left(pt)':<10} {'Text preview'}")
    print("-" * 80)

    for i in range(1, min(doc.Paragraphs.Count + 1, 35)):
        para = doc.Paragraphs(i)
        rng = para.Range
        page_num = rng.Information(3)  # wdActiveEndPageNumber

        top = -1
        left = -1
        try:
            word.Selection.SetRange(rng.Start, rng.Start)
            top = float(word.Selection.Information(10))  # wdVerticalPositionRelativeToPage
            left = float(word.Selection.Information(5))   # wdHorizontalPositionRelativeToPage
        except:
            pass

        text = rng.Text[:40].replace('\r', '\r').replace('\n', '\n')
        if len(rng.Text) > 40:
            text += "..."

        print(f"{i:<6} {page_num:<5} {top:<10.2f} {left:<10.2f} {text}")

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
