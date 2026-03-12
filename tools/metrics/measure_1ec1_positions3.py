"""Measure all paragraph Y positions (Info(6) = vertical position relative to page)."""
import win32com.client
import os
import time

docx_path = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx"
assert os.path.exists(docx_path)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    ps = doc.Sections(1).PageSetup
    print(f"Page: {ps.PageWidth:.1f}pt x {ps.PageHeight:.1f}pt")
    print(f"Margins: T={ps.TopMargin:.1f}, B={ps.BottomMargin:.1f}, L={ps.LeftMargin:.1f}, R={ps.RightMargin:.1f}")
    print(f"Total paragraphs: {doc.Paragraphs.Count}\n")

    print(f"{'Para':<5} {'Page':<5} {'Y(pt)':<10} {'X(pt)':<10} {'Text'}")
    print("-" * 80)

    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(i)
        rng = para.Range
        page_num = rng.Information(3)  # wdActiveEndPageNumber
        
        word.Selection.SetRange(rng.Start, rng.Start)
        x = float(word.Selection.Information(5))   # wdHorizontalPositionRelativeToPage
        y = float(word.Selection.Information(6))   # wdVerticalPositionRelativeToPage
        
        text = rng.Text[:35].replace('\r', '\r').replace('\n', '\n')

        print(f"{i:<5} {page_num:<5} {y:<10.2f} {x:<10.2f} {text}")

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
