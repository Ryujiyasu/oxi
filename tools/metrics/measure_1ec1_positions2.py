"""Measure paragraph Y positions using Range.Information constants."""
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

    # Try using Window.GetPoint which returns pixel coordinates
    wnd = word.ActiveWindow
    
    print(f"{'Para#':<6} {'Page':<5} {'TopPx':<10} {'LeftPx':<10} {'Text'}")
    print("-" * 80)

    for i in range(1, min(doc.Paragraphs.Count + 1, 35)):
        para = doc.Paragraphs(i)
        rng = para.Range
        page_num = rng.Information(3)  # wdActiveEndPageNumber
        
        # Try getting position via selection and pane
        word.Selection.SetRange(rng.Start, rng.Start)
        
        # wdVerticalPositionRelativeToPage = 10 returns line number, not points
        # Let's try wdVerticalPositionRelativeToTextBoundary = 12
        # and other constants
        info_vals = {}
        for const_id in [5, 6, 9, 10, 11, 12, 13]:
            try:
                val = word.Selection.Information(const_id)
                info_vals[const_id] = val
            except:
                info_vals[const_id] = "ERR"
        
        text = rng.Text[:30].replace('\r', '\r').replace('\n', '\n')
        
        if i <= 5 or i in [27, 28, 29]:
            print(f"Para {i}: page={page_num}")
            for k, v in info_vals.items():
                print(f"  Info({k}) = {v}")
            print(f"  Text: {text}")
            print()

    doc.Close(False)
finally:
    word.Quit()
    print("Done.")
