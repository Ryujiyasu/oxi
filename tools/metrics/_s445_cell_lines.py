"""S445b: count text lines + measure content extent per row5-7 cell of 7ead52.
Want to know: is the cell single-line? What is Word's content height
(line-box top to bottom) that makes atLeast row render 44.25 not 43.0?
"""
import sys, win32com.client as win32
sys.stdout.reconfigure(encoding="utf-8")
DOCX = r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\7ead52b63f0e_000067058.docx"
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    tbl = doc.Tables(1)
    cells = tbl.Range.Cells
    for ci in range(1, cells.Count + 1):
        cell = cells(ci)
        ri = cell.RowIndex; col = cell.ColumnIndex
        if ri < 4:  # focus on the clean 860tw rows (4-8)
            continue
        crng = cell.Range
        # count display lines: iterate paragraphs, each para's line count via
        # Information(10)=wdFirstCharacterLineNumber differences not reliable.
        # Use: para count + char count; and measure first/last line Y.
        startr = doc.Range(crng.Start, crng.Start)
        endr = doc.Range(crng.End - 1, crng.End - 1)
        y0 = float(startr.Information(6))
        y1 = float(endr.Information(6))
        nparas = crng.Paragraphs.Count
        txt = crng.Text.replace("\r", "|").replace("\x07", "")[:30]
        # font size of first run
        try:
            sz = crng.Font.Size
        except Exception:
            sz = "?"
        print(f"r{ri}c{col}: nparas={nparas} y0={y0:.2f} y1={y1:.2f} (span={y1-y0:.2f}) sz={sz} txt='{txt}'")
finally:
    doc.Close(False); word.Quit()
