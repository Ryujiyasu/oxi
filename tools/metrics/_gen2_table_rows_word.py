import win32com.client as win32
DOCX=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\gen2_lineheight\table_5x3_cambria11.docx"
VPOS=6
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True)
rows=[]
try:
    tbl=doc.Tables(1)
    for r in range(1,tbl.Rows.Count+1):
        cell=tbl.Cell(r,1).Range
        start=doc.Range(cell.Start,cell.Start)
        rows.append(round(start.Information(VPOS),3))
finally:
    doc.Close(False);word.Quit()
print("WORD col1 row tops:",rows)
print("WORD row pitches:",[round(rows[i+1]-rows[i],3) for i in range(len(rows)-1)])
print("OXI  row pitches: [15.16, 15.17, 15.16, 15.16]")
