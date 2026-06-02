import win32com.client as w
import os, json, pythoncom
pythoncom.CoInitialize()
DOC = os.path.abspath(r"tools/golden-test/documents/docx/1636d28e2c46_tokumei_08_04.docx")
app = w.Dispatch("Word.Application"); app.Visible=False
doc = app.Documents.Open(DOC, ReadOnly=True)
ps = doc.Sections(1).PageSetup
LM=float(ps.LeftMargin); TM=float(ps.TopMargin)
print("leftMargin=%.2f topMargin=%.2f"%(LM,TM))
RELH={0:"Margin",1:"Page",2:"Column",3:"Char"}
RELV={0:"Margin",1:"Page",2:"Para",3:"Line"}
for s in doc.Shapes:
    relH=int(s.RelativeHorizontalPosition); relV=int(s.RelativeVerticalPosition)
    L=round(float(s.Left),2); T=round(float(s.Top),2)   # raw, native frame, NOT mutated
    a=s.Anchor; para=a.Paragraphs(1); pstart=para.Range.Start
    rng0=doc.Range(pstart,pstart)
    para_top=round(float(rng0.Information(6)),2)
    cell_content_left=None; cell_top=None
    if para.Range.Information(12):
        try:
            c=para.Range.Cells(1); cr=c.Range
            cs=doc.Range(cr.Start,cr.Start)
            cell_content_left=round(float(cs.Information(5)),2)
            cell_top=round(float(cs.Information(6)),2)
        except Exception as e: pass
    print(json.dumps({"name":s.Name,"relH":RELH.get(relH,relH),"relV":RELV.get(relV,relV),
        "rawLeft":L,"rawTop":T,"para_top":para_top,"cell_cleft":cell_content_left,
        "cell_top":cell_top,"W":round(float(s.Width),1),"H":round(float(s.Height),1)},ensure_ascii=False))
doc.Close(False); app.Quit()
