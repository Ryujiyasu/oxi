import json,glob
import win32com.client as win32
VPOS=6
DOCS={"gen2_025(JP regress)":glob.glob(r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\gen2_025*.docx")[0],
      "gen2_054(EN improve)":r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\gen2_054_Audit_Report.docx"}
DUMPS={"gen2_025(JP regress)":("C:/tmp/s463cmp/off25.json","C:/tmp/s463cmp/on25.json"),
       "gen2_054(EN improve)":("C:/tmp/s463cmp/off54.json","C:/tmp/s463cmp/on54.json")}
def oxi_rowpitch(dump):
    d=json.load(open(dump,encoding="utf-8"))
    rows={}
    for pg in d["pages"]:
        if pg["page"]!=1: continue
        for el in pg["elements"]:
            if el["type"]!="text": continue
            if el.get("cell_col_idx")==0 and el.get("cell_row_idx") is not None:
                r=el["cell_row_idx"]; rows[r]=min(rows.get(r,1e9),el["y"])
    ys=[rows[k] for k in sorted(rows)]
    return [round(ys[i+1]-ys[i],2) for i in range(len(ys)-1)]
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
for name,docx in DOCS.items():
    doc=word.Documents.Open(docx,ReadOnly=True)
    try:
        t=doc.Tables(1); ys=[]
        for r in range(1,min(t.Rows.Count+1,7)):
            c=t.Cell(r,1).Range; st=doc.Range(c.Start,c.Start)
            ys.append(round(st.Information(VPOS),2))
        wp=[round(ys[i+1]-ys[i],2) for i in range(len(ys)-1)]
        cell_txt=t.Cell(1,1).Range.Text.strip()[:14]
    finally:
        doc.Close(False)
    off,on=DUMPS[name]
    print(f"\n{name}  (cell[1,1]='{cell_txt}')")
    print(f"  Word row pitch: {wp}  mean {round(sum(wp)/len(wp),2) if wp else 0}")
    print(f"  Oxi OFF pitch : {oxi_rowpitch(off)}")
    print(f"  Oxi ON  pitch : {oxi_rowpitch(on)}")
word.Quit()
