import win32com.client as w, json, os, pythoncom
pythoncom.CoInitialize()
app=w.Dispatch('Word.Application'); app.Visible=False
path=os.path.abspath('tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx')
doc=app.Documents.Open(path, ReadOnly=True)
out=[]
n=doc.Paragraphs.Count
for i in range(1,n+1):
    p=doc.Paragraphs(i); rng=p.Range
    c=w.constants
    st=doc.Range(rng.Start,rng.Start)
    y=st.Information(6)   # wdVerticalPositionRelativeToPage
    pg=st.Information(3)  # wdActiveEndPageNumber
    txt=rng.Text.strip()[:18]
    fnt=rng.Font.Name; sz=rng.Font.Size
    out.append((i,pg,round(y,2),fnt,sz,txt))
doc.Close(False); app.Quit()
json.dump(out, open('tools/metrics/_s459_lod_word.json','w',encoding='utf-8'), ensure_ascii=False)
# print pitch within pages
prev=None
for i,pg,y,fnt,sz,txt in out:
    d=''
    if prev and prev[1]==pg: d=f'pitch={y-prev[2]:+.2f}'
    if pg in (3,4,5):
        print(f'p{pg} #{i} y={y:.2f} {d} sz={sz} {fnt[:10]} | {txt}')
    prev=(i,pg,y)
