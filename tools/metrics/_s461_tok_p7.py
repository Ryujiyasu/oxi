import win32com.client as w, json, os, pythoncom
pythoncom.CoInitialize()
app=w.Dispatch('Word.Application'); app.Visible=False
path=os.path.abspath('tools/golden-test/documents/docx/de6e32b5960b_tokumei_08_01-1.docx')
doc=app.Documents.Open(path, ReadOnly=True)
out=[]
for i in range(1,doc.Paragraphs.Count+1):
    p=doc.Paragraphs(i); rng=p.Range
    st=doc.Range(rng.Start,rng.Start)
    pg=st.Information(3); y=st.Information(6)
    if pg==7:
        out.append((i,round(y,2),rng.Font.Size,rng.Text.strip()[:20]))
doc.Close(False); app.Quit()
print('Word p7 paragraphs (idx, y, sz, text):')
prev=None
for i,y,sz,t in out:
    d=f'{y-prev:+.2f}' if prev is not None else ''
    print(f'  #{i} y={y:.2f} gap={d:>7} sz={sz} | {t}')
    prev=y
