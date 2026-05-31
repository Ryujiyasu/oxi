import win32com.client as w, os, pythoncom
pythoncom.CoInitialize()
app=w.Dispatch('Word.Application'); app.Visible=False
doc=app.Documents.Open(os.path.abspath('tools/golden-test/documents/docx/de6e32b5960b_tokumei_08_01-1.docx'), ReadOnly=True)
# page setup
ps=doc.Sections(1).PageSetup
pw=ps.PageWidth; lm=ps.LeftMargin; rm=ps.RightMargin
print(f'page_width={pw:.1f}pt left_margin={lm:.1f} right_margin={rm:.1f} content_w={pw-lm-rm:.1f}pt')
for i in range(1,doc.Paragraphs.Count+1):
    p=doc.Paragraphs(i); rng=p.Range
    st=doc.Range(rng.Start,rng.Start)
    if st.Information(3)!=7: continue
    txt=rng.Text.replace('\r','').replace('\x07','')
    if len(txt.strip())<5: continue
    pf=p.Format
    li=pf.LeftIndent; ri=pf.RightIndent; fli=pf.FirstLineIndent
    # number of lines: compute via ComputeStatistics or line range
    # use the paragraph's vertical extent
    nchar=len(txt.strip())
    fs=rng.Font.Size
    print(f'#{i} nchar={nchar} fs={fs} LI={li:.1f} RI={ri:.1f} FLI={fli:.1f} | {txt.strip()[:38]}')
doc.Close(False); app.Quit()
