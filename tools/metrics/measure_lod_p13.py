"""COM: Measure LOD_Handbook P12-P14 spacing details."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
path = os.path.abspath("tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

for pi in range(10, 18):
    para = doc.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    y = rng.Information(6)
    fs = rng.Font.Size
    fn = rng.Font.Name
    ls = para.Format.LineSpacing
    lsr = para.Format.LineSpacingRule
    sb = para.Format.SpaceBefore
    sa = para.Format.SpaceAfter
    style = para.Style.NameLocal if hasattr(para.Style, 'NameLocal') else '?'
    rules = {0:'auto',1:'atLeast',2:'exact',3:'1.5lines',4:'double',5:'multiple'}
    print(f"P{pi}: y={y:.2f} fs={fs} fn={fn} ls={ls:.1f}({rules.get(lsr,'?')}) sb={sb:.1f} sa={sa:.1f} style='{style}' empty={len(text)==0} '{text[:20]}'")

doc.Close(SaveChanges=False)
word.Quit()
