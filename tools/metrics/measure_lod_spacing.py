"""COM: Measure LOD_Handbook paragraph spacing."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
path = os.path.abspath("tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

print("LOD_Handbook paragraph spacing:")
for pi in range(1, min(doc.Paragraphs.Count + 1, 30)):
    para = doc.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    if not text:
        continue
    y = rng.Information(6)
    page = rng.Information(3)
    fs = rng.Font.Size
    fn = rng.Font.Name
    ls = para.Format.LineSpacing
    lsr = para.Format.LineSpacingRule
    sb = para.Format.SpaceBefore
    sa = para.Format.SpaceAfter
    rules = {0: 'auto', 1: 'atLeast', 2: 'exactly', 3: '1.5', 4: 'double', 5: 'multiple'}
    print(f"P{pi} pg={page} y={y:.1f} fs={fs} ls={ls:.1f}({rules.get(lsr,'?')}) sb={sb:.1f} sa={sa:.1f} \"{text[:30]}\"")

doc.Close(SaveChanges=False)
word.Quit()
