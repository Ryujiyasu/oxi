"""Measure 1ec1 paragraph Y positions to find layout drift."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(0.5)

WD_Y = 6
WD_IN_TABLE = 12

print(f"Paragraphs: {doc.Paragraphs.Count}")
for pi in range(1, min(30, doc.Paragraphs.Count + 1)):
    p = doc.Paragraphs(pi)
    y = p.Range.Information(WD_Y)
    in_tbl = p.Range.Information(WD_IN_TABLE)
    txt = p.Range.Text[:40].replace('\r', '\\r').replace('\x07', '\\BEL')
    marker = " [TBL]" if in_tbl else ""
    fs = p.Range.Font.Size
    ls = p.Format.LineSpacing
    sa = p.Format.SpaceAfter
    sb = p.Format.SpaceBefore
    print(f"  P{pi}: y={y:.2f}{marker} fs={fs} ls={ls:.1f} sa={sa:.1f} sb={sb:.1f} '{txt}'")

doc.Close(SaveChanges=False)
word.Quit()
