"""Measure gen2_046_Travel_Report.docx paragraph positions"""
import win32com.client
import time, os

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    docx = os.path.abspath("tools/golden-test/documents/docx/gen2_046_Travel_Report.docx")
    doc = word.Documents.Open(docx, ReadOnly=True)
    time.sleep(1)

    try:
        sec = doc.Sections(1)
        print(f"LayoutMode: {sec.PageSetup.LayoutMode}")
        print(f"TopMargin: {sec.PageSetup.TopMargin:.2f}pt")

        for i in range(1, min(doc.Paragraphs.Count + 1, 8)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:30]
            ls = p.Format.LineSpacing
            lsr = p.Format.LineSpacingRule
            sb = p.Format.SpaceBefore
            sa = p.Format.SpaceAfter
            fn = p.Range.Font.Name
            fs = p.Range.Font.Size
            style = p.Style.NameLocal
            print(f"P{i}: y={py:.2f} ls={ls:.2f} lsr={lsr} sb={sb:.2f} sa={sa:.2f} fn={fn} fs={fs:.1f} style={style}")
            print(f"    text='{text}'")

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
