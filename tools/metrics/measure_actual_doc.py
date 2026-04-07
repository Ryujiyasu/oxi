"""Measure actual Y positions in gen2_068_SOW_Template.docx"""
import win32com.client
import time, os

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    docx = os.path.abspath("tools/golden-test/documents/docx/gen2_068_SOW_Template.docx")
    doc = word.Documents.Open(docx, ReadOnly=True)
    time.sleep(1)

    try:
        sec = doc.Sections(1)
        lm = sec.PageSetup.LayoutMode
        print(f"LayoutMode: {lm}")

        # Measure table row Y positions
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            print(f"\nTable {ti} ({tbl.Rows.Count} rows):")
            for r in range(1, min(tbl.Rows.Count + 1, 8)):
                try:
                    y = tbl.Cell(r, 1).Range.Information(6)
                    text = tbl.Cell(r, 1).Range.Text.strip()[:30]
                    print(f"  Row {r}: y={y:.2f}pt, text='{text}'")
                except:
                    print(f"  Row {r}: error")

        # First 10 body paragraphs
        print(f"\nParagraphs (total {doc.Paragraphs.Count}):")
        for i in range(1, min(doc.Paragraphs.Count + 1, 15)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:40]
            page = p.Range.Information(3)  # wdActiveEndPageNumber
            if text:
                print(f"  P{i}: y={py:.2f}pt, page={page}, text='{text}'")

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
