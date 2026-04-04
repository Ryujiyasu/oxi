"""Measure chars per line for b837808d0555 (data_guideline_02).
linesAndChars mode. Compare with Oxi's line char counts."""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "b837808d0555_20240705_resources_data_guideline_02.docx"))

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    print(f"LayoutMode: {doc.PageSetup.LayoutMode}")
    print(f"CharsLine: {doc.PageSetup.CharsLine}")
    print(f"LinesPage: {doc.PageSetup.LinesPage}")

    ps = doc.Paragraphs
    print(f"\nFirst 15 paragraphs:")
    for i in range(1, min(16, ps.Count + 1)):
        p = ps(i)
        r = p.Range
        y = r.Information(6)
        page = r.Information(3)
        text = r.Text.replace('\r','').replace('\x07','')
        text_len = len(text)
        print(f"P{i:2d}: pg={page} y={y:7.2f} chars={text_len:3d} text='{text[:30]}'")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
