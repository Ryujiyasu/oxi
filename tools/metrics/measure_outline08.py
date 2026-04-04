"""Measure line heights in d77a58485f16 (outline_08).

Key question: does grid snap apply? Word P1→P2 = 15.5pt (not 18pt grid snap).
"""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "d77a58485f16_20240705_resources_data_outline_08.docx"))

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    ps = doc.Paragraphs
    print(f"Total paragraphs: {ps.Count}")
    print(f"Page setup: top={doc.PageSetup.TopMargin}")

    for i in range(1, min(20, ps.Count + 1)):
        p = ps(i)
        r = p.Range
        y = r.Information(6)
        page = r.Information(3)
        fmt = p.Format
        ls = fmt.LineSpacing
        lr = fmt.LineSpacingRule
        try:
            snap = p.Range.ParagraphFormat.NoLineNumber  # placeholder
            snap = "?"
        except:
            snap = "?"
        sz = r.Font.Size
        fn = r.Font.Name
        text = r.Text[:25].replace('\r','').replace('\x07','')
        print(f"P{i:2d}: pg={page} y={y:7.2f} sz={sz:4.1f} font={fn:15s} ls={ls:5.1f} lr={lr} snap={snap} text={text}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
