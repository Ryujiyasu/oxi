"""Measure every paragraph gap in LOD_Handbook p1 to find 20.0 vs 20.5 pattern."""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx", "e3c545fac7a7_LOD_Handbook.docx"))

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    ps = doc.Paragraphs
    prev_y = None
    for i in range(1, min(30, ps.Count + 1)):
        p = ps(i)
        r = p.Range
        y = r.Information(6)
        page = r.Information(3)
        if page > 1: break
        gap = y - prev_y if prev_y is not None else 0
        text = r.Text[:20].replace('\r','').replace('\x07','')
        is_empty = len(text.strip()) == 0
        fn = r.Font.Name
        sz = r.Font.Size
        sb = p.Format.SpaceBefore
        sa = p.Format.SpaceAfter
        print(f'P{i:2d} y={y:7.1f} gap={gap:5.1f} empty={is_empty} sb={sb:.1f} sa={sa:.1f} {fn} {sz}pt text={text}')
        prev_y = y

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
