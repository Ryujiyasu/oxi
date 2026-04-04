"""Check snapToGrid for each paragraph in outline_08."""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "d77a58485f16_20240705_resources_data_outline_08.docx"))

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    ps = doc.Paragraphs
    for i in range(1, min(15, ps.Count + 1)):
        p = ps(i)
        r = p.Range
        y = r.Information(6)
        fmt = p.Format
        # SnapToGrid is on ParagraphFormat in newer Word
        # Try various property names
        snap = "?"
        try:
            snap = fmt.SnapToGrid
        except:
            pass
        # Also check via XML
        try:
            snap2 = fmt.DisableLineHeightGrid
        except:
            snap2 = "?"

        fn = r.Font.Name
        fn_ea = r.Font.NameFarEast
        sz = r.Font.Size
        text = r.Text[:20].replace('\r','')
        style = p.Style.NameLocal
        print(f"P{i:2d}: y={y:7.2f} snap={snap} disGrid={snap2} style={style:10s} font={fn}/{fn_ea} sz={sz} text='{text}'")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
