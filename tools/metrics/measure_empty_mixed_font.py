"""Measure empty paragraph line height with mixed fonts (Century + MS Mincho).

docDefaults: ascii=Century eastAsia=MS Mincho, Normal sz=10.5pt.
docGrid lines linePitch=360.
Word P1(empty) → P2 gap = 15.5pt. Why not 18pt (grid snap)?
"""
import win32com.client
import time

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Create fresh document
    doc = word.Documents.Add()
    time.sleep(1)

    # Set page setup to match outline_08
    ps = doc.PageSetup
    ps.TopMargin = 70.9

    # Set docGrid
    # wdLineGridHeight = 1 (lines mode)
    doc.PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid

    sel = word.Selection

    # Test 1: Century 10.5pt empty paragraph (default docGrid)
    sel.Font.Name = "Century"
    sel.Font.Size = 10.5
    sel.TypeText("Century 10.5pt line 1")
    sel.TypeParagraph()
    sel.TypeParagraph()  # empty
    sel.TypeText("Century 10.5pt line 3")
    sel.TypeParagraph()

    # Test 2: MS Mincho 10.5pt
    sel.Font.Name = "ＭＳ 明朝"
    sel.Font.Size = 10.5
    sel.TypeText("MS Mincho 10.5pt line")
    sel.TypeParagraph()
    sel.TypeParagraph()  # empty
    sel.TypeText("MS Mincho 10.5pt line")

    time.sleep(1)

    paras = doc.Paragraphs
    print(f"LayoutMode: {doc.PageSetup.LayoutMode}")
    print(f"LinePitch: {doc.Sections(1).PageSetup.CharsLine}")
    print(f"Total paragraphs: {paras.Count}")
    for i in range(1, paras.Count + 1):
        p = paras(i)
        r = p.Range
        y = r.Information(6)
        fn = r.Font.Name
        sz = r.Font.Size
        ls = p.Format.LineSpacing
        text = r.Text[:30].replace('\r','')
        print(f"P{i}: y={y:7.2f} font={fn:15s} sz={sz} ls={ls} text='{text}'")

    print("\nGaps:")
    for i in range(1, paras.Count):
        y1 = paras(i).Range.Information(6)
        y2 = paras(i+1).Range.Information(6)
        print(f"  P{i}->P{i+1}: {y2-y1:.2f}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
