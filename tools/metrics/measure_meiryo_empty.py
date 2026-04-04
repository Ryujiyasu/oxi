"""Measure empty paragraph line height for Meiryo font.

Question: does empty paragraph use floor or ceil for 10tw rounding?
LOD_Handbook P2(empty) → P3 gap = 20.5pt (ceil), but code says floor→20.0pt.
"""
import win32com.client
import time

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Create fresh document with Meiryo empty paragraphs
    doc = word.Documents.Add()
    time.sleep(1)

    # Set font to Meiryo 10.5pt
    sel = word.Selection
    sel.Font.Name = "メイリオ"
    sel.Font.Size = 10.5

    # Type text, empty, text pattern
    sel.TypeText("テスト行1")
    sel.TypeParagraph()
    # Empty paragraph
    sel.TypeParagraph()
    sel.TypeText("テスト行3")
    sel.TypeParagraph()
    sel.TypeText("テスト行4")
    sel.TypeParagraph()
    # Two empty paragraphs
    sel.TypeParagraph()
    sel.TypeParagraph()
    sel.TypeText("テスト行7")

    time.sleep(1)

    ps = doc.Paragraphs
    print(f"Total paragraphs: {ps.Count}")
    for i in range(1, ps.Count + 1):
        p = ps(i)
        r = p.Range
        y = r.Information(6)
        text = r.Text[:30].replace('\r', '').replace('\x07', '')
        is_empty = len(text.strip()) == 0
        print(f"P{i}: y={y:7.2f} empty={is_empty} text='{text}'")

    # Calculate gaps
    print("\nGaps:")
    for i in range(1, ps.Count):
        y1 = ps(i).Range.Information(6)
        y2 = ps(i+1).Range.Information(6)
        gap = y2 - y1
        print(f"  P{i}→P{i+1}: gap={gap:.2f}")

    # Now test with MS Mincho 10.5pt for comparison
    doc2 = word.Documents.Add()
    time.sleep(1)
    sel = word.Selection
    sel.Font.Name = "ＭＳ 明朝"
    sel.Font.Size = 10.5
    sel.TypeText("テスト行1")
    sel.TypeParagraph()
    sel.TypeParagraph()
    sel.TypeText("テスト行3")
    time.sleep(1)

    print("\n--- MS Mincho 10.5pt ---")
    ps2 = doc2.Paragraphs
    for i in range(1, ps2.Count + 1):
        p = ps2(i)
        r = p.Range
        y = r.Information(6)
        text = r.Text[:30].replace('\r', '').replace('\x07', '')
        is_empty = len(text.strip()) == 0
        print(f"P{i}: y={y:7.2f} empty={is_empty}")
    for i in range(1, ps2.Count):
        y1 = ps2(i).Range.Information(6)
        y2 = ps2(i+1).Range.Information(6)
        print(f"  P{i}→P{i+1}: gap={y2-y1:.2f}")

    doc2.Close(False)
    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
