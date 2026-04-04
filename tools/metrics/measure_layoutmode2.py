"""Compare line heights between LayoutMode=1 (lines) and LayoutMode=2 (linesAndChars).

Key question: does LayoutMode=2 change grid snap behavior for non-CJK fonts?
outline_08: type="lines" but Word reports LayoutMode=2, Century 10.5pt gap=15.5pt.
"""
import win32com.client
import time

def test_layout_mode(word, mode, label):
    doc = word.Documents.Add()
    time.sleep(1)

    # Set layout mode
    doc.PageSetup.LayoutMode = mode
    # Set top margin to match outline_08
    doc.PageSetup.TopMargin = 70.9

    sel = word.Selection

    # Test fonts at 10.5pt
    fonts = [
        ("Century", 10.5),
        ("ＭＳ 明朝", 10.5),
        ("ＭＳ ゴシック", 10.5),
        ("Century", 12.0),
        ("ＭＳ ゴシック", 12.0),
    ]

    for fn, sz in fonts:
        sel.Font.Name = fn
        sel.Font.Size = sz
        sel.TypeText(f"{fn} {sz}pt text")
        sel.TypeParagraph()
        sel.TypeParagraph()  # empty
        sel.TypeText(f"{fn} {sz}pt text2")
        sel.TypeParagraph()

    time.sleep(1)

    print(f"\n=== LayoutMode={mode} ({label}) ===")
    print(f"  Actual LayoutMode: {doc.PageSetup.LayoutMode}")

    paras = doc.Paragraphs
    for i in range(1, paras.Count + 1):
        p = paras(i)
        r = p.Range
        y = r.Information(6)
        fn = r.Font.Name
        sz = r.Font.Size
        text = r.Text[:30].replace('\r', '')
        is_empty = len(text.strip()) == 0
        print(f"  P{i:2d}: y={y:7.2f} sz={sz:4.1f} font={fn:15s} empty={is_empty}")

    print(f"\n  Gaps:")
    for i in range(1, paras.Count):
        y1 = paras(i).Range.Information(6)
        y2 = paras(i + 1).Range.Information(6)
        fn = paras(i).Range.Font.Name
        sz = paras(i).Range.Font.Size
        print(f"    P{i:2d}->P{i+1:2d}: gap={y2-y1:5.2f}  ({fn} {sz}pt)")

    doc.Close(False)

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    test_layout_mode(word, 1, "lines")
    test_layout_mode(word, 2, "linesAndChars")

    word.Quit()

if __name__ == "__main__":
    main()
