"""Measure b837 P2 '公共データは...' paragraph line 1 char positions."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    # Find the para starting with '公共データは'
    target_idx = None
    for i in range(1, doc.Paragraphs.Count + 1):
        txt = doc.Paragraphs(i).Range.Text
        if txt.startswith('公共データは'):
            target_idx = i
            print(f"Found at paragraph {i}")
            print(f"Text: {txt[:80]}...")
            break
    if target_idx is None:
        print("Not found")
    else:
        rng = doc.Paragraphs(target_idx).Range
        text = rng.Text
        # Dump first 45 chars with x/y positions
        print("\nChar positions:")
        prev_x = None
        for ci in range(min(50, len(text))):
            c = rng.Characters(ci + 1)
            try:
                x = c.Information(5)  # wdHorizontalPositionRelativeToPage
                y = c.Information(6)  # wdVerticalPositionRelativeToPage
                ln = c.Information(10)  # line number
                ch = text[ci]
                adv = (x - prev_x) if prev_x is not None else None
                print(f"  C{ci:2d} L{ln}: x={x:6.1f} y={y:6.1f} '{ch}' U+{ord(ch):04X}" +
                      (f"  adv={adv:.1f}" if adv is not None else ""))
                prev_x = x
            except Exception as e:
                print(f"  C{ci}: ERR {e}")
                break
finally:
    doc.Close(False)
    word.Quit()
