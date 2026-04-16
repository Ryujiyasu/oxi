"""Measure b837 P1 brackets in 'ネットワーク社会推進戦略本部決定）及び「オープンデータ2.0」...' paragraph."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    # Find 'ネットワーク社会推進戦略本部決定' in doc
    for i in range(1, doc.Paragraphs.Count + 1):
        txt = doc.Paragraphs(i).Range.Text
        if 'ネットワーク社会推進戦略本部決定' in txt and '及び' in txt and 'オープンデータ' in txt:
            print(f"Paragraph {i}: {txt[:80]}...")
            rng = doc.Paragraphs(i).Range
            text = rng.Text
            prev_x = None
            prev_y = None
            for ci in range(len(text)):
                c = rng.Characters(ci + 1)
                try:
                    x = c.Information(5)
                    y = c.Information(6)
                    ln = c.Information(10)
                    ch = text[ci]
                    adv = (x - prev_x) if (prev_x is not None and prev_y == y) else None
                    marker = '  '
                    if ch in '（）「」『』、。' or ord(ch) in (0xFF08, 0xFF09, 0x300C, 0x300D, 0x3001, 0x3002):
                        marker = 'Y:'
                    print(f"  {marker} C{ci:3d} L{ln}: x={x:6.1f} y={y:6.1f} '{ch}' U+{ord(ch):04X}" +
                          (f"  adv={adv:.1f}" if adv is not None else "  (new_line)"))
                    prev_x = x
                    prev_y = y
                except Exception as e:
                    print(f"  C{ci}: ERR {e}")
                    break
            break
finally:
    doc.Close(False)
    word.Quit()
