"""
Measure Word's Meiryo char widths for chars on e3c545 idx 30 line.

Create a test doc with controlled characters in Meiryo 10.5pt and measure
each char's X position to derive actual char widths.
"""
import win32com.client
import os
import time

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Open e3c545 to measure existing chars in-place
    docx_path = os.path.abspath(
        "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx"
    )
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    try:
        # Measure P30 (1-indexed) = the long paragraph that wraps divergently
        para = doc.Paragraphs(30)
        r = para.Range
        text = r.Text
        print(f"Text (len={len(text)}): {repr(text)}")

        # Per-char X positions
        positions = []
        for i in range(len(text)):
            sub = doc.Range(r.Start + i, r.Start + i + 1)
            try:
                x = sub.Information(5)
                y = sub.Information(6)
            except:
                x = None
                y = None
            positions.append((i, text[i], x, y))

        # Print each char with width (delta from previous)
        print("\nPer-char measurements:")
        print(f"{'i':>3s} {'ch':>5s} {'cp':>6s} {'X':>7s} {'Y':>7s} {'dX':>7s}")
        prev_x = None
        prev_y = None
        for (i, ch, x, y) in positions:
            if x is None:
                continue
            cp = f"{ord(ch):04X}"
            if prev_x is None or abs(y - prev_y) > 3:
                dx = 0
                prev_x = x
            else:
                dx = x - prev_x
                prev_x = x
            prev_y = y
            print(f"{i:>3d} {ch!r:>5s} {cp:>6s} {x:>7.2f} {y:>7.2f} {dx:>+7.2f}")
    finally:
        doc.Close(False)
        word.Quit()

if __name__ == "__main__":
    main()
