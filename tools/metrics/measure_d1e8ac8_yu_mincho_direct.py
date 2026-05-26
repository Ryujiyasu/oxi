"""Directly measure d1e8ac8 paragraph Y/Font via COM, per-paragraph.

S322 finding said: 殿 paragraph (IR para 5, XML p6 = 6th <w:p>) has
first run with eastAsiaTheme="minorEastAsia" → theme resolves to
Yu Mincho. The paragraph renders at 14.5pt in Word.

Verify directly via COM:
  - Each paragraph's Y stride to next paragraph
  - Each paragraph's font (Range.Font.Name + Font.NameAscii etc.)

This isolates the hypothesis: does Word actually use Yu Mincho for
this paragraph or does Word use something else?
"""
import os
import win32com.client

DOC_PATH = os.path.join(
    os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "d1e8ac8fd1cc_kyodokenkyuyoushiki06.docx",
)

WD_VERT_POS_REL_TO_PAGE = 6


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        path = os.path.abspath(DOC_PATH)
        print(f"Opening: {path}")
        doc = word.Documents.Open(path, ReadOnly=True)
        n = doc.Paragraphs.Count
        print(f"Total paragraphs: {n}")
        # We're interested in paragraphs around the 殿 area.
        # From IoU data: i=5 wy=154 (国税庁長官), i=6 wy=168 (殿),
        # i=7 wy=182.5 (empty), i=8 wy=195.5 (私は).
        # So paragraphs 5..8 (1-indexed COM) span this area.
        print(f"\n{'idx':<4}{'y':<10}{'h_stride':<12}{'font_name':<25}{'font_ea':<25}{'text':<40}")
        prev_y = None
        for i in range(1, min(n+1, 15)):
            rng = doc.Paragraphs(i).Range
            collapsed = doc.Range(rng.Start, rng.Start)
            y = collapsed.Information(WD_VERT_POS_REL_TO_PAGE)
            font_name = rng.Font.Name or ""
            font_ascii = rng.Font.NameAscii or ""
            font_ea = rng.Font.NameFarEast or ""
            text = (rng.Text or "").rstrip("\r\n").strip()[:35]
            stride = (y - prev_y) if prev_y is not None else 0.0
            print(f"{i:<4}{y:<10.2f}{stride:<12.2f}{font_name:<25}{font_ea:<25}{text!r}")
            prev_y = y
        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
