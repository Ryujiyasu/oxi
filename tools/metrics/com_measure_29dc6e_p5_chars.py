"""Word COM: get the x coordinate of each character of 29dc6e i=265 to
determine the actual wrap point.

This resolves the S300 open question: does Word really wrap at 21 chars/line
(which would imply a ~220pt effective wrap budget — far narrower than the
~342pt cell width budget Oxi uses), or is COM `Information(10)`
(`wdFirstCharacterLineNumber`) returning misleading values?

Per-char measurement plan:
  for each char in i=265:
    rng = Range(start+i, start+i)
    x = Information(2)  # wdHorizontalPositionRelativeToPage
    y = Information(6)  # wdVerticalPositionRelativeToPage
    line = Information(10)  # wdFirstCharacterLineNumber

Output: per-char table of (i, char, x, y, line). Visible wrap point =
where (y, line) jumps to next line.

Run from repo root:
  python tools/metrics/com_measure_29dc6e_p5_chars.py
"""
from __future__ import annotations

import os

DOC = "tools/golden-test/documents/docx/29dc6e8943fe_order_01.docx"


def main() -> None:
    import win32com.client as win32

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
        p265 = doc.Paragraphs(265)
        rng = p265.Range
        # rng.Text includes paragraph mark; we want just the visible chars
        full_text = rng.Text or ""
        # Strip trailing paragraph mark
        text = full_text.rstrip("\r\x07")
        print(f"i=265 text length: {len(text)} visible chars (full Range.Text len {len(full_text)})")
        print()

        # Iterate each visible char
        start = rng.Start
        prev_line = None
        prev_y = None
        line_chars_x = []
        for offset, ch in enumerate(text):
            pos = start + offset
            cr = doc.Range(pos, pos)
            try:
                x = cr.Information(2)  # wdHorizontalPositionRelativeToPage
                y = cr.Information(6)
                line = cr.Information(10)
            except Exception as e:
                print(f"  offset={offset} char={ch!r}: error {e}")
                continue
            # Detect line break
            if prev_line is not None and line != prev_line:
                # Print previous line summary
                if line_chars_x:
                    first_x = line_chars_x[0][1]
                    last_x = line_chars_x[-1][1]
                    print(f"  LINE {prev_line} on y={prev_y:.2f}: "
                          f"{len(line_chars_x)} chars, x range [{first_x:.2f}..{last_x:.2f}], "
                          f"width={last_x-first_x:.2f}pt")
                line_chars_x = []
            line_chars_x.append((offset, x, y, ch))
            prev_line = line
            prev_y = y

        # Last line
        if line_chars_x:
            first_x = line_chars_x[0][1]
            last_x = line_chars_x[-1][1]
            print(f"  LINE {prev_line} on y={prev_y:.2f}: "
                  f"{len(line_chars_x)} chars, x range [{first_x:.2f}..{last_x:.2f}], "
                  f"width={last_x-first_x:.2f}pt")

        print()
        # Also: range start char and end char to confirm
        rs = doc.Range(rng.Start, rng.Start)
        re_ = doc.Range(rng.End - 1, rng.End - 1)
        print(f"i=265 START (Range.Start): x={rs.Information(2):.2f} y={rs.Information(6):.2f} line={rs.Information(10)}")
        print(f"i=265 END   (Range.End-1): x={re_.Information(2):.2f} y={re_.Information(6):.2f} line={re_.Information(10)}")
        doc.Close(SaveChanges=False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
