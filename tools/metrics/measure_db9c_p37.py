"""Measure Word P37 (1-indexed) line positions for db9ca18368cd document.
Specifically check if P37 starts on page 2 or page 3."""

import win32com.client
import os
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc_path = os.path.abspath("tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    # P37 (1-indexed)
    p37 = doc.Paragraphs(37)
    rng = p37.Range

    print(f"P37 text: {rng.Text[:80]}")
    print(f"P37 char count: {len(rng.Text)-1}")
    print(f"P37 range: {rng.Start}-{rng.End}")

    # Information at START of paragraph
    start_rng = doc.Range(rng.Start, rng.Start + 1)
    print(f"\nP37 first char:")
    print(f"  char: {repr(start_rng.Text)}")
    print(f"  y: {start_rng.Information(6)}")
    print(f"  x: {start_rng.Information(5)}")
    print(f"  page: {start_rng.Information(3)}")

    # Information at END of paragraph
    end_rng = doc.Range(rng.End - 2, rng.End - 1)
    print(f"\nP37 last char:")
    print(f"  char: {repr(end_rng.Text)}")
    print(f"  y: {end_rng.Information(6)}")
    print(f"  x: {end_rng.Information(5)}")
    print(f"  page: {end_rng.Information(3)}")

    # Check P36 last char
    p36 = doc.Paragraphs(36)
    r36 = p36.Range
    end36 = doc.Range(r36.End - 2, r36.End - 1)
    print(f"\nP36 last char:")
    print(f"  char: {repr(end36.Text)}")
    print(f"  y: {end36.Information(6)}")
    print(f"  page: {end36.Information(3)}")

    # Track line breaks in P37 by character Y positions
    print(f"\nP37 character-level Y positions:")
    para_start = rng.Start
    para_end = rng.End

    prev_y = None
    line_count = 0

    for ci in range(para_start, min(para_end, para_start + 500)):
        cr = doc.Range(ci, ci + 1)
        cy = cr.Information(6)
        cp = cr.Information(3)
        ch = cr.Text

        if prev_y is None or abs(cy - prev_y) > 1.0:
            line_count += 1
            print(f"  Line {line_count}: y={cy:.1f} page={cp} offset={ci-para_start} char={repr(ch)}")
            prev_y = cy

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
