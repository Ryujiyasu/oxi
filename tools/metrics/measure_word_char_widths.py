"""Measure Word's actual character placement widths via COM.

Creates a document with test strings and measures character x positions.
This gives the EXACT widths Word uses internally (which differ from GDI ABC widths).

Usage: python measure_word_char_widths.py <font_name> <font_sizes> [output_json]
Example: python measure_word_char_widths.py Calibri 11,14,26 output.json
"""
import win32com.client
import pythoncom
import json
import sys
import os

# ASCII printable + common punctuation
TEST_CHARS = ''.join(chr(c) for c in range(32, 127))  # space to ~


def measure_font_widths(font_name, font_sizes):
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    result = {}

    try:
        for fs in font_sizes:
            doc = word.Documents.Add()
            sel = word.Selection
            sel.Font.Name = font_name
            sel.Font.Size = fs

            # Type test string: pairs of characters to measure width
            # For each char, type "X<char>X" and measure positions
            widths = {}
            for ch in TEST_CHARS:
                if ch == '\r' or ch == '\n':
                    continue
                # Clear document
                doc.Content.Delete()
                sel = word.Selection
                sel.Font.Name = font_name
                sel.Font.Size = fs

                # Type: A + char + A (to avoid edge effects)
                test = 'H' + ch + 'H'
                sel.TypeText(test)

                # Measure character positions
                rng = doc.Paragraphs(1).Range
                chars = rng.Characters
                if chars.Count >= 3:
                    x1 = chars(1).Information(5)  # 'H' position
                    x2 = chars(2).Information(5)  # target char position
                    x3 = chars(3).Information(5)  # 'H' after
                    char_width_pt = round(x3 - x2, 2)
                    widths[ch] = char_width_pt

            ppem = round(fs * 96 / 72)
            result[str(ppem)] = widths
            print(f"  {font_name} {fs}pt (ppem={ppem}): {len(widths)} chars measured")
            doc.Close(False)

    finally:
        word.Quit()

    return result


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python measure_word_char_widths.py <font_name> <sizes> [output.json]")
        sys.exit(1)

    font_name = sys.argv[1]
    sizes = [float(s) for s in sys.argv[2].split(',')]
    output = sys.argv[3] if len(sys.argv) > 3 else f"pipeline_data/word_widths_{font_name}.json"

    print(f"Measuring {font_name} at sizes {sizes}...")
    result = measure_font_widths(font_name, sizes)

    # Convert to pixel widths for storage
    pixel_result = {}
    for ppem_str, widths in result.items():
        ppem = int(ppem_str)
        px_widths = {}
        for ch, pt_w in widths.items():
            px_w = round(pt_w * 96 / 72)
            px_widths[str(ord(ch))] = px_w
        pixel_result[ppem_str] = px_widths

    with open(output, 'w') as f:
        json.dump({"font": font_name, "widths": pixel_result}, f, indent=2)
    print(f"Saved to {output}")
