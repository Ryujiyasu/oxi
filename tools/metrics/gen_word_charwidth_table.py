"""Generate character width table by measuring Word COM character positions.

For UPM=256 fonts (MS Mincho, MS Gothic), Word uses GDI hinting widths
that differ from GetCharWidth32. This script measures actual placement
widths by inserting characters into a document and reading X positions.

Output: JSON with {font: {ppem: {codepoint: width_px}}}
"""
import win32com.client
import os, time, json, sys

# Characters to measure: common CJK + hiragana + katakana + symbols
# Only fullwidth chars that might differ from fontSize
HIRAGANA = [chr(c) for c in range(0x3041, 0x3097)]  # 86 chars
KATAKANA = [chr(c) for c in range(0x30A1, 0x30FB)]  # 90 chars
COMMON_CJK = [chr(c) for c in range(0x4E00, 0x4E00 + 500)]  # first 500 CJK
SYMBOLS = [chr(c) for c in range(0x2460, 0x2474)]  # circled numbers ①-⑳
FULLWIDTH = [chr(c) for c in range(0xFF01, 0xFF5F)]  # fullwidth ASCII
PUNCTUATION = [chr(c) for c in range(0x3000, 0x3040)]  # CJK punctuation

ALL_CHARS = HIRAGANA + KATAKANA + COMMON_CJK + SYMBOLS + FULLWIDTH + PUNCTUATION


def measure_widths(word, font_name, font_size, chars_to_test):
    """Measure character widths via Word COM X position differences."""
    # Create a temp document
    doc = word.Documents.Add()
    time.sleep(0.5)

    # Set font
    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = font_size

    # Insert chars in batches (one char per line to avoid justify interference)
    # Use pairs: reference char + test char, measure X diff
    # Better: insert all chars on one left-aligned paragraph
    batch_size = 100
    results = {}

    for batch_start in range(0, len(chars_to_test), batch_size):
        batch = chars_to_test[batch_start:batch_start + batch_size]

        # Clear document
        doc.Content.Delete()

        # Insert batch as single paragraph, left-aligned
        text = ''.join(batch)
        rng = doc.Range()
        rng.InsertAfter(text)
        rng = doc.Range()
        rng.Font.Name = font_name
        rng.Font.Size = font_size

        # Set left alignment (no justify stretching)
        for p in range(1, doc.Paragraphs.Count + 1):
            doc.Paragraphs(p).Alignment = 0

        time.sleep(0.05)

        # Measure X positions
        chars = doc.Range().Characters
        prev_x = None
        prev_ch = None

        for i in range(1, min(chars.Count + 1, len(batch) + 2)):
            try:
                c = chars(i)
                ch = c.Text
                if ch in ('\r', '\x07', '\n'):
                    continue
                cx = c.Information(5)  # wdHorizontalPositionRelativeToPage

                if prev_x is not None and prev_ch is not None:
                    w = cx - prev_x
                    if 0 < w < 50:  # reasonable width
                        cp = ord(prev_ch)
                        width_px = round(w * 96 / 72)  # pt -> px
                        if cp not in results:
                            results[cp] = width_px

                prev_x = cx
                prev_ch = ch
            except:
                pass

    doc.Close(SaveChanges=False)
    return results


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    fonts_and_sizes = [
        ("ＭＳ 明朝", [7, 7.5, 8, 8.5, 9, 9.5, 10, 10.5, 11, 12, 14, 16]),
        ("ＭＳ ゴシック", [7, 7.5, 8, 8.5, 9, 9.5, 10, 10.5, 11, 12, 14, 16]),
    ]

    all_results = {}

    for font_name, sizes in fonts_and_sizes:
        print(f"\n=== {font_name} ===")
        font_results = {}

        for fs in sizes:
            ppem = round(fs * 96 / 72)
            print(f"  {fs}pt (ppem={ppem})...", end="", flush=True)

            widths = measure_widths(word, font_name, fs, ALL_CHARS)

            # Count how many differ from fontSize (fullwidth = fontSize)
            expected_px = round(fs * 96 / 72)
            diff_count = sum(1 for w in widths.values() if w != expected_px)

            # Filter: only keep chars that differ from the expected fullwidth size
            narrower = {cp: w for cp, w in widths.items() if w < expected_px}
            wider = {cp: w for cp, w in widths.items() if w > expected_px}

            print(f" {len(widths)} chars, {len(narrower)} narrower, {len(wider)} wider")

            if narrower:
                font_results[str(ppem)] = {
                    'expected_px': expected_px,
                    'narrower': narrower,
                    'wider': wider,
                    'total_measured': len(widths),
                }

                # Show sample narrower chars
                for cp, w in list(narrower.items())[:5]:
                    try:
                        print(f"    U+{cp:04X} '{chr(cp)}': {w}px (expected {expected_px}px)")
                    except:
                        print(f"    U+{cp:04X}: {w}px (expected {expected_px}px)")

        all_results[font_name] = font_results

    word.Quit()

    out = "tools/metrics/output/word_charwidth_table.json"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
