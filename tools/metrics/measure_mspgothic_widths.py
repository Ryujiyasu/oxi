"""Measure MS PGothic character widths via Word COM for common CJK codepoints.

GDI GetCharABCWidths returns monospaced widths (14px) for most ideographs,
but Word actually uses proportional widths. Measure via COM Information(5)
on unjustified lines to get the true advance widths.

Output: com_tw_overrides.json additions for MS PGothic
"""
import win32com.client
import os
import json
import time
import tempfile

def create_test_doc(word, font_name, font_size, chars_per_line=10):
    """Create a test document with single characters per line for width measurement."""
    doc = word.Documents.Add()

    # Set up: no justify, no grid, simple layout
    sec = doc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72  # 1 inch
    ps.RightMargin = 72
    ps.TopMargin = 72

    # Remove document grid
    # Can't easily set docGrid via COM, but jc=left is enough

    return doc

def measure_char_widths(word, font_name, font_size_pt, codepoints):
    """Measure advance widths for a list of codepoints.

    Strategy: Create a document with pairs of characters "XA" where X is the
    measured char and A is a reference. Measure X position and A position,
    advance = A_x - X_x.
    """
    doc = word.Documents.Add()

    # Set page to left-aligned, no grid effects
    sec = doc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72

    results = {}
    batch_size = 50  # characters per measurement batch

    for batch_start in range(0, len(codepoints), batch_size):
        batch = codepoints[batch_start:batch_start + batch_size]

        # Clear document
        doc.Content.Delete()

        # Insert each char on its own line: "X\n" pattern
        # with left alignment and specific font
        text_parts = []
        for cp in batch:
            ch = chr(cp)
            text_parts.append(ch + "A")  # char + reference 'A'

        # Join with paragraph breaks
        full_text = "\r".join(text_parts)

        rng = doc.Range(0, 0)
        rng.Text = full_text
        rng.Font.Name = font_name
        rng.Font.Size = font_size_pt
        rng.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
        rng.ParagraphFormat.SpaceAfter = 0
        rng.ParagraphFormat.SpaceBefore = 0

        time.sleep(0.5)

        # Measure each character
        pos = 0
        for i, cp in enumerate(batch):
            try:
                # Position of the measured char
                rng1 = doc.Range(pos, pos + 1)
                x1 = rng1.Information(5)

                # Position of the reference 'A'
                rng2 = doc.Range(pos + 1, pos + 2)
                x2 = rng2.Information(5)

                advance_pt = x2 - x1
                # Convert to twips for storage
                advance_tw = round(advance_pt * 20)

                results[cp] = advance_tw

                if i < 5 or i % 100 == 0:
                    print(f"  U+{cp:04X} ({chr(cp)}): {advance_pt:.2f}pt = {advance_tw}tw")

                pos += 3  # char + 'A' + \r
            except Exception as e:
                print(f"  U+{cp:04X}: error {e}")
                pos += 3

    doc.Close(False)
    return results

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # CJK codepoint ranges to measure
    # Focus on the most common characters in our test documents
    codepoints = []

    # Hiragana (U+3040-U+309F)
    codepoints.extend(range(0x3041, 0x3097))

    # Katakana (U+30A0-U+30FF)
    codepoints.extend(range(0x30A1, 0x30FB))
    codepoints.extend(range(0x30FC, 0x3100))

    # CJK Punctuation (U+3000-U+303F)
    codepoints.extend(range(0x3000, 0x3040))

    # Fullwidth forms (U+FF00-U+FF5F)
    codepoints.extend(range(0xFF01, 0xFF5F))

    # Common CJK Unified Ideographs (U+4E00-U+9FFF)
    # Too many to measure all, pick the most common ~2000
    # Start with the ones that appear in our test documents
    common_kanji = list(range(0x4E00, 0x4E60)) + list(range(0x5000, 0x5100)) + \
                   list(range(0x5200, 0x5300)) + list(range(0x5400, 0x5500)) + \
                   list(range(0x5600, 0x5700)) + list(range(0x5800, 0x5900)) + \
                   list(range(0x5B00, 0x5C00)) + list(range(0x5E00, 0x5F00)) + \
                   list(range(0x6000, 0x6200)) + list(range(0x6300, 0x6600)) + \
                   list(range(0x6700, 0x6A00)) + list(range(0x6B00, 0x6E00)) + \
                   list(range(0x6F00, 0x7200)) + list(range(0x7300, 0x7600)) + \
                   list(range(0x7800, 0x7B00)) + list(range(0x7C00, 0x7F00)) + \
                   list(range(0x8000, 0x8300)) + list(range(0x8400, 0x8700)) + \
                   list(range(0x8800, 0x8B00)) + list(range(0x8C00, 0x8F00)) + \
                   list(range(0x9000, 0x9300)) + list(range(0x9400, 0x9700)) + \
                   list(range(0x9800, 0x9B00)) + list(range(0x9C00, 0x9FFF))
    codepoints.extend(common_kanji)

    # Remove duplicates and sort
    codepoints = sorted(set(codepoints))
    print(f"Total codepoints to measure: {len(codepoints)}")

    # Measure at 10.5pt (ppem=14) - most common size
    font_name = "\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af"  # ＭＳ Ｐゴシック
    font_size = 10.5

    print(f"\nMeasuring {font_name} {font_size}pt...")
    results = measure_char_widths(word, font_name, font_size, codepoints)

    print(f"\nMeasured {len(results)} characters")

    # Compare with GDI table
    with open('crates/oxidocs-core/src/font/data/gdi_width_overrides.json', encoding='utf-8') as f:
        gdi = json.load(f)

    pg14 = gdi.get('MS PGothic', {}).get('14', {})

    mismatches = 0
    for cp, tw in results.items():
        gdi_px = pg14.get(str(cp))
        if gdi_px is not None:
            gdi_tw = round(float(gdi_px) * 72 / 96 * 20)
            if abs(tw - gdi_tw) > 1:  # more than 1 twip difference
                mismatches += 1
                if mismatches <= 20:
                    print(f"  MISMATCH U+{cp:04X} ({chr(cp)}): COM={tw}tw GDI={gdi_tw}tw (px={gdi_px})")

    print(f"\nTotal mismatches: {mismatches}/{len(results)}")

    # Save results
    out = {"MS PGothic": {"10.5": {}}}
    for cp, tw in results.items():
        out["MS PGothic"]["10.5"][str(cp)] = tw / 20.0  # store as pt

    out_path = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data",
                           "mspgothic_com_widths.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_path}")

    word.Quit()

if __name__ == "__main__":
    main()
