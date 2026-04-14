"""
COM measurement: MS Mincho/Gothic half-width character widths in twips.
These UPM=256 fonts have actual advance widths that differ from fontSize/2.
"""
import win32com.client
import json
import os
import sys

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Create a temporary document with known text
    doc = word.Documents.Add()

    results = {}

    # Test chars: ASCII printable (32-126)
    test_chars = [chr(i) for i in range(32, 127)]

    for font_name in ["ＭＳ 明朝", "ＭＳ ゴシック"]:
        for font_size in [9.0, 10.5, 11.0, 12.0, 14.0]:
            family_key = "MS Mincho" if "明朝" in font_name else "MS Gothic"
            size_key = f"{font_size:.1f}" if font_size != int(font_size) else f"{int(font_size)}"

            overrides = {}

            for ch in test_chars:
                # Clear document
                doc.Content.Delete()

                # Insert test string: same char repeated to avoid autoSpaceDE
                # Use "{ch}{ch}{ch}" pattern - 3 copies, measure middle one
                test_text = f"{ch}{ch}{ch}"
                doc.Content.InsertAfter(test_text)

                # Set font
                doc.Content.Font.Name = font_name
                doc.Content.Font.Size = font_size

                # Measure x positions
                r0 = doc.Range(0, 1)  # あ
                r1 = doc.Range(1, 2)  # test char
                r2 = doc.Range(2, 3)  # い

                x0 = r0.Information(5)
                x1 = r1.Information(5)
                x2 = r2.Information(5)

                char_width_pt = x2 - x1
                char_width_tw = char_width_pt * 20

                # Only record if different from fontSize/2
                expected = font_size / 2.0
                if abs(char_width_pt - expected) > 0.01:
                    cp = ord(ch)
                    overrides[cp] = round(char_width_tw, 1)
                    if ch.isprintable():
                        print(f"  {family_key} {size_key}pt U+{cp:04X} '{ch}': "
                              f"{char_width_tw:.1f}tw = {char_width_pt:.2f}pt "
                              f"(expected {expected:.2f}pt, diff={char_width_pt-expected:+.2f})")

            if overrides:
                results.setdefault(family_key, {})[size_key] = overrides
                print(f"{family_key} {size_key}pt: {len(overrides)} overrides")
            else:
                print(f"{family_key} {size_key}pt: all match fontSize/2")

    doc.Close(False)
    word.Quit()

    # Save results
    out_path = "tools/metrics/output/ms_mincho_halfwidth_tw.json"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w') as f:
        json.dump(results, f, indent=2)
    print(f"\nSaved to {out_path}")

if __name__ == "__main__":
    main()
