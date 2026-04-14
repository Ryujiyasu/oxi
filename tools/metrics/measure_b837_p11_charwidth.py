"""Measure per-character advance widths for b837 P11 (71 chars in Word, 38 in Oxi).
linesAndChars grid, MS Gothic 12pt.
"""
import win32com.client
import os
import json
import time

DOCX = os.path.join(os.path.dirname(__file__), "..", "golden-test", "documents", "docx",
                     "b837808d0555_20240705_resources_data_guideline_02.docx")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    time.sleep(1)

    # Page setup
    sec = doc.Sections(1)
    ps = sec.PageSetup
    cw = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    print(f"Content width: {cw:.2f}pt")
    print(f"CharsLine: {ps.CharsLine}")
    print(f"GridDistanceHorizontal (COM): {sec.PageSetup.GridDistanceHorizontal if hasattr(sec.PageSetup, 'GridDistanceHorizontal') else 'N/A'}")

    # Find paragraph 11 (0-indexed = DML P11, approximately the one with ~71ch on L1)
    # From DML: P11 y=232, text starts with Japanese + has mixed content
    # The paragraphs with indent (\u3000) are body paragraphs
    target = None
    for i in range(1, min(30, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        text = p.Range.Text.rstrip("\r\n")
        y = p.Range.Information(6)
        if abs(y - 232.0) < 2.0 and len(text) > 50:
            target = i
            print(f"\nFound target P{i}: y={y:.1f}, len={len(text)}")
            print(f"  text: {text[:80]}")
            break

    if not target:
        # Try a broader search
        for i in range(10, min(20, doc.Paragraphs.Count + 1)):
            p = doc.Paragraphs(i)
            text = p.Range.Text.rstrip("\r\n")
            y = p.Range.Information(6)
            if len(text) > 60:
                target = i
                print(f"\nUsing P{i}: y={y:.1f}, len={len(text)}")
                print(f"  text: {text[:80]}")
                break

    if target:
        p = doc.Paragraphs(target)
        text = p.Range.Text.rstrip("\r\n")
        start = p.Range.Start

        print(f"\n=== Per-char advances for P{target} ({len(text)} chars) ===")
        prev_x = None
        advances = []
        for j in range(min(len(text), 80)):
            rng = doc.Range(start + j, start + j + 1)
            try:
                cx = rng.Information(5)  # horizontal pos
                char = text[j]
                advance = cx - prev_x if prev_x is not None and j > 0 else 0
                # Detect line break (x resets)
                if prev_x is not None and cx < prev_x - 10:
                    print(f"  --- LINE BREAK at char {j} ---")
                advances.append({"pos": j, "char": char, "x": round(cx, 2), "adv": round(advance, 2)})
                print(f"  [{j:3d}] '{char}' (U+{ord(char):04X}) x={cx:.2f} adv={advance:.2f}")
                prev_x = cx
            except Exception as e:
                print(f"  [{j:3d}] error: {e}")
                break

    # Also measure a few paragraphs before P11 for reference
    print("\n=== Paragraph Y positions ===")
    for i in range(1, min(20, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        text = p.Range.Text.rstrip("\r\n")
        font = p.Range.Font.Name
        sz = p.Range.Font.Size
        print(f"  P{i:2d}: y={y:.1f} font={font} sz={sz}pt len={len(text)} \"{text[:30]}\"")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
