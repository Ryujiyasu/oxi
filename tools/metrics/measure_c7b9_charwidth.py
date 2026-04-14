"""Measure per-char advances for c7b9 P4 (W=77ch, O=45ch).
Century/MS Mincho 10.5pt, grid=lines (no charGrid).
"""
import win32com.client
import os
import time

DOCX = os.path.join(os.path.dirname(__file__), "..", "golden-test", "documents", "docx",
                     "c7b923e5c616_20240705_resources_data_outline_06.docx")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    time.sleep(1)

    sec = doc.Sections(1)
    ps = sec.PageSetup
    cw = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    print(f"Content width: {cw:.2f}pt")

    # Find P4 (y~130, ~77ch)
    target = None
    for i in range(1, 20):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        text = p.Range.Text.rstrip("\r\n")
        if len(text) > 60 and abs(y - 130) < 5:
            target = i
            break
        if len(text) > 100 and target is None:
            target = i

    if not target:
        target = 5  # fallback

    p = doc.Paragraphs(target)
    text = p.Range.Text.rstrip("\r\n")
    y = p.Range.Information(6)
    print(f"\nP{target}: y={y:.1f}, len={len(text)}")
    print(f"text: {text[:100]}")
    print(f"font: {p.Range.Font.Name} {p.Range.Font.Size}pt")

    start = p.Range.Start
    prev_x = None
    line_break_pos = None
    advances = {}
    for j in range(min(len(text), 90)):
        rng = doc.Range(start + j, start + j + 1)
        try:
            cx = rng.Information(5)
            char = text[j]
            advance = cx - prev_x if prev_x is not None and j > 0 else 0
            if prev_x is not None and cx < prev_x - 10:
                if line_break_pos is None:
                    line_break_pos = j
                    print(f"\n--- LINE BREAK at char {j} ---\n")
            cp = ord(char)
            is_latin = cp < 0x80
            is_fullwidth_digit = 0xFF10 <= cp <= 0xFF19
            cat = "LATIN" if is_latin else ("FWDIGIT" if is_fullwidth_digit else "CJK")
            if j < 15 or (j > 40 and j < 55) or (line_break_pos and j >= line_break_pos and j < line_break_pos + 5):
                print(f"  [{j:3d}] '{char}' (U+{cp:04X}) {cat:8s} x={cx:.2f} adv={advance:.2f}")
            # Collect advance stats by category
            if advance > 0:
                advances.setdefault(cat, []).append(advance)
            prev_x = cx
        except Exception as e:
            print(f"  [{j:3d}] error: {e}")
            break

    print(f"\nLine break position: {line_break_pos}")
    print(f"\nAdvance statistics:")
    for cat, vals in sorted(advances.items()):
        avg = sum(vals) / len(vals) if vals else 0
        print(f"  {cat}: n={len(vals)}, avg={avg:.2f}, min={min(vals):.2f}, max={max(vals):.2f}")

    # Also measure P9 (W=217ch, O=53ch) - longest mismatch
    for i in range(1, 30):
        p2 = doc.Paragraphs(i)
        y2 = p2.Range.Information(6)
        text2 = p2.Range.Text.rstrip("\r\n")
        if len(text2) > 200 and abs(y2 - 274) < 5:
            print(f"\n=== P{i} (long paragraph): y={y2:.1f}, len={len(text2)} ===")
            start2 = p2.Range.Start
            prev_x2 = None
            for j in range(min(len(text2), 60)):
                rng2 = doc.Range(start2 + j, start2 + j + 1)
                cx2 = rng2.Information(5)
                char2 = text2[j]
                adv2 = cx2 - prev_x2 if prev_x2 is not None and j > 0 else 0
                cp2 = ord(char2)
                is_latin2 = cp2 < 0x80
                cat2 = "LAT" if is_latin2 else "CJK"
                if j < 10 or (j > 48 and j < 58):
                    print(f"  [{j:3d}] '{char2}' (U+{cp2:04X}) {cat2} x={cx2:.2f} adv={adv2:.2f}")
                prev_x2 = cx2
            break

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
