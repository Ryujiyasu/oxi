"""COM-measure per-character X positions for 15076df L12 paragraph.

15076df L12 = `１．提供を受けた匿名データの名称` in a narrow table cell
(tcW=1968dxa=98.4pt, cellMar=12dxa=0.6pt, indent left=215tw right=76tw
hanging=192tw, fs=10.5pt).

LLA diff shows:
  Word fits  L12: `１．提供を受けた匿名` (10 chars) / L13: `データの名称` (4 chars)
  Oxi  fits  L12: `１．提供を受けた匿`   (9 chars)  / L13: `名データの名称` (5 chars)

Question: what is the exact x position of each char in Word? Specifically:
  - Where does Word place 「名」 on line 1? (Oxi puts it on line 2)
  - Does it overshoot the indent_right boundary?
  - What is Word's effective right-edge for this paragraph?

This pins down whether the bug is:
  (a) Oxi computes a smaller available_width than Word
  (b) Oxi's wrap policy is more conservative for the same width
  (c) Char width is computed differently (e.g. "．" half-width)
"""
import os
import sys
import io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.normpath(os.path.join(
    REPO, "tools/golden-test/documents/docx/15076df085f5_tokumei_08_09.docx"))

print(f"DOCX: {DOCX}")

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    # Walk paragraphs and find the one containing '提供を受けた匿名データの名称'
    target_text = "提供を受けた匿名データの名称"
    target_para = None
    target_idx = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        if target_text in p.Range.Text:
            target_para = p
            target_idx = i
            break
    if target_para is None:
        print("FAIL: target paragraph not found")
        sys.exit(1)

    p = target_para
    txt = p.Range.Text
    print(f"\nFound L12 paragraph at index {target_idx}: {txt[:50]!r}")
    print(f"Length: {p.Range.End - p.Range.Start} chars")

    wdHorizontal = 5  # wdHorizontalPositionRelativeToPage (pt)
    wdVertical = 6    # wdVerticalPositionRelativeToPage (pt)

    rng_start = p.Range.Start
    rng_end = p.Range.End

    print(f"\nPer-character positions (Information(5)/(6) in pt):")
    print(f"{'i':>3} {'X(pt)':>9} {'Y(pt)':>9}  {'advance':>9}  char  on_line")
    prev_x = None
    prev_y = None
    line_num = 1
    chars = []
    for i in range(rng_start, rng_end):
        r = doc.Range(i, i)  # collapsed-start range (R30 fix)
        x = r.Information(wdHorizontal)
        y = r.Information(wdVertical)
        # next char text
        nr = doc.Range(i, i + 1)
        ch = nr.Text
        if prev_x is not None:
            if y != prev_y:
                line_num += 1
                adv = None
                adv_str = "  newline"
            else:
                adv = x - prev_x
                adv_str = f"{adv:+9.3f}"
        else:
            adv_str = "       -"
        line_str = f"L{line_num}"
        print(f"{i - rng_start:>3} {x:>9.3f} {y:>9.3f}  {adv_str}  {ch!r:6}  {line_str}")
        chars.append({"i": i - rng_start, "x": x, "y": y, "ch": ch, "line": line_num})
        prev_x = x
        prev_y = y

    # Summary: line break boundaries
    print(f"\n=== Line break summary ===")
    lines = {}
    for c in chars:
        lines.setdefault(c["line"], []).append(c)
    for lno, lchars in sorted(lines.items()):
        text = "".join(c["ch"] for c in lchars).rstrip("\r\n")
        x_start = lchars[0]["x"]
        x_end = lchars[-1]["x"]
        print(f"  L{lno}: x={x_start:.2f}..{x_end:.2f}, {len(lchars)} chars: {text!r}")

    # Identify the char-width pattern
    print(f"\n=== Per-char advance widths (same-line only) ===")
    for j in range(1, len(chars)):
        prev_ch = chars[j - 1]["ch"]
        if chars[j]["line"] == chars[j - 1]["line"]:
            adv = chars[j]["x"] - chars[j - 1]["x"]
            print(f"  {prev_ch!r}: advance = {adv:.3f}pt")
finally:
    doc.Close(SaveChanges=False)
    word.Quit()
