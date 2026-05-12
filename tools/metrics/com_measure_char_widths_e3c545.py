"""COM-measure per-character X positions in e3c545 w_i=30 paragraph.

Goal: verify whether Word's kerning compresses Latin chars (space, (, 1, ))
in this Meiryo 10.5pt paragraph, freeing up enough space for "い" to fit
on line 1 alongside "ください。".

Method: walk the paragraph character-by-character using Word's Range API,
get each char's `wdHorizontalPositionRelativeToPage`, compute consecutive
diffs (= rendered char advance width).

Compare to Oxi's expected widths from font_metrics_compact.json (Meiryo):
- " " (space): 3.568pt
- "(": 4.614pt
- "1": 6.521pt
- ")": 4.614pt
- Sum (1)+space = 19.317pt

If Word's measured widths are smaller than this, kerning is active.
"""
import os
import sys
import win32com.client

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.normpath(os.path.join(
    REPO, "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx"))

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    p = doc.Paragraphs(30)
    print(f"Paragraph text: {p.Range.Text[:60]!r}")
    print(f"Length: {p.Range.End - p.Range.Start} chars")

    wdHorizontal = 5
    wdVertical = 6

    # Walk char-by-char
    chars = []
    rng_start = p.Range.Start
    rng_end = p.Range.End
    print(f"\nChar positions:")
    print(f"{'i':>3} {'X(pt)':>8} {'Y(pt)':>8}  {'advance':>8}  char")
    prev_x = None
    prev_y = None
    for i in range(rng_start, min(rng_end, rng_start + 60)):
        r = doc.Range(i, i)
        x = r.Information(wdHorizontal)
        y = r.Information(wdVertical)
        # Get the next char's text
        if i + 1 <= rng_end:
            next_r = doc.Range(i, i + 1)
            ch = next_r.Text
        else:
            ch = "(end)"
        if prev_x is not None:
            advance_x = x - prev_x if y == prev_y else None
            adv_str = f"{advance_x:+8.3f}" if advance_x is not None else "  newline"
        else:
            adv_str = "       —"
        print(f"{i-rng_start:>3} {x:>8.2f} {y:>8.2f}  {adv_str}  {ch!r}")
        chars.append({"i": i - rng_start, "x": x, "y": y, "ch": ch})
        prev_x = x
        prev_y = y

    # Identify Latin chars and their advance widths
    print(f"\n=== Latin char advance widths ===")
    for j in range(1, len(chars)):
        prev_ch = chars[j - 1]["ch"]
        if prev_ch.isascii() and prev_ch not in ("\r", "\n"):
            if chars[j]["y"] == chars[j - 1]["y"]:
                adv = chars[j]["x"] - chars[j - 1]["x"]
                print(f"  {prev_ch!r}: advance = {adv:.3f}pt")
finally:
    doc.Close(SaveChanges=False)
    word.Quit()
