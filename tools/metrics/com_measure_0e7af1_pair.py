"""COM-measure 0e7af1's 、（ pair to determine if Word compresses it.

Strategy: scan paragraphs for one containing 、（ adjacent pair, walk
char-by-char via Range.Information(wdHorizontalPositionRelativeToPage),
report the advance widths of "、" and "（".

If Word's "、" before "（" has width ≈ 5.25pt (50% compressed): Word
compresses → R7.37 was correct, the regression is page-boundary
tipping (other docs would also regress on small horizontal shifts).

If Word's "、" stays ≈ 10.5pt fullwidth: Word does NOT compress here →
R7.37 incorrectly applied compression → need more nuanced gate.
"""
import os
import sys
import win32com.client

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.normpath(os.path.join(
    REPO, "tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"))

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    target = '、（'
    target_para_i = None
    for i in range(1, doc.Paragraphs.Count + 1):
        if target in doc.Paragraphs(i).Range.Text:
            target_para_i = i
            break
    if target_para_i is None:
        print("No paragraph with 、（ found", file=sys.stderr)
        sys.exit(1)

    p = doc.Paragraphs(target_para_i)
    text = p.Range.Text
    print(f"Paragraph {target_para_i}: {text[:100]!r}")

    # Find position of '、' followed by '（'
    rel_idx = text.find('、（')
    print(f"  '、（' at char offset {rel_idx} in paragraph")

    wdH = 5
    wdV = 6
    base = p.Range.Start

    # Sample positions: prev_char, '、', '（', next_char
    sample = []
    for k in range(rel_idx - 1, rel_idx + 4):
        pos = base + k
        r = doc.Range(pos, pos + 1)
        ch = r.Text
        x = doc.Range(pos, pos).Information(wdH)
        y = doc.Range(pos, pos).Information(wdV)
        sample.append((k, pos, ch, x, y))
        print(f"  k={k} char_pos={pos} ch={ch!r} x={x:.2f} y={y:.2f}")

    # Compute advances
    print("\nAdvance widths:")
    for j in range(1, len(sample)):
        ch_prev = sample[j-1][2]
        x_prev = sample[j-1][3]
        x_curr = sample[j][3]
        y_prev = sample[j-1][4]
        y_curr = sample[j][4]
        if y_prev == y_curr:
            print(f"  {ch_prev!r} → advance {x_curr - x_prev:+.3f}pt")
        else:
            print(f"  {ch_prev!r} → newline (Y changed)")

finally:
    doc.Close(SaveChanges=False)
    word.Quit()
