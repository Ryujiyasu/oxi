"""Round 30: Measure b837808d0555 character widths to verify Oxi vs Word.

User observation: Oxi text is WIDER than Word's. Suspected font width discrepancy
on MS Mincho 12pt (the body font for this document).

Test approach:
1. Open the doc in Word, measure each char's advance via Range.Information(5) (X position).
2. For a representative body paragraph, list (char, X) pairs.
3. Compare with Oxi's char-by-char positions extracted via layout_json output.
4. Identify which chars differ and by how much.
"""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True  # Some Information() values only work when Word window is rendered
word.DisplayAlerts = False

doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(0.5)

# Find a substantial body paragraph (after header). Paragraph 5 or 6 is typically body.
print(f"Total paragraphs: {doc.Paragraphs.Count}")
for pi in range(1, min(15, doc.Paragraphs.Count + 1)):
    p = doc.Paragraphs(pi)
    txt = p.Range.Text[:60]
    print(f"  P{pi}: {txt!r}")

# Pick a long body paragraph
TARGET_P = 11  # first body paragraph "我が国においては..."
print(f"\n=== Measuring P{TARGET_P} char positions ===")
p = doc.Paragraphs(TARGET_P)
chars = p.Range.Characters
print(f"  char count: {chars.Count}, text: {p.Range.Text[:80]!r}")
positions = []
for i in range(1, min(50, chars.Count + 1)):
    try:
        c = chars(i)
        x = c.Range.Information(5)
        y = c.Range.Information(6)
        font = c.Font.Name
        sz = c.Font.Size
        positions.append((i, c.Range.Text, x, y, font, sz))
    except Exception as e:
        positions.append((i, "?", None, None, None, None))
        break
for i, ch, x, y, font, sz in positions[:40]:
    if x is None:
        print(f"  {i:3d}: {ch!r:>4} ERR")
    else:
        # Compute width as difference from previous char
        prev_x = positions[i-2][2] if i > 1 and positions[i-2][2] is not None else x
        width = x - prev_x if i > 1 else 0
        print(f"  {i:3d}: {ch!r:>4} x={x:7.2f} y={y:6.2f} font={font:<14} sz={sz:5} delta_x={width:6.2f}")

doc.Close(SaveChanges=False)
word.Quit()
