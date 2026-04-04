"""COM: Check actual chars per line for linesAndChars grid document."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

sec = doc.Sections(1)
ps = sec.PageSetup
print(f"Page: {ps.PageWidth:.1f}x{ps.PageHeight:.1f}")
print(f"Margins: L={ps.LeftMargin:.1f} R={ps.RightMargin:.1f}")
print(f"Content: {ps.PageWidth - ps.LeftMargin - ps.RightMargin:.1f}pt")
print(f"CharsLine: {ps.CharsLine}")

# Check P11-P15 chars per line via Y positions
for pi in [11, 12, 13, 14, 15, 18, 21]:
    para = doc.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')

    # Font info
    fs = rng.Font.Size
    fn = rng.Font.Name

    # Check Y of first few chars and last char
    chars = rng.Characters
    n = chars.Count

    # Sample positions
    positions = {}
    for pos in [1, min(n, 37), min(n, 38), min(n, 50)]:
        try:
            c = chars(pos)
            y = c.Information(6)
            x = c.Information(5)
            positions[pos] = (x, y, c.Text)
        except:
            pass

    print(f"\nP{pi}: {len(text)}ch font={fn} size={fs}")
    print(f"  text: \"{text[:60]}\"")
    for pos, (x, y, ch) in sorted(positions.items()):
        print(f"  char[{pos}] x={x:.1f} y={y:.1f} '{ch}'")

doc.Close(SaveChanges=False)
word.Quit()
