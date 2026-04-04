"""Test: does Information(6) return per-character Y or per-paragraph Y?"""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

# P196 on page 3 — a body paragraph with font_size=8
# Find paragraph 196
para = doc.Paragraphs(196)
rng = para.Range
text = rng.Text.rstrip('\r')
print(f"P196: {len(text)} chars")
print(f"Para Y (rng.Information(6)): {rng.Information(6):.1f}")

# Check Y for chars at different positions
chars = rng.Characters
n = chars.Count
print(f"Total chars (incl markers): {n}")

# Sample: char 1, 20, 40, 60, 80
for pos in [1, 10, 20, 30, 40, 50, 60, 70, 80, min(n, 84)]:
    try:
        c = chars(pos)
        cy = c.Information(6)
        cx = c.Information(5)
        ch = c.Text
        print(f"  char[{pos}] = '{ch}' x={cx:.1f} y={cy:.1f}")
    except Exception as e:
        print(f"  char[{pos}] error: {e}")

doc.Close(SaveChanges=False)
word.Quit()
