"""COM-measure ruby_text_01.docx — mixed-font (游ゴシック + eastAsia fallback) line layout."""
import win32com.client
import time
import os
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

path = os.path.abspath("pipeline_data/docx/ruby_text_01.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.4)

ps = doc.PageSetup
print(f"Page={ps.PageWidth:.1f}x{ps.PageHeight:.1f}pt L={ps.LeftMargin:.1f} R={ps.RightMargin:.1f} T={ps.TopMargin:.1f} B={ps.BottomMargin:.1f}")
print(f"body width = {ps.PageWidth - ps.LeftMargin - ps.RightMargin:.2f}pt")
print()

para = doc.Paragraphs(1)
text = para.Range.Text
print(f"Paragraph 1: len={len(text)} text={text[:40]!r}...")

chars = para.Range.Characters
print(f"Characters: count={chars.Count}\n")

# Per-char layout
rows = []
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        rows.append({
            "i": ci,
            "ch": ch,
            "x": c.Information(5),
            "y": c.Information(6),
            "font": c.Font.Name,
            "size": c.Font.Size,
            "ascii": ord(ch) < 128,
        })
    except Exception:
        continue

# Group by line (round y to 1 decimal)
lines = {}
for r in rows:
    lines.setdefault(round(r["y"], 1), []).append(r)

for y in sorted(lines.keys()):
    ln = lines[y]
    first_x = min(r["x"] for r in ln)
    last_x = max(r["x"] for r in ln)
    text_line = "".join(r["ch"] for r in ln)
    fonts = set((r["font"], r["size"]) for r in ln)
    print(f"y={y} chars={len(ln)} first_x={first_x:.2f} last_x={last_x:.2f} width={last_x-first_x:.2f}pt fonts={fonts}")
    print(f"  text: {text_line!r}")
    # Per-char advances by char type
    ln_sorted = sorted(ln, key=lambda r: r["x"])
    print(f"  per-char (ch, font, advance):")
    for i in range(len(ln_sorted) - 1):
        adv = round(ln_sorted[i+1]["x"] - ln_sorted[i]["x"], 3)
        ch = ln_sorted[i]["ch"]
        fn = ln_sorted[i]["font"]
        marker = " <-- ascii" if ord(ch) < 128 else ""
        print(f"    {i}: {ch!r} font={fn!r} adv={adv}{marker}")

doc.Close(SaveChanges=False)
word.Quit()
