"""Verify that Word in linesAndChars mode does NOT snap half-width Latin to full cells."""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

import os
path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.5)

# Find a paragraph with mostly Latin and measure char advances
for pi in range(1, 30):
    try:
        p = doc.Paragraphs(pi)
        text = p.Range.Text
        # find paragraphs with significant Latin content
        latin_count = sum(1 for c in text if 0x20 <= ord(c) < 0x7F)
        print(f"P{pi} len={len(text)} latin={latin_count}")
        if latin_count >= 5:
            chars = p.Range.Characters
            xs = []
            for ci in range(1, min(chars.Count + 1, 80)):
                try:
                    c = chars(ci)
                    ch = c.Text
                    if ch in ("\r","\x07"):
                        continue
                    xs.append((ch, c.Information(5), c.Information(6), c.Font.Name, c.Font.Size))
                except Exception:
                    continue
            print(f"P{pi} text: {text[:60]!r}")
            print(f"  font: {xs[0][3]!r} size: {xs[0][4]}")
            # Group by line
            lines = {}
            for r in xs:
                lines.setdefault(round(r[2], 1), []).append(r)
            for y in sorted(lines.keys())[:2]:
                ln = sorted(lines[y], key=lambda r: r[1])
                advs = []
                for i in range(len(ln)-1):
                    a = round(ln[i+1][1] - ln[i][1], 2)
                    advs.append((ln[i][0], a))
                first_x = ln[0][1]
                last_x = ln[-1][1]
                print(f"  y={y} chars={len(ln)} width={last_x-first_x:.2f}pt")
                for c, a in advs:
                    is_latin = ord(c) < 128 if c else False
                    print(f"    {c!r:>6} adv={a} {'(latin)' if is_latin else ''}")
            break
    except Exception as e:
        continue

doc.Close(SaveChanges=False)
word.Quit()
