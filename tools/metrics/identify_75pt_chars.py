"""Identify which characters are 7.5pt wide in MS Mincho 8pt."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
doc = word.Documents.Open(path, ReadOnly=False)
time.sleep(1)

para = doc.Paragraphs(196)
rng = para.Range

# Switch to LEFT to get natural widths
para.Alignment = 0
time.sleep(0.1)

chars = rng.Characters
line1_y = chars(1).Information(6)

# Collect ALL chars with their widths
char_widths = {}
prev_x = None
for i in range(1, 85):
    try:
        c = chars(i)
        ch = c.Text
        if ch in ('\r', '\x07'):
            continue
        cx = c.Information(5)

        if prev_x is not None:
            w = round(cx - prev_x, 2)
            if w > 0 and w < 20:  # reasonable width
                cp = ord(prev_ch)
                if cp not in char_widths:
                    char_widths[cp] = {'char': prev_ch, 'width': w, 'count': 1}
                else:
                    char_widths[cp]['count'] += 1

        prev_x = cx
        prev_ch = ch
    except:
        pass

# Restore
para.Alignment = 3

# Print chars grouped by width
w_groups = {}
for cp, info in char_widths.items():
    w = info['width']
    if w not in w_groups:
        w_groups[w] = []
    w_groups[w].append(info)

for w in sorted(w_groups.keys()):
    print(f"\n=== Width {w}pt ===")
    for info in sorted(w_groups[w], key=lambda x: -x['count']):
        cp = ord(info['char'])
        print(f"  U+{cp:04X} '{info['char']}' count={info['count']}")

doc.Close(SaveChanges=False)
word.Quit()
