"""Measure Word right-aligned paragraph X positions in b837 P3-P9."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

ps = doc.PageSetup
print(f"page={ps.PageWidth} x {ps.PageHeight}")
print(f"margins L={ps.LeftMargin} R={ps.RightMargin} T={ps.TopMargin}")
content_right = ps.PageWidth - ps.RightMargin
print(f"content_right edge = {content_right}")
print()

# P3-P9 are the date block. Check first char and last char X
for pi in range(1, 12):
    p = doc.Paragraphs(pi)
    txt = p.Range.Text.strip()
    if not txt:
        print(f"P{pi}: (empty)")
        continue
    align = p.Alignment
    align_name = {0:'left', 1:'center', 2:'right', 3:'both'}.get(align, f'?{align}')
    chars = p.Range.Characters
    if chars.Count > 0:
        try:
            first_x = chars(1).Information(5)
            last_idx = chars.Count
            # Skip trailing \r
            while last_idx > 1:
                t = chars(last_idx).Text
                if t in ('\r', '\x07', '\x0b'): last_idx -= 1
                else: break
            last_x = chars(last_idx).Information(5)
            print(f"P{pi}: align={align_name} text={txt[:30]!r}")
            print(f"   first_x={first_x:.2f} last_x={last_x:.2f}")
        except Exception as e:
            print(f"P{pi}: error {e}")

doc.Close(SaveChanges=False)
word.Quit()
