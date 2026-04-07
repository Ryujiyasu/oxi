"""Extract Word's line-by-line text content for a docx via COM."""
import win32com.client
import os
import sys

docx = os.path.abspath(sys.argv[1])

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

doc = word.Documents.Open(docx, ReadOnly=True)

import time
time.sleep(1)

for pi in range(1, min(doc.Paragraphs.Count + 1, 6)):
    p = doc.Paragraphs(pi)
    rng = p.Range
    chars = rng.Characters
    n = chars.Count
    print(f"--- P{pi} ({n} chars) ---")
    prev_y = None
    line_text = ""
    line_no = 0
    for ci in range(1, n + 1):
        try:
            c = chars(ci)
            ch = c.Text
            # \r=para mark, \x07=cell mark, \x0b=soft line break (<w:br/>)
            if ch in ("\r", "\x07", "\x0b"):
                continue
            cy = c.Information(6)
            if prev_y is None or abs(cy - prev_y) > 0.5:
                if prev_y is not None:
                    sys.stdout.buffer.write(
                        (f"  L{line_no} y={prev_y:.1f} ({len(line_text)}ch) " + line_text + "\n").encode("utf-8", errors="replace")
                    )
                    line_no += 1
                    line_text = ""
                prev_y = cy
            line_text += ch
        except Exception:
            continue
    if line_text:
        sys.stdout.buffer.write(
            (f"  L{line_no} y={prev_y:.1f} ({len(line_text)}ch) " + line_text + "\n").encode("utf-8", errors="replace")
        )

doc.Close(SaveChanges=False)
word.Quit()
