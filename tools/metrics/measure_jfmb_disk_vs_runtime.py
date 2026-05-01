"""Measure jfmb on-disk vs runtime-saved space advance.

Spec §4.6.3 claimed:
  jfmb (on-disk) → 3.5pt natural
  jfmb (runtime saved) → 6.0pt half-em

Verify both measurements with current Word state.
"""
import win32com.client
import os
import time
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = [
    ("on_disk",  "pipeline_data/docx/japanese_font_mixing_baseline.docx"),
    ("runtime",  "pipeline_data/docx/japanese_font_mixing_baseline_runtime.docx"),
]


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        for label, rel in DOCS:
            path = os.path.abspath(rel)
            d = word.Documents.Open(path, ReadOnly=True)
            time.sleep(0.4)
            chars = d.Range().Characters
            xs = []
            for ci in range(1, min(chars.Count + 1, 200)):
                try:
                    c = chars(ci)
                    t = c.Text
                    if t in ("\r", "\x07"):
                        continue
                    xs.append((ci, t,
                               float(c.Information(5)),
                               c.Font.Name, c.Font.Size))
                except Exception:
                    continue
            d.Close(SaveChanges=False)
            print(f"\n=== {label} ({len(xs)} chars) ===")
            # Find spaces and their advances
            for i in range(len(xs) - 1):
                ci, ch, x, fn, sz = xs[i]
                next_ch = xs[i + 1][1]
                next_x = xs[i + 1][2]
                adv = round(next_x - x, 4)
                if ch == " " or next_ch == " ":
                    marker = " <-- SPACE CHAR"
                else:
                    marker = ""
                if (ch == " " and ord(next_ch) > 127) or (next_ch == " " and
                                                           ord(ch) > 127):
                    marker += " (CJK adj)"
                print(f"  [{ci:3d}] {ch!r:>5} adv={adv:5.2f} font={fn!r}"
                      + marker)
                if i > 60:
                    break
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
