"""Measure per-char x in middle_dot_context_repro to derive ・ spacing rule."""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(r"pipeline_data\middle_dot_context_repro.docx")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True); time.sleep(0.3)
        doc.Repaginate()
        n = doc.Paragraphs.Count
        print(f"Total paragraphs: {n}", flush=True)
        for i in range(1, n + 1):
            p = doc.Paragraphs(i)
            txt = p.Range.Text.replace("\r","").replace("\x07","")
            if not (txt.startswith("S") and len(txt) > 3 and txt[1].isdigit()):
                continue
            label = txt[:3 if txt[2].isdigit() else 2]
            pr = p.Range
            nc = pr.Characters.Count
            chars = []
            for ci in range(1, nc + 1):
                try:
                    ch = pr.Characters(ci)
                    x = ch.Information(5)
                    y = ch.Information(6)
                    c = ch.Text
                    chars.append((ci, c, round(x, 2), round(y, 2)))
                except Exception:
                    pass
            print(f"\n=== {label}: {txt[:40]!r}")
            prev_x = None
            prev_y = None
            for ci, c, x, y in chars:
                adv = f"{x-prev_x:.2f}" if prev_x is not None and y == prev_y else "WRAP"
                print(f"  {ci:2} {c!r:>6} x={x:7.2f} y={y:6.1f} adv={adv}")
                prev_x, prev_y = x, y
        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
