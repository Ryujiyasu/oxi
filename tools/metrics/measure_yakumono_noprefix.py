"""Measure yakumono_noprefix variants — verify ・ at paragraph start."""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = ["yakumono_noprefix_A.docx", "yakumono_noprefix_B.docx"]
DOC_DIR = os.path.abspath(r"pipeline_data")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        for name in DOCS:
            path = os.path.join(DOC_DIR, name)
            if not os.path.exists(path):
                print(f"[{name}] missing"); continue
            doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
            doc.Repaginate()
            print(f"\n=== {name} ===")
            for pi in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(pi)
                txt = p.Range.Text.replace("\r","").replace("\x07","")
                if not txt.startswith("・"):
                    continue
                pr = p.Range
                n = pr.Characters.Count
                # First 3 chars
                advances = []
                for ci in range(1, min(n + 1, 5)):
                    try:
                        ch = pr.Characters(ci)
                        x = ch.Information(5)
                        y = ch.Information(6)
                        c = ch.Text
                        advances.append({"ci": ci, "c": c, "x": round(x,2), "y": round(y,2)})
                    except Exception:
                        pass
                if len(advances) >= 2:
                    dot_x = advances[0]["x"]
                    next_x = advances[1]["x"]
                    dot_adv = next_x - dot_x
                    print(f"  para {pi}: ・x={dot_x} next={advances[1]['c']!r}x={next_x} ・adv={dot_adv:.2f}")
            doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
