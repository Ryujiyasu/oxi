"""Measure ・ advance in hint variants."""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VARIANTS = ["H1_plain", "H2_hint", "H3_pprRpr_hint", "H4_pprRpr_nohint"]
DOC_DIR = os.path.abspath(r"pipeline_data")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        for v in VARIANTS:
            path = os.path.join(DOC_DIR, f"yakumono_hint_{v}.docx")
            if not os.path.exists(path):
                continue
            doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
            doc.Repaginate()
            # Find ・利 para
            for pi in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(pi)
                txt = p.Range.Text.replace("\r","").replace("\x07","")
                if txt.startswith("・"):
                    pr = p.Range
                    # First 3 chars
                    x1 = pr.Characters(1).Information(5)
                    x2 = pr.Characters(2).Information(5)
                    x3 = pr.Characters(3).Information(5)
                    c1 = pr.Characters(1).Text
                    c2 = pr.Characters(2).Text
                    c3 = pr.Characters(3).Text
                    print(f"[{v}]")
                    print(f"  {c1}x={x1:.2f} -> {c2}x={x2:.2f} (adv={x2-x1:.2f}) -> {c3}x={x3:.2f} (adv={x3-x2:.2f})")
                    break
            doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
