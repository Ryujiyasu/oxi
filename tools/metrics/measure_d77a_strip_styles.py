import os, sys, time
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VARIANTS = ["d77a_S6_minimal_styles.docx", "d77a_S7_minimal_theme.docx", "d77a_S8_no_docgrid.docx", "d77a_S9_no_hint.docx"]
DOC_DIR = os.path.abspath(r"pipeline_data")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False; word.DisplayAlerts = False
try:
    for name in VARIANTS:
        path = os.path.join(DOC_DIR, name)
        if not os.path.exists(path): continue
        try:
            doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
            doc.Repaginate()
            for pi in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(pi)
                txt = p.Range.Text.replace("\r","").replace("\x07","")
                if "・利用規約名" in txt[:10]:
                    pr = p.Range
                    try:
                        c1 = pr.Characters(1); c2 = pr.Characters(2)
                        x1 = c1.Information(5); x2 = c2.Information(5)
                        print(f"[{name}] {c1.Text}={x1:.2f} {c2.Text}={x2:.2f} adv={x2-x1:.2f}")
                    except Exception as e: print(f"  err: {e}")
                    break
            doc.Close(False)
        except Exception as e: print(f"[{name}] open err: {e}")
finally:
    word.Quit()
