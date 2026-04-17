import os, sys, time
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VARIANTS = ["yakumono_kern_K1_nokern.docx", "yakumono_kern_K2_kern2.docx", "yakumono_kern_K3_kern20.docx"]
DOC_DIR = os.path.abspath(r"pipeline_data")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False; word.DisplayAlerts = False
try:
    for name in VARIANTS:
        path = os.path.join(DOC_DIR, name)
        if not os.path.exists(path): continue
        doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
        doc.Repaginate()
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            txt = p.Range.Text.replace("\r","").replace("\x07","")
            if txt.startswith("・"):
                pr = p.Range
                try:
                    c1 = pr.Characters(1); c2 = pr.Characters(2); c3 = pr.Characters(3)
                    x1 = c1.Information(5); x2 = c2.Information(5); x3 = c3.Information(5)
                    print(f"[{name}] {c1.Text}={x1:.2f} {c2.Text}={x2:.2f} adv1={x2-x1:.2f} {c3.Text}={x3:.2f} adv2={x3-x2:.2f}")
                except Exception as e:
                    print(f"  err: {e}")
                break
        doc.Close(False)
finally:
    word.Quit()
