"""Measure ・ advance in stripped d77a variants."""
import os, sys, time
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VARIANTS = [
    "d77a_S1_only_para28.docx",
    "d77a_S3_no_tables.docx",
    "d77a_S5_minimal_compat.docx",
]
DOC_DIR = os.path.abspath(r"pipeline_data")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False; word.DisplayAlerts = False
try:
    for name in VARIANTS:
        path = os.path.join(DOC_DIR, name)
        if not os.path.exists(path):
            print(f"[{name}] missing"); continue
        try:
            doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
            doc.Repaginate()
            found = False
            for pi in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(pi)
                txt = p.Range.Text.replace("\r","").replace("\x07","")
                if "・利用規約名" in txt[:10]:
                    pr = p.Range
                    try:
                        c1 = pr.Characters(1); c2 = pr.Characters(2); c3 = pr.Characters(3)
                        x1 = c1.Information(5); x2 = c2.Information(5); x3 = c3.Information(5)
                        print(f"[{name}] {c1.Text}={x1:.2f} {c2.Text}={x2:.2f} adv1={x2-x1:.2f} {c3.Text}={x3:.2f} adv2={x3-x2:.2f}")
                        found = True
                    except Exception as e:
                        print(f"  err: {e}")
                    break
            if not found:
                print(f"[{name}] para not found")
            doc.Close(False)
        except Exception as e:
            print(f"[{name}] open err: {e}")
finally:
    word.Quit()
