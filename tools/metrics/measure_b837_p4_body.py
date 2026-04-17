"""Find Word's last body paragraph on b837 p.4."""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

docx = os.path.abspath(
    r"tools\golden-test\documents\docx\b837808d0555_20240705_resources_data_guideline_02.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx, ReadOnly=True); time.sleep(0.5)
    paras = list(doc.Paragraphs)
    p4_paras = []
    for i, p in enumerate(paras, 1):
        try:
            pg = p.Range.Information(3)
            if pg == 4:
                y = p.Range.Information(6)
                txt = p.Range.Text[:40].replace('\r','').replace('\n',' ').replace('\x07','')
                p4_paras.append({"idx": i, "y": round(y, 1), "text": txt})
        except: pass

    print(f"Word p4 paras: {len(p4_paras)}")
    if p4_paras:
        print(f"FIRST idx={p4_paras[0]['idx']} y={p4_paras[0]['y']}: {p4_paras[0]['text']!r}")
        print(f"LAST  idx={p4_paras[-1]['idx']} y={p4_paras[-1]['y']}: {p4_paras[-1]['text']!r}")
        # Print all
        for p in p4_paras:
            print(f"  idx={p['idx']:<3} y={p['y']:<6} {p['text'][:50]!r}")

    doc.Close(False)
finally:
    word.Quit()
