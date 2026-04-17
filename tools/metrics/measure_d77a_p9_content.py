"""Measure which paragraphs Word places on d77a page 9.

Iterate all body paragraphs; use Range.Information(3) to get page number.
Output list of paragraphs on p9 with index and first 30 chars.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

docx = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx, ReadOnly=True); time.sleep(0.5)
    paras = list(doc.Paragraphs)
    print(f"Total body paras: {len(paras)}")
    print(f"Total tables: {doc.Tables.Count}")
    n_pages = doc.ComputeStatistics(2)  # wdStatisticPages
    print(f"Word reports {n_pages} pages\n")

    p9_paras = []
    for i, p in enumerate(paras, 1):
        try:
            pg = p.Range.Information(3)
            if pg == 9:
                y = p.Range.Information(6)
                txt = p.Range.Text[:40].replace('\r', '').replace('\n', ' ').replace('\x07', '')
                p9_paras.append({"idx": i, "y": round(y, 1), "text": txt})
        except Exception as e:
            pass

    print(f"Paragraphs on p9: {len(p9_paras)}")
    if p9_paras:
        print(f"FIRST (idx={p9_paras[0]['idx']}, y={p9_paras[0]['y']}): {p9_paras[0]['text']!r}")
        print(f"LAST  (idx={p9_paras[-1]['idx']}, y={p9_paras[-1]['y']}): {p9_paras[-1]['text']!r}")
        print(f"\nAll para indices on p9: {[p['idx'] for p in p9_paras]}")

    # Save
    with open("pipeline_data/d77a_word_p9_paras.json", "w", encoding="utf-8") as f:
        json.dump(p9_paras, f, ensure_ascii=False, indent=2)

    doc.Close(False)
finally:
    word.Quit()
