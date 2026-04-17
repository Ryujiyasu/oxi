"""Measure ALL Word paragraphs for d77a: page + y + text (encoded as shift_jis bytes hex).

Uses raw Shift-JIS byte output to avoid UTF-8 encoding artifacts.
"""
import os, sys, time, json
import win32com.client

docx = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx, ReadOnly=True); time.sleep(0.5)
    paras = list(doc.Paragraphs)
    data = []
    for i, p in enumerate(paras, 1):
        try:
            pg = p.Range.Information(3)
            y = p.Range.Information(6)
            text = p.Range.Text[:60].replace('\r', '').replace('\n', ' ').replace('\x07', '')
            # Encode to shift_jis bytes; store as hex string to avoid JSON encoding issues
            sjis = text.encode('shift_jis', errors='replace')
            data.append({
                "idx": i, "page": int(pg), "y": round(y, 1),
                "sjis_hex": sjis.hex(),
                "text_utf8": text,  # also keep utf-8 for reference
            })
        except Exception as e:
            data.append({"idx": i, "error": str(e)})

    with open("pipeline_data/d77a_word_all_paras.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Saved {len(data)} paragraphs to pipeline_data/d77a_word_all_paras.json")

    doc.Close(False)
finally:
    word.Quit()
