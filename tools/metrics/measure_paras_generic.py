"""Generic Word paragraph measurement tool.

Usage: python measure_paras_generic.py <docx_path> <output_json>

Outputs list of {idx, page, y, sjis_hex, text_utf8} for all body paragraphs.
"""
import os, sys, time, json
import win32com.client

if len(sys.argv) < 3:
    print("Usage: measure_paras_generic.py <docx> <output.json>")
    sys.exit(1)

docx = os.path.abspath(sys.argv[1])
out_path = sys.argv[2]

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
            sjis = text.encode('shift_jis', errors='replace')
            data.append({
                "idx": i, "page": int(pg), "y": round(y, 1),
                "sjis_hex": sjis.hex(),
                "text_utf8": text,
            })
        except Exception as e:
            data.append({"idx": i, "error": str(e)})

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Saved {len(data)} paragraphs to {out_path}")

    doc.Close(False)
finally:
    word.Quit()
