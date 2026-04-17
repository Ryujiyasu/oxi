"""Generic Word paragraph extractor: reads docx path from CLI, writes JSON.

Usage: python measure_word_paras_generic.py <docx_path> <out_json>
"""
import os, sys, time, json
import win32com.client

if len(sys.argv) < 3:
    print("Usage: measure_word_paras_generic.py <docx_path> <out_json>")
    sys.exit(1)

docx = os.path.abspath(sys.argv[1])
out = sys.argv[2]

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    doc = word.Documents.Open(docx, ReadOnly=True); time.sleep(0.5)
    doc.Repaginate()
    n_paras = doc.Paragraphs.Count
    print(f"Total paragraphs: {n_paras}", flush=True)
    data = []
    # Iterate by index (avoid holding the full Paragraphs collection COM object)
    for i in range(1, n_paras + 1):
        try:
            p = doc.Paragraphs(i)
            pg = p.Range.Information(3)
            y = p.Range.Information(6)
            text = p.Range.Text[:60].replace('\r', '').replace('\n', ' ').replace('\x07', '')
            sjis = text.encode('shift_jis', errors='replace')
            data.append({
                "idx": i, "page": int(pg), "y": round(y, 1),
                "sjis_hex": sjis.hex(),
                "text_utf8": text,
            })
            if i % 50 == 0:
                # Save progress so far in case RPC dies later
                with open(out, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f"  saved {i}/{n_paras} paras so far", flush=True)
        except Exception as e:
            data.append({"idx": i, "error": str(e)})

    os.makedirs(os.path.dirname(out), exist_ok=True) if os.path.dirname(out) else None
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Saved {len(data)} paragraphs to {out}")

    doc.Close(False)
finally:
    word.Quit()
