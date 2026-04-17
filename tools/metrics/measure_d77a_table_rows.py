"""Measure Word's actual table row heights for d77a p.2–p.9 tables.

Uses COM: doc.Tables(i).Rows(j).HeightRule + .Height.
Also Range.Information(6) at cell start + next para start for raw Y.

Output: per-table (start_y, content_end_y, word_height_pt) so we can
compute per-table over-alloc vs Oxi.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

docx_path = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True); time.sleep(0.5)
    print(f"Tables: {doc.Tables.Count}")
    print(f"{'#':>3} {'page':>5} {'start_y':>8} {'end_y':>8} {'height':>8} {'n_paras':>8} {'first_text':>30}")
    print("-" * 75)
    data = []
    for i in range(1, doc.Tables.Count + 1):
        t = doc.Tables(i)
        rng = t.Range
        try:
            page = rng.Information(3)
            start_y = rng.Information(6)
            # Find the paragraph AFTER this table (next non-table para)
            end_pos = rng.End
            after = doc.Range(end_pos, end_pos + 1)
            end_y_raw = after.Information(6)
            end_page = after.Information(3)
            # Handle page crossing
            if end_page != page:
                end_y = 841.9  # page bottom; height is rough
            else:
                end_y = end_y_raw
            height = end_y - start_y
            # Count paragraphs in table
            n_paras = t.Range.Paragraphs.Count
            first_text = t.Range.Paragraphs(1).Range.Text[:28].replace('\r', ' ')
            data.append({
                "idx": i, "page": int(page), "start_y": round(start_y, 2),
                "end_y": round(end_y, 2), "height_pt": round(height, 2),
                "n_paras": n_paras, "first_text": first_text,
            })
            print(f"{i:>3} {int(page):>5} {start_y:>8.2f} {end_y:>8.2f} {height:>8.2f} {n_paras:>8} {first_text!r:>30}")
        except Exception as e:
            print(f"{i:>3} ERROR: {e}")

    out = "pipeline_data/d77a_word_table_rows.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out}")

    doc.Close(False)
finally:
    word.Quit()
