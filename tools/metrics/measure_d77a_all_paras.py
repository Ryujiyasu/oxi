"""
Measure ALL paragraphs in d77a with their (idx, page, y, in_table, text).
Used to identify the block_idx correspondence to pi=126/127 in the cursor_y
FALSIFIED session, and to understand what Word does at the p.10→p.11 boundary.

Output: pipeline_data/d77a_all_paras_measurement.json
"""
import json
from pathlib import Path
import win32com.client as w32


DOC = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\d77a_all_paras_measurement.json")


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": DOC.name}
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            result["page_count"] = doc.ComputeStatistics(2)

            all_paras = []
            for i, p in enumerate(doc.Paragraphs, start=1):
                r = p.Range
                try:
                    y = r.Information(6)
                    pg = r.Information(3)
                    in_table = bool(r.Information(12))
                except Exception:
                    continue
                text = r.Text[:60].replace("\r", "\\r").replace("\n", "\\n").replace("\x07", "\\x07")
                is_empty = (r.Text.strip("\r\n\x07\t ") == "")
                all_paras.append({
                    "idx": i,
                    "page": pg,
                    "y_pt": round(y, 3),
                    "in_table": in_table,
                    "is_empty": is_empty,
                    "text": text,
                })

            result["paras"] = all_paras
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    print(f"Wrote {OUT}")

    # Print pp.10-12 for phantom-page analysis
    for pg in [10, 11, 12]:
        print(f"--- Page {pg} (Word) ---")
        for p in all_paras:
            if p["page"] == pg:
                marker = "[CELL]" if p["in_table"] else ("[EMPT]" if p["is_empty"] else "[BODY]")
                print(f"  idx={p['idx']:3d} {marker} y={p['y_pt']:6.2f} text={p['text']!r}")


if __name__ == "__main__":
    main()
