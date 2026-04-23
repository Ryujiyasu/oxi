"""
Measure Word's placement of the body paragraph immediately AFTER tbl5 on p.7
in d77a. Oxi places it at y=79 which overlaps the tbl5 continuation box
(y=71 to y=89). Hypothesis: Word places it below the continuation box.

Also measures how many body paragraphs follow immediately and their y positions
so we can check against Oxi's pi=69..72 positions from the layout dump.
"""
import json
from pathlib import Path
import win32com.client as w32


DOC = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\d77a_post_tbl5_measurement.json")


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": DOC.name}
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            result["page_count"] = doc.ComputeStatistics(2)

            # Iterate all body paragraphs. For each, capture y, page, and text sample.
            # Also flag whether paragraph is inside a table (we want the ones just
            # before and after the tbl5 split on p.6/p.7).
            all_paras = []
            for i, p in enumerate(doc.Paragraphs, start=1):
                r = p.Range
                try:
                    y = r.Information(6)  # wdVerticalPositionRelativeToPage
                    pg = r.Information(3)  # wdActiveEndPageNumber
                    in_table = bool(r.Information(12))  # wdWithInTable
                except Exception:
                    continue
                text = r.Text[:40].replace("\r", "\\r").replace("\n", "\\n").replace("\x07", "\\x07")
                if pg in (6, 7, 8):
                    all_paras.append({
                        "idx": i,
                        "page": pg,
                        "y_pt": round(y, 3),
                        "in_table": in_table,
                        "text": text,
                    })

            result["page_6_7_8_paras"] = all_paras
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    print(f"Wrote {OUT}")

    # Print paras on p.7 for quick review
    print("--- Page 7 paragraphs (Word) ---")
    for p in all_paras:
        if p["page"] == 7:
            marker = "[CELL]" if p["in_table"] else "[BODY]"
            print(f"  idx={p['idx']} {marker} y={p['y_pt']:.2f} text={p['text']!r}")


if __name__ == "__main__":
    main()
