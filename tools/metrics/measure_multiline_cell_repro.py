"""COM-measure Word's row height for each table in multiline_cell_repro.docx.

Output: per-table (n_lines, row_height_pt) data. Plot or derive formula.
"""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    r"pipeline_data\multiline_cell_repro.docx"
)
OUT = r"pipeline_data/multiline_cell_repro_data.json"


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True); time.sleep(0.3)
        doc.Repaginate()
        results = []
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            cell = tbl.Cell(1, 1)
            # Row top Y
            top_y = cell.Range.Information(6)
            top_pg = cell.Range.Information(3)
            # Bottom Y: use the paragraph after the table
            tbl_end = tbl.Range.End
            try:
                after = doc.Range(tbl_end, tbl_end)
                bottom_y = after.Information(6)
                bottom_pg = after.Information(3)
            except Exception:
                bottom_y = None
                bottom_pg = None
            if top_y and bottom_y and top_pg == bottom_pg:
                row_h = round(bottom_y - top_y, 2)
            else:
                row_h = None
            # Also count cell lines
            pr = cell.Range
            n = pr.Characters.Count
            ys = set()
            step = max(1, n // 80)
            for i in range(1, n + 1, step):
                try:
                    ys.add(round(pr.Characters(i).Information(6), 1))
                except Exception:
                    pass
            try:
                ys.add(round(pr.Characters(n).Information(6), 1))
            except Exception:
                pass
            n_lines = len(ys)
            # Also measure pitch (diff between consecutive line ys)
            sorted_ys = sorted(ys)
            pitches = [round(sorted_ys[i+1]-sorted_ys[i], 2) for i in range(len(sorted_ys)-1)]
            result = {
                "ti": ti,
                "n_lines": n_lines,
                "row_h": row_h,
                "top_y": round(top_y, 2) if top_y else None,
                "bottom_y": round(bottom_y, 2) if bottom_y else None,
                "pitches": pitches,
            }
            results.append(result)
            print(f"T{ti}: n={n_lines} row_h={row_h}pt pitches={pitches[:3]}", file=sys.stderr)
        doc.Close(False)
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUT}")

    # Formula analysis
    print("\n=== Formula derivation ===")
    print(f"{'n':4} {'row_h':8} {'h/n':8} {'(h-P1)/(n-1) for P1=row_h[0]':30}")
    if results and results[0]["row_h"]:
        p1_h = results[0]["row_h"]
        for r in results:
            if r["row_h"] and r["n_lines"] > 0:
                n = r["n_lines"]
                h = r["row_h"]
                per_n = round(h/n, 2)
                extra = round((h - p1_h) / (n - 1), 2) if n > 1 else "-"
                print(f"{n:4} {h:8.2f} {per_n:8} {extra}")


if __name__ == "__main__":
    main()
