"""Measure the LM1 repro (docGrid=lines 350tw) — same as measure_multiline_cell_repro
but reads the LM1 docx and writes to _lm1.json.
"""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(r"pipeline_data\multiline_cell_repro_lm1.docx")
OUT = r"pipeline_data/multiline_cell_repro_lm1_data.json"


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
            top_y = cell.Range.Information(6)
            top_pg = cell.Range.Information(3)
            tbl_end = tbl.Range.End
            try:
                after = doc.Range(tbl_end, tbl_end)
                bottom_y = after.Information(6)
                bottom_pg = after.Information(3)
            except Exception:
                bottom_y = None; bottom_pg = None
            row_h = round(bottom_y - top_y, 2) if (top_y and bottom_y and top_pg == bottom_pg) else None
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
            sorted_ys = sorted(ys)
            pitches = [round(sorted_ys[i+1]-sorted_ys[i], 2) for i in range(len(sorted_ys)-1)]
            results.append({"ti": ti, "n_lines": n_lines, "row_h": row_h, "pitches": pitches})
            print(f"T{ti}: n={n_lines} row_h={row_h}pt pitches={pitches[:4]}", file=sys.stderr)
        doc.Close(False)
    finally:
        word.Quit()

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUT}")

    print("\n=== LM1 Formula derivation (linePitch=17.5pt expected if grid-snap) ===")
    if results and results[0]["row_h"]:
        p1_h = results[0]["row_h"]
        print(f"{'n':4} {'row_h':8} {'h/n':8} {'extra/(n-1)':12}")
        for r in results:
            if r["row_h"] and r["n_lines"] > 0:
                n = r["n_lines"]; h = r["row_h"]
                per_n = round(h/n, 2)
                extra = round((h - p1_h) / (n - 1), 2) if n > 1 else "-"
                print(f"{n:4} {h:8.2f} {per_n:8} {extra}")


if __name__ == "__main__":
    main()
