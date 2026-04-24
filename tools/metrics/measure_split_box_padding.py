"""Measure SB_* repros to confirm the formula:
    first_body_after_y = last_new_page_y + line_height * (1 + has_trailing_empty)

For each repro:
  - Find the table (1 per doc)
  - Find the split row
  - Measure continuation y-positions, line_height, last_new_page_y
  - Measure has_trailing_empty (cell's last paragraph has no runs)
  - Measure first body paragraph after table (y, page)
  - Predict using formula, compare with actual

Output: pipeline_data/split_box_padding_measurements.json
"""
import json
import sys
from pathlib import Path
from statistics import median
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPRO_DIR = Path(r"c:\Users\ryuji\oxi-main\tools\metrics\split_box_padding_repro")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\split_box_padding_measurements.json")

REPROS = ["SB_A", "SB_B", "SB_C", "SB_D", "SB_E", "SB_F"]


def measure_one(word, docx: Path) -> dict:
    doc = word.Documents.Open(str(docx.resolve()), ReadOnly=True)
    try:
        result = {"file": docx.name}
        if doc.Tables.Count == 0:
            return {**result, "error": "no tables"}
        tbl = doc.Tables(1)
        rows = tbl.Rows.Count
        cols = tbl.Columns.Count
        split_row_idx = None
        row_start_page = None
        for ri in range(1, rows + 1):
            row = tbl.Rows(ri)
            rs = doc.Range(row.Range.Start, row.Range.Start + 1).Information(3)
            re = doc.Range(row.Range.End - 1, row.Range.End).Information(3)
            if re > rs:
                split_row_idx = ri
                row_start_page = rs
                break
        if split_row_idx is None:
            return {**result, "error": "no split row"}

        row = tbl.Rows(split_row_idx)
        cell = row.Cells(1)
        cell_range = cell.Range

        seen = {}
        for off in range(cell_range.Start, cell_range.End):
            try:
                r = doc.Range(off, off + 1)
                pg = r.Information(3)
                y = r.Information(6)
            except Exception:
                continue
            if pg <= row_start_page:
                continue
            key = (int(pg), round(y, 1))
            if key not in seen:
                seen[key] = off
        ys_sorted = sorted(seen.keys())
        if not ys_sorted:
            return {**result, "error": "no continuation y"}
        first_pg = ys_sorted[0][0]
        page_ys = [y for p, y in ys_sorted if p == first_pg]
        filtered = [page_ys[0]]
        for y in page_ys[1:]:
            if y - filtered[-1] < 50:
                filtered.append(y)
            else:
                break
        lh = None
        if len(filtered) >= 2:
            diffs = [filtered[i+1] - filtered[i] for i in range(len(filtered)-1)]
            lh = round(median(diffs), 2)

        # has_trailing_empty: last cell paragraph has no runs/text
        paras = cell.Range.Paragraphs
        last_p = paras(paras.Count)
        last_p_text = last_p.Range.Text.strip("\r\n\x07 \t")
        has_trailing = len(last_p_text) == 0

        tbl_end = tbl.Range.End
        after_body = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            pr = p.Range
            if pr.Start < tbl_end:
                continue
            y = pr.Information(6)
            pg = pr.Information(3)
            in_table = bool(pr.Information(12))
            if in_table:
                continue
            after_body = {
                "idx": pi,
                "page": int(pg),
                "y_pt": round(y, 2),
                "text": pr.Text.strip("\r\n\x07 \t")[:30],
            }
            break

        pred = None
        resid = None
        if filtered and lh is not None and after_body:
            last_y = filtered[-1]
            te = 1 if has_trailing else 0
            pred = round(last_y + lh * (1 + te), 2)
            resid = round(after_body["y_pt"] - pred, 2)

        return {
            **result,
            "continuation_lines": len(filtered),
            "first_new_page_y": filtered[0] if filtered else None,
            "last_new_page_y": filtered[-1] if filtered else None,
            "line_height": lh,
            "has_trailing_empty": has_trailing,
            "first_body_after": after_body,
            "formula_prediction": pred,
            "formula_residual": resid,
        }
    finally:
        doc.Close(SaveChanges=0)


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = {}
    try:
        for name in REPROS:
            path = REPRO_DIR / f"{name}.docx"
            if not path.exists():
                print(f"SKIP: {name} not found")
                continue
            print(f"Measuring {name}...", flush=True)
            r = measure_one(word, path)
            results[name] = r
            by = r.get('first_body_after', {}).get('y_pt') if r.get('first_body_after') else '?'
            print(f"  cont={r.get('continuation_lines')} lh={r.get('line_height')} last_y={r.get('last_new_page_y')} TE={r.get('has_trailing_empty')} body_y={by} pred={r.get('formula_prediction')} Δ={r.get('formula_residual')}", flush=True)
    finally:
        word.Quit()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
