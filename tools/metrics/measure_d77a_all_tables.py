"""Measure Word's actual geometry for ALL tables in d77a.

Goal (per memory project_d77a_p9_table_height_corrected.md):
Confirm 24pt/table under-allocation hypothesis and localize root cause to
one of: (1) line count (wrap mismatch), (2) cell padding, (3) line pitch.

For each table:
- page, row count, total height (y of row1 → y after last row)
- per-row: rendered height (y-delta), cell count
- for (1,1) cell: rendered line count (count distinct line y values in cell range)
- tblGrid widths (sum)

Output JSON: pipeline_data/d77a_tables_word_com.json
"""
import win32com.client, os, json, sys

DOC = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
OUT = r"C:\Users\ryuji\oxi-1\pipeline_data\d77a_tables_word_com.json"

def cell_line_count_and_y(cell):
    """Count distinct line-y values inside a cell by walking chars."""
    try:
        rng = cell.Range
    except Exception:
        return None, None, None
    n = rng.Characters.Count
    if n == 0:
        return 0, None, None
    ys = []
    # Sample up to ~200 chars to avoid pathological runtime
    step = max(1, n // 200)
    for i in range(1, n + 1, step):
        try:
            ch = rng.Characters(i)
            y = ch.Information(6)
            ys.append(round(y, 1))
        except Exception:
            pass
    # Count unique y values (each distinct y = one line)
    uniq = sorted(set(ys))
    if not uniq:
        return 0, None, None
    first_y = uniq[0]
    last_y = uniq[-1]
    return len(uniq), first_y, last_y


def main():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
    out = {"doc": os.path.basename(DOC), "tables": []}
    try:
        wdoc.Repaginate()
        tcount = wdoc.Tables.Count
        print(f"[INFO] total tables: {tcount}", file=sys.stderr)
        for ti in range(1, tcount + 1):
            tbl = wdoc.Tables(ti)
            nrows = tbl.Rows.Count
            try:
                start_page = tbl.Rows(1).Range.Information(3)
                start_y = tbl.Rows(1).Range.Information(6)
            except Exception:
                start_page = start_y = None

            row_ys = []
            for ri in range(1, nrows + 1):
                try:
                    rng = tbl.Rows(ri).Range
                    pg = rng.Information(3)
                    y = rng.Information(6)
                    row_ys.append((ri, pg, y))
                except Exception:
                    row_ys.append((ri, None, None))

            # Heights between consecutive rows on same page
            row_heights = []
            for i in range(len(row_ys) - 1):
                ri, pg, y = row_ys[i]
                nri, npg, ny = row_ys[i + 1]
                if pg == npg and y is not None and ny is not None:
                    row_heights.append({"ri": ri, "pg": pg, "y": round(y, 2), "h": round(ny - y, 2)})
                else:
                    row_heights.append({"ri": ri, "pg": pg, "y": round(y, 2) if y else None, "h": None, "pagebreak": True})

            # Last row height: estimate from next paragraph after table
            last_ri, last_pg, last_y = row_ys[-1]
            last_h = None
            try:
                # Word range end of table + 1 = first thing after
                tbl_end = tbl.Range.End
                after = wdoc.Range(tbl_end, tbl_end)
                after_y = after.Information(6)
                after_pg = after.Information(3)
                if after_pg == last_pg and last_y is not None and after_y is not None:
                    last_h = round(after_y - last_y, 2)
            except Exception:
                pass
            row_heights.append({"ri": last_ri, "pg": last_pg, "y": round(last_y, 2) if last_y else None, "h": last_h, "last": True})

            # Line count in cell(1,1) of first row
            try:
                c11 = tbl.Cell(1, 1)
                lc, fy, ly = cell_line_count_and_y(c11)
                c11_text = c11.Range.Text.replace("\r", "").replace("\x07", "").strip()[:40]
            except Exception as e:
                lc = fy = ly = None
                c11_text = f"<err:{e}>"

            # Total table height
            total_h = None
            if start_y is not None and last_y is not None and start_page == last_pg:
                # sum row heights on same page
                total_h = round(sum(rh["h"] for rh in row_heights if rh.get("h") is not None), 2)

            t_entry = {
                "ti": ti,
                "nrows": nrows,
                "ncols": tbl.Columns.Count if nrows > 0 else None,
                "start_page": start_page,
                "start_y": round(start_y, 2) if start_y else None,
                "c11_text": c11_text,
                "c11_line_count": lc,
                "total_h": total_h,
                "rows": row_heights[:5],  # only first 5 rows to keep file size down
            }
            out["tables"].append(t_entry)
            print(
                f"[T{ti:02d}] p{start_page} rows={nrows} cols={t_entry['ncols']} "
                f"total_h={total_h} c11_lines={lc} text='{c11_text}'",
                file=sys.stderr,
            )
    finally:
        wdoc.Close(False)
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] wrote {OUT}")


if __name__ == "__main__":
    main()
