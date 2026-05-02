"""
Measure cell-border-absorption fixtures.

For each fixture, opens in Word and records text_x of:
  - Row 1 first paragraph (cell with target border)
  - Row 2 first paragraph (reference cell, no border)

The diff between R1 and R2 text_x reveals border absorption behavior.
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "cell_border_absorption_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_cell_border_absorption.json")


RPC_REJECTED_CODES = {-2147418111, -2147023174, -2147023170}

def retry(fn, *args, retries=15, delay=0.3, **kwargs):
    last_exc = None
    for i in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            code = e.args[0] if hasattr(e, "args") and len(e.args) >= 1 else None
            if code in RPC_REJECTED_CODES or "rejected" in str(e).lower():
                pythoncom.PumpWaitingMessages()
                time.sleep(delay * (1.3 ** i))
                continue
            raise
    raise last_exc


def measure(word, path):
    wdoc = retry(lambda: word.Documents.Open(path))
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        # Find both rows' first paragraph in the table
        if wdoc.Tables.Count == 0:
            return {"path": os.path.basename(path), "error": "no table"}
        tbl = wdoc.Tables(1)
        row_data = []
        for r in range(1, tbl.Rows.Count + 1):
            cell = tbl.Cell(r, 1)
            # Use cell.Range start to get cell content top-left
            cell_x_para = round(cell.Range.Information(5), 4)
            cell_y_para = round(cell.Range.Information(6), 4)
            # Find first character of body text — measure ITS x via sub-range
            cell_text = cell.Range.Text
            sample_idx = None
            for ch in ("R", "t", ":"):
                i = cell_text.find(ch)
                if i >= 0:
                    sample_idx = i
                    break
            char_x = None
            char_glyph = None
            if sample_idx is not None:
                sub = wdoc.Range(cell.Range.Start + sample_idx,
                                 cell.Range.Start + sample_idx + 1)
                try:
                    char_x = round(sub.Information(5), 4)
                    char_glyph = cell_text[sample_idx]
                except Exception:
                    pass
            # Also try a position deeper in the text (e.g. char 5)
            char_x_5 = None
            if len(cell_text.strip()) >= 5:
                sub5 = wdoc.Range(cell.Range.Start + 5, cell.Range.Start + 6)
                try:
                    char_x_5 = round(sub5.Information(5), 4)
                except Exception:
                    pass
            # Also: cell.Borders(LeftEdge=1).LineStyle / .LineWidth
            try:
                left_border = cell.Borders(1)  # wdBorderLeft
                lb_lw = left_border.LineWidth
                lb_ls = left_border.LineStyle
            except Exception:
                lb_lw = None
                lb_ls = None
            row_data.append({
                "row": r,
                "para_start_x": cell_x_para,
                "para_start_y": cell_y_para,
                "char0_glyph": char_glyph,
                "char0_x": char_x,
                "char5_x": char_x_5,
                "left_border_lineWidth": lb_lw,
                "left_border_lineStyle": lb_ls,
            })
        return {
            "path": os.path.basename(path),
            "rows": row_data,
        }
    finally:
        wdoc.Close(False)


def parse_filename(name):
    m = re.search(r"sz(\d+)hp", name)
    if m:
        return int(m.group(1))
    return None


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    paths = sorted(glob.glob(os.path.join(FIX_DIR, "*.docx")))
    print(f"Measuring {len(paths)} fixtures from {FIX_DIR}")

    results = []
    try:
        for p in paths:
            try:
                r = measure(word, p)
                sz_hp = parse_filename(os.path.basename(p))
                r["sz_halfpt"] = sz_hp
                r["sz_pt"] = sz_hp / 2.0 if sz_hp is not None else None
                results.append(r)
                if "rows" in r:
                    r1 = r["rows"][0] if r["rows"] else {}
                    r2 = r["rows"][1] if len(r["rows"]) > 1 else {}
                    diff_para = (r1.get("para_start_x", 0) - r2.get("para_start_x", 0)) if (r1 and r2) else None
                    diff_char0 = ((r1.get("char0_x") or 0) - (r2.get("char0_x") or 0)) if (r1 and r2) else None
                    diff_char5 = ((r1.get("char5_x") or 0) - (r2.get("char5_x") or 0)) if (r1 and r2) else None
                    print(f"  sz={sz_hp:3}hp ({r['sz_pt']:5}pt) "
                          f"R1.para={r1.get('para_start_x')} ch0={r1.get('char0_x')} ch5={r1.get('char5_x')} "
                          f"lb_lw={r1.get('left_border_lineWidth')} | "
                          f"R2.para={r2.get('para_start_x')} ch0={r2.get('char0_x')} | "
                          f"diff(para,ch0,ch5)={diff_para},{diff_char0},{diff_char5}")
                else:
                    print(f"  sz={sz_hp}: ERR {r.get('error')}")
            except Exception as e:
                print(f"  ERR {os.path.basename(p)}: {e}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records to {OUT_JSON}")
        try:
            word.Quit()
        except Exception as e:
            print(f"  (word.Quit failed: {e})")


if __name__ == "__main__":
    main()
