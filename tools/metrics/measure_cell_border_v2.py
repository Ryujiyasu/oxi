"""Measure cell border absorption v2 fixtures (built via Word COM)."""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "cell_border_absorption_v2")
OUT_JSON = os.path.join(OUT_DIR, "ra2_cell_border_absorption_v2.json")


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
        if wdoc.Tables.Count == 0:
            return {"path": os.path.basename(path), "error": "no table"}
        tbl = wdoc.Tables(1)
        rows = []
        for r in range(1, tbl.Rows.Count + 1):
            cell = tbl.Cell(r, 1)
            text = cell.Range.Text
            idx = text.find("R")
            char_x = None
            if idx >= 0:
                sub = wdoc.Range(cell.Range.Start + idx,
                                 cell.Range.Start + idx + 1)
                try:
                    char_x = round(sub.Information(5), 4)
                except Exception:
                    pass
            try:
                b = cell.Borders(1)  # wdBorderLeft
                ls = b.LineStyle
                lw = b.LineWidth
            except Exception:
                ls = None
                lw = None
            try:
                lp = round(cell.LeftPadding, 4)
            except Exception:
                lp = None
            rows.append({
                "row": r,
                "text": text.strip()[:30],
                "first_char_x": char_x,
                "left_padding": lp,
                "left_border_LineStyle": ls,
                "left_border_LineWidth": lw,
            })
        return {"path": os.path.basename(path), "rows": rows}
    finally:
        wdoc.Close(False)


def parse_filename(name):
    """Extract (padding, border_width) from CBV2_pad{X}_w{Y}pt.docx."""
    m = re.search(r"pad(\d+)p(\d+)_w(\d+)p(\d+)pt", name)
    if m:
        pad = float(f"{m.group(1)}.{m.group(2)}")
        w = float(f"{m.group(3)}.{m.group(4)}")
        return pad, w
    # Fallback for old naming
    m2 = re.search(r"w(\d+)p(\d+)pt", name)
    if m2:
        return None, float(f"{m2.group(1)}.{m2.group(2)}")
    return None, None


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    def sort_key(p):
        pad, w = parse_filename(os.path.basename(p))
        return (pad if pad is not None else -1, w if w is not None else -1)
    paths = sorted(glob.glob(os.path.join(FIX_DIR, "*.docx")), key=sort_key)
    print(f"Measuring {len(paths)} fixtures from {FIX_DIR}\n")
    print(f"  {'pad':>5}  {'border':>6}  {'R1.x':>6} ls lw  {'R2.x':>6}  {'diff':>6}")
    print("  " + "-"*48)

    results = []
    try:
        last_pad = None
        for p in paths:
            try:
                r = measure(word, p)
                pad, w_pt = parse_filename(os.path.basename(p))
                r["set_padding_pt"] = pad
                r["set_width_pt"] = w_pt
                results.append(r)
                if "rows" in r:
                    r1 = r["rows"][0]
                    r2 = r["rows"][1] if len(r["rows"]) > 1 else {}
                    diff = (r1.get("first_char_x") or 0) - (r2.get("first_char_x") or 0)
                    if pad != last_pad:
                        print()
                        last_pad = pad
                    print(f"  {pad:>5}  {w_pt:>6}  {r1.get('first_char_x'):>6} {r1.get('left_border_LineStyle')}  {r1.get('left_border_LineWidth'):>2}  {r2.get('first_char_x'):>6}  {diff:+.2f}")
            except Exception as e:
                print(f"  ERR {os.path.basename(p)}: {e}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT_JSON}")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
