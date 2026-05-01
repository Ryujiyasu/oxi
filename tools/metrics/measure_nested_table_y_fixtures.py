"""
Measure nested-table Y positions.

For each fixture:
  - Body P1 (reference) Y
  - Outer table row top Y (via Tables(1).Rows(1).Range.Information(6))
  - Inner table first paragraph Y
  - Body P-end Y

Compute deltas to characterize:
  - inner_table_top_y - outer_cell_top_y = effective inner-table offset
  - effect of outer cell topPadding on inner table position
  - effect of leading paragraph in outer cell
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "nested_table_y_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_nested_table_y.json")


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

        # Body paragraphs (top-level, not in tables)
        body_paras = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            try:
                in_table = p.Range.Information(12)  # wdWithInTable
                body_paras.append({
                    "i": i,
                    "y": round(p.Range.Information(6), 4),
                    "x": round(p.Range.Information(5), 4),
                    "in_table": bool(in_table),
                    "text": p.Range.Text.strip()[:30],
                })
            except Exception as e:
                body_paras.append({"i": i, "err": str(e)})

        # Tables enumeration
        tables = []
        for tk in range(1, wdoc.Tables.Count + 1):
            t = wdoc.Tables(tk)
            row1 = t.Rows(1)
            cell = t.Cell(1, 1)
            tables.append({
                "k": tk,
                "row1_y": round(row1.Range.Information(6), 4),
                "row1_x": round(row1.Range.Information(5), 4),
                "cell_x": round(cell.Range.Information(5), 4),
                "cell_y": round(cell.Range.Information(6), 4),
                "cell_text": cell.Range.Text.strip()[:30],
                "nested_count": cell.Tables.Count,
            })

        return {
            "path": os.path.basename(path),
            "body_paragraphs": body_paras,
            "tables": tables,
        }
    finally:
        wdoc.Close(False)


def parse_filename(name):
    m_top = re.search(r"top(\d+)", name)
    m_bot = re.search(r"bot(\d+)", name)
    has_lead = "_lead" in name
    return {
        "top_tw": int(m_top.group(1)) if m_top else None,
        "bot_tw": int(m_bot.group(1)) if m_bot else None,
        "has_lead": has_lead,
    }


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
                meta = parse_filename(os.path.basename(p))
                r.update(meta)
                results.append(r)
                print(f"\n  {os.path.basename(p)}: top_tw={meta.get('top_tw')} bot_tw={meta.get('bot_tw')} lead={meta.get('has_lead')}")
                for t in r["tables"]:
                    print(f"    Tbl{t['k']}: row1_y={t['row1_y']} cell_y={t['cell_y']} cell_x={t['cell_x']} nested={t['nested_count']}")
                for bp in r["body_paragraphs"][:8]:
                    if "err" in bp:
                        continue
                    print(f"    P{bp['i']:2}@in_table={bp['in_table']:>5}: y={bp['y']:>7} x={bp['x']:>6} '{bp['text']}'")
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
