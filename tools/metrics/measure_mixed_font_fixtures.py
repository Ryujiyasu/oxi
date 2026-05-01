"""
Measure mixed-font line height fixtures.

For each fixture: report Y of P1, P2, P3, P4. Compute:
  - P1→P2 gap = pure font A line height
  - P2→P3 gap = pure font B line height (since P3 starts after P2)
  - P3→P4 gap = MIXED line height (the value we want to verify)

Verify: P3→P4 gap == max(P1→P2 gap, P2→P3 gap) (the §1.7 max rule)
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "mixed_font_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_mixed_font.json")


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
        paras = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            paras.append({
                "i": i,
                "y": round(p.Range.Information(6), 4),
                "text": p.Range.Text.strip()[:30],
            })
        return {
            "path": os.path.basename(path),
            "paragraphs": paras,
        }
    finally:
        wdoc.Close(False)


def parse_filename(name):
    m = re.match(r"MF_([A-Za-z]+)(\d+)_([A-Za-z]+)(\d+)_(.+)\.docx", name)
    if m:
        return {
            "font_a": m.group(1),
            "size_a": int(m.group(2)),
            "font_b": m.group(3),
            "size_b": int(m.group(4)),
            "grid_label": m.group(5),
        }
    return {}


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
                if len(r["paragraphs"]) >= 4:
                    p1y, p2y, p3y, p4y = (r["paragraphs"][i]["y"] for i in range(4))
                    g_p1p2 = round(p2y - p1y, 4)
                    g_p2p3 = round(p3y - p2y, 4)
                    g_p3p4 = round(p4y - p3y, 4)
                    expected_max = max(g_p1p2, g_p2p3)
                    match = abs(g_p3p4 - expected_max) < 0.6
                    mark = "OK" if match else "FAIL"
                    print(f"  {meta.get('font_a')[:7]:7s}{meta.get('size_a')}/{meta.get('font_b')[:7]:7s}{meta.get('size_b')} "
                          f"({meta.get('grid_label')}): "
                          f"A={g_p1p2} B={g_p2p3} mix={g_p3p4} max={expected_max} {mark}")
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
