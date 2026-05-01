"""
Measure default-tab fixture: find x position of M1, M2, M3, M4, M5 markers.
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "default_tab_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_default_tab.json")


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


def measure_marker_x(wdoc, p, marker_text):
    """Find x of marker_text inside paragraph p."""
    text = p.Range.Text
    idx = text.find(marker_text)
    if idx < 0:
        return None
    sub = wdoc.Range(p.Range.Start + idx, p.Range.Start + idx + 1)
    try:
        return round(sub.Information(5), 4)
    except Exception:
        return None


def measure(word, path):
    wdoc = retry(lambda: word.Documents.Open(path))
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        paras = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            text = p.Range.Text.strip()[:60]
            marker = f"M{i}"
            mx = measure_marker_x(wdoc, p, marker)
            paras.append({
                "i": i,
                "text": text,
                "para_x": round(p.Range.Information(5), 4),
                "marker_x": mx,
            })
        return {
            "path": os.path.basename(path),
            "paragraphs": paras,
        }
    finally:
        wdoc.Close(False)


def parse_filename(name):
    m = re.search(r"DT_(?:default)?(\d+)tw", name)
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
                tw = parse_filename(os.path.basename(p))
                r["default_tab_tw"] = tw
                r["default_tab_pt"] = tw / 20.0 if tw else None
                results.append(r)
                print(f"\n  {os.path.basename(p)}: defaultTab = {tw}tw ({r['default_tab_pt']}pt)")
                for para in r["paragraphs"]:
                    print(f"    P{para['i']}: '{para['text']:40s}' marker_x={para['marker_x']}")
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
