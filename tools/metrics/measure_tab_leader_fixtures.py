"""
Measure tab leader fixtures.

For each paragraph in each fixture:
  - Find the position of "Item" (start) → x_start
  - Find the position of the digits at the end (e.g., "99") → x_after_tab
  - Document the gap (= leader region)

We can't directly query "leader char positions" via COM, but we can document:
  - What character Word uses for each leader keyword (via XML inspection)
  - Where the after-tab text lands (confirms tab alignment)
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "tab_leader_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_tab_leader.json")


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


def measure_paragraph(p):
    """Find x of paragraph start, x of the last text run after the tab."""
    rng = p.Range
    start_x = round(rng.Information(5), 4)
    start_y = round(rng.Information(6), 4)

    # Iterate words to find the position right after the tab
    # The text is "Item leader=X<TAB>99\r"
    # Find the digit "9" position: select Range starting where we expect
    text = rng.Text
    after_tab_x = None
    if "99" in text:
        idx = text.rfind("99")
        # Build a sub-range for that position
        sub = p.Range.Document.Range(rng.Start + idx, rng.Start + idx + 1)
        try:
            after_tab_x = round(sub.Information(5), 4)
        except Exception:
            after_tab_x = None

    return {
        "text": text.strip()[:60],
        "start_x": start_x,
        "start_y": start_y,
        "after_tab_x": after_tab_x,
    }


def measure(word, path):
    wdoc = retry(lambda: word.Documents.Open(path))
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        paras = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            try:
                d = measure_paragraph(p)
                d["i"] = i
                paras.append(d)
            except Exception as e:
                paras.append({"i": i, "error": str(e)})
        return {
            "path": os.path.basename(path),
            "paragraphs": paras,
        }
    finally:
        wdoc.Close(False)


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
                results.append(r)
                print(f"\n  {os.path.basename(p)}")
                for para in r["paragraphs"]:
                    if "error" in para:
                        print(f"    P{para['i']}: ERR {para['error']}")
                    else:
                        print(f"    P{para['i']}: '{para['text']:50s}' "
                              f"start_x={para['start_x']} after_tab_x={para['after_tab_x']}")
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
