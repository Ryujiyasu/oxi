"""
Stage 2 of footnote multi-line decomposition: measure pre-authored fixtures.

Reads .docx fixtures from output/footnote_multiline_fixtures/ (built by
build_footnote_multiline_fixtures.py), opens each in Word, records
footnote position. Outputs JSON.
"""
import os
import json
import time
import glob

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "footnote_multiline_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_footnote_multiline_decompose.json")


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
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        page_h = float(ps.PageHeight)
        bm = float(ps.BottomMargin)

        fns = wdoc.Footnotes
        fn_records = []
        for k in range(1, fns.Count + 1):
            fn = fns(k)
            r = fn.Range
            fn_records.append({
                "i": k,
                "y": round(r.Information(6), 4),
                "x": round(r.Information(5), 4),
                "page": r.Information(3),
                "text_len": len(r.Text),
            })

        rec = {
            "path": os.path.basename(path),
            "page_height": round(page_h, 4),
            "bottom_margin": round(bm, 4),
            "footnotes": fn_records,
        }
        if fn_records:
            fn1 = fn_records[0]
            rec["fn1_y"] = fn1["y"]
            rec["block_h"] = round((page_h - bm) - fn1["y"], 4)
        return rec
    finally:
        wdoc.Close(False)


def parse_filename(name):
    """Extract vl, sz from FN_vl{N}_sz{X}.docx."""
    base = os.path.splitext(name)[0]
    parts = base.split("_")
    vl = None
    sz = None
    for p in parts:
        if p.startswith("vl"):
            vl = int(p[2:])
        elif p.startswith("sz"):
            sz_str = p[2:].replace("p", ".")
            sz = float(sz_str)
    return vl, sz


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
                vl, sz = parse_filename(os.path.basename(p))
                r["vl_target"] = vl
                r["fn_font_size"] = sz
                results.append(r)
                bh = r.get("block_h", "?")
                print(f"  {os.path.basename(p):28s} vl={vl} sz={sz}pt block_h={bh}")
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
