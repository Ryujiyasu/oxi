"""
Measure shape-wrap Y effects on body paragraphs.

For each fixture, opens in Word and records:
  - Each body paragraph's (i, x, y, page)
  - Shape position (via Shapes collection)

We can then identify:
  - Paragraphs above shape (y < shape_y) — should be unaffected
  - Paragraphs beside shape (shape_y <= y <= shape_y + shape_h)
  - Paragraphs below shape (y > shape_y + shape_h)

For each wrap type, the affected behavior differs.
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "shape_wrap_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_shape_wrap.json")


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

        # Shape info
        shapes = []
        try:
            for k in range(1, wdoc.Shapes.Count + 1):
                s = wdoc.Shapes(k)
                shapes.append({
                    "i": k,
                    "name": s.Name,
                    "left": round(s.Left, 4),
                    "top": round(s.Top, 4),
                    "width": round(s.Width, 4),
                    "height": round(s.Height, 4),
                    "wrap": str(s.WrapFormat.Type),
                })
        except Exception as e:
            shapes.append({"err": str(e)})

        # Body paragraph positions
        paras = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            try:
                paras.append({
                    "i": i,
                    "x": round(p.Range.Information(5), 4),
                    "y": round(p.Range.Information(6), 4),
                    "page": p.Range.Information(3),
                    "text": p.Range.Text.strip()[:20],
                })
            except Exception as e:
                paras.append({"i": i, "err": str(e)})
        return {
            "path": os.path.basename(path),
            "shapes": shapes,
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
                for s in r["shapes"]:
                    print(f"    shape {s.get('i')}: name={s.get('name')} pos=({s.get('left')},{s.get('top')}) "
                          f"size=({s.get('width')}x{s.get('height')}) wrap={s.get('wrap')}")
                # Print body paragraphs with annotation
                shape_top = r["shapes"][0]["top"] if r["shapes"] and "top" in r["shapes"][0] else None
                shape_bot = (shape_top + r["shapes"][0]["height"]) if shape_top is not None else None
                for para in r["paragraphs"][:30]:
                    annot = ""
                    if shape_top is not None and "y" in para:
                        y = para["y"]
                        if y < shape_top - 5:
                            annot = "(above)"
                        elif y > shape_bot + 5:
                            annot = "(below)"
                        else:
                            annot = "(beside)"
                    print(f"    P{para['i']:2}@p{para.get('page')} y={para.get('y'):>7} x={para.get('x'):>6} {annot:8s} '{para.get('text', '')}'")
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
