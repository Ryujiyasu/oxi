"""
Measure §2.4 beforeLines fixtures — record gap = P2.y - P1.y.
"""
import os
import json
import time
import glob
import re

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "before_lines_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_before_lines.json")


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
    m_pitch = re.search(r"pitch(\d+)tw", name)
    m_bl = re.search(r"bl(\d+)", name)
    return (
        int(m_pitch.group(1)) if m_pitch else None,
        int(m_bl.group(1)) if m_bl else None,
    )


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
                pitch, bl = parse_filename(os.path.basename(p))
                r["pitch_tw"] = pitch
                r["before_lines"] = bl
                pitch_pt = pitch / 20.0 if pitch else None
                results.append(r)
                if len(r["paragraphs"]) >= 3:
                    p1y, p2y, p3y = (r["paragraphs"][0]["y"],
                                      r["paragraphs"][1]["y"],
                                      r["paragraphs"][2]["y"])
                    p1p2 = round(p2y - p1y, 4)
                    p2p3 = round(p3y - p2y, 4)
                    pred_before_pt = bl / 100 * pitch_pt if (bl and pitch_pt) else 0
                    pred_p1p2 = pitch_pt + pred_before_pt
                    print(f"  pitch={pitch}tw bl={bl:3} | p1y={p1y} p2y={p2y} p3y={p3y} | "
                          f"p1→p2={p1p2} p2→p3={p2p3} | pred(p1→p2)={pred_p1p2:.2f}")
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
