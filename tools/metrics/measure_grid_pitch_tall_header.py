"""
Stage 2: measure grid-pitch tall-header fixtures (build_grid_pitch_tall_header_fixtures.py).
"""
import os
import json
import time
import glob

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "grid_pitch_tall_header_fixtures")
OUT_JSON = os.path.join(OUT_DIR, "ra2_grid_pitch_tall_header.json")


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

        # Header paragraphs
        hdr = sec.Headers(1)
        hdr_paras = []
        for i in range(1, hdr.Range.Paragraphs.Count + 1):
            p = hdr.Range.Paragraphs(i)
            hdr_paras.append({
                "i": i,
                "y": round(p.Range.Information(6), 4),
                "text": p.Range.Text.strip()[:20],
            })

        # Body paragraphs
        body_paras = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            try:
                body_paras.append({
                    "i": i,
                    "y": round(p.Range.Information(6), 4),
                    "x": round(p.Range.Information(5), 4),
                    "page": p.Range.Information(3),
                    "text": p.Range.Text.strip()[:20],
                })
            except Exception:
                pass

        rec = {
            "path": os.path.basename(path),
            "page_height": round(ps.PageHeight, 4),
            "top_margin": round(ps.TopMargin, 4),
            "header_distance": round(ps.HeaderDistance, 4),
            "header_paragraphs": hdr_paras,
            "body_paragraphs": body_paras,
        }
        if hdr_paras and body_paras:
            rec["hdr_p1_y"] = hdr_paras[0]["y"]
            rec["hdr_pN_y"] = hdr_paras[-1]["y"]
            rec["body_p1_y"] = body_paras[0]["y"]
            rec["delta_body_topmargin"] = round(body_paras[0]["y"] - ps.TopMargin, 4)
            if len(body_paras) >= 2:
                rec["body_p2_p1_diff"] = round(body_paras[1]["y"] - body_paras[0]["y"], 4)
        return rec
    finally:
        wdoc.Close(False)


def parse_filename(name):
    base = os.path.splitext(name)[0]
    parts = base.split("_")
    pitch = None
    hdr = None
    for p in parts:
        if p.startswith("pitch"):
            pitch_str = p[5:].rstrip("tw")
            pitch = int(pitch_str)
        elif p.startswith("hdr"):
            hdr = int(p[3:])
    return pitch, hdr


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
                pitch, hdr_lines = parse_filename(os.path.basename(p))
                r["pitch_tw"] = pitch
                r["pitch_pt"] = pitch / 20.0 if pitch else None
                r["hdr_lines"] = hdr_lines
                results.append(r)
                print(f"  {os.path.basename(p):28s} hdr_p1={r.get('hdr_p1_y')} hdr_pN={r.get('hdr_pN_y')} "
                      f"body_p1={r.get('body_p1_y')} delta_tm={r.get('delta_body_topmargin')} "
                      f"body_p2-p1={r.get('body_p2_p1_diff')}")
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
