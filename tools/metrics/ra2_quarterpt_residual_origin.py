"""
Ra2: 0.25pt residual physical origin investigation (§3.3).

Round 26 hypothesized Latin-floor/CJK-ceil for the ±0.25pt residual in LM=2
first-paragraph centering; refuted. Current spec hypothesis (un-confirmed):
the residual is **Word's pixel-snap of the absolute Y coordinate**
(topMargin + raw_offset → rounded to nearest 0.5pt grid).

This script tests that hypothesis directly. Strategy:

1. Pick a (font, size) that produces a quarter-pt expected offset, e.g.,
   TNR 18pt grid 18pt → expected (36 - 20.5)/2 = 7.75pt.
2. Vary topMargin in 0.25pt steps from 72.0 to 75.0.
3. Measure P0_y at each step.
4. Predict three models:
     M1: P0_y_pred = topMargin + 7.75 (no snap)
     M2: P0_y_pred = round_to_half(topMargin + 7.75)  (round-half-up)
     M3: P0_y_pred = round_to_half_even(topMargin + 7.75)  (banker's)
5. Determine which model matches measured P0_y.

We ALSO sweep multiple fonts (Calibri, TNR, Garamond, Times if available, MS
Mincho, MS Gothic, Yu Mincho) at sizes that produce quarter-pt offsets, to
confirm font-independence (if hypothesis holds).

Output:
  pipeline_data/ra2_quarterpt_residual_origin.json
"""
import os
import json
import time

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_quarterpt_residual_origin.json")


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


def make_doc_grid(word, *, top_margin, font, size, grid_pitch_tw):
    wdoc = retry(word.Documents.Add)
    sec = retry(lambda: wdoc.Sections(1))
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = top_margin
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    ps.LayoutMode = 2  # wdLayoutModeLineGrid
    ps.LinesPage = int(round(
        (ps.PageHeight - ps.TopMargin - ps.BottomMargin) * 20 / grid_pitch_tw
    ))

    wdoc.Content.Text = "P1\rP2"
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = font
        try:
            p.Range.Font.NameFarEast = font
        except Exception:
            pass
        p.Range.Font.Size = size
        p.Format.LineSpacingRule = 0  # single
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    wdoc.Repaginate()
    time.sleep(0.05)
    p1_y = round(wdoc.Paragraphs(1).Range.Information(6), 4)
    p2_y = round(wdoc.Paragraphs(2).Range.Information(6), 4)
    rec = {
        "top_margin": top_margin,
        "font": font,
        "size": size,
        "grid_pitch_tw": grid_pitch_tw,
        "grid_pitch_pt": grid_pitch_tw / 20.0,
        "p1_y": p1_y,
        "p2_y": p2_y,
        "p2_minus_p1": round(p2_y - p1_y, 4),
        "delta_p1_topmargin": round(p1_y - top_margin, 4),
    }
    wdoc.Close(False)
    return rec


def round_to_half(x):
    """Standard round half away from zero to 0.5pt."""
    return round(x * 2) / 2


def round_to_half_down(x):
    """Round half down (banker's-like)."""
    import math
    n = x * 2
    return math.floor(n + 0.5 - 1e-9) / 2  # half always floors


def round_to_half_even(x):
    """Banker's rounding to 0.5pt."""
    import math
    n = x * 2
    rounded = round(n)  # round-half-even (Python default)
    return rounded / 2


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    # Phase 1: TNR 18pt grid 18pt — expected raw offset 7.75 (quarter-pt)
    # Vary topMargin in 0.25pt steps to see which model matches.
    phase1_records = []
    print("=== Phase 1: TNR 18pt grid 18pt, topMargin sweep 72.00→75.00 by 0.25pt ===")
    for tm_quarter in range(0, 13):
        tm = 72.0 + tm_quarter * 0.25
        try:
            r = make_doc_grid(word, top_margin=tm, font="Times New Roman",
                              size=18, grid_pitch_tw=360)
            r["phase"] = "P1_TNR_18pt"
            phase1_records.append(r)
            # Compute predictions
            offset_exact = (r["p2_minus_p1"] - 20.5) / 2 if r["p2_minus_p1"] >= 20.5 else None
            # Actually offset = (P0_h - lm0_lh)/2, where P0_h = 18*ceil(20.5/18) = 18*2 = 36
            # but the gap p2-p1 is the LH for P0 in lines, which differs.
            # Easier: use known constants: offset_exact = (36 - 20.5) / 2 = 7.75
            predicted_M1 = tm + 7.75
            predicted_M2 = round_to_half(tm + 7.75)
            predicted_M3 = round_to_half_even(tm + 7.75)
            print(f"  tm={tm:6.2f} p1_y={r['p1_y']:7.2f} delta={r['delta_p1_topmargin']:+5.2f} "
                  f"M1={predicted_M1:7.2f} M2={predicted_M2:7.2f} M3={predicted_M3:7.2f}")
        except Exception as e:
            print(f"  ERR tm={tm}: {e}")

    # Phase 2: same sweep but on TNR 24pt grid 18pt — expected raw offset 4.25
    phase2_records = []
    print("\n=== Phase 2: TNR 24pt grid 18pt, topMargin sweep 72.00→75.00 by 0.25pt ===")
    for tm_quarter in range(0, 13):
        tm = 72.0 + tm_quarter * 0.25
        try:
            r = make_doc_grid(word, top_margin=tm, font="Times New Roman",
                              size=24, grid_pitch_tw=360)
            r["phase"] = "P2_TNR_24pt"
            phase2_records.append(r)
            predicted_M1 = tm + 4.25
            predicted_M2 = round_to_half(tm + 4.25)
            predicted_M3 = round_to_half_even(tm + 4.25)
            print(f"  tm={tm:6.2f} p1_y={r['p1_y']:7.2f} delta={r['delta_p1_topmargin']:+5.2f} "
                  f"M1={predicted_M1:7.2f} M2={predicted_M2:7.2f} M3={predicted_M3:7.2f}")
        except Exception as e:
            print(f"  ERR tm={tm}: {e}")

    # Phase 3: cross-font check — same topMargin, multiple fonts at sizes producing quarter offsets.
    # If hypothesis "snap is on absolute Y" holds, residual direction depends on
    # absolute Y, NOT font.
    phase3_records = []
    print("\n=== Phase 3: Cross-font quarter-pt cases at tm=72.00 ===")
    cross_cases = [
        # (font, size, grid_pitch_tw, expected_offset_when_quarter, lm0_lh_approx)
        ("Times New Roman", 18, 360, 7.75, 20.5),
        ("Times New Roman", 24, 360, 4.25, 27.5),
        ("Calibri",         18, 360, 8.75, 18.5),  # nat_lh ≈ 18.5 (per §8.2 table)
        ("Garamond",        13, 360, None, None),  # exploratory
        ("Garamond",        18, 360, None, None),
        ("MS Mincho",       14, 360, None, None),
        ("Yu Mincho",       14, 360, None, None),
    ]
    for font, size, pitch_tw, exp_off, exp_lm0 in cross_cases:
        try:
            r = make_doc_grid(word, top_margin=72.0, font=font, size=size,
                              grid_pitch_tw=pitch_tw)
            r["phase"] = "P3_cross_font"
            r["expected_offset"] = exp_off
            r["expected_lm0_lh"] = exp_lm0
            phase3_records.append(r)
            if exp_off is not None:
                m1 = 72 + exp_off
                m2 = round_to_half(72 + exp_off)
                print(f"  {font:18s} {size:5}pt: p1_y={r['p1_y']:7.2f} delta={r['delta_p1_topmargin']:+5.2f} "
                      f"M1={m1:7.2f} M2={m2:7.2f}")
            else:
                print(f"  {font:18s} {size:5}pt: p1_y={r['p1_y']:7.2f} delta={r['delta_p1_topmargin']:+5.2f}")
        except Exception as e:
            print(f"  ERR {font} {size}: {e}")

    all_records = phase1_records + phase2_records + phase3_records

    # Save before quit
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(all_records, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(all_records)} records to {OUT_JSON}")

    try:
        word.Quit()
    except Exception as e:
        print(f"  (word.Quit failed, ignoring: {e})")


if __name__ == "__main__":
    main()
