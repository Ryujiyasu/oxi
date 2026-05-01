"""§4.7 — Test interaction between adjacency compression (always-on)
and justify-time tightening.

Master's commit 1f8b5f2 ("yakumono architectural-validation CLOSED")
found yakumono compression is a Word LINE-WRAP HEURISTIC (reactive).
But our 4-font sweep found Type-A/B/C adjacency rules ARE applied
in left-aligned 4-char paragraphs (no line pressure, no justify).

Hypothesis: there are TWO mechanisms:
1. Adjacency rule (Type A/B/C) — always-on, intrinsic per-pair
2. Reactive line-fit tightening — additional compression when
   line overflows (and possibly justify)

This script tests:
  A. Same probe at jc=left vs jc=both (no overflow)
  B. Same probe at jc=both with content_width close to overflow
  C. Same probe at jc=both with content_width forcing overflow
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Probe chars known to compress under FINAL RULE (B→A pair)
# Build text that's exactly N chars wide so we can test under various
# content widths.

FONT = "ＭＳ 明朝"
SIZE = 12.0  # 12pt natural CJK width = 12pt → 30 chars per line at 360pt content

# We'll set page width so content_w ≈ N × 12pt for varying N
# Default A4 portrait: pgW=595.3, l=85, r=85 → content=425.3pt → 35.4 chars
# To test overflow vs fit, vary right margin:

PROBES = [
    # Each probe is a string we measure under different layouts.
    ("close_open_pair",
     "漢" * 8 + "」（" + "漢" * 8 + "」（" + "漢" * 8),  # ~30 chars wide
    ("comma_close_pair",
     "漢" * 5 + "、）" + "漢" * 5 + "、）" + "漢" * 5),  # 19 chars
    ("paren_chain",
     "漢" + "（（（（" + "漢" + "））））" + "漢"),  # 11 chars
]

# jc options:
#   0 = wdAlignParagraphLeft
#   1 = wdAlignParagraphCenter
#   2 = wdAlignParagraphRight
#   3 = wdAlignParagraphJustify
LAYOUTS = [
    ("wide_left",       0, 595.3, 85, 85),   # A4 portrait left-align
    ("wide_justify",    3, 595.3, 85, 85),   # A4 portrait justify
    ("narrow_left",     0, 595.3, 85, 280),  # narrow → forces wrap
    ("narrow_justify",  3, 595.3, 85, 280),  # narrow + justify
]

RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, text in PROBES:
            results[label] = {}
            for layout_name, jc, pgW, lm, rm in LAYOUTS:
                try:
                    d = word.Documents.Add()
                    time.sleep(0.3)
                    sec = d.Sections(1)
                    ps = sec.PageSetup
                    ps.PageWidth = pgW
                    ps.PageHeight = 841.9
                    ps.LeftMargin = lm
                    ps.RightMargin = rm
                    ps.TopMargin = 72
                    ps.BottomMargin = 72
                    rng = d.Range()
                    rng.Text = text
                    rng = d.Range()
                    rng.Font.Name = FONT
                    rng.Font.Size = SIZE
                    p = d.Paragraphs(1)
                    p.Alignment = jc
                    p.Format.SpaceBefore = 0
                    p.Format.SpaceAfter = 0
                    time.sleep(0.2)
                    chars = d.Range().Characters
                    xs = []
                    for ci in range(1, chars.Count + 1):
                        try:
                            c = chars(ci)
                            t = c.Text
                            if t in ("\r", "\x07"):
                                continue
                            xs.append((t,
                                       float(c.Information(5)),
                                       float(c.Information(6))))
                        except Exception:
                            continue
                    d.Close(SaveChanges=False)
                    # Group by line (Information(6) y)
                    lines = {}
                    for ch, x, y in xs:
                        lines.setdefault(round(y, 1), []).append((ch, x))
                    line_data = []
                    for y in sorted(lines.keys()):
                        sorted_chars = sorted(lines[y], key=lambda t: t[1])
                        advs = []
                        for i in range(len(sorted_chars) - 1):
                            advs.append((sorted_chars[i][0],
                                         round(sorted_chars[i + 1][1]
                                               - sorted_chars[i][1], 4)))
                        if sorted_chars:
                            advs.append((sorted_chars[-1][0], None))
                        line_data.append({
                            "y": y,
                            "n_chars": len(sorted_chars),
                            "first_x": sorted_chars[0][1] if sorted_chars else None,
                            "last_x": sorted_chars[-1][1] if sorted_chars else None,
                            "advs": advs,
                        })
                    results[label][layout_name] = {
                        "jc": jc,
                        "content_w": round(pgW - lm - rm, 2),
                        "lines": line_data,
                    }
                    print(f"[{label}][{layout_name}] content_w="
                          f"{round(pgW-lm-rm,2)} jc={jc} "
                          f"lines={len(line_data)}", flush=True)
                    for ln in line_data:
                        # Find compressed yakumono in line
                        compressed = [(c, a) for c, a in ln["advs"][:-1]
                                      if a is not None and a < SIZE * 0.7
                                      and c in "、。」』）】〕｝〉》］，．"
                                      "「『（【〔｛〈《［—"]
                        full = [(c, a) for c, a in ln["advs"][:-1]
                                if a is not None and a >= SIZE * 0.9
                                and c in "、。」』）】〕｝〉》］，．"
                                "「『（【〔｛〈《［—"]
                        print(f"  y={ln['y']} n={ln['n_chars']} "
                              f"compr={compressed} full={full}", flush=True)
                except Exception as e:
                    results[label][layout_name] = {"error": str(e)}
                    print(f"[{label}][{layout_name}] ERR: {e}", flush=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    if os.path.exists(RESULT_PATH):
        try:
            with open(RESULT_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    else:
        existing = {}
    existing["yakumono_justify_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
