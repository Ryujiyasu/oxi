"""§4.7 — Characterize Mechanism 2 (justify-time additional compression)
slack-distribution algorithm.

Earlier finding: in narrow_justify, （ compressed to 7.0 and 7.5pt
(non-uniform). Need to determine Word's slack-distribution algorithm.

This script:
1. Holds yakumono content constant, varies content_w to control slack
2. Measures per-char advance to see compression distribution
3. Tests char priority: which yakumono compresses first? In what order?

Probe text: `漢漢漢漢漢漢漢」（漢漢漢漢漢漢漢漢漢漢」（漢漢漢` (24 chars)
At MS Mincho 12pt natural width = 24×12 = 288pt + autoSpaceDE? 0 (no Latin).
Actually with 」（ pairs: 」 compresses to 6pt (Mech 1), （ stays 12pt
(no Mech 1 trigger, A→CJK).

Natural compressed line width:
  21 × 12 (CJK) + 2 × 6 (compressed 」) + 2 × 12 (full （) = 252+12+24 = 288pt?
  Wait let me recount: 24 chars total.

Actually simpler: I'll use 30 chars CJK with periodic 「」 pairs to enable
Type-A/B/C compression and have many candidates for Mechanism 2 priority.

Probe: `漢×3 + 「漢漢漢」 + 漢×3 + 「漢漢漢」 + 漢×3 + 、 + 漢×3 + 、 + 漢×3`
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Probe: include various yakumono types
PROBE_TEXT = (
    "漢漢漢「漢漢漢」漢漢漢「漢漢漢」漢漢漢、漢漢漢、漢漢漢"
)
# Length: 3+1+3+1+3+1+3+1+3+1+3+1+3 = 27 chars

FONT = "ＭＳ 明朝"
SIZE = 12.0

# Sweep content_w from "fits comfortably" to "very tight"
# Natural width: 27 × 12 = 324pt
# After Mech 1: 」 compresses (B→CJK no, B→「 -> hmm let's check)
# Actually 「 in our text: prev=漢, next=漢漢漢」 — A→CJK no compress
# 」: prev=漢, next=漢 — single B between CJK = no compress
# 、: prev=漢, next=漢 — single B between CJK = no compress
# So after Mech 1, NO compression. All 27 chars natural = 324pt.
# We need to compress via Mech 2 only.

CONTENT_WIDTHS = [
    400.0,  # plenty of room, no overflow
    330.0,  # slight overflow, ~6pt slack needed
    324.0,  # exactly natural width
    320.0,  # 4pt slack
    310.0,  # 14pt slack
    300.0,  # 24pt slack
    290.0,  # 34pt slack
    280.0,  # 44pt slack — heavy compression needed
]

RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for content_w in CONTENT_WIDTHS:
            # PageWidth must accommodate content_w + margins
            page_w = content_w + 170  # 85pt margins each side
            try:
                d = word.Documents.Add()
                time.sleep(0.3)
                sec = d.Sections(1)
                ps = sec.PageSetup
                ps.PageWidth = page_w
                ps.PageHeight = 841.9
                ps.LeftMargin = 85
                ps.RightMargin = 85
                ps.TopMargin = 72
                ps.BottomMargin = 72
                rng = d.Range()
                rng.Text = PROBE_TEXT
                rng = d.Range()
                rng.Font.Name = FONT
                rng.Font.Size = SIZE
                p = d.Paragraphs(1)
                p.Alignment = 3  # justify
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
                # Get line 1 only
                if not xs:
                    continue
                y0 = xs[0][2]
                line1 = [(c, x) for c, x, y in xs if abs(y - y0) < 0.5]
                line1_sorted = sorted(line1, key=lambda t: t[1])
                advs = []
                for i in range(len(line1_sorted) - 1):
                    advs.append((line1_sorted[i][0],
                                 round(line1_sorted[i + 1][1]
                                       - line1_sorted[i][1], 4)))
                results[f"cw{content_w}"] = {
                    "content_w": content_w,
                    "line1_advances": advs,
                    "line1_n": len(line1_sorted),
                }
                # Print summary: which yakumono got compressed?
                yak_advs = [(c, a) for c, a in advs
                            if c in "「」（）、。"]
                print(f"\n[cw={content_w}] line1 has {len(line1_sorted)} chars")
                print(f"  yakumono advances: {yak_advs}")
                # Per-char-type summary
                for ch_type in ["「", "」", "（", "）", "、", "。"]:
                    matches = [a for c, a in advs if c == ch_type]
                    if matches:
                        print(f"  {ch_type!r}: {matches}")
            except Exception as e:
                results[f"cw{content_w}"] = {"error": str(e)}
                print(f"[cw={content_w}] ERR: {e}", flush=True)
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
    existing["mechanism2_slack_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
