"""§4.7 Mech 2 trigger pair characterization.

Per Session 51 R0 (session_51_mechanism2_slack_algorithm.md):
  Mech 2 = justify-time slack-distribution. Fires at jc=both AND overflow.
  Tested only with `「」（）、。` chars in homogeneous CJK line.

Open questions (per user):
  1. Which char triggers fire (8.0 / 7.5 / 5.5pt usage)?
  2. "mid-line" position % threshold?
  3. Trigger pair list (vs Mech 1 Type A/B/C)?
  4. Overflow gating ("reactive" semantics)?

This script tests 8-10 character pairs in mid-line position with
controlled overflow (jc=both). Probe template:
  CJK×5 + [PAIR] + CJK×N (line natural width slightly exceeds content_w)

For each pair, measure per-char advance via Information(5) deltas, identify
which char(s) compressed and to what value.

Pairs tested (NOT firing Mech 1 Type A/B/C, so any compression must be Mech 2):
  「漢   - A→CJK (Type A bracket followed by CJK)
  漢「   - CJK→A
  」漢   - B→CJK (Type B bracket followed by CJK)
  漢」   - CJK→B
  、漢   - B→CJK (comma)
  漢、   - CJK→B
  。漢   - B→CJK (period)
  漢。   - CJK→B
  ）漢   - B→CJK (right paren)
  漢）   - CJK→B

Plus controls (Mech 1 SHOULD fire):
  」）   - B→B (Mech 1 → 5.5pt)
  ）（   - B→A (Mech 1 → 5.5pt)
"""
import json, os, sys, time
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT = "ＭＳ 明朝"
SIZE = 12.0  # 1 char = 12pt CJK fullwidth
RESULT_PATH = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m2_trigger_pairs.json")

# Test pairs. Each pair embedded at chars 6-7 of a 20-char probe.
# Probe = "漢"×5 + pair + "漢"×(20-5-2) = 5 + 2 + 13 = 20 chars
# Natural width = 20 × 12 = 240pt
# Set content_w = 236pt → 4pt slack (forces Mech 2 if applicable)
PAIRS = [
    # (label, pair_str, expected_mech, notes)
    ("A_CJK",   "「漢",   "Mech2?",   "A → CJK, no Mech 1 trigger"),
    ("CJK_A",   "漢「",   "Mech2?",   "CJK → A"),
    ("B_CJK",   "」漢",   "Mech2?",   "B → CJK"),
    ("CJK_B",   "漢」",   "Mech2?",   "CJK → B"),
    ("Bcomma_CJK",  "、漢", "Mech2?",  "B(comma) → CJK"),
    ("CJK_Bcomma",  "漢、", "Mech2?",  "CJK → B(comma)"),
    ("Bperiod_CJK", "。漢", "Mech2?",  "B(period) → CJK"),
    ("CJK_Bperiod", "漢。", "Mech2?",  "CJK → B(period)"),
    ("Bparen_CJK",  "）漢", "Mech2?",  "B(paren) → CJK"),
    ("CJK_Bparen",  "漢）", "Mech2?",  "CJK → B(paren)"),
    # Controls — Mech 1 should fire
    ("BB_Mech1", "」）",  "Mech1",   "Type B → B (Mech 1 fires)"),
    ("BA_Mech1", "）（",  "Mech1",   "Type B → A (Mech 1 fires)"),
]

# Probe layout: 5 CJK + pair + 13 CJK = 20 chars. Pair at positions 6-7.
def make_probe(pair: str) -> str:
    n_pre = 5
    n_post = 20 - n_pre - len(pair)
    return "漢" * n_pre + pair + "漢" * n_post


# Vary content_w to control slack. Natural=240pt for 20 CJK.
# Test slack values [0, 2, 4, 8].
CONTENT_WIDTHS = [240.0, 238.0, 236.0, 232.0]


def measure_one(word, content_w, probe):
    page_w = content_w + 170
    d = word.Documents.Add()
    time.sleep(0.2)
    try:
        ps = d.Sections(1).PageSetup
        ps.PageWidth = page_w
        ps.PageHeight = 841.9
        ps.LeftMargin = 85
        ps.RightMargin = 85
        ps.TopMargin = 72
        ps.BottomMargin = 72
        rng = d.Range()
        rng.Text = probe
        rng = d.Range()
        rng.Font.Name = FONT
        rng.Font.Size = SIZE
        p = d.Paragraphs(1)
        p.Alignment = 3   # wdAlignParagraphJustify
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        time.sleep(0.15)
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
        if not xs:
            return None
        y0 = xs[0][2]
        line1 = [(t, x) for t, x, y in xs if abs(y - y0) < 0.5]
        line1_sorted = sorted(line1, key=lambda t: t[1])
        advs = []
        for i in range(len(line1_sorted) - 1):
            advs.append((line1_sorted[i][0],
                         round(line1_sorted[i + 1][1] - line1_sorted[i][1], 3)))
        # Last char advance: estimate from line position vs content_w (skip)
        return {
            "n_chars_line1": len(line1_sorted),
            "advances": advs,
            "first_x": line1_sorted[0][1] if line1_sorted else None,
            "last_x": line1_sorted[-1][1] if line1_sorted else None,
        }
    finally:
        try:
            d.Close(SaveChanges=False)
        except Exception:
            pass


def main():
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, pair, exp, notes in PAIRS:
            probe = make_probe(pair)
            print(f"\n=== {label} pair={pair!r} ({notes}) ===")
            results[label] = {
                "pair": pair,
                "probe": probe,
                "expected_mech": exp,
                "notes": notes,
                "by_content_w": {},
            }
            for cw in CONTENT_WIDTHS:
                try:
                    r = measure_one(word, cw, probe)
                    if r is None:
                        continue
                    # Find the pair char advances:
                    # Pair is at positions 6-7 (1-indexed) in probe
                    # i.e. line1_sorted[5] and [6] (0-indexed)
                    pair_advs = []
                    for i, (t, a) in enumerate(r["advances"]):
                        if 4 <= i <= 6:  # check chars at probe pos 5/6/7
                            pair_advs.append((i, t, a))
                    natural = len(probe) * SIZE  # 20 × 12 = 240
                    slack = natural - cw
                    print(f"  [cw={cw} slack={slack:.1f}]  n_line1={r['n_chars_line1']}  pair_chars(idx 4-6): {pair_advs}")
                    results[label]["by_content_w"][f"cw{cw}"] = {
                        "content_w": cw,
                        "slack": slack,
                        "n_chars_line1": r["n_chars_line1"],
                        "advances_all": r["advances"],
                        "pair_advances": pair_advs,
                    }
                except Exception as e:
                    print(f"  [cw={cw}] ERR: {e}")
                    results[label]["by_content_w"][f"cw{cw}"] = {"error": str(e)}
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}")


if __name__ == "__main__":
    main()
