"""§4.7b Mech 2 — wrap-budget × line-break interaction.

Goal: reverse-engineer Word's line-fit algorithm when Mech 2 is available.
Specifically: at what content_w threshold does Word
  (a) keep N chars on line + apply Mech 2 distribution
  (b) drop a char (N-1) + apply lighter Mech 2

Probes:
  P_yak3:  3 yak + filler CJK (low Mech 2 capacity ≈ 3 × 4pt = 12pt)
  P_yak6:  6 yak + filler CJK (mid capacity ≈ 24pt)
  P_yak12: 12 yak + filler CJK (high capacity ≈ 48pt)

For each probe, sweep cw in 0.5pt steps around the natural line width.
Identify:
  - Threshold cw where N → N-1 transition occurs
  - Slack at threshold (does it match `n_uncomp_yak × cap`?)
  - Multiple-drop thresholds (N-1 → N-2)

This pins:
  Q1: When Mech 2 capacity factors into wrap-fit
  Q2: Whether `width + n_yak × cap > avail` is allowed
  Q3: Drop-char trigger condition exactly
"""
import json, os, sys, time
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT = "ＭＳ 明朝"
SIZE = 12.0
RESULT_PATH = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m2_wrap_budget.json")

PROBES = {
    # Each probe is a fixed string with controlled # of yakumono.
    # Yak NOT triggering Mech 1: surrounded by CJK with single Mech-2-only neighbors.
    "P_yak3":  "漢漢漢漢漢「漢漢漢漢漢」漢漢漢漢漢、漢漢漢",   # 3 yak, 20 chars
    "P_yak6":  "漢漢「漢漢」漢漢「漢漢」漢漢、漢漢、漢漢漢",     # 6 yak, 20 chars
    "P_yak12": "「漢」漢「漢」漢「漢」漢、漢、漢、漢、漢漢漢",   # 12 yak, 20 chars (max density)
}


def measure_one(word, content_w, probe):
    page_w = content_w + 170
    d = word.Documents.Add()
    time.sleep(0.12)
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
        p.Alignment = 3   # justify
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        time.sleep(0.1)
        chars = d.Range().Characters
        xs = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                xs.append((t, float(c.Information(5)), float(c.Information(6))))
            except Exception:
                continue
        if not xs: return None
        y0 = xs[0][2]
        line1 = sorted([(t, x) for t, x, y in xs if abs(y - y0) < 0.5],
                       key=lambda v: v[1])
        advs = [(line1[i][0], round(line1[i + 1][1] - line1[i][1], 3))
                for i in range(len(line1) - 1)]
        # Yak compressed total
        yak_compressed = sum(12.0 - a for t, a in advs if t in "「」（）、。" and a < 12.0)
        return {
            "n_chars_line1": len(line1),
            "advances": advs,
            "yak_compression_total": round(yak_compressed, 3),
        }
    finally:
        try: d.Close(SaveChanges=False)
        except: pass


def sweep_probe(word, label, probe):
    natural = len(probe) * SIZE  # all chars 12pt fullwidth
    print(f"\n=== {label} probe={probe!r} natural={natural}pt ===")
    # Sweep cw in 1pt steps from natural+5 down to 0.5×natural
    # Focus on transitions
    results = []
    cw_values = []
    cw = natural + 5
    while cw >= natural * 0.5:
        cw_values.append(cw)
        cw -= 1.0
    for cw in cw_values:
        try:
            r = measure_one(word, cw, probe)
            if r is None:
                continue
            slack = max(0.0, natural - cw)
            print(f"  cw={cw:6.1f}  slack={slack:5.1f}  n_line1={r['n_chars_line1']:2d}  yak_compress_total={r['yak_compression_total']:5.2f}")
            results.append({
                "content_w": cw,
                "natural": natural,
                "slack_natural": slack,
                "n_chars_line1": r["n_chars_line1"],
                "advances": r["advances"],
                "yak_compression_total": r["yak_compression_total"],
            })
        except Exception as e:
            print(f"  cw={cw:6.1f} ERR: {e}")
            results.append({"content_w": cw, "error": str(e)})
    return results


def main():
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = {}
    try:
        for label, probe in PROBES.items():
            out[label] = {
                "probe": probe,
                "n_chars": len(probe),
                "n_yak": sum(1 for c in probe if c in "「」（）、。"),
                "results": sweep_probe(word, label, probe),
            }
    finally:
        try: word.Quit()
        except: pass

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n=== Threshold summary ===")
    for label, info in out.items():
        n_yak = info["n_yak"]
        results = info["results"]
        # Find transitions in n_chars_line1
        transitions = []
        prev_n = None
        for r in results:
            if "error" in r: continue
            n = r["n_chars_line1"]
            if prev_n is not None and n != prev_n:
                transitions.append((r["content_w"], prev_n, n, r["slack_natural"]))
            prev_n = n
        print(f"  {label} (n_yak={n_yak}):")
        for cw, prev, cur, slack in transitions:
            print(f"    cw={cw:6.1f} slack={slack:5.1f}  {prev} → {cur} chars")


if __name__ == "__main__":
    main()
