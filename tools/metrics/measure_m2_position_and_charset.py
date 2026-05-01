"""§4.7 Mech 2 follow-up: position dependence + extended char set + jc=left control.

Round 1 (measure_m2_trigger_pairs.py) confirmed all 10 yakumono-CJK pairs
compress identically under jc=both with overflow. Now characterize:

  Q2: Does Mech 2 compression depend on where the yakumono is on the line?
       Test: pair at start (pos 1-2), middle (pos 10-11), end (pos 19-20)
  Q3: Which chars are compressible by Mech 2?
       Test broader yakumono set: ［］【】〔〕「」（）、。 plus non-yakumono
       (CJK, ASCII) controls.
  Q4: Overflow gating — does Mech 2 fire under jc=left?
       Test: same probe + jc=left, with overflow.
"""
import json, os, sys, time
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT = "ＭＳ 明朝"
SIZE = 12.0
RESULT_PATH = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m2_position_charset.json")


def make_probe(pair: str, position: str) -> str:
    """20-char probe, pair at start/mid/end position."""
    n_pre, n_post = {
        "start": (0, 18),
        "mid":   (9, 9),
        "end":   (18, 0),
    }[position]
    return "漢" * n_pre + pair + "漢" * n_post


def measure_one(word, content_w, probe, alignment=3):
    """alignment: 0=left, 1=center, 2=right, 3=justify"""
    page_w = content_w + 170
    d = word.Documents.Add()
    time.sleep(0.15)
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
        p.Alignment = alignment
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
        return {"n_chars_line1": len(line1), "advances": advs}
    finally:
        try: d.Close(SaveChanges=False)
        except: pass


def main():
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {"position": {}, "charset": {}, "alignment": {}}
    try:
        # ===== Q2 / Q4: position × alignment =====
        print("=== Q2/Q4: 「漢 pair × position × alignment ===")
        for pos in ["start", "mid", "end"]:
            probe = make_probe("「漢", pos)
            for align_name, align_val in [("justify", 3), ("left", 0)]:
                key = f"pos={pos}/{align_name}"
                results["position"][key] = {}
                for cw in [240.0, 236.0, 232.0]:  # slack 0, 4, 8
                    r = measure_one(word, cw, probe, align_val)
                    if r is None:
                        continue
                    # find positions of yakumono
                    yak_pos = [(i, t, a) for i, (t, a) in enumerate(r["advances"])
                               if t in "「」（）、。［］【】〔〕"]
                    print(f"  [{key} cw={cw}] n={r['n_chars_line1']}  yak: {yak_pos}")
                    results["position"][key][f"cw{cw}"] = {
                        "content_w": cw,
                        "n_chars_line1": r["n_chars_line1"],
                        "yak_positions": yak_pos,
                    }

        # ===== Q3: extended charset (non-Mech-1-trigger, single in mid-line) =====
        print("\n=== Q3: extended yakumono char set (mid-line, jc=both, slack=4) ===")
        EXTRA_CHARS = [
            ("LBracket_a",   "「漢"),
            ("RBracket_a",   "」漢"),
            ("LParen_a",     "（漢"),
            ("RParen_a",     "）漢"),
            ("LSqBr_a",      "［漢"),
            ("RSqBr_a",      "］漢"),
            ("LCurlyBr_a",   "【漢"),
            ("RCurlyBr_a",   "】漢"),
            ("LCornerBr_a",  "〔漢"),
            ("RCornerBr_a",  "〕漢"),
            ("Comma_a",      "、漢"),
            ("Period_a",     "。漢"),
            ("Em_a",         "―漢"),  # em-dash U+2015 (Type C, was font-dependent)
            ("Hyphen_a",     "-漢"),  # ASCII hyphen — should not compress
            ("CJK_pair",     "漢漢"),  # control — no yakumono
            ("Latin_a",      "a漢"),  # Latin → CJK control
        ]
        for label, pair in EXTRA_CHARS:
            probe = make_probe(pair, "mid")
            try:
                r = measure_one(word, 236.0, probe, 3)
                if r is None: continue
                yak_pos = [(i, t, a) for i, (t, a) in enumerate(r["advances"])
                           if t == pair[0] or t == pair[1]]
                print(f"  [{label} pair={pair!r}] n={r['n_chars_line1']}  pair: {yak_pos}")
                results["charset"][label] = {
                    "pair": pair,
                    "n_chars_line1": r["n_chars_line1"],
                    "pair_positions": yak_pos,
                }
            except Exception as e:
                print(f"  [{label}] ERR: {e}")
                results["charset"][label] = {"error": str(e)}

    finally:
        try: word.Quit()
        except: pass

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}")


if __name__ == "__main__":
    main()
