"""§4.7 Mech 2 — exact position threshold for "mid-line" gating.

Round 1+2 showed:
  pos=start (yak at line pos 1) — Mech 2 NOT fire (Word drops char)
  pos=mid (yak at line pos 10) — Mech 2 fires
  pos=end (yak at line pos 19) — Mech 2 fires

Sweep yak position 1..19 (each producing 20-char probe) at jc=both
content_w=236 (slack=4pt). Find the threshold where Mech 2 starts firing.
"""
import json, os, sys, time
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT = "ＭＳ 明朝"
SIZE = 12.0
RESULT_PATH = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m2_position_sweep.json")


def make_probe(yak_pos: int) -> str:
    """yak_pos = 1-indexed position of 「 in 20-char line."""
    return "漢" * (yak_pos - 1) + "「" + "漢" * (20 - yak_pos)


def measure_one(word, content_w, probe):
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
        p.Alignment = 3
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
    results = {}
    try:
        for yak_pos in range(1, 21):
            probe = make_probe(yak_pos)
            try:
                r = measure_one(word, 236.0, probe)
                if r is None: continue
                yak_advs = [(i, t, a) for i, (t, a) in enumerate(r["advances"])
                            if t == "「"]
                results[f"yak_pos{yak_pos}"] = {
                    "yak_pos": yak_pos,
                    "n_chars_line1": r["n_chars_line1"],
                    "yak_advs": yak_advs,
                }
                yak_a = yak_advs[0][2] if yak_advs else "N/A"
                fired = "✓ M2 FIRED" if (yak_advs and yak_advs[0][2] < 12.0) else "✗ no compression"
                print(f"  yak_pos={yak_pos:2d}  n_line1={r['n_chars_line1']:2d}  yak_adv={yak_a}  {fired}")
            except Exception as e:
                results[f"yak_pos{yak_pos}"] = {"error": str(e)}
                print(f"  yak_pos={yak_pos:2d} ERR: {e}")
    finally:
        try: word.Quit()
        except: pass

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}")


if __name__ == "__main__":
    main()
