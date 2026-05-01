"""§4.7b Mech 1 → Mech 2 precedence interaction.

Question: when a yakumono is already Mech-1-compressed (advance = fontSize/2),
does Mech 2 ALSO redistribute slack onto it? Or only onto uncompressed
yakumono on the line?

Probe template: contains BOTH Mech 1 firing pair AND uncompressed yakumono.

Probe: "漢漢漢」）漢漢「漢漢漢" (11 chars)
  - 」 followed by ）: Mech 1 B→B fires → 」=fontSize/2
  - ）   preceded by 」 but followed by 漢 (CJK): NO Mech 1 (B→CJK)
  - 「 between CJK: NO Mech 1 (A surrounded by CJK)

Mech 1 produces:
  natural = 11 × 12 = 132pt; minus 」 saving 6pt → 126pt
  Each char advance: 12,12,12, 6 (」), 12, 12, 12, 12 (「), 12, 12, 12

For Mech 2: vary content_w to control slack:
  cw = 126 (slack=0): no Mech 2
  cw = 124 (slack=2): Mech 2 fires with 2pt slack
  cw = 122 (slack=4): Mech 2 fires with 4pt slack

Measure each yakumono's advance under each slack:
  - If 」 stays at 6.0pt → Mech 2 does NOT touch Mech-1-compressed yak
  - If 」 < 6.0pt → Mech 2 DOES include Mech-1-compressed yak in distribution
"""
import json, os, sys, time
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT = "ＭＳ 明朝"
SIZE = 12.0
PROBE = "漢漢漢」）漢漢「漢漢漢"  # 11 chars
RESULT_PATH = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m1_m2_precedence.json")


def measure_one(word, content_w, alignment=3):
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
        rng.Text = PROBE
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
    results = {}

    # Sweep content_w to vary slack post-Mech1
    # post-Mech1 natural = 11×12 - 6 (」 compressed) = 126pt
    CONTENT_WIDTHS = [200.0, 132.0, 126.0, 125.0, 124.0, 123.0, 122.0, 120.0, 118.0]
    try:
        for cw in CONTENT_WIDTHS:
            try:
                r = measure_one(word, cw, 3)
                if r is None: continue
                # Identify yakumono advances by char and position
                yak_advs = [(i, t, a) for i, (t, a) in enumerate(r["advances"])
                            if t in "「」（）、。"]
                results[f"cw{cw}"] = {
                    "content_w": cw,
                    "n_chars_line1": r["n_chars_line1"],
                    "advances": r["advances"],
                    "yak_positions": yak_advs,
                }
                print(f"cw={cw}  n_line1={r['n_chars_line1']}  yak: {yak_advs}")
            except Exception as e:
                results[f"cw{cw}"] = {"error": str(e)}
                print(f"cw={cw} ERR: {e}")
    finally:
        try: word.Quit()
        except: pass

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}")


if __name__ == "__main__":
    main()
