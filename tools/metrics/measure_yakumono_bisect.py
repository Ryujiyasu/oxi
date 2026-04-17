"""Measure yakumono advances in each bisect variant.

For each variant, open the docx in Word, find the test paragraph (starts
with "VN: ・利用規約名..."), measure per-char x, report ・ advance and
・ + CJK-ideograph distance.
"""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

VARIANTS = ["V1", "V2", "V3", "V4", "V5", "V6"]
DOC_DIR = os.path.abspath(r"pipeline_data")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for v in VARIANTS:
            path = os.path.join(DOC_DIR, f"yakumono_bisect_{v}.docx")
            if not os.path.exists(path):
                print(f"[{v}] missing", file=sys.stderr); continue
            doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
            doc.Repaginate()
            measurements = []
            for pi in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(pi)
                txt = p.Range.Text.replace("\r","").replace("\x07","")
                if txt.startswith(v + ":") and "・利用規約名" in txt:
                    pr = p.Range
                    n = pr.Characters.Count
                    for ci in range(1, n + 1):
                        try:
                            ch = pr.Characters(ci)
                            x = ch.Information(5)
                            y = ch.Information(6)
                            c = ch.Text
                            measurements.append({"ci": ci, "c": c, "x": round(x,2), "y": round(y,2)})
                        except Exception:
                            pass
                    break
            results[v] = measurements
            doc.Close(False)
            print(f"[{v}] {len(measurements)} chars measured", file=sys.stderr)
            # Save progress after each variant
            with open("pipeline_data/yakumono_bisect_results.json", "w", encoding="utf-8") as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
    finally:
        word.Quit()

    # Analyze: find first ・ and its advance
    print("\n=== Per-variant ・ advance + critical char patterns ===")
    print(f"{'V':>3} {'・_x':>8} {'next_x':>8} {'・adv':>8} {'last_char':>10} {'last_x':>8} {'line_end_x':>10} {'total_chars':>12}")
    for v in VARIANTS:
        m = results.get(v, [])
        if not m: continue
        # find ・
        dot = None; next_ch = None
        for i, mm in enumerate(m):
            if mm["c"] == "・" and mm["ci"] >= 5:  # skip label prefix
                dot = mm
                if i + 1 < len(m):
                    next_ch = m[i + 1]
                break
        if dot and next_ch and dot["y"] == next_ch["y"]:
            adv = next_ch["x"] - dot["x"]
        else:
            adv = None
        # line info
        y_groups = {}
        for mm in m:
            y_groups.setdefault(mm["y"], []).append(mm)
        # Use dot's line
        if dot:
            line_chars = y_groups.get(dot["y"], [])
            last_ch = max(line_chars, key=lambda x: x["x"])
            last_x = last_ch["x"]
            last_char = last_ch["c"]
            total_in_line = len(line_chars)
        else:
            last_x = None; last_char = None; total_in_line = None
        adv_s = f"{adv:.2f}" if adv is not None else "-"
        lx_s = f"{last_x:.2f}" if last_x is not None else "-"
        lc_s = repr(last_char) if last_char else "-"
        print(f"{v:>3} {dot['x']:>8.2f} {next_ch['x']:>8.2f} {adv_s:>8} {lc_s:>10} {lx_s:>8} {total_in_line if total_in_line else '-':>12}")

    # Full per-char table for key variants (V1 vs V6)
    print("\n=== V1 vs V6 per-char advance diff ===")
    v1 = results.get("V1", [])
    v6 = results.get("V6", [])
    if v1 and v6:
        print(f"{'#':>3} {'char':>4} {'V1_x':>8} {'V6_x':>8} {'V1_adv':>8} {'V6_adv':>8} {'diff':>7}")
        prev_v1 = None; prev_v6 = None
        for i in range(min(len(v1), len(v6))):
            a = v1[i]; b = v6[i]
            v1a = f"{a['x']-prev_v1['x']:.2f}" if prev_v1 and a['y']==prev_v1['y'] else "WRAP"
            v6a = f"{b['x']-prev_v6['x']:.2f}" if prev_v6 and b['y']==prev_v6['y'] else "WRAP"
            try:
                diff = (b['x']-prev_v6['x']) - (a['x']-prev_v1['x']) if prev_v6 and prev_v1 and b['y']==prev_v6['y'] and a['y']==prev_v1['y'] else None
                diff_s = f"{diff:+.2f}" if diff is not None else "-"
            except:
                diff_s = "-"
            marker = " *" if diff_s.startswith("-") or diff_s.startswith("+") and diff_s != "+0.00" else ""
            print(f"{i+1:>3} {a['c']:>4} {a['x']:>8.2f} {b['x']:>8.2f} {v1a:>8} {v6a:>8} {diff_s:>7}{marker}")
            prev_v1 = a; prev_v6 = b

    with open("pipeline_data/yakumono_bisect_results.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] pipeline_data/yakumono_bisect_results.json")


if __name__ == "__main__":
    main()
