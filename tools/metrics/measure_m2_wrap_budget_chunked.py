"""Resume P_yak6 and P_yak12 sweeps with Word-restart on RPC failure."""
import json, os, sys, time, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FONT = "ＭＳ 明朝"
SIZE = 12.0
RESULT_PATH = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m2_wrap_budget.json")

PROBES_TO_RUN = {
    "P_yak6":  "漢漢「漢漢」漢漢「漢漢」漢漢、漢漢、漢漢漢",
    "P_yak12": "「漢」漢「漢」漢「漢」漢、漢、漢、漢、漢漢漢",
}


def kill_word():
    try:
        subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    except Exception: pass
    time.sleep(3)


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
        rng = d.Range(); rng.Text = probe
        rng = d.Range(); rng.Font.Name = FONT; rng.Font.Size = SIZE
        p = d.Paragraphs(1); p.Alignment = 3
        p.Format.SpaceBefore = 0; p.Format.SpaceAfter = 0
        time.sleep(0.1)
        chars = d.Range().Characters
        xs = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci); t = c.Text
                if t in ("\r","\x07"): continue
                xs.append((t, float(c.Information(5)), float(c.Information(6))))
            except Exception: continue
        if not xs: return None
        y0 = xs[0][2]
        line1 = sorted([(t, x) for t, x, y in xs if abs(y - y0) < 0.5], key=lambda v: v[1])
        advs = [(line1[i][0], round(line1[i+1][1] - line1[i][1], 3)) for i in range(len(line1)-1)]
        yak_compressed = sum(12.0 - a for t, a in advs if t in "「」（）、。" and a < 12.0)
        return {
            "n_chars_line1": len(line1), "advances": advs,
            "yak_compression_total": round(yak_compressed, 3),
        }
    finally:
        try: d.Close(SaveChanges=False)
        except: pass


def sweep_probe_chunked(label, probe, chunk_size=10):
    natural = len(probe) * SIZE
    print(f"\n=== {label} probe={probe!r} natural={natural}pt ===")
    cw_values = []
    cw = natural + 5
    while cw >= natural * 0.5:
        cw_values.append(cw); cw -= 1.0

    results = []
    chunk_idx = 0
    while chunk_idx < len(cw_values):
        kill_word()
        try:
            word = w32.Dispatch("Word.Application")
            word.Visible = False; word.DisplayAlerts = False
        except Exception as e:
            print(f"  Word start failed: {e}"); break
        try:
            for cw in cw_values[chunk_idx:chunk_idx + chunk_size]:
                try:
                    r = measure_one(word, cw, probe)
                    if r is None: continue
                    slack = max(0.0, natural - cw)
                    print(f"  cw={cw:6.1f} slack={slack:5.1f} n_line1={r['n_chars_line1']:2d} yak_compr_total={r['yak_compression_total']:5.2f}")
                    results.append({
                        "content_w": cw, "natural": natural, "slack_natural": slack,
                        "n_chars_line1": r["n_chars_line1"], "advances": r["advances"],
                        "yak_compression_total": r["yak_compression_total"],
                    })
                except Exception as e:
                    print(f"  cw={cw:6.1f} ERR: {e}")
                    results.append({"content_w": cw, "error": str(e)})
                    if "RPC" in str(e) or "サーバー" in str(e):
                        break
        finally:
            try: word.Quit()
            except: pass
        chunk_idx += chunk_size
    return results


def main():
    # Load existing partial results
    if os.path.exists(RESULT_PATH):
        with open(RESULT_PATH, "r", encoding="utf-8") as f:
            out = json.load(f)
    else:
        out = {}

    for label, probe in PROBES_TO_RUN.items():
        # Skip if already done well
        prev = out.get(label, {})
        prev_results = prev.get("results", [])
        prev_valid = [r for r in prev_results if "error" not in r]
        if len(prev_valid) >= 30:
            print(f"\n{label}: already has {len(prev_valid)} valid results — skip")
            continue
        out[label] = {
            "probe": probe,
            "n_chars": len(probe),
            "n_yak": sum(1 for c in probe if c in "「」（）、。"),
            "results": sweep_probe_chunked(label, probe),
        }

    os.makedirs(os.path.dirname(RESULT_PATH), exist_ok=True)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n=== Threshold summary ===")
    for label, info in out.items():
        n_yak = info["n_yak"]
        results = [r for r in info["results"] if "error" not in r]
        transitions = []
        prev_n = None
        for r in results:
            n = r["n_chars_line1"]
            if prev_n is not None and n != prev_n:
                transitions.append((r["content_w"], prev_n, n, r["slack_natural"], r["yak_compression_total"]))
            prev_n = n
        print(f"  {label} (n_yak={n_yak}, valid={len(results)}):")
        for cw, prev, cur, slack, yc in transitions:
            print(f"    cw={cw:6.1f} slack={slack:5.1f} {prev}→{cur}  yak_comp_total={yc}")


if __name__ == "__main__":
    main()
