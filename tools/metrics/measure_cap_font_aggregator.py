"""Aggregator: measure existing docx in cap_font_sweep_repro/ and write JSON.

The original script crashed before saving. Re-measure all existing docx
with incremental save and Word restart per variant.
"""
import json, os, sys, time, subprocess, re
from pathlib import Path
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_font_sweep_repro")
RESULT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_font_sweep.json")
PROBE_LEN = 24
YAKUMONO = set("「")


def kill_word():
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(2)


def measure_one(path):
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(str(path), ReadOnly=True)
        time.sleep(0.2)
        try:
            chars = d.Range().Characters
            xs = []
            for ci in range(1, chars.Count + 1):
                try:
                    c = chars(ci)
                    t = c.Text
                    if t in ("\r", "\x07"): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6)),
                               float(c.Font.Size if c.Font.Size else 0)))
                except Exception: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        y0 = xs[0][2]
        line1 = sorted([(t, x, sz) for t, x, y, sz in xs if abs(y - y0) < 0.5],
                       key=lambda v: v[1])
        n_line1 = len(line1)
        total_comp = 0.0
        n_yak = 0
        n_yak_comp = 0
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            if t in YAKUMONO:
                n_yak += 1
                if sz > 0 and adv < sz * 0.99:
                    n_yak_comp += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_total_line1": n_yak,
            "n_yak_compressed": n_yak_comp,
        }
    finally:
        try: word.Quit()
        except: pass


def parse_label(name):
    """A_fs10.5_N3_cw244.8.docx → (suite, fs, n_yak, cw)"""
    m = re.match(r'(A|B)_fs([\d.]+)_N(\d+)_cw([\d.]+)\.docx', name)
    if not m: return None
    return m.group(1), float(m.group(2)), int(m.group(3)), float(m.group(4))


def main():
    docs = sorted(REPRO_DIR.glob("*.docx"))
    print(f"Found {len(docs)} docx", file=sys.stderr)
    out = {}
    for d in docs:
        meta = parse_label(d.name)
        if meta is None: continue
        suite, fs, n_yak, cw = meta
        natural = PROBE_LEN * fs
        slack = round(natural - cw, 1)
        kill_word()
        try:
            r = measure_one(d)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        key = f"{suite}_fs{fs}_N{n_yak}"
        if key not in out:
            out[key] = {
                "suite": suite,
                "font_size_pt": fs,
                "n_yak": n_yak,
                "natural": natural,
                "expected_cap": fs / 2,
                "sweep": [],
            }
        out[key]["sweep"].append({"cw": cw, "slack": slack, **r})
        n_line = r.get("n_chars_line1", "?")
        comp = r.get("total_compression_pt", "?")
        print(f"  {d.name:<35} slack={slack:>+6.1f}  n_line1={n_line}  comp={comp}", file=sys.stderr)
        # Incremental save
        RESULT.parent.mkdir(parents=True, exist_ok=True)
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {RESULT}", file=sys.stderr)

    # Summary
    print("\n========== SUMMARY ==========")
    print(f"{'key':<22} {'fs':>5} {'N':>3} {'exp_cap':>8} {'max_comp_at_24':>15} {'per_yak':>9} {'first_drop':>11}")
    for key in sorted(out.keys()):
        info = out[key]
        sweep = info["sweep"]
        max_comp = 0
        first_drop = None
        for r in sorted(sweep, key=lambda x: x.get("slack", 0)):
            n = r.get("n_chars_line1")
            if n == PROBE_LEN:
                tc = r.get("total_compression_pt", 0) or 0
                if tc > max_comp: max_comp = tc
            elif first_drop is None and isinstance(n, int) and n < PROBE_LEN and r.get("slack", -1) > 0:
                first_drop = r["slack"]
        per_yak = max_comp / info["n_yak"] if info["n_yak"] else 0
        print(f"{key:<22} {info['font_size_pt']:>5} {info['n_yak']:>3} {info['expected_cap']:>8.2f} "
              f"{max_comp:>15.2f} {per_yak:>9.3f} {str(first_drop):>11}")


if __name__ == "__main__":
    main()
