"""Measure M1A_* alignment dependency for Mech 1 ）（ pair.

For each docx, iterate paragraphs (5 alignment variants), measure per-char
advance via Information(5). Identify ） advance (= 5.5pt = Mech 1 fired
for B→A pair, or 10.5pt = no Mech 1).
"""
import json, sys, time
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\m1_alignment_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\m1_alignment_measurements.json")
ALIGN_NAMES = ["both", "left", "center", "right", "(no jc)"]


def measure_doc(word, docx_path):
    d = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    out = []
    try:
        for pi in range(1, d.Paragraphs.Count + 1):
            try:
                p = d.Paragraphs(pi)
                chars = p.Range.Characters
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
                if not xs: continue
                xs.sort(key=lambda v: v[1])
                advs = [(xs[i][0], round(xs[i+1][1] - xs[i][1], 3))
                        for i in range(len(xs) - 1)]
                # Look for the ） char — is it compressed?
                paren_advs = [(i, t, a) for i, (t, a) in enumerate(advs) if t == "）"]
                out.append({
                    "para_idx": pi,
                    "alignment_label": ALIGN_NAMES[pi-1] if pi-1 < len(ALIGN_NAMES) else f"p{pi}",
                    "alignment_int": p.Alignment,
                    "n_chars": len(xs),
                    "advances": advs,
                    "paren_advs": paren_advs,
                })
            except Exception as e:
                out.append({"para_idx": pi, "error": str(e)})
    finally:
        d.Close(SaveChanges=0)
    return out


def main():
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for d in sorted(REPRO_DIR.glob("M1A_*.docx")):
            print(f"\n=== {d.name} ===")
            try:
                results[d.name] = measure_doc(word, d)
                for r in results[d.name]:
                    if 'error' in r:
                        print(f"  p{r['para_idx']} ERROR: {r['error']}")
                        continue
                    pa = r['paren_advs']
                    pa_str = ", ".join(f"pos{i}={a}pt" for i,_,a in pa) if pa else "no )"
                    print(f"  {r['alignment_label']:>8s} (Align={r['alignment_int']}) ) advance: {pa_str}")
                    print(f"    advances: {r['advances']}")
            except Exception as e:
                results[d.name] = {"error": str(e)}
                print(f"  ERROR: {e}")
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
