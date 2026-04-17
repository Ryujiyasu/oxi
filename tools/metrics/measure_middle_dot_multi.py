"""Measure Word's ・ (U+30FB) advance across multiple docs with MS Gothic/Mincho.

Hypothesis: ・ renders at ~0.75-0.79× em (NOT fullwidth) regardless of doc/size.
If confirmed on 3+ docs, the `fix/middle-dot-width` commit meets Ra protocol
evidence requirement (≥3 docs + minimal repro).

Strategy: search every body paragraph starting with ・ in these docs:
- d77a (already measured at 12pt)
- 683f (MS Gothic/Mincho bullets)
- 0e7a (MS Gothic bullets)
- b35 (MS Mincho)
"""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = [
    ("d77a58485f16_20240705_resources_data_outline_08", "d77a"),
    ("683fcab86e22_20230315_resources_data_contract_sample_02", "683f"),
    ("0e7af1ae8f21_20230331_resources_open_data_contract_sample_00", "0e7a"),
    ("b35123fe8efc_tokumei_08_01", "b35"),
]
DOC_DIR = r"tools\golden-test\documents\docx"


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for stem, label in DOCS:
            path = os.path.abspath(os.path.join(DOC_DIR, stem + ".docx"))
            if not os.path.exists(path):
                print(f"[skip] {path} missing", file=sys.stderr)
                continue
            doc = word.Documents.Open(path, ReadOnly=True); time.sleep(0.3)
            doc.Repaginate()
            dot_measurements = []
            n = doc.Paragraphs.Count
            print(f"[{label}] {n} paragraphs", file=sys.stderr)
            for i in range(1, min(n + 1, 400)):  # cap at 400 for speed
                try:
                    p = doc.Paragraphs(i)
                    txt = p.Range.Text[:3]
                    if "・" not in txt:
                        continue
                    pr = p.Range
                    # Get the ・ char's x and next char's x → advance
                    for cidx in range(1, min(pr.Characters.Count + 1, 5)):
                        ch = pr.Characters(cidx)
                        if ch.Text == "・":
                            x1 = ch.Information(5)
                            # Find next char for advance
                            if cidx + 1 <= pr.Characters.Count:
                                ch2 = pr.Characters(cidx + 1)
                                x2 = ch2.Information(5)
                                y1 = ch.Information(6)
                                y2 = ch2.Information(6)
                                if y1 == y2:  # same line
                                    adv = x2 - x1
                                    # Get font size — use rPr or character style
                                    fs = ch.Font.Size or 10.5
                                    dot_measurements.append({
                                        "para_idx": i, "position": cidx,
                                        "x1": round(x1, 2), "x2": round(x2, 2),
                                        "adv": round(adv, 2), "font_size": float(fs),
                                    })
                            break
                except Exception:
                    pass
            results.append({"doc": label, "measurements": dot_measurements})
            doc.Close(False)
            print(f"[{label}] found {len(dot_measurements)} ・ measurements", file=sys.stderr)
    finally:
        word.Quit()

    out = "pipeline_data/middle_dot_multi.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {out}")

    # Summary
    print("\n=== Summary ===")
    print(f"{'doc':>6} {'size':>6} {'n':>4} {'min':>6} {'med':>6} {'max':>6} {'expected_fw':>12}")
    for r in results:
        doc_label = r["doc"]
        # Group by font_size
        by_size = {}
        for m in r["measurements"]:
            by_size.setdefault(m["font_size"], []).append(m["adv"])
        for fs in sorted(by_size):
            advs = sorted(by_size[fs])
            n = len(advs)
            med = advs[n // 2]
            mn = advs[0]
            mx = advs[-1]
            fw_expected = fs  # fullwidth
            ratio = med / fs if fs else 0
            print(f"{doc_label:>6} {fs:>6.1f} {n:>4} {mn:>6.2f} {med:>6.2f} {mx:>6.2f} {fw_expected:>8.2f} ratio={ratio:.2%}")


if __name__ == "__main__":
    main()
