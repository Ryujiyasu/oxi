"""Measure per-char x in yakumono_context_repro for each text sample.

For each character, record:
  - context (body or cell)
  - sample id (B1..B5 / C1..C5)
  - char index within text
  - character
  - x position (HorizontalPosition)

Compare x-advances between chars → derive actual char widths.
Specifically check whether yakumono pairs get compressed identically
in body vs cell.
"""
import os, sys, time, json
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(r"pipeline_data\yakumono_context_repro.docx")
OUT = r"pipeline_data/yakumono_context_measurements.json"

wdHorizontalPositionRelativeToPage = 5  # Information constant


def measure_para_chars(pr, label):
    """Return list of (idx, char, x, y)."""
    n = pr.Characters.Count
    if n == 0:
        return []
    data = []
    for i in range(1, n + 1):
        try:
            ch = pr.Characters(i)
            x = ch.Information(wdHorizontalPositionRelativeToPage)
            y = ch.Information(6)
            c = ch.Text
            data.append({"idx": i, "char": c, "x": round(x, 2), "y": round(y, 2)})
        except Exception:
            pass
    return data


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {"body": [], "cell": []}
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True); time.sleep(0.3)
        doc.Repaginate()

        # Iterate ALL paragraphs (body + cell); classify by label prefix
        n = doc.Paragraphs.Count
        print(f"[info] total paragraphs: {n}", file=sys.stderr)
        for i in range(1, n + 1):
            try:
                p = doc.Paragraphs(i)
                txt = p.Range.Text.replace("\r", "").replace("\x07", "")
                if len(txt) > 3 and txt[0] in ("B", "C") and txt[1].isdigit() and txt[2] == ":":
                    label = txt[:2]
                    chars = measure_para_chars(p.Range, label)
                    which = "body" if label[0] == "B" else "cell"
                    results[which].append({"label": label, "text": txt[3:50], "chars": chars})
                    print(f"[{which}] {label}: {len(chars)} chars", file=sys.stderr)
            except Exception as e:
                print(f"[para {i}] err: {e}", file=sys.stderr)

        doc.Close(False)
    finally:
        word.Quit()

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUT}")

    # Analysis: compare body B1 to cell C1 per char
    print("\n=== Body vs Cell width comparison (per char advance) ===")
    body_by_label = {r["label"].replace("B", ""): r for r in results["body"]}
    cell_by_label = {r["label"].replace("C", ""): r for r in results["cell"]}
    for key in sorted(body_by_label.keys() & cell_by_label.keys()):
        b = body_by_label[key]["chars"]
        c = cell_by_label[key]["chars"]
        if len(b) != len(c):
            print(f"Sample {key}: char count differs body={len(b)} cell={len(c)}")
            continue
        print(f"\n--- Sample {key} ({len(b)} chars): {body_by_label[key]['text'][:40]}")
        print(f"{'#':>3} {'char':>4} {'b_x':>7} {'c_x':>7} {'b_adv':>7} {'c_adv':>7} {'diff':>6}")
        for i in range(len(b)):
            b_x = b[i]["x"]; c_x = c[i]["x"]
            b_adv = (b[i+1]["x"] - b[i]["x"]) if i+1 < len(b) else None
            c_adv = (c[i+1]["x"] - c[i]["x"]) if i+1 < len(c) else None
            diff = (c_adv - b_adv) if b_adv is not None and c_adv is not None else None
            mark = ""
            if diff is not None and abs(diff) > 0.5:
                mark = " ***"
            ba = f"{b_adv:.2f}" if b_adv is not None else "-"
            ca = f"{c_adv:.2f}" if c_adv is not None else "-"
            df = f"{diff:+.2f}" if diff is not None else "-"
            print(f"{i+1:3} {b[i]['char']:>4} {b_x:7.2f} {c_x:7.2f} {ba:>7} {ca:>7} {df:>6}{mark}")


if __name__ == "__main__":
    main()
