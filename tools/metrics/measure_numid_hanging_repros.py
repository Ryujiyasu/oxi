"""COM-measure the minimal repros built by build_numid_hanging_repros.py."""
import glob
import json
import os
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def measure(word, path):
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.2)
    entries = []
    try:
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            fmt = p.Format
            left = fmt.LeftIndent
            fli = fmt.FirstLineIndent
            lf = p.Range.ListFormat
            list_str = lf.ListString or ""
            list_value = lf.ListValue
            chars = p.Range.Characters
            if chars.Count == 0:
                continue
            c1 = chars(1)
            x = c1.Information(7)
            y = c1.Information(6)
            entries.append({
                "para": i,
                "left": left,
                "fli": fli,
                "marker": list_str,
                "list_value": list_value,
                "char1": c1.Text,
                "char1_x": x,
                "char1_y": y,
                "expected_at_left": left,
                "delta": x - left,
            })
    finally:
        doc.Close(SaveChanges=False)
    return entries


def main():
    repros = sorted(glob.glob(os.path.abspath("tools/metrics/numid_hanging_repro/NH_*.docx")))
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    all_results = {}
    try:
        for rpath in repros:
            label = os.path.splitext(os.path.basename(rpath))[0]
            print(f"\n=== {label} ===")
            entries = measure(word, rpath)
            all_results[label] = entries
            for e in entries:
                print(f"  P{e['para']} left={e['left']:.2f} fli={e['fli']:+.2f} "
                      f"marker={e['marker']!r} char1={e['char1']!r} "
                      f"x={e['char1_x']:.2f} expected={e['expected_at_left']:.2f} "
                      f"delta={e['delta']:+.2f}")
    finally:
        word.Quit()

    out = os.path.abspath("pipeline_data/numid_hanging_repro_measurements.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {out}")


if __name__ == "__main__":
    main()
