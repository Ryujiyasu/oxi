"""Measure Word's X position of the first-text-character in numbered list
paragraphs that use hanging indent.

Hypothesis (from pixel diff on e3c545 p.1):
  For paragraphs with numId + <w:ind left=L hanging=H>,
  Word places the list marker at (margin + L - H) and the first TEXT
  character at (margin + L) — i.e. the hanging applies to the marker but
  the text stays at `left`.

Oxi currently places BOTH marker and first-line text at (margin + L - H),
causing them to overlap (visible on e3c545 p.1 "3．基本的な考え方").

Method: use Range.Information(7) which returns the horizontal distance in
points from the left edge of the text column. Character 1 of a list
paragraph is the first body-text character (markers are rendered by Word
but not included in Range.Characters).

Usage: run from oxi-main root with bundled bottom-5 docs, or pass custom
doc path + para indices via CLI.
"""
import argparse
import json
import sys
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# (path, [(label, 1-based para index)])
DEFAULT_TARGETS = [
    (r"tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx",
     [
        ("1. はじめに", 3),
        ("2. 本資料の範囲", 7),
        ("3. 基本的な考え方", 14),
        ("4. 公開するデータ", 18),
     ]),
    (r"tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx",
     # Auto-detect paras with numId + hanging; or try common ones.
     []),
]


def measure_doc(word, path, targets, auto_scan=False):
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.3)
    result = []
    try:
        n_paras = doc.Paragraphs.Count

        if auto_scan:
            # Scan all paragraphs. For each with >0 LeftIndent and hanging
            # (FirstLineIndent<0), measure.
            scan_targets = []
            for i in range(1, min(n_paras + 1, 80)):
                p = doc.Paragraphs(i)
                # Check for ListFormat
                lf = p.Range.ListFormat
                if lf.ListType == 0:  # wdListNoNumbering
                    continue
                fmt = p.Format
                left = fmt.LeftIndent  # in points
                fli = fmt.FirstLineIndent  # in points (negative = hanging)
                if fli < 0 and left > 0:
                    text = p.Range.Text.strip()[:20]
                    scan_targets.append((f"auto:{text}", i))
            targets = scan_targets

        for label, idx in targets:
            if idx > n_paras:
                print(f"  skip {label}: idx={idx} > paras={n_paras}")
                continue
            p = doc.Paragraphs(idx)
            fmt = p.Format
            left = fmt.LeftIndent
            fli = fmt.FirstLineIndent
            range_text = p.Range.Text.strip()[:30]

            # First char X
            chars = p.Range.Characters
            if chars.Count == 0:
                continue
            c1 = chars(1)
            x_char1 = c1.Information(7)  # horizontal distance from text boundary (margin)
            y_char1 = c1.Information(6)
            ch_text = c1.Text

            # ListFormat info
            lf = p.Range.ListFormat
            list_type = lf.ListType
            list_str = lf.ListString  # rendered marker text like "3．" or "1."
            list_value = lf.ListValue

            entry = {
                "label": label,
                "idx": idx,
                "left_pt": left,
                "fli_pt": fli,
                "expected_marker_x_pt": left + fli,  # left - hanging_size
                "expected_text_x_pt_if_word": left,
                "expected_text_x_pt_if_oxi": left + fli,
                "char1": ch_text,
                "char1_x_pt": x_char1,
                "char1_y_pt": y_char1,
                "list_type": list_type,
                "list_str": list_str,
                "list_value": list_value,
                "para_text_preview": range_text,
            }
            result.append(entry)

            print(f"  Para {idx} '{range_text}':")
            print(f"    LeftIndent={left:.2f}pt  FirstLineIndent={fli:.2f}pt")
            print(f"    ListString='{list_str}'  ListValue={list_value}")
            print(f"    char[1]='{ch_text}' x={x_char1:.2f}pt y={y_char1:.2f}pt")
            print(f"    Expected text x if Word spec: {left:.2f}pt")
            print(f"    Measured text x:              {x_char1:.2f}pt")
            print(f"    Delta (measured - expected):  {x_char1 - left:+.2f}pt")
    finally:
        doc.Close(SaveChanges=False)
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--auto", action="store_true", help="Auto-scan all numId+hanging paras")
    ap.add_argument("--docs", nargs="*", default=None)
    ap.add_argument("--out", default="pipeline_data/numid_hanging_text_x.json")
    args = ap.parse_args()

    import os
    targets = DEFAULT_TARGETS
    if args.docs:
        targets = [(d, []) for d in args.docs]
        args.auto = True

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    all_results = {}
    try:
        for path, tlist in targets:
            abspath = os.path.abspath(path)
            print(f"\n=== {os.path.basename(abspath)} ===")
            if not os.path.exists(abspath):
                print("  NOT FOUND")
                continue
            auto = args.auto or not tlist
            data = measure_doc(word, abspath, tlist, auto_scan=auto)
            all_results[os.path.basename(abspath)] = data
    finally:
        word.Quit()

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {args.out}")


if __name__ == "__main__":
    main()
