"""Measure IC_* indent-chars resolution variants."""
import json, sys
from pathlib import Path
import win32com.client as w32

REPRO_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\ic_repro")
OUT = Path(r"C:\Users\ryuji\oxi-1\pipeline_data\ic_indent_chars_measurements.json")


def main():
    docs = sorted(REPRO_DIR.glob("IC_*.docx"))
    print(f"Measuring {len(docs)} docs...", file=sys.stderr)
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for d in docs:
            try:
                doc = word.Documents.Open(str(d.resolve()), ReadOnly=True)
            except Exception as e:
                results.append({"file": d.name, "error": f"open: {e}"})
                continue
            try:
                p = doc.Paragraphs(1)
                # Word reports LeftIndent in points
                left_indent_pt = p.Format.LeftIndent
                # Information(5) returns wdHorizontalPositionRelativeToTextBoundary (in pt)
                # First-character X position relative to page
                x_first_char = p.Range.Information(7)  # wdHorizontalPositionRelativeToPage
                # Page left margin = pgMar.left = 1134tw = 56.7pt
                # x_first_char - 56.7 = effective indent
                results.append({
                    "file": d.name,
                    "format_left_indent_pt": left_indent_pt,
                    "x_first_char_pt": x_first_char,
                    "effective_indent_pt": x_first_char - 56.7,
                })
                print(f"  done {d.name}", file=sys.stderr)
            except Exception as e:
                results.append({"file": d.name, "error": f"measure: {e}"})
            finally:
                try: doc.Close(SaveChanges=0)
                except Exception: pass
    finally:
        try: word.Quit()
        except Exception: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print()
    print(f"{'file':30s} {'fmt.LI':>8} {'x_first':>9} {'eff_ind':>9}")
    print("-" * 65)
    for r in results:
        if 'error' in r:
            print(f"  {r['file'][:28]:28s}  ERROR: {r['error']}")
            continue
        print(f"  {r['file'][:28]:28s}"
              f" {r['format_left_indent_pt']:>8.2f}"
              f" {r['x_first_char_pt']:>9.2f}"
              f" {r['effective_indent_pt']:>9.2f}")

    print()
    print("Reference: leftChars=100 (= 1 char). char_width interpretations:")
    print("  10pt run/pPr  -> indent = 10.0pt")
    print("  10.5pt run/pPr -> indent = 10.5pt")
    print("  14pt run/pPr  -> indent = 14.0pt")


if __name__ == "__main__":
    main()
