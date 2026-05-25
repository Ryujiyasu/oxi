"""COM-measure Word's wrap behavior on the v4 family of minimal repros.

For each fixture, report:
  - paragraph 1 START y
  - paragraph 1 END y (i.e., where text ends — paragraph mark position)
  - line_in_page for first and last char (= line count)
  - in_table flag

Result interpretation:
  v4 (baseline): if Word reports line count = 3 AND Oxi (already shown) = 3
    → MINIMAL REPRO STILL NOT REPRODUCING THE BUG. Oxi ALSO produces 3 lines.
       Need different repro structure that breaks the Oxi-3-line behavior
       on a doc that already matches 29dc6e's style chain.
  v4 (baseline): if Word reports 2 lines AND Oxi = 3
    → REVERSE bug: Oxi over-wraps where Word doesn't.
  v4 (baseline): if Word reports 3 lines but Oxi = 2 on the FULL 29dc6e (which
    we already know) → the trigger is something not yet captured in v4.

Run from repo root:
  python tools/metrics/com_measure_v4_family.py
"""
from __future__ import annotations

import os
import sys

FIXTURE_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures", "phase2_wrap_samples"
)

FIXTURES = [
    "v4_style_ac_inherited.docx",
    "v4a_wordwrap1.docx",
    "v4b_autospace1.docx",
    "v4c_no_kinsoku_comma.docx",
    "v4d_charspacing0.docx",
]


def measure(doc_path: str) -> dict:
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
        rs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            txt = (rng.Text or "").rstrip("\r\x07").strip()
            if not txt:
                continue
            cs = doc.Range(rng.Start, rng.Start)
            ce = doc.Range(rng.End - 1, rng.End - 1) if rng.End > rng.Start else cs
            try:
                y_s = cs.Information(6)
                line_s = cs.Information(10)
                y_e = ce.Information(6)
                line_e = ce.Information(10)
                page = cs.Information(1)
            except Exception as e:
                rs.append({"i": i, "error": str(e)})
                continue
            rs.append({
                "i": i, "page": page,
                "y_start": float(y_s), "line_start": int(line_s),
                "y_end": float(y_e), "line_end": int(line_e),
                "n_lines": int(line_e) - int(line_s) + 1,
                "text": txt[:40],
            })
        doc.Close(SaveChanges=False)
        return rs
    finally:
        word.Quit()


def main() -> None:
    for f in FIXTURES:
        path = os.path.join(FIXTURE_DIR, f)
        if not os.path.exists(path):
            print(f"MISSING: {f}")
            continue
        print(f"=== {f} ===")
        rs = measure(path)
        for r in rs:
            if "error" in r:
                print(f"  i={r['i']}: ERROR {r['error']}")
                continue
            print(
                f"  i={r['i']} page={r['page']} "
                f"y_start={r['y_start']:.2f} line_start={r['line_start']} "
                f"y_end={r['y_end']:.2f} line_end={r['line_end']} "
                f"n_lines={r['n_lines']} text={r['text']!r}"
            )
        print()


if __name__ == "__main__":
    main()
