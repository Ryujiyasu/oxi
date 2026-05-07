"""COM-measure Word's actual x_start for 31420a (一太郎 PASS doc) paragraphs
that have firstLineChars+firstLine. Goal: see if Word ignores firstLine
in PASS docs the same way it does in FAIL docs (bd90b00, 191cb5).

If YES → 一太郎 doc indent rule is consistent (Oxi already handles
PASS docs correctly via some other compensation, would break with naive
"ignore indent" implementation = compensation triangle).

If NO → 一太郎 PASS docs DON'T have indent ignored, so the trigger is
something more specific within FAIL docs.
"""
import json
import os
import sys
import win32com.client

WD_HPOS = 5
WD_VPOS = 6
WD_PAGE = 3

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx", "31420af1a08f_tokumei_08_07.docx")
OUT = os.path.join(REPO, "pipeline_data", "31420a_indent_word.json")

TARGETS = [11, 21, 22, 30, 32, 34, 52, 75]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    rows = []
    try:
        # Find docx
        docx_dir = os.path.dirname(DOCX)
        for f in os.listdir(docx_dir):
            if f.startswith("31420a"):
                docx_path = os.path.join(docx_dir, f)
                break
        else:
            print("31420a docx not found")
            return
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        try:
            sec = doc.Sections(1)
            pm_left = sec.PageSetup.LeftMargin
            print(f"Page margin left: {pm_left}pt")
            for pi in TARGETS:
                try:
                    para = doc.Paragraphs(pi)
                    rng = para.Range
                    first = doc.Range(rng.Start, rng.Start)
                    actual_x = first.Information(WD_HPOS)
                    actual_y = first.Information(WD_VPOS)
                    actual_pg = first.Information(WD_PAGE)
                    fmt = para.Format
                    # Word reports indent in points
                    word_left_pt = fmt.LeftIndent
                    word_first_line_pt = fmt.FirstLineIndent
                    expected_full = pm_left + word_left_pt + word_first_line_pt
                    expected_no_first = pm_left + word_left_pt
                    diff_full = actual_x - expected_full
                    diff_no_first = actual_x - expected_no_first
                    if abs(diff_full) < 1.0:
                        cls = "FULL_INDENT"
                    elif abs(diff_no_first) < 1.0:
                        cls = "FIRSTLINE_IGNORED"
                    else:
                        cls = f"OTHER({diff_full:+.2f}/{diff_no_first:+.2f})"
                    row = {
                        "pi": pi,
                        "actual_x": round(actual_x, 3),
                        "word_left": round(word_left_pt, 3),
                        "word_first_line": round(word_first_line_pt, 3),
                        "expected_full": round(expected_full, 3),
                        "expected_no_first": round(expected_no_first, 3),
                        "diff_full": round(diff_full, 3),
                        "diff_no_first": round(diff_no_first, 3),
                        "classification": cls,
                    }
                    rows.append(row)
                    print(
                        f"pi={pi:3d}  L={word_left_pt:>6.2f}  FL={word_first_line_pt:>6.2f}  "
                        f"x_act={actual_x:>6.2f}  x_full={expected_full:>6.2f}  x_noFL={expected_no_first:>6.2f}  "
                        f"=> {cls}"
                    )
                except Exception as e:
                    print(f"pi={pi}: ERR {e}")
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"rows": rows}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
