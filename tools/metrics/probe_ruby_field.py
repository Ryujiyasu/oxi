"""Probe Word's Field representation of ruby — can we extract per-glyph X
of the rendered ruby annotation?

The previous probe revealed that Range.Duplicate + SetRange across the
ruby placeholder shows characters 'E', 'Q', ' ', '\\', '*' — Word stores
ruby as an EQ field internally. EQ field result text contains base text +
ruby annotation displayed via \\o (overstrike) and \\ad (advance) operators.

This probe enumerates the field, its code, and whether Word exposes
per-glyph result positions.
"""
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_PATH = os.path.abspath("pipeline_data/docx/RUBY_V2_align_variants.docx")


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(DOCX_PATH, ReadOnly=True)
    time.sleep(0.4)

    print(f"Doc.Fields.Count = {doc.Fields.Count}")
    for fi in range(1, doc.Fields.Count + 1):
        f = doc.Fields(fi)
        print(f"\n=== Field {fi} ===")
        try:
            print(f"  Type = {f.Type}")
            # Field code
            try:
                print(f"  Code.Text = {f.Code.Text!r}")
                print(f"  Code.Start = {f.Code.Start} End = {f.Code.End}")
            except Exception as e:
                print(f"  Code: ERROR {e}")
            # Field result
            try:
                print(f"  Result.Text = {f.Result.Text!r}")
                print(f"  Result.Start = {f.Result.Start} End = {f.Result.End}")
            except Exception as e:
                print(f"  Result: ERROR {e}")
            # Field range
            try:
                print(f"  Range.Text = {f.Range.Text!r}")
                rngx = f.Range.Information(5)
                rngy = f.Range.Information(6)
                print(f"  Range.Information(5,6) = ({rngx}, {rngy})")
            except Exception as e:
                print(f"  Range: ERROR {e}")
            # If field has a Result range, try to iterate its characters
            try:
                rr = f.Result
                if rr is not None:
                    print(f"  Result.Characters.Count = {rr.Characters.Count}")
                    for ci in range(1, rr.Characters.Count + 1):
                        c = rr.Characters(ci)
                        ct = c.Text
                        cx = c.Information(5)
                        cy = c.Information(6)
                        cf = c.Font.Name
                        cs = c.Font.Size
                        print(f"    R[{ci}] text={ct!r} x={cx} y={cy} font={cf!r} sz={cs}")
            except Exception as e:
                print(f"  Result.Characters: ERROR {e}")
        except Exception as e:
            print(f"  field ERROR: {e}")

    doc.Close(SaveChanges=False)
    word.Quit()


if __name__ == "__main__":
    main()
