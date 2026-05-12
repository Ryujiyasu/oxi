"""Measure Word's actual line count + Y positions for db9ca18 paragraph 37
(pi=36 in Oxi = "This Terms of Use herein does not apply...").

Pre-fix Oxi: pi=36 breaks BEFORE rendering line 0 (at cy=755 + 18 > 771).
Post-fix Oxi: pi=36 line 0 fits page 2 (cy=755 + 12.7 < 771).

This script verifies how many lines paragraph 37 has in Word and where they
are positioned. If Word fits ALL lines on page 2, the natural_lh fix is
correct. If Word also wraps some lines to page 3, the fix might be wrong.

Run: python tools/metrics/measure_db9ca18_pi36_wrap.py
"""

import sys
import os
import time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_PATH = os.path.join(
    REPO_ROOT,
    "tools",
    "golden-test",
    "documents",
    "docx",
    "db9ca18368cd_20241122_resource_open_data_01.docx",
)


def main() -> int:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOCX_PATH, ReadOnly=True)
        time.sleep(0.3)
        p = doc.Paragraphs(37)
        rng = p.Range
        text = rng.Text.replace("\r", "").replace("\x07", "")
        print(f"Paragraph 37 text length: {len(text)}")
        print(f"First 80: {text[:80]}")

        start_rng = doc.Range(rng.Start, rng.Start)
        end_rng = doc.Range(rng.End - 1, rng.End - 1)
        print(f"\nStart: page={start_rng.Information(3)} Y={start_rng.Information(6):.2f}")
        print(f"End:   page={end_rng.Information(3)} Y={end_rng.Information(6):.2f}")

        n_chars = rng.End - rng.Start
        sample_step = max(1, n_chars // 60)
        unique = set()
        for offset in range(0, n_chars, sample_step):
            r = doc.Range(rng.Start + offset, rng.Start + offset)
            pg = r.Information(3)
            y = round(r.Information(6) * 2) / 2
            unique.add((pg, y))

        print(f"\nUnique (page, line_y) positions:")
        for pg, y in sorted(unique):
            print(f"  page={pg} y={y:.1f}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
