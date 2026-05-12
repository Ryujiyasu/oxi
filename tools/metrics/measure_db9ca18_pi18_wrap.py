"""Measure Word's actual line count + Y positions for db9ca18 pi=18 paragraph
"b. If the user has edited..."

Hypothesis (Day 33 part 64): Oxi over-wraps pi=18 by 1 line (renders 7 lines
vs Word 6 lines). If confirmed, the 18pt savings would tip db9ca18 PASS.

This script uses Word COM to:
1. Open db9ca18 read-only
2. Find paragraph starting with "b. If the user has edited"
3. Measure its line count via Word.LineCount or per-character Information(7)
4. Report start/end Y positions

Run: python tools/metrics/measure_db9ca18_pi18_wrap.py
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
    if not os.path.exists(DOCX_PATH):
        print(f"[NG] file not found: {DOCX_PATH}")
        return 1

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOCX_PATH, ReadOnly=True)
        time.sleep(0.3)

        n = doc.Paragraphs.Count
        print(f"Total paragraphs: {n}")

        target_idx = None
        for i in range(1, n + 1):
            p = doc.Paragraphs(i)
            text = p.Range.Text.replace("\r", "").replace("\x07", "")
            if text.startswith("b. If the user has edited"):
                target_idx = i
                print(f"Found target at paragraph {i}")
                print(f"  Text length: {len(text)}")
                print(f"  Text first 80: {text[:80]}")
                break

        if target_idx is None:
            print("[NG] target paragraph not found")
            return 1

        p = doc.Paragraphs(target_idx)
        rng = p.Range
        start_rng = doc.Range(rng.Start, rng.Start)
        end_rng = doc.Range(rng.End - 1, rng.End - 1)

        # Information(3) = wdActiveEndPageNumber
        # Information(5) = wdHorizontalPositionRelativeToPage (X in points)
        # Information(6) = wdVerticalPositionRelativeToPage (Y in points)
        start_page = start_rng.Information(3)
        start_x = start_rng.Information(5)
        start_y = start_rng.Information(6)
        end_page = end_rng.Information(3)
        end_x = end_rng.Information(5)
        end_y = end_rng.Information(6)

        print()
        print(f"Paragraph start: page={start_page} X={start_x:.2f} Y={start_y:.2f}")
        print(f"Paragraph end:   page={end_page} X={end_x:.2f} Y={end_y:.2f}")

        # Walk through characters to count distinct Y positions (lines).
        # For multi-page paragraphs, lines may span pages.
        n_chars = rng.End - rng.Start
        print(f"\nChar count: {n_chars}")

        prev_y_per_page: dict[int, list[float]] = {}
        lines_per_page: dict[int, int] = {}
        char_y_samples: list[tuple[int, float]] = []
        # Sample every 5 chars to keep COM round-trips reasonable.
        sample_step = max(1, n_chars // 60)
        for offset in range(0, n_chars, sample_step):
            char_rng = doc.Range(rng.Start + offset, rng.Start + offset)
            pg = char_rng.Information(3)
            y = char_rng.Information(6)
            char_y_samples.append((pg, y))

        # Deduplicate consecutive same (pg, round-to-0.5pt y)
        unique_lines: list[tuple[int, float]] = []
        for pg, y in char_y_samples:
            y_rounded = round(y * 2) / 2
            if not unique_lines or unique_lines[-1] != (pg, y_rounded):
                unique_lines.append((pg, y_rounded))
        # Re-deduplicate (in case of zig-zag)
        unique_lines = sorted(set(unique_lines))

        print(f"\nUnique (page, line_y) positions found (sample step={sample_step}):")
        for pg, y in unique_lines:
            print(f"  page={pg} y={y:.1f}")

        by_page: dict[int, list[float]] = {}
        for pg, y in unique_lines:
            by_page.setdefault(pg, []).append(y)
        print(f"\nLines per page:")
        total_lines = 0
        for pg in sorted(by_page.keys()):
            ys = sorted(by_page[pg])
            print(f"  page {pg}: {len(ys)} lines at y={[round(y,1) for y in ys]}")
            total_lines += len(ys)
        print(f"Total lines: {total_lines}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
