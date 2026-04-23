"""Measure per-line Y for every paragraph on d77a Word p6 via COM.

For each paragraph whose start or end is on page 6, walk characters in
increments and record each distinct Y position = line start.

Output: JSON list of {idx, text_preview, page_start, line_ys, char_counts}
"""
import win32com.client
import json
from pathlib import Path

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
OUT = r"C:\Users\ryuji\oxi-main\pipeline_data\d77a_p6_line_ys.json"
TARGET_PAGES = (5, 6, 7)  # context around p6
STEP = 1  # char step for Y sampling


def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        total = doc.Paragraphs.Count

        results = []
        for i in range(1, total + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            try:
                ps = rng.Information(3)
                pe = doc.Range(max(rng.Start, rng.End - 1), max(rng.Start, rng.End - 1)).Information(3)
            except Exception:
                continue
            if ps not in TARGET_PAGES and pe not in TARGET_PAGES:
                continue

            # Walk the paragraph range char-by-char, sample Y at each char.
            start = rng.Start
            end = rng.End
            line_ys = []  # unique Ys in order
            line_starts = []  # char offset where each line begins
            last_y = None
            for pos in range(start, end, STEP):
                sub = doc.Range(pos, pos)
                try:
                    y = sub.Information(6)
                    pg = sub.Information(3)
                except Exception:
                    continue
                if y is None or not isinstance(y, float):
                    continue
                if last_y is None or abs(y - last_y) > 2.0:
                    line_ys.append(round(y, 2))
                    line_starts.append(pos - start)
                    last_y = y

            text_preview = rng.Text[:50].replace("\r", " ").replace("\n", " ").replace("\x07", "|")
            results.append({
                "idx": i,
                "page_start": ps,
                "page_end": pe,
                "line_count": len(line_ys),
                "line_ys": line_ys,
                "line_starts": line_starts,
                "length": end - start,
                "text": text_preview,
            })
            print(f"  idx={i:3d} pg {ps}-{pe} lines={len(line_ys)} ys={line_ys[:5]}{'...' if len(line_ys)>5 else ''} {text_preview!r}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    Path(OUT).parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {len(results)} paragraphs to {OUT}")


if __name__ == "__main__":
    main()
