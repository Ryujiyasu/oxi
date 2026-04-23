"""Measure which paragraph ends p6 in Word for d77a, and properties of nearby paras."""
import win32com.client
import os
import json
from pathlib import Path

DOCX = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True)

        out = []
        total = doc.Paragraphs.Count
        print(f"Total paragraphs: {total}")

        for i in range(1, total + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            # wdHorizontalPositionRelativeToPage(5), wdVerticalPositionRelativeToPage(6)
            # wdActiveEndAdjustedPageNumber(1) — page where range starts
            try:
                page_start = rng.Information(3)  # wdActiveEndPageNumber at start
            except Exception:
                page_start = -1
            try:
                y_start = rng.Information(6)
            except Exception:
                y_start = -1.0

            try:
                end_pos = rng.End
                # Clamp end_pos to safe range (subtract 1 to avoid overshoot)
                safe_end = max(rng.Start, end_pos - 1)
                end_rng = doc.Range(safe_end, safe_end)
                page_end = end_rng.Information(3)
                y_end = end_rng.Information(6)
            except Exception:
                page_end = -1
                y_end = -1.0

            text_preview = rng.Text[:40].replace("\r", " ").replace("\n", " ")
            out.append({
                "idx": i,
                "page_start": page_start,
                "page_end": page_end,
                "y_start": round(y_start, 2),
                "y_end": round(y_end, 2),
                "keep_lines": p.KeepWithNext,  # actually keepNext
                "widow_control": p.WidowControl,
                "text": text_preview,
            })

        # Find paragraphs around p6/p7 boundary
        print("\n=== p5-p8 paragraphs ===")
        for r in out:
            if 5 <= r["page_start"] <= 8 or 5 <= r["page_end"] <= 8:
                print(f"  idx={r['idx']:3d} pg {r['page_start']}-{r['page_end']}  y_start={r['y_start']}  y_end={r['y_end']}  widow={r['widow_control']} keepNext={r['keep_lines']}  {r['text']!r}")

        out_path = r"C:\Users\ryuji\oxi-main\pipeline_data\d77a_p6_p7_boundary.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"\nSaved: {out_path}")

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()

if __name__ == "__main__":
    main()
