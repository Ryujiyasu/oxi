"""
Measure Word's actual wrap of b837 paras[48] ("更に、オープンデータの...").

Opens b837 directly in Word COM and scans the paragraph's text positions
character by character to determine how many lines it wraps into.
Compares with Oxi's 5-line wrap claim.
"""
import json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"C:/Users/ryuji/oxi-4/tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
OUT = Path(__file__).with_name("output") / "b837_para48_wrap.json"
OUT.parent.mkdir(parents=True, exist_ok=True)


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
        time.sleep(1.5)

        # Document paragraphs — XML body index 48 = COM Paragraphs(49) (1-based)
        # 1-based: XML paras[N] = Word Paragraphs(N+1)
        target_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 49

        p = doc.Paragraphs(target_idx)
        rng = p.Range

        print(f"Paragraph {target_idx}: text chars = {rng.End - rng.Start}")
        print(f"  Start={rng.Start}, End={rng.End}")
        print(f"  First char text = {rng.Text[:30]!r}")

        # Walk each character, get y + page
        sel = word.Selection
        line_ys = {}  # (page, y_key) -> char count
        line_order = []
        line_start_x = {}
        for ci in range(rng.Start, rng.End):
            sel.SetRange(ci, ci + 1)
            try:
                y = sel.Information(6)   # Vertical position, pt
                x = sel.Information(5)   # Horizontal position, pt
                pg = int(sel.Information(3))  # wdActiveEndPageNumber
            except Exception:
                continue
            y_key = round(y * 2) / 2
            key = (pg, y_key)
            if key not in line_ys:
                line_ys[key] = 0
                line_order.append(key)
                line_start_x[key] = x
            line_ys[key] += 1

        doc.Close(False)

        line_order.sort()
        print(f"\n{len(line_order)} lines total:")
        for k in line_order:
            pg, y = k
            print(f"  page={pg} y={y:.1f} x_start={line_start_x[k]:.2f} chars={line_ys[k]}")

        result = {
            "doc": "b837",
            "paragraph_index_1based": target_idx,
            "n_lines": len(line_order),
            "lines": [{"page": k[0], "y": k[1], "x_start": line_start_x[k], "chars": line_ys[k]} for k in line_order],
        }
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        print(f"\nSaved -> {OUT}")
    finally:
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
