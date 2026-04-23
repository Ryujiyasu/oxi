"""Measure Word's exact wrapping of d77a tbl6 (p.7, 1-col 1-row table).

Cell has 3-4 paragraphs; the long "イ" paragraph wraps to N lines in Word.
Oxi wraps the same paragraph to 3 lines (18pt extra table height).

Per-character sampling: for each character offset in the cell, record (page,
y, char). New line detected when y changes.
"""
import json
from pathlib import Path
import win32com.client as w32


DOC = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\d77a_tbl6_wrap_measurement.json")


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": DOC.name}
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            tbl = doc.Tables(6)
            print(f"Table 6: rows={tbl.Rows.Count} cols={tbl.Columns.Count}")
            cell = tbl.Cell(1, 1)
            cell_range = cell.Range
            print(f"Cell range: {cell_range.Start}..{cell_range.End} chars={cell_range.End - cell_range.Start}")

            # Per-paragraph line dump
            cell_paras = []
            for pi in range(1, cell_range.Paragraphs.Count + 1):
                para = cell_range.Paragraphs(pi)
                pr = para.Range
                lines = []
                prev_y = None
                prev_page = None
                for off in range(pr.Start, pr.End):
                    r = doc.Range(off, off + 1)
                    try:
                        pg = r.Information(3)
                        y = r.Information(6)
                        x = r.Information(1)
                    except Exception:
                        continue
                    if prev_y is None or abs(y - prev_y) > 0.3 or pg != prev_page:
                        ch = r.Text[:1].replace('\r', '¶').replace('\n', '↵').replace('\t', '→').replace('\x07', '⌂')
                        lines.append({"offset": off, "page": pg, "x_pt": round(x, 2), "y_pt": round(y, 2), "char": ch})
                        prev_y = y
                        prev_page = pg
                preview = pr.Text[:60].replace("\r", "¶").replace("\x07", "⌂")
                cell_paras.append({
                    "p_idx": pi,
                    "chars": pr.End - pr.Start,
                    "preview": preview,
                    "line_count": len(lines),
                    "lines": lines,
                })

            result["cell_paragraphs"] = cell_paras
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print("\n=== Word tbl6 cell paragraphs ===")
    for p in result["cell_paragraphs"]:
        print(f"\npara{p['p_idx']}: {p['chars']} chars, {p['line_count']} lines")
        for ln in p["lines"]:
            print(f"  off={ln['offset']} p{ln['page']} x={ln['x_pt']:6.2f} y={ln['y_pt']:6.2f} ch={ln['char']!r}")


if __name__ == "__main__":
    main()
