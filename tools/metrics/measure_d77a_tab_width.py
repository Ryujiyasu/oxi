"""Measure Word's tab character width in d77a tbl8 カ paragraph.

Target: カ paragraph starts with "カ\t本利用ルール...". Measure x position
of 'カ', '\t' (no x), '本' (first char after tab), then compute:
  tab_width = x('本') - x('カ') - width('カ')

Compare against Oxi which has tab ≈ 5pt (too narrow).
"""
import json
from pathlib import Path
import win32com.client as w32


DOC = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\d77a_tab_width_measurement.json")


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": DOC.name}
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            # tbl8 is Table 8, cell (1,1), paragraph 7 (カ)
            tbl = doc.Tables(8)
            cell = tbl.Cell(1, 1)
            cr = cell.Range
            para_count = cr.Paragraphs.Count
            print(f"tbl8 cell has {para_count} paragraphs")

            # Find カ paragraph by content
            target_para = None
            for pi in range(1, para_count + 1):
                p = cr.Paragraphs(pi)
                txt = p.Range.Text[:10]
                if txt.startswith("カ") or "カ\t" in txt:
                    target_para = p
                    print(f"Found カ paragraph at pi={pi}: text={p.Range.Text[:40]!r}")
                    break

            if target_para is None:
                print("Could not find カ paragraph")
                return

            pr = target_para.Range
            # Measure first 15 characters' x positions
            # Information(1) = wdHorizontalPositionRelativeToPage
            measurements = []
            for offset in range(pr.Start, min(pr.Start + 20, pr.End)):
                r = doc.Range(offset, offset + 1)
                try:
                    x = r.Information(1)
                    y = r.Information(6)
                    ch = r.Text[:1]
                    if ch == '\t': ch_display = '<TAB>'
                    elif ch == '\r': ch_display = '<CR>'
                    else: ch_display = ch
                    measurements.append({"offset": offset, "char": ch_display, "x_pt": round(x, 2), "y_pt": round(y, 2)})
                    print(f"  off={offset} ch={ch_display!r} x={x:.2f} y={y:.2f}")
                except Exception as e:
                    print(f"  off={offset} ERR: {e}")

            result["measurements"] = measurements

            # Compute tab width
            # Find カ (first non-empty) and 本 (char after tab)
            ka_m = None
            hon_m = None
            tab_m = None
            for i, m in enumerate(measurements):
                if m['char'] == 'カ' and ka_m is None:
                    ka_m = m
                elif m['char'] == '<TAB>' and tab_m is None:
                    tab_m = m
                elif m['char'] == '本' and hon_m is None:
                    hon_m = m

            if ka_m and hon_m:
                tab_width = hon_m['x_pt'] - ka_m['x_pt']
                print(f"\nWord tab consumption: 本.x={hon_m['x_pt']} - カ.x={ka_m['x_pt']} = {tab_width:.2f}pt (includes カ char width)")
                # Also estimate カ alone
                result["ka_x"] = ka_m['x_pt']
                result["hon_x"] = hon_m['x_pt']
                result["tab_advance"] = round(tab_width, 2)

        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
