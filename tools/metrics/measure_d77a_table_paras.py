"""Measure Y position of each paragraph inside d77a tables.

Each table is 1-row 1-cell with multi-paragraph content.
Get: (table top, each paragraph y, after-table y) -> compute internal height.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

docx_path = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True); time.sleep(0.5)
    n = doc.Tables.Count
    print(f"Tables: {n}\n")

    data = []
    for i in range(1, n + 1):
        t = doc.Tables(i)
        cell = t.Cell(1, 1)
        tbl_page = t.Range.Information(3)

        paras_data = []
        for p in cell.Range.Paragraphs:
            try:
                y = p.Range.Information(6)
                pg = p.Range.Information(3)
                txt = p.Range.Text[:25].replace('\r', '').replace('\n', ' ').replace('\x07', '')
                paras_data.append({"y": round(y, 2), "page": int(pg), "text": txt})
            except Exception as e:
                paras_data.append({"error": str(e)})

        # After-table Y (paragraph following the table)
        try:
            after = doc.Range(t.Range.End, t.Range.End + 1)
            after_y = after.Information(6)
            after_pg = after.Information(3)
        except Exception:
            after_y = None; after_pg = None

        # Inline height estimate: paras[-1].y - paras[0].y + last paragraph height
        if len(paras_data) >= 2:
            y0 = paras_data[0].get('y', 0)
            yn = paras_data[-1].get('y', 0)
            inline_span = yn - y0
        else:
            inline_span = 0

        entry = {
            "idx": i,
            "page": int(tbl_page),
            "n_paras": len(paras_data),
            "paras": paras_data,
            "after_y": round(after_y, 2) if after_y else None,
            "after_page": int(after_pg) if after_pg else None,
            "inline_span": round(inline_span, 2),
        }
        data.append(entry)

        print(f"#{i} p{int(tbl_page)}: {len(paras_data)} paras; span={inline_span:.1f}pt; after_y={after_y}")
        for pi, p in enumerate(paras_data[:4]):
            print(f"  p{pi}: y={p.get('y')} page={p.get('page')} text={p.get('text')!r}")
        if len(paras_data) > 4:
            print(f"  ... ({len(paras_data)-4} more)")
        # Gap pattern: y diffs
        if len(paras_data) > 1:
            gaps = []
            for j in range(1, len(paras_data)):
                y_prev = paras_data[j-1].get('y', 0)
                y_cur = paras_data[j].get('y', 0)
                pg_prev = paras_data[j-1].get('page', 0)
                pg_cur = paras_data[j].get('page', 0)
                if pg_prev == pg_cur:
                    gaps.append(round(y_cur - y_prev, 2))
                else:
                    gaps.append(f"[pg {pg_prev}->{pg_cur}]")
            print(f"  gaps: {gaps}")

    out = "pipeline_data/d77a_word_table_paras.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out}")

    doc.Close(False)
finally:
    word.Quit()
