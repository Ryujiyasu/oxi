"""Count distinct Y positions (≈ lines) inside each d77a table cell.

Uses Range.Characters iteration — get y of each character, count unique ys per cell.
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
        # Step through paragraphs; within each, capture start and end char Y
        para_lines = []
        total_lines = 0
        for p in cell.Range.Paragraphs:
            pr = p.Range
            # Step every ~5 chars for speed
            n_chars = pr.Characters.Count
            ys = set()
            if n_chars == 0:
                ys.add(round(pr.Information(6), 1))
            else:
                step = max(1, n_chars // 30)  # up to ~30 samples
                for c_idx in range(1, n_chars + 1, step):
                    try:
                        ch = pr.Characters(c_idx)
                        y = ch.Information(6)
                        ys.add(round(y, 1))
                    except Exception:
                        pass
                # Also last char
                try:
                    y = pr.Characters(n_chars).Information(6)
                    ys.add(round(y, 1))
                except Exception:
                    pass
            para_lines.append(len(ys))
            total_lines += len(ys)
        print(f"#{i}: paras={len(para_lines)} lines/para={para_lines} total={total_lines}")
        data.append({"idx": i, "lines_per_para": para_lines, "total_lines": total_lines})

    out = "pipeline_data/d77a_word_table_line_counts.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out}")

    doc.Close(False)
finally:
    word.Quit()
