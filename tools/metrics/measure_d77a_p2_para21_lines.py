"""Measure Word's per-line Y positions WITHIN PARA 21 of d77a p.2.

Hypothesis: Oxi wraps this 12pt MS Gothic paragraph in 7 lines; Word in 8.
Walk char-by-char within the paragraph range, emit a new event when Y jumps.
"""
import win32com.client
import json
import os

docx_path = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    # Body paragraph index: XML body has 200 top-level paras; PARA 21 (0-indexed) = Paragraphs(22)
    # But Word COM's Paragraphs collection may include paragraphs inside tables.
    # Strategy: find paragraph whose first chars are "平成26年"
    target = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        t = p.Range.Text[:10]
        if t.startswith("平成26年"):
            target = p
            target_idx = i
            print(f"Target para idx={i} y_start={p.Range.Information(6):.2f}")
            print(f"  text[:80]={p.Range.Text[:80]!r}")
            print(f"  length={p.Range.End - p.Range.Start}")
            break
    if target is None:
        print("平成26年 paragraph not found")
        raise SystemExit(1)

    # Walk chars within range, capture Y jumps
    start = target.Range.Start
    end = target.Range.End - 1  # exclude trailing \r
    lines = []
    prev_y = None
    for off in range(start, end):
        r = doc.Range(off, off + 1)
        y = r.Information(6)
        if prev_y is None or abs(y - prev_y) > 0.3:
            ch = r.Text[:1].replace("\r", "\\r").replace("\n", "\\n")
            lines.append({"offset": off, "y_pt": round(y, 2), "char": ch})
            prev_y = y

    print(f"\nDetected {len(lines)} lines in PARA 21:")
    for i, L in enumerate(lines):
        if i > 0:
            gap = L["y_pt"] - lines[i - 1]["y_pt"]
            print(f"  line {i + 1:2d}  off={L['offset']:6d}  y={L['y_pt']:7.2f}  gap={gap:+6.2f}  [{L['char']}]")
        else:
            print(f"  line {i + 1:2d}  off={L['offset']:6d}  y={L['y_pt']:7.2f}                 [{L['char']}]")

    # Also report y of next paragraph (confirms p22 start)
    nxt = doc.Paragraphs(target_idx + 1)
    print(f"\nNext para (idx {target_idx + 1}) starts at y={nxt.Range.Information(6):.2f}")

    out = {
        "para_text_len": end - start,
        "first_y": lines[0]["y_pt"] if lines else None,
        "last_y": lines[-1]["y_pt"] if lines else None,
        "num_lines": len(lines),
        "next_para_y": round(nxt.Range.Information(6), 2),
        "lines": lines,
    }
    with open("pipeline_data/d77a_p2_para21_word_lines.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to pipeline_data/d77a_p2_para21_word_lines.json")

    doc.Close(False)
finally:
    word.Quit()
