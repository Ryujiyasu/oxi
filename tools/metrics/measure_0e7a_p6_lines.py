"""Measure Y positions of EACH wrapped line on page 6 of 0e7a (not just paragraph starts).

Uses Word's Information(6) on character-level Range subsets to find where Y jumps.
Writes JSON listing (line_idx, y_pt, first_char_offset, first_char).
"""
import win32com.client
import json
import os

docx_path = os.path.abspath(
    r"tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)

    total_chars = doc.Range().End
    print(f"Total chars: {total_chars}")

    # Find char offset range that corresponds to page 6
    # Walk paragraphs, find first paragraph on page 6 and first on page 7
    start_off = None
    end_off = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        pg = p.Range.Information(3)
        if pg == 6 and start_off is None:
            start_off = p.Range.Start
        elif pg == 7 and start_off is not None:
            end_off = p.Range.Start
            break
    if end_off is None:
        end_off = total_chars
    print(f"Page 6 range: {start_off} to {end_off}")

    # Walk through char range; query Y for each single-char range
    lines_seen = []
    prev_y = None
    for off in range(start_off, end_off):
        r = doc.Range(off, off + 1)
        pg = r.Information(3)
        if pg != 6:
            continue
        y = r.Information(6)
        if prev_y is None or abs(y - prev_y) > 0.3:
            ch = r.Text[:1].replace('\r', '\\r').replace('\n', '\\n').replace('\t', '\\t')
            lines_seen.append({"offset": off, "y_pt": round(y, 2), "char": ch})
            prev_y = y

    print(f"Detected {len(lines_seen)} lines on page 6")
    for i, L in enumerate(lines_seen):
        if i > 0:
            gap = L["y_pt"] - lines_seen[i-1]["y_pt"]
            print(f"  line {i:3d}  off={L['offset']:6d}  y={L['y_pt']:7.2f}  gap={gap:+6.2f}  [{L['char']}]")
        else:
            print(f"  line {i:3d}  off={L['offset']:6d}  y={L['y_pt']:7.2f}                   [{L['char']}]")

    outfn = "pipeline_data/0e7a_p6_word_lines.json"
    with open(outfn, "w", encoding="utf-8") as f:
        json.dump(lines_seen, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {outfn}")

    doc.Close(False)
finally:
    word.Quit()
