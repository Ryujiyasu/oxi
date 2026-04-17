"""Measure all paragraphs on pages 6-8 of 0e7a to find where Oxi diverges.

Output: for each paragraph that touches page 6,7,8 — the paragraph index, page,
Y position (pt from page top), line spacing, space before/after, text snippet,
keepNext/keepTogether flags, whether it's inside a table.
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

    results = []
    total = doc.Paragraphs.Count
    print(f"Total paragraphs: {total}")

    prev_page = None
    for i in range(1, total + 1):
        para = doc.Paragraphs(i)
        rng = para.Range
        page = rng.Information(3)  # wdActiveEndPageNumber

        if page < 6:
            continue
        if page > 8:
            break

        y = rng.Information(6)  # wdVerticalPositionRelativeToPage
        fmt = para.Format
        ls_rule = fmt.LineSpacingRule  # 0=single, 1=1.5, 2=double, 3=atLeast, 4=exactly, 5=multiple
        ls = fmt.LineSpacing
        sb = fmt.SpaceBefore
        sa = fmt.SpaceAfter
        keep_next = fmt.KeepWithNext
        keep_together = fmt.KeepTogether
        widow = fmt.WidowControl
        outline_level = fmt.OutlineLevel
        in_table = rng.Information(12)  # wdWithInTable (bool)

        text = rng.Text[:60].replace('\r', '\\r').replace('\n', '\\n')

        marker = ""
        if prev_page is not None and page != prev_page:
            marker = f" <<< PAGE BREAK {prev_page}->{page}"

        print(f"P{i:3d} pg={page} y={y:7.1f} lsR={ls_rule} ls={ls:.1f} sb={sb:4.1f} sa={sa:4.1f} "
              f"kn={int(keep_next)} kt={int(keep_together)} w={int(widow)} ol={outline_level} "
              f"tbl={int(in_table)} [{text[:45]}]{marker}")

        results.append({
            "para": i,
            "page": page,
            "y_pt": round(y, 2),
            "ls_rule": ls_rule,
            "line_spacing": round(ls, 2),
            "space_before": round(sb, 2),
            "space_after": round(sa, 2),
            "keep_next": bool(keep_next),
            "keep_together": bool(keep_together),
            "widow": bool(widow),
            "outline_level": outline_level,
            "in_table": bool(in_table),
            "text": text,
        })

        prev_page = page

    doc.Close(False)

    outfn = "pipeline_data/0e7a_p678_y_positions.json"
    with open(outfn, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {len(results)} paragraphs to {outfn}")

finally:
    word.Quit()
