"""Measure Word's wrap positions for d77a Table 1 paragraph 2 (5 lines).

Step through every char; capture (x, y). Lines identified by Y change.
Output each line's char count, start/end x, and text.
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

    # Target: Table 1, paragraph 2 (index 2 in 1-based Word; 5 lines)
    t = doc.Tables(1)
    cell = t.Cell(1, 1)
    paras = list(cell.Range.Paragraphs)
    print(f"Cell has {len(paras)} paragraphs")
    for pi, p in enumerate(paras, 1):
        txt = p.Range.Text[:40].replace('\r', '').replace('\n', ' ').replace('\x07', '')
        nc = p.Range.Characters.Count
        print(f"  para {pi}: chars={nc} text={txt!r}")

    print("\n--- Measure para 2 (the long CJK body) ---")
    target = paras[1]  # 0-indexed; second paragraph
    pr = target.Range
    n_chars = pr.Characters.Count
    print(f"Total chars: {n_chars}")

    # Walk every char; collect (idx, char, x, y)
    chars = []
    for ci in range(1, n_chars + 1):
        try:
            c = pr.Characters(ci)
            x = c.Information(5)  # wdHorizontalPositionRelativeToPage
            y = c.Information(6)  # wdVerticalPositionRelativeToPage
            ch = c.Text
            chars.append({"i": ci, "ch": ch, "x": round(x, 2), "y": round(y, 2)})
        except Exception as e:
            chars.append({"i": ci, "error": str(e)})

    # Group by Y (lines)
    lines = []
    cur = []
    cur_y = None
    for c in chars:
        if 'error' in c:
            continue
        if cur_y is None or abs(c['y'] - cur_y) < 2:
            cur.append(c); cur_y = c['y']
        else:
            lines.append(cur); cur = [c]; cur_y = c['y']
    if cur:
        lines.append(cur)

    print(f"\nLines detected: {len(lines)}")
    for li, line in enumerate(lines):
        chars_str = ''.join(c['ch'] for c in line).replace('\r', '').replace('\n', '')
        if line:
            print(f"  L{li+1}: n={len(line)} y={line[0]['y']:.1f} x={line[0]['x']:.1f}..{line[-1]['x']:.1f} | {chars_str[:50]!r}")

    out = "pipeline_data/d77a_t1p2_word_chars.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump({
            "n_lines": len(lines),
            "lines": [{"y": l[0]['y'] if l else None, "x_start": l[0]['x'] if l else None,
                       "x_end": l[-1]['x'] if l else None, "n_chars": len(l),
                       "text": ''.join(c['ch'] for c in l).replace('\r','').replace('\n','')}
                      for l in lines],
            "chars": chars,
        }, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out}")

    doc.Close(False)
finally:
    word.Quit()
