"""Measure Word's wrap positions for d77a Table 1 paragraph 4.

Word says this para is 3 lines; Oxi renders 2. Find wrap chars.
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

    t = doc.Tables(1)
    cell = t.Cell(1, 1)
    paras = list(cell.Range.Paragraphs)
    print(f"Cell has {len(paras)} paragraphs\n")

    for pi in [1, 3]:  # para 2 and 4 (0-idx 1 and 3)
        target = paras[pi]
        pr = target.Range
        n_chars = pr.Characters.Count
        print(f"--- para {pi+1}: n_chars={n_chars} ---")
        # Paragraph Format info
        pf = target.Format
        left_indent = pf.LeftIndent
        first_indent = pf.FirstLineIndent
        right_indent = pf.RightIndent
        print(f"  leftIndent={left_indent:.2f} firstLine={first_indent:.2f} rightIndent={right_indent:.2f}")

        chars = []
        for ci in range(1, n_chars + 1):
            try:
                c = pr.Characters(ci)
                x = c.Information(5)
                y = c.Information(6)
                ch = c.Text
                chars.append({"i": ci, "ch": ch, "x": round(x, 2), "y": round(y, 2)})
            except Exception as e:
                chars.append({"i": ci, "error": str(e)})

        # Group by y
        lines = []
        cur = []
        cur_y = None
        for c in chars:
            if 'error' in c: continue
            if cur_y is None or abs(c['y'] - cur_y) < 2:
                cur.append(c); cur_y = c['y']
            else:
                lines.append(cur); cur = [c]; cur_y = c['y']
        if cur: lines.append(cur)

        print(f"  Lines: {len(lines)}")
        for li, line in enumerate(lines):
            txt = ''.join(c['ch'] for c in line).replace('\r','').replace('\n','')
            if line:
                print(f"    L{li+1}: n={len(line)} y={line[0]['y']:.1f} x={line[0]['x']:.1f}..{line[-1]['x']:.1f} {txt[:60]!r}")

    doc.Close(False)
finally:
    word.Quit()
