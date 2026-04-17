"""Measure Word's table row heights for b35123fe8efc_tokumei_08_01.

b35 has docGrid=linesAndChars (linePitch=350tw=17.5pt) and the
adjustLineHeightInTable flag set. Oxi's current row grid-snap is gated on
`grid_char_pitch.is_none()`, so b35 skips the snap → rows at ~14pt content
height. Hypothesis: Word still snaps b35 rows to 17.5pt despite linesAndChars.
If confirmed on b35 + 1 other linesAndChars-table doc, remove the gate.
"""
import win32com.client, os, json

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    n_tables = wdoc.Tables.Count
    print(f"Document has {n_tables} table(s)")

    all_diffs = []
    for ti in range(1, n_tables + 1):
        tbl = wdoc.Tables(ti)
        first_cell = tbl.Cell(1,1).Range.Text.strip()[:30]
        n_rows = tbl.Rows.Count
        print(f"\n=== Table {ti} ({n_rows} rows): [{first_cell}]")
        positions = []
        for ri in range(1, n_rows + 1):
            try:
                rng = tbl.Rows(ri).Range
                pg = rng.Information(3)
                y = rng.Information(6)
            except Exception as e:
                print(f"  row{ri} Range err: {e}")
                continue
            try: h = tbl.Rows(ri).Height
            except Exception as e: h = None
            try: rule = tbl.Rows(ri).HeightRule
            except Exception as e: rule = None
            positions.append((ri, pg, y, h, rule))
        for idx in range(len(positions) - 1):
            ri, pg, y, h, rule = positions[idx]
            nri, npg, ny, nh, nrule = positions[idx+1]
            rh = round(ny - y, 2) if (y and ny and pg==npg) else None
            rule_str = {0:'auto', 1:'atLeast', 2:'exact'}.get(rule, str(rule))
            print(f"  row{ri:3d}  p{pg} y={y:.2f} rule={rule_str} rendered_h={rh}")
            if rh is not None and 0 < rh < 60:
                all_diffs.append(rh)
        if positions:
            ri, pg, y, h, rule = positions[-1]
            rule_str = {0:'auto', 1:'atLeast', 2:'exact'}.get(rule, str(rule))
            print(f"  row{ri:3d}  p{pg} y={y:.2f} rule={rule_str} rendered_h=None (last)")

    if all_diffs:
        print(f"\n=== Summary: {len(all_diffs)} rendered heights")
        from collections import Counter
        c = Counter(all_diffs)
        for h, cnt in sorted(c.items()):
            print(f"  {h:.2f}pt: {cnt}x")
        print(f"  min={min(all_diffs):.2f} median={sorted(all_diffs)[len(all_diffs)//2]:.2f} max={max(all_diffs):.2f}")
        # Check: is any cluster of heights at linePitch multiples (17.5pt)?
        LINEPITCH = 17.5
        snapped = sum(1 for h in all_diffs if abs(h % LINEPITCH) < 0.5 or abs((h % LINEPITCH) - LINEPITCH) < 0.5)
        print(f"  rows at linePitch={LINEPITCH}pt multiples: {snapped}/{len(all_diffs)}")
finally:
    wdoc.Close(False)
    word.Quit()
