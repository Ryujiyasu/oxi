"""Measure ALL paragraph Y positions in gen2_001 and extract exact advances."""
import win32com.client
import os
import glob
import math

docx = glob.glob(os.path.join(os.path.abspath("tools/golden-test/documents/docx"), "gen2_001*"))[0]

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx, ReadOnly=True)

    paras = []
    for i in range(1, 40):
        p = doc.Paragraphs(i)
        rng = p.Range
        page = rng.Information(3)
        if page > 1:
            break
        y = rng.Information(6)
        ls = p.Format.LineSpacing
        rule = p.Format.LineSpacingRule
        fs = rng.Font.Size
        sa = p.Format.SpaceAfter
        sb = p.Format.SpaceBefore

        end_rng = doc.Range(rng.End - 1, rng.End)
        end_y = end_rng.Information(6)
        n_lines = max(1, round((end_y - y) / 13) + 1) if end_y > y + 5 else 1

        # Extract style from XML
        import re
        xml_snippet = rng.XML[:1000]
        pstyle_m = re.search(r'w:pStyle w:val="([^"]+)"', xml_snippet)
        pstyle = pstyle_m.group(1) if pstyle_m else "Normal"

        paras.append({
            'i': i, 'y': y, 'end_y': end_y, 'fs': fs, 'ls': ls, 'rule': rule,
            'sa': sa, 'sb': sb, 'n_lines': n_lines, 'style': pstyle
        })

    doc.Close(False)

    # Print and compute advances
    print(f"{'P':>3} {'y':>7} {'end_y':>7} {'gap':>6} {'fs':>4} {'ls':>5} {'rule':>2} {'sa':>4} {'sb':>4} {'ln':>2} {'style':>15}")
    for k, r in enumerate(paras):
        gap = r['y'] - paras[k-1]['end_y'] if k > 0 else 0
        print(f"{r['i']:3d} {r['y']:7.1f} {r['end_y']:7.1f} {gap:6.1f} {r['fs']:4.0f} {r['ls']:5.1f} {r['rule']:2d} {r['sa']:4.0f} {r['sb']:4.0f} {r['n_lines']:2d} {r['style']:>15}")

    # Now try to find j offset using ROUND
    print("\n=== Trying ROUND cumulative with shared j ===")
    raw_tw = 14.25 * 1.15 * 20  # body 11pt

    # Count ALL paragraphs as contributing to j
    j = 0
    for k, r in enumerate(paras):
        if r['rule'] != 5:  # only Multiple spacing
            continue

        if k > 0:
            gap = r['y'] - paras[k-1]['end_y']
            collapsed = max(paras[k-1]['sa'], r['sb'])

            # Use this para's own raw_tw for advance
            para_raw_tw = math.floor(r['fs'] * 83/64 * 8) / 8 * 1.15 * 20

            # ROUND advance at current j
            if j == 0:
                adv_ceil = math.ceil(para_raw_tw / 10) * 10 / 20
                adv_round = round(para_raw_tw / 10) * 10 / 20
                predicted = adv_ceil  # j=0 uses CEIL
            else:
                cn = round((j+1) * para_raw_tw / 10) * 10
                cc = round(j * para_raw_tw / 10) * 10
                predicted = (cn - cc) / 20

            # Check if spacing was contextually suppressed
            actual_advance = gap - collapsed if collapsed > 0 and gap > collapsed + 5 else gap
            if collapsed > 0 and abs(gap - predicted) < 0.5:
                actual_advance = gap  # spacing was suppressed

            diff = actual_advance - predicted
            ok = "OK" if abs(diff) < 0.5 else f"DIFF({diff:+.1f})"
            print(f"  P{r['i']:2d} j={j:2d} fs={r['fs']:4.0f} gap={gap:5.1f} advance={actual_advance:5.1f} predicted={predicted:5.1f} {ok}")

        j += r['n_lines']

finally:
    word.Quit()
