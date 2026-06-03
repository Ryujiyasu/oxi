# -*- coding: utf-8 -*-
"""S492 big-job: CONTENT-based per-glyph render-truth matcher (replaces the broken order-match).

Word EMF gives per-RUN (x,y)+text+Dx(inter-glyph advances) -> expand to per-GLYPH (char,x,y).
Oxi dump gives per-CHAR (char,x,y). Both in reading order. Align the two char streams with
difflib (handles Word/Oxi wrapping differently -> wn!=on). For matched glyphs compute (ox-wx),
(oy-wy); CALIBRATE by subtracting the median (= content-relative page-margin offset, the EMF is
selection-relative). Residual dx/dy = the REAL per-glyph error. Decides horizontal-vs-vertical cap.

Usage: rendertruth_match_glyph.py <emf_pos.json> <oxi_dump.json> <page_idx0> <out.json>
cp932-safe: UTF-8 file, results to JSON, ASCII verdicts only."""
import json, sys, difflib, statistics

SC = 0.12  # EMF raw device units -> points (calibrated S492, scale_device_pt_per_unit)


def cluster_lines(glyphs, tol):
    """glyphs: list of dicts with 'y'. Cluster into lines by y proximity, sort each by x."""
    gs = sorted(glyphs, key=lambda g: (round(g['y'], 1), g['x']))
    lines = []
    for g in gs:
        if lines and abs(g['y'] - lines[-1]['y']) < tol:
            lines[-1]['gs'].append(g)
            lines[-1]['y'] = (lines[-1]['y'] * len(lines[-1]['gs'][:-1]) + g['y']) / len(lines[-1]['gs'])
        else:
            lines.append({'y': g['y'], 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
    return lines


def main():
    emf_path, dump_path, pidx, out = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]
    W = json.load(open(emf_path, encoding='utf-8'))['records']
    O = json.load(open(dump_path, encoding='utf-8'))['pages'][pidx]['elements']

    # --- Word per-glyph stream: expand each run via Dx ---
    wg = []
    for r in W:
        t = r['text']
        if not t.strip():
            continue
        dx = r.get('dx', [])
        x = r['x']
        y = r['y']
        for k, ch in enumerate(t):
            wg.append({'char': ch, 'x': x * SC, 'y': y * SC})
            x += dx[k] if k < len(dx) else 75
    # reading order: cluster Word glyphs into lines (tol ~ half line pitch in pt)
    wlines = cluster_lines(wg, tol=5.0)

    # --- Oxi per-glyph stream ---
    og = [{'char': e['text'], 'x': e['x'], 'y': e['y'], 'w': e.get('w', 0), 'fs': e.get('font_size', 0)}
          for e in O if e.get('type') == 'text' and e.get('text', '').strip()]
    olines = cluster_lines(og, tol=5.0)

    # flatten in reading order
    wseq = [g for L in wlines for g in L['gs']]
    oseq = [g for L in olines for g in L['gs']]
    wchars = [g['char'] for g in wseq]
    ochars = [g['char'] for g in oseq]

    sm = difflib.SequenceMatcher(None, wchars, ochars, autojunk=False)
    matched = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for d in range(i2 - i1):
                w, o = wseq[i1 + d], oseq[j1 + d]
                matched.append({'char': w['char'], 'wx': w['x'], 'wy': w['y'],
                                'ox': o['x'], 'oy': o['y'], 'ow': o['w'], 'fs': o['fs']})

    ratio = sm.ratio()
    cal_x = statistics.median(m['ox'] - m['wx'] for m in matched)
    cal_y = statistics.median(m['oy'] - m['wy'] for m in matched)
    for m in matched:
        m['rdx'] = round((m['ox'] - m['wx']) - cal_x, 2)  # residual horizontal (Oxi-Word, calib-removed)
        m['rdy'] = round((m['oy'] - m['wy']) - cal_y, 2)  # residual vertical

    rdxs = [m['rdx'] for m in matched]
    rdys = [m['rdy'] for m in matched]
    summary = {
        'word_glyphs': len(wg), 'oxi_glyphs': len(og), 'word_lines': len(wlines), 'oxi_lines': len(olines),
        'matched': len(matched), 'match_ratio': round(ratio, 4),
        'cal_x_pt': round(cal_x, 2), 'cal_y_pt': round(cal_y, 2),
        'rdx_mean': round(statistics.mean(rdxs), 2), 'rdx_std': round(statistics.pstdev(rdxs), 2),
        'rdx_absmean': round(statistics.mean(abs(v) for v in rdxs), 2),
        'rdx_range': [round(min(rdxs), 1), round(max(rdxs), 1)],
        'rdy_mean': round(statistics.mean(rdys), 2), 'rdy_std': round(statistics.pstdev(rdys), 2),
        'rdy_absmean': round(statistics.mean(abs(v) for v in rdys), 2),
        'rdy_range': [round(min(rdys), 1), round(max(rdys), 1)],
    }
    print("=== 0e7af content-match render-truth ===")
    for k, v in summary.items():
        print("  %-14s %s" % (k, v))
    print("\nINTERP: rdx = residual horizontal err (Oxi-Word) after removing left-margin median.")
    print("        rdy = residual vertical. If rdx_absmean >> rdy_absmean -> HORIZONTAL cap (char-x/wrap).")
    print("        If both small -> sub-pixel scatter. If rdy structured -> vertical.")

    json.dump({'summary': summary, 'matched': matched}, open(out, 'w', encoding='utf-8'),
              ensure_ascii=False)
    print("\nwrote", out)


if __name__ == '__main__':
    main()
