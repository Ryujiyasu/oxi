"""Word vs Oxi per-paragraph divergence framework v2.

Improvements over v1:
- Extract Oxi first_x per line (apples-to-apples X comparison)
- Per-line n_chars on both sides
- Better Word-Oxi line matching by text prefix
- Structured JSON output for cross-doc aggregation
- Skip empty paragraphs and noise-only matches

Usage:
    python tools/metrics/word_oxi_diff_v2.py <docx_path> [--out path.json] [--quiet]

For multi-doc batch:
    python tools/metrics/word_oxi_diff_v2.py --batch doc1.docx doc2.docx --out-dir dir/
"""
import sys, os, json, argparse, time, subprocess
from collections import defaultdict

WD_VPOS = 6
WD_HPOS = 5
WD_PAGE = 3

def extract_word(docx_path):
    """Extract per-paragraph Word data via COM."""
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    paras = []
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            doc.Repaginate()
            time.sleep(0.3)
            n_paras = doc.Paragraphs.Count
            for pi in range(1, n_paras + 1):
                try:
                    p = doc.Paragraphs(pi)
                    pr = p.Range
                    cs = doc.Range(pr.Start, pr.Start)
                    page = int(cs.Information(WD_PAGE))
                    if page != 1: continue
                    y0 = float(cs.Information(WD_VPOS))
                    x0 = float(cs.Information(WD_HPOS))
                    txt = pr.Text.replace('\r','').replace('\x07','').replace('\n','/')
                    if not txt:
                        continue  # skip empty
                    n_chars = len(txt)
                    # Walk chars to detect line breaks
                    line_groups = []
                    cur = {"y": y0, "first_x": x0, "last_x": x0, "n_chars": 1, "first_char_idx": 0}
                    for ci in range(1, n_chars):
                        crng = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                        cy = float(crng.Information(WD_VPOS))
                        cx = float(crng.Information(WD_HPOS))
                        if abs(cy - cur["y"]) > 1.0:
                            line_groups.append(cur)
                            cur = {"y": cy, "first_x": cx, "last_x": cx, "n_chars": 1, "first_char_idx": ci}
                        else:
                            cur["last_x"] = cx
                            cur["n_chars"] += 1
                    line_groups.append(cur)
                    paras.append({
                        "idx": pi,
                        "text": txt[:60],
                        "n_chars_total": n_chars,
                        "n_lines": len(line_groups),
                        "lines": line_groups,
                        "in_table": p.Range.Tables.Count > 0,
                        "line_spacing": float(p.Format.LineSpacing),
                    })
                except Exception as ex:
                    pass
        finally:
            doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return paras

def extract_oxi(docx_path, layout_out):
    """Extract per-line Oxi text data from layout dump."""
    RENDERER = os.path.abspath(os.path.join(
        os.path.dirname(__file__), "..", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe"
    ))
    proc = subprocess.run(
        [RENDERER, docx_path, "/tmp/dummy.png", f"--dump-layout={layout_out}"],
        capture_output=True, timeout=120,
    )
    if proc.returncode != 0:
        return None
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    if not layout.get('pages'):
        return []
    page1 = layout['pages'][0]
    text_elems = sorted([e for e in page1['elements'] if e.get('type') == 'text'],
                        key=lambda e: (e['y'], e['x']))
    # Group by approximate Y (within 1pt)
    lines = []
    cur_y = None
    cur_elems = []
    for e in text_elems:
        if cur_y is None or abs(e['y'] - cur_y) < 1.0:
            cur_elems.append(e)
            cur_y = e['y']
        else:
            lines.append(_finalize_oxi_line(cur_y, cur_elems))
            cur_elems = [e]
            cur_y = e['y']
    if cur_elems:
        lines.append(_finalize_oxi_line(cur_y, cur_elems))
    return lines

def _finalize_oxi_line(y, elems):
    elems_sorted = sorted(elems, key=lambda e: e['x'])
    first_x = elems_sorted[0]['x']
    last_x = max(e['x'] + e.get('w', 0) for e in elems)
    text = ''.join(e.get('text', '') for e in elems_sorted)
    n_chars = len(text)
    return {"y": y, "first_x": first_x, "last_x": last_x, "n_chars": n_chars, "text": text[:60]}

def cross_join(word_paras, oxi_lines):
    """Match Word paragraphs to Oxi lines by Y proximity + text similarity.
    Returns list of {word_para, oxi_lines: [matched lines], divergence}."""
    matches = []
    used_oxi = set()
    for wp in word_paras:
        wp_y = wp["lines"][0]["y"]
        wp_text = wp["text"][:8]
        # Find best Oxi line: small y diff + text overlap
        best_idx = None
        best_score = 999
        for oi, ol in enumerate(oxi_lines):
            if oi in used_oxi: continue
            y_diff = abs(ol["y"] - wp_y)
            if y_diff > 60: continue  # too far
            # text similarity: longest common substring of first chars
            text_score = 0
            if wp_text and ol["text"]:
                # check if any 3-char prefix of word matches anywhere in oxi text first 30 chars
                for i in range(len(wp_text) - 2):
                    if wp_text[i:i+3] in ol["text"][:30]:
                        text_score = 5
                        break
            score = y_diff - text_score
            if score < best_score:
                best_score = score
                best_idx = oi
        oxi_for_wp = []
        if best_idx is not None:
            oxi_for_wp.append(oxi_lines[best_idx])
            used_oxi.add(best_idx)
            # If Word has multi-line, grab next N-1 oxi lines that aren't used and within reasonable y
            if wp["n_lines"] > 1:
                last_y = oxi_lines[best_idx]["y"]
                for oi in range(best_idx + 1, len(oxi_lines)):
                    if oi in used_oxi: continue
                    if oxi_lines[oi]["y"] - last_y > 30: break
                    oxi_for_wp.append(oxi_lines[oi])
                    used_oxi.add(oi)
                    last_y = oxi_lines[oi]["y"]
                    if len(oxi_for_wp) >= wp["n_lines"]: break
        # Compute divergence
        div = {"y_diff": None, "n_lines_diff": None, "first_x_diff": None, "last_x_diff_line1": None}
        if oxi_for_wp:
            o1 = oxi_for_wp[0]
            div["y_diff"] = round(o1["y"] - wp_y, 2)
            div["n_lines_diff"] = len(oxi_for_wp) - wp["n_lines"]
            div["first_x_diff"] = round(o1["first_x"] - wp["lines"][0]["first_x"], 2)
            div["last_x_diff_line1"] = round(o1["last_x"] - wp["lines"][0]["last_x"], 2)
        matches.append({"word": wp, "oxi_lines": oxi_for_wp, "divergence": div})
    return matches

def print_summary(doc_id, matches):
    print(f"\n=== {doc_id} ===")
    print(f"{'pi':>3} {'text':35s} {'Δy':>6} {'Δlines':>7} {'Δfx':>6} {'Δlx1':>6}")
    for m in matches:
        wp = m["word"]; div = m["divergence"]
        if div["y_diff"] is None:
            mark = " (no oxi match)"
            print(f"{wp['idx']:>3} {wp['text'][:35]:35s} {'?':>6s} {'?':>7s} {'?':>6s} {'?':>6s}{mark}")
            continue
        dy = div["y_diff"]; dn = div["n_lines_diff"]; dfx = div["first_x_diff"]; dlx = div["last_x_diff_line1"]
        marks = []
        if abs(dy) > 3: marks.append("Y")
        if dn != 0: marks.append(f"L{dn:+d}")
        if abs(dfx) > 5: marks.append("FX")
        if abs(dlx) > 10: marks.append("LX")
        mark = " " + ",".join(marks) if marks else ""
        print(f"{wp['idx']:>3} {wp['text'][:35]:35s} {dy:>+6.2f} {dn:>+7d} {dfx:>+6.2f} {dlx:>+6.2f}{mark}")

def aggregate(matches, doc_id):
    """Reduce per-para matches to summary stats for cross-doc aggregation."""
    n_paras = len(matches)
    n_matched = sum(1 for m in matches if m["divergence"]["y_diff"] is not None)
    y_diffs = [m["divergence"]["y_diff"] for m in matches if m["divergence"]["y_diff"] is not None]
    n_lines_diffs = [m["divergence"]["n_lines_diff"] for m in matches if m["divergence"]["n_lines_diff"] is not None]
    n_well_aligned = sum(1 for d in y_diffs if abs(d) <= 3)
    n_drifting = sum(1 for d in y_diffs if abs(d) > 10)
    n_extra_line = sum(1 for d in n_lines_diffs if d > 0)
    n_short_line = sum(1 for d in n_lines_diffs if d < 0)
    return {
        "doc_id": doc_id,
        "n_paras_p1": n_paras,
        "n_matched": n_matched,
        "max_y_drift": max(y_diffs) if y_diffs else 0,
        "min_y_drift": min(y_diffs) if y_diffs else 0,
        "well_aligned_pct": round(100 * n_well_aligned / max(n_matched, 1), 1),
        "drifting_pct": round(100 * n_drifting / max(n_matched, 1), 1),
        "n_extra_line": n_extra_line,
        "n_short_line": n_short_line,
    }

def process_one(docx_path, quiet=False):
    doc_id = os.path.basename(docx_path).replace('.docx', '')
    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/diffv2_{doc_id}.json"
    word_data = extract_word(docx_path)
    oxi_data = extract_oxi(docx_path, layout_out)
    if oxi_data is None:
        print(f"  {doc_id}: Oxi render FAILED")
        return None
    matches = cross_join(word_data, oxi_data)
    if not quiet:
        print_summary(doc_id, matches)
    summary = aggregate(matches, doc_id)
    return {"doc_id": doc_id, "summary": summary, "matches": matches}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", nargs="*", help="docx file(s)")
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    if not args.docx:
        print("Usage: word_oxi_diff_v2.py <docx_path>...")
        sys.exit(1)
    summaries = []
    for path in args.docx:
        result = process_one(os.path.abspath(path), quiet=args.quiet)
        if result:
            summaries.append(result["summary"])
            if args.out_dir:
                os.makedirs(args.out_dir, exist_ok=True)
                out = os.path.join(args.out_dir, f"{result['doc_id']}.json")
                with open(out, "w", encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                print(f"  Saved {out}")

    if len(summaries) > 1:
        print(f"\n=== Cross-doc summary ===")
        print(f"  {'doc_id':45s} {'n_p1':>5} {'matched':>7} {'aligned%':>9} {'drift%':>7} {'+lines':>7} {'-lines':>7} {'maxΔy':>7}")
        for s in summaries:
            print(f"  {s['doc_id'][:45]:45s} {s['n_paras_p1']:>5} {s['n_matched']:>7} {s['well_aligned_pct']:>9.1f} {s['drifting_pct']:>7.1f} {s['n_extra_line']:>7} {s['n_short_line']:>7} {s['max_y_drift']:>7.2f}")

if __name__ == "__main__":
    main()
