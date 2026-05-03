"""Word vs Oxi per-paragraph divergence framework.

For a given docx:
1. Word COM: extract per-paragraph y, x, line_count, last_x per line, line_spacing
2. Oxi GDI layout dump: extract same fields
3. Cross-join by paragraph order; compute diffs
4. Output JSON + tabular summary

Usage:
    python tools/metrics/word_oxi_diff.py <docx_path> [--out path.json]

The diff structure surfaces:
- Y-position drift (cumulative or local)
- X-position drift (indent / wrap_w issues)
- Line count diff (wrap differences)
- Per-line last_x diff (wrap point differences)
"""
import sys, os, json, argparse, time, subprocess
from collections import defaultdict

WD_VPOS = 6
WD_HPOS = 5
WD_PAGE = 3

def extract_word_paragraphs(docx_path):
    """Extract per-paragraph Word data via COM."""
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    paras_data = []
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            doc.Repaginate()
            time.sleep(0.3)
            n = doc.Paragraphs.Count
            for pi in range(1, n + 1):
                try:
                    p = doc.Paragraphs(pi)
                    pr = p.Range
                    cs_start = doc.Range(pr.Start, pr.Start)
                    page = int(cs_start.Information(WD_PAGE))
                    if page != 1:
                        continue  # only page 1 for now
                    y = float(cs_start.Information(WD_VPOS))
                    x = float(cs_start.Information(WD_HPOS))
                    txt = pr.Text.replace('\r', '').replace('\x07', '').replace('\n', '/')
                    if not txt:
                        # empty para
                        paras_data.append({
                            "idx": pi,
                            "text_prefix": "",
                            "y": y, "x": x,
                            "n_chars": 0,
                            "n_lines_word": 0,
                            "lines_word": [],
                            "line_spacing": float(p.Format.LineSpacing),
                            "in_table": p.Range.Tables.Count > 0,
                        })
                        continue
                    # Find line breaks: iterate chars, group by y
                    n_chars = len(txt)
                    line_groups = []  # list of (y_first, last_x)
                    cur_y = y
                    last_x = x
                    n_lines = 1
                    for ci in range(n_chars):
                        crng = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                        cy = float(crng.Information(WD_VPOS))
                        cx = float(crng.Information(WD_HPOS))
                        if abs(cy - cur_y) > 1.0:
                            # New line
                            line_groups.append({"y": cur_y, "last_x": last_x, "n_chars": ci - sum(g.get("n_chars", 0) for g in line_groups)})
                            cur_y = cy
                            n_lines += 1
                        last_x = cx
                    # Final line
                    n_chars_so_far = sum(g.get("n_chars", 0) for g in line_groups)
                    line_groups.append({"y": cur_y, "last_x": last_x, "n_chars": n_chars - n_chars_so_far})
                    paras_data.append({
                        "idx": pi,
                        "text_prefix": txt[:30],
                        "y": y, "x": x,
                        "n_chars": n_chars,
                        "n_lines_word": n_lines,
                        "lines_word": line_groups,
                        "line_spacing": float(p.Format.LineSpacing),
                        "in_table": p.Range.Tables.Count > 0,
                    })
                except Exception as ex:
                    paras_data.append({"idx": pi, "error": str(ex)[:80]})
        finally:
            doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return paras_data

def extract_oxi_paragraphs(docx_path, layout_out):
    """Extract per-paragraph Oxi data via layout JSON dump."""
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
    # Group text elements by approximate paragraph (paragraph_index when present, else by x/y locality)
    # For now: group by para_idx if present; else by clustering
    text_elems = [e for e in page1['elements'] if e.get('type') == 'text']
    # Sort by y then x
    text_elems.sort(key=lambda e: (e['y'], e['x']))
    # Group consecutive same-y as same line
    lines = []  # list of (y, [elements])
    cur_y = None
    cur_line = []
    for e in text_elems:
        if cur_y is None or abs(e['y'] - cur_y) < 1.0:
            cur_line.append(e)
            cur_y = e['y']
        else:
            lines.append({"y": cur_y, "elements": cur_line})
            cur_line = [e]
            cur_y = e['y']
    if cur_line:
        lines.append({"y": cur_y, "elements": cur_line})
    # Each line: compute first_x, last_x
    for ln in lines:
        ln["first_x"] = min(e['x'] for e in ln['elements'])
        ln["last_x"] = max(e['x'] + e.get('w', 0) for e in ln['elements'])
        ln["text"] = ''.join(e.get('text', '') for e in ln['elements'])[:30]
        ln["n_chars"] = sum(len(e.get('text', '')) for e in ln['elements'])
    return lines

def cross_join(word_paras, oxi_lines):
    """Match Word paragraphs to Oxi lines by text prefix + Y proximity."""
    matches = []
    used_oxi = set()
    for wp in word_paras:
        if "error" in wp: continue
        if wp.get("n_lines_word", 0) == 0:
            matches.append({"word": wp, "oxi_lines": []})
            continue
        # Find Oxi line(s) starting near wp.y with matching text prefix
        wp_first_line_text = wp["lines_word"][0].get("n_chars", 0)
        # Match by Y proximity (within 30pt) and text prefix
        best_match = None
        best_score = 999
        for oi, ol in enumerate(oxi_lines):
            if oi in used_oxi: continue
            y_diff = abs(ol["y"] - wp["y"])
            if y_diff > 30: continue
            text_match = wp["text_prefix"][:5] in ol["text"][:30] or ol["text"][:5] in wp["text_prefix"][:30]
            score = y_diff - (10 if text_match else 0)
            if score < best_score:
                best_score = score
                best_match = oi
        if best_match is not None:
            # Greedily collect the wrap lines too
            oxi_for_wp = [oxi_lines[best_match]]
            used_oxi.add(best_match)
            # Look for next N-1 lines if word has > 1 line
            n_word_lines = wp["n_lines_word"]
            if n_word_lines > 1:
                for oi in range(best_match + 1, len(oxi_lines)):
                    if oi in used_oxi: continue
                    if oxi_lines[oi]["y"] - oxi_lines[best_match]["y"] > 100: break
                    oxi_for_wp.append(oxi_lines[oi])
                    used_oxi.add(oi)
                    if len(oxi_for_wp) >= n_word_lines: break
            matches.append({"word": wp, "oxi_lines": oxi_for_wp})
        else:
            matches.append({"word": wp, "oxi_lines": []})
    return matches

def print_summary(matches):
    print(f"{'pi':>3} {'text':30s} {'W_y':>7} {'O_y':>7} {'Δy':>6} {'W_n':>4} {'O_n':>4} {'Δn':>4} {'W_lx':>7} {'O_lx':>7} {'Δlx':>6}")
    for m in matches:
        wp = m["word"]
        if "error" in wp: continue
        w_y = wp.get("y", 0)
        w_n = wp.get("n_lines_word", 0)
        w_lx = wp["lines_word"][0]["last_x"] if w_n > 0 else 0
        ol = m["oxi_lines"]
        o_y = ol[0]["y"] if ol else None
        o_n = len(ol)
        o_lx = ol[0]["last_x"] if ol else None
        if o_y is None:
            print(f"{wp['idx']:>3} {wp['text_prefix']:30s} {w_y:>7.2f} {'?':>7s} {'?':>6s} {w_n:>4} {'?':>4s} {'?':>4s} {w_lx:>7.2f} {'?':>7s} {'?':>6s}")
        else:
            dy = o_y - w_y
            dn = o_n - w_n
            dlx = o_lx - w_lx
            mark = ""
            if abs(dy) > 2 or abs(dlx) > 5 or dn != 0:
                mark = " <<<"
            print(f"{wp['idx']:>3} {wp['text_prefix']:30s} {w_y:>7.2f} {o_y:>7.2f} {dy:>+6.2f} {w_n:>4} {o_n:>4} {dn:>+4} {w_lx:>7.2f} {o_lx:>7.2f} {dlx:>+6.2f}{mark}")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", help="path to docx file")
    ap.add_argument("--out", default=None, help="JSON output path")
    args = ap.parse_args()
    docx_path = os.path.abspath(args.docx)

    print(f"=== Word COM extraction: {os.path.basename(docx_path)} ===")
    word_data = extract_word_paragraphs(docx_path)
    print(f"  Got {len(word_data)} paragraphs on page 1")

    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/wod_oxi_{os.path.basename(docx_path)}.json"
    print(f"=== Oxi layout extraction ===")
    oxi_data = extract_oxi_paragraphs(docx_path, layout_out)
    if oxi_data is None:
        print("  Oxi render FAILED")
        return
    print(f"  Got {len(oxi_data)} text-bearing lines on page 1")

    print(f"\n=== Cross-join + diff ===")
    matches = cross_join(word_data, oxi_data)
    print_summary(matches)

    if args.out:
        with open(args.out, "w", encoding='utf-8') as f:
            json.dump({"doc": docx_path, "matches": matches}, f, ensure_ascii=False, indent=2)
        print(f"\nSaved {args.out}")

if __name__ == "__main__":
    main()
