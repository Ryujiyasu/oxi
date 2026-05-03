"""Day 4: Char-level Word vs Oxi divergence framework.

For each paragraph on page 1:
- Word COM: per-char X position via doc.Range(start, start+1).Information(WD_HPOS)
- Oxi: per-text-element X+W from layout JSON dump
- Match chars by position-in-paragraph + line membership
- Report per-char delta_x and per-char inferred-width

Findings will reveal:
- Per-char width discrepancies (Word vs Oxi)
- Wrap budget allocation differences
- Cumulative drift sources within a single paragraph

Usage:
    python tools/metrics/word_oxi_char_diff.py <docx> [--max-paras N]
"""
import sys, os, json, argparse, time, subprocess
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

WD_VPOS = 6
WD_HPOS = 5
WD_PAGE = 3

def extract_word_chars(docx_path, max_paras=None):
    """For each paragraph on page 1, extract per-char (idx, x, y, char)."""
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
            if max_paras: n_paras = min(n_paras, max_paras)
            for pi in range(1, n_paras + 1):
                try:
                    p = doc.Paragraphs(pi)
                    pr = p.Range
                    cs = doc.Range(pr.Start, pr.Start)
                    page = int(cs.Information(WD_PAGE))
                    if page != 1: continue
                    txt = pr.Text.replace('\r', '').replace('\x07', '').replace('\n', '/')
                    if not txt or len(txt) > 200: continue  # skip empty + skip very long for speed
                    chars = []
                    for ci in range(len(txt)):
                        crng = doc.Range(pr.Start + ci, pr.Start + ci + 1)
                        x = float(crng.Information(WD_HPOS))
                        y = float(crng.Information(WD_VPOS))
                        chars.append({"i": ci, "ch": txt[ci], "x": x, "y": y})
                    paras.append({
                        "idx": pi,
                        "text": txt[:60],
                        "n_chars": len(txt),
                        "chars": chars,
                        "in_table": p.Range.Tables.Count > 0,
                    })
                except Exception as ex:
                    pass
        finally:
            doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return paras

def extract_oxi_lines(docx_path, layout_out):
    """Extract per-line (y, [text_elements with x+w])."""
    RENDERER = os.path.abspath(os.path.join(
        os.path.dirname(__file__), "..", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe"
    ))
    proc = subprocess.run(
        [RENDERER, docx_path, "/tmp/dummy.png", f"--dump-layout={layout_out}"],
        capture_output=True, timeout=120,
    )
    if proc.returncode != 0: return None
    with open(layout_out, encoding='utf-8') as f:
        layout = json.load(f)
    if not layout.get('pages'): return []
    text_elems = sorted(
        [e for e in layout['pages'][0]['elements'] if e.get('type') == 'text'],
        key=lambda e: (e['y'], e['x']),
    )
    # Group by Y
    lines = []
    cur_y = None; cur = []
    for e in text_elems:
        if cur_y is None or abs(e['y'] - cur_y) < 1.0:
            cur.append(e); cur_y = e['y']
        else:
            lines.append({"y": cur_y, "elems": cur})
            cur = [e]; cur_y = e['y']
    if cur:
        lines.append({"y": cur_y, "elems": cur})
    # For each line, expand multi-char text elements into per-char x positions
    for ln in lines:
        chars = []
        for e in ln['elems']:
            # Each element may have multiple chars; estimate per-char x by even split
            txt = e.get('text', '')
            n = len(txt)
            if n == 0: continue
            w = e.get('w', 0)
            for i, ch in enumerate(txt):
                # approximate x = e.x + (i / n) * w (linear; not perfect for variable width but OK first-order)
                chars.append({"ch": ch, "x": e['x'] + (i / n) * w if n > 0 else e['x']})
        ln['chars'] = chars
    return lines

def match_paragraph_to_oxi_lines(wp, oxi_lines, used):
    """Find Oxi line(s) matching Word paragraph wp by Y proximity + text match."""
    wp_y = wp['chars'][0]['y']
    wp_text_prefix = wp['text'][:8]
    best_idx = None
    best_score = 999
    for oi, ol in enumerate(oxi_lines):
        if oi in used: continue
        y_diff = abs(ol['y'] - wp_y)
        if y_diff > 60: continue
        text_score = 0
        ol_text = ''.join(c['ch'] for c in ol.get('chars', []))[:30]
        for i in range(len(wp_text_prefix) - 2):
            if wp_text_prefix[i:i+3] in ol_text:
                text_score = 5
                break
        score = y_diff - text_score
        if score < best_score:
            best_score = score
            best_idx = oi
    return best_idx

def compare_para(wp, oxi_lines, used):
    """Match wp to oxi line(s); compute per-char x diff in line 1."""
    best_idx = match_paragraph_to_oxi_lines(wp, oxi_lines, used)
    if best_idx is None:
        return {"matched": False}
    used.add(best_idx)
    ol = oxi_lines[best_idx]
    # Match chars between wp.chars (line 1 only) and ol.chars
    # Word line 1: chars before any y change
    wp_y0 = wp['chars'][0]['y']
    wp_line1 = [c for c in wp['chars'] if abs(c['y'] - wp_y0) < 1.0]
    ol_chars = ol['chars']
    n_w = len(wp_line1)
    n_o = len(ol_chars)
    n_min = min(n_w, n_o)
    diffs = []
    for i in range(n_min):
        wc = wp_line1[i]
        oc = ol_chars[i]
        diffs.append({
            "i": i, "ch_w": wc['ch'], "ch_o": oc['ch'],
            "x_w": wc['x'], "x_o": oc['x'],
            "dx": round(oc['x'] - wc['x'], 2),
        })
    return {
        "matched": True,
        "n_w_line1": n_w, "n_o_line1": n_o,
        "diffs": diffs,
    }

def print_para_summary(wp, cmp):
    if not cmp.get("matched"):
        print(f"  P{wp['idx']:>3} '{wp['text'][:30]}' UNMATCHED")
        return
    diffs = cmp["diffs"]
    if not diffs: return
    # Find first divergence point
    for d in diffs:
        if abs(d["dx"]) > 1.5:
            first_div = d
            break
    else:
        first_div = None
    last = diffs[-1]
    print(f"  P{wp['idx']:>3} '{wp['text'][:30]:30s}' nW={cmp['n_w_line1']:>3} nO={cmp['n_o_line1']:>3} | first_div={'i' + str(first_div['i']) + ' dx=' + str(first_div['dx']) if first_div else 'NONE'} | last_dx={last['dx']:+.2f}")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--max-paras", type=int, default=20)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    doc_path = os.path.abspath(args.docx)

    print(f"=== Word per-char extract: {os.path.basename(doc_path)} (limit {args.max_paras} paras) ===")
    word_paras = extract_word_chars(doc_path, max_paras=args.max_paras)
    print(f"  Got {len(word_paras)} non-empty paras with chars on page 1")

    layout_out = f"C:/Users/ryuji/AppData/Local/Temp/cdiff_{os.path.basename(doc_path)}.json"
    print(f"=== Oxi per-char extract ===")
    oxi_lines = extract_oxi_lines(doc_path, layout_out)
    if oxi_lines is None:
        print("  Oxi render FAILED"); return
    print(f"  Got {len(oxi_lines)} text-bearing lines on page 1")

    print(f"\n=== Per-paragraph compare ===")
    used = set()
    results = []
    for wp in word_paras:
        cmp = compare_para(wp, oxi_lines, used)
        print_para_summary(wp, cmp)
        results.append({"word_idx": wp['idx'], "text": wp['text'], "compare": cmp})

    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            json.dump({"doc": doc_path, "results": results}, f, ensure_ascii=False, indent=2)
        print(f"\nSaved {args.out}")

if __name__ == "__main__":
    main()
