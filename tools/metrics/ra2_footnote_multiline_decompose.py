"""
Ra2: Decompose footnote multi-line 43pt block height (§9.2 follow-up).

Prior finding (ra2_footnote_separator_formula.py):
  - 1-visual-line fn (default Calibri 10.5pt): 17.5pt block
  - 2-visual-line fn: 43pt block

Hypothesis candidates:
  H_FIXED_SA: each fn = (visual_lines × nat_lh) + sa_per_fn
              For sa=4pt: 1-line=13.5+4=17.5 ✓; 2-line=27+4=31 ✗
              For sa=16pt: 1-line=13.5+16=29.5 ✗; 2-line=27+16=43 ✓
              → constant sa doesn't fit both.

  H_LINE_PLUS_SA_HEAD: each fn = (visual_lines × line_h) + paragraph_sa
              line_h could be different from nat_lh (e.g., 13.5 inter-line, with
              a "trailing line full-height" of e.g. 17.5pt that is the LAST line.
              Like header: block = (N-1) × inter_gap + last_line_full_h.
              1-line: 0×? + 17.5 = 17.5pt ✓
              2-line: 1×? + ? + 17.5 = 43pt → if last_line_full=17.5 then
                      remaining 25.5pt for 1 inter-gap.
              That means inter-line gap (between wrap lines) ≈ 25.5pt? Suspicious.

  H_SCALED: 1-line=17.5 = something_fixed_per_fn (e.g. 17.5pt overhead per fn)
            2-line: extra 25.5 (= 43-17.5) for 1 wrap line
            → wrap line height = 25.5pt? Let's check by adding more wrap.

  H_PARA_STYLE: footnote paragraphs may have a different lineSpacingRule
                applied by Word's Footnote Text style than just "Single 12pt"
                claimed by §9.1.

Strategy:
  Vary fn text length to produce 1, 2, 3, 4 visual lines. Measure each
  block height. Solve for line_h_inter + line_h_last_or_sa.

Output:
  output/ra2_footnote_multiline_decompose.json
"""
import os
import json
import time

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_footnote_multiline_decompose.json")


RPC_REJECTED_CODES = {-2147418111, -2147023174, -2147023170}

def retry(fn, *args, retries=15, delay=0.3, **kwargs):
    last_exc = None
    for i in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            code = e.args[0] if hasattr(e, "args") and len(e.args) >= 1 else None
            if code in RPC_REJECTED_CODES or "rejected" in str(e).lower():
                pythoncom.PumpWaitingMessages()
                time.sleep(delay * (1.3 ** i))
                continue
            raise
    raise last_exc


def build_doc_with_one_footnote(word, *, fn_text, fn_font_size=10.5,
                                 body_font_size=11, n_body_paras=60):
    wdoc = retry(word.Documents.Add)
    sec = retry(lambda: wdoc.Sections(1))
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    ps.LayoutMode = 0

    body_text = "\r".join(f"B{i+1}" for i in range(n_body_paras))
    wdoc.Content.Text = body_text
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = body_font_size
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0

    fns = wdoc.Footnotes
    anchor_para = wdoc.Paragraphs(1)
    anchor_range = anchor_para.Range
    ins_pt = wdoc.Range(anchor_range.End - 1, anchor_range.End - 1)
    fn = fns.Add(Range=ins_pt, Text=fn_text)
    fn.Range.Font.Name = "Calibri"
    fn.Range.Font.Size = fn_font_size

    wdoc.Repaginate()
    time.sleep(0.05)

    fn_rng = fn.Range
    rec = {
        "fn_text_chars": len(fn_text),
        "fn_font_size": fn_font_size,
        "page_height": round(ps.PageHeight, 4),
        "bottom_margin": round(ps.BottomMargin, 4),
        "fn1_y": round(fn_rng.Information(6), 4),
        "fn1_x": round(fn_rng.Information(5), 4),
    }
    # Block height = (page_h - bm) - fn1.y
    block_h = (ps.PageHeight - ps.BottomMargin) - rec["fn1_y"]
    rec["block_h"] = round(block_h, 4)

    # Skip ComputeStatistics — it's flaky and can crash Word

    wdoc.Close(False)
    return rec


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    results = []
    try:
        # Generate footnote texts of increasing length to force more visual lines.
        # Default fn font is 10.5pt Calibri, body width = 612pt - 144pt margins = 468pt.
        # Footnote available width may be 468pt minus indent. Each char ~5-6pt.
        # So ~80 chars per line. We'll target 1, 2, 3, 4, 5 visual lines.
        # Calibrate empirically.
        # Calibrate: 1-line ≈ short, multi-line via repeated content.
        # Footnote area width = body width - indent ≈ 460pt.
        # Calibri 10.5pt avg char width ~5pt → ~92 chars per line.
        # Use words with spaces so Word can wrap naturally.
        wd = "lorem "  # avoid shadowing `word` (the COM Application)
        cases = [
            (1, "x"),
            (2, wd * 25),       # ~150 chars → ~2 lines
            (3, wd * 45),       # ~270 chars → ~3 lines
            (4, wd * 65),       # ~390 chars → ~4 lines
            (5, wd * 85),       # ~510 chars → ~5 lines
        ]

        print("=== fn block height vs visual-line count ===")
        for target_lines, text in cases:
            try:
                r = build_doc_with_one_footnote(word, fn_text=text)
                r["target_lines"] = target_lines
                results.append(r)
                vl = r.get("fn_visual_lines", "?")
                print(f"  target={target_lines:2} actual_vl={vl} fn1_y={r['fn1_y']:7.2f} "
                      f"block_h={r['block_h']:6.2f}pt chars={r['fn_text_chars']}")
            except Exception as e:
                print(f"  ERR target={target_lines}: {e}")

        # Also test font size variations to confirm scaling pattern
        print("\n=== fn block_h vs (visual_lines, font_size) ===")
        size_cases = [
            (10.5, "x"),
            (10.5, wd * 25),
            (10.5, wd * 45),
            (14, "x"),
            (14, wd * 18),     # ~108 chars → ~2 lines at 14pt
            (8, "x"),
            (8, wd * 35),
        ]
        for sz, text in size_cases:
            try:
                r = build_doc_with_one_footnote(word, fn_text=text, fn_font_size=sz)
                r["phase"] = "size_var"
                results.append(r)
                vl = r.get("fn_visual_lines", "?")
                print(f"  sz={sz:5}pt vl={vl} fn1_y={r['fn1_y']:7.2f} "
                      f"block_h={r['block_h']:6.2f}pt chars={r['fn_text_chars']}")
            except Exception as e:
                print(f"  ERR sz={sz}: {e}")

    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records to {OUT_JSON}")
        try:
            word.Quit()
        except Exception as e:
            print(f"  (word.Quit failed: {e})")


if __name__ == "__main__":
    main()
