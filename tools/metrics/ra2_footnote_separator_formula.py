"""
Ra2: Footnote separator height + area formula investigation (§9.2).

Spec §9.2 has a placeholder formula:
  area_height = separator_height + sum(fn.line_height for fn in footnotes)
but `separator_height` is never numerically defined. The 2026-03-29 note
("Single footnote: y=752.5pt → 17.4pt above body bottom 769.9") implies
separator_height ≈ ?, but the breakdown of "17.4 = separator + line_h" is
ambiguous (could be sep=4 + line_h=13.4, or sep=12 + line_h=5.4, etc.).

Goal:
  - Pin down the **separator gap** (vertical space between body bottom and
    first footnote line top).
  - Confirm the **per-footnote line_height** (should match GDI nat_lh for
    10.5pt body default).
  - Confirm body_area shrinkage matches: body_max_y = footnote_area_top_y.

Variables:
  - n_footnotes: 1, 2, 3, 5
  - footnote text length: 1-line vs 2-line wrap (force via long text)
  - body font/size for footnote: default (10.5pt) and override (14pt)

Output:
  pipeline_data/ra2_footnote_separator_formula.json
"""
import os
import json
import time

import win32com.client
import pythoncom

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_footnote_separator_formula.json")


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


def build_doc_with_footnotes(word, *, n_footnotes, fn_text, fn_font_size,
                              body_font_size=11, n_body_paras=60):
    """Create a doc with N footnotes, each with the same fn_text content."""
    wdoc = retry(word.Documents.Add)
    sec = retry(lambda: wdoc.Sections(1))
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    ps.LayoutMode = 0  # default (no grid)

    # Body content with N footnote anchors
    body_text = "\r".join(f"B{i+1}" for i in range(n_body_paras))
    wdoc.Content.Text = body_text
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = body_font_size
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0  # single

    # Add footnotes anchored at the END of body paragraphs i (i=1..n)
    fns = wdoc.Footnotes
    for k in range(n_footnotes):
        anchor_para = wdoc.Paragraphs(k + 1)
        anchor_range = anchor_para.Range
        # Move insertion point to just before paragraph mark
        ins_pt = wdoc.Range(anchor_range.End - 1, anchor_range.End - 1)
        fn = fns.Add(Range=ins_pt, Text=fn_text)
        fn.Range.Font.Name = "Calibri"
        fn.Range.Font.Size = fn_font_size

    wdoc.Repaginate()
    time.sleep(0.05)

    # Measure footnotes
    fn_records = []
    for k in range(1, fns.Count + 1):
        fn = fns(k)
        rng = fn.Range
        fn_records.append({
            "i": k,
            "y": round(rng.Information(6), 4),
            "x": round(rng.Information(5), 4),
            "page": rng.Information(3),
            "text": rng.Text.strip()[:30],
        })

    # Find last body paragraph on page 1
    last_body_y_p1 = None
    last_body_idx = None
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        try:
            page = p.Range.Information(3)
            y = round(p.Range.Information(6), 4)
        except Exception:
            break
        if page == 1:
            last_body_y_p1 = y
            last_body_idx = i
        elif page > 1:
            break

    rec = {
        "n_footnotes": n_footnotes,
        "fn_font_size": fn_font_size,
        "body_font_size": body_font_size,
        "fn_text": fn_text[:50],
        "page_height": round(ps.PageHeight, 4),
        "bottom_margin": round(ps.BottomMargin, 4),
        "body_max_y_default": round(ps.PageHeight - ps.BottomMargin, 4),
        "footnotes": fn_records,
        "last_body_y_page1": last_body_y_p1,
        "last_body_idx": last_body_idx,
    }
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
        # Phase A — single-line footnotes, N varies
        print("=== Phase A: single-line footnote, N=1..5 ===")
        short_text = "FN_short"
        for n in [1, 2, 3, 5]:
            try:
                r = build_doc_with_footnotes(
                    word, n_footnotes=n, fn_text=short_text, fn_font_size=10.5,
                )
                r["phase"] = "A"
                results.append(r)
                print(f"  n={n}: last_body_y_p1={r['last_body_y_page1']}")
                for f in r["footnotes"][:5]:
                    print(f"    fn{f['i']}@p{f['page']} y={f['y']:7.2f} x={f['x']:6.2f} '{f['text']}'")
            except Exception as e:
                print(f"  ERR n={n}: {e}")

        # Phase B — multi-line footnotes (force wrap)
        print("\n=== Phase B: long-text footnotes (force 2+ lines), N=1, 2, 3 ===")
        long_text = "This is a long footnote text that should wrap onto multiple lines because it contains many words. " * 3
        for n in [1, 2, 3]:
            try:
                r = build_doc_with_footnotes(
                    word, n_footnotes=n, fn_text=long_text, fn_font_size=10.5,
                )
                r["phase"] = "B"
                results.append(r)
                print(f"  n={n}: last_body_y_p1={r['last_body_y_page1']}")
                for f in r["footnotes"][:5]:
                    print(f"    fn{f['i']}@p{f['page']} y={f['y']:7.2f} x={f['x']:6.2f} '{f['text']}'")
            except Exception as e:
                print(f"  ERR n={n}: {e}")

        # Phase C — different footnote font sizes
        print("\n=== Phase C: single footnote, font size varies (8, 10.5, 14, 18) ===")
        for sz in [8, 10.5, 14, 18]:
            try:
                r = build_doc_with_footnotes(
                    word, n_footnotes=1, fn_text="FN_size", fn_font_size=sz,
                )
                r["phase"] = "C"
                results.append(r)
                print(f"  fn_size={sz}: last_body_y_p1={r['last_body_y_page1']}")
                for f in r["footnotes"]:
                    print(f"    fn{f['i']}@p{f['page']} y={f['y']:7.2f} '{f['text']}'")
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
