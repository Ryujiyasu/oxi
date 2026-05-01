"""
Ra2: Section break Y handoff — §11 word_layout_spec_ra.md.

Three sub-questions answered:

§11.1 Continuous Section Break
  Q1.1 Y of section-2 first paragraph relative to section-1 last paragraph.
       Spec note: "Section 1 last y=110.5, Section 2 first y=146.5" (gap=36).
       Pin down whether gap is line_h, line_h * factor, or break-paragraph-mark.

§11.3 Continuous + Column Change
  Q3.1 Section 2 column-1 first line Y vs section 1 last line Y.
  Q3.2 Section 2 column-2 first line Y — same as column-1 start (top of column
       region) or top of page?
  Q3.3 Where does column-2 start when column-1 doesn't fill? Word "balances"
       columns — does that affect Y?
  Q3.4 What if section 1 is empty (continuous break right at page top)?

Variables:
  - body font/size (Calibri 11pt baseline; CJK MS Mincho 10.5pt)
  - section 1 body paragraph count (0, 1, 5, 20)
  - column count for section 2 (1, 2, 3)
  - column spacing (default 36pt = 0.5in)
  - balanced/unbalanced second-section last column
"""
import win32com.client
import pythoncom
import json
import os
import time


RPC_REJECTED_CODES = {-2147418111, -2147023174, -2147023170}

def retry(fn, *args, retries=20, delay=0.3, **kwargs):
    """Retry COM call when Word raises a transient RPC error.
    Pumps COM messages between retries to let Word complete pending work."""
    last_exc = None
    for i in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            code = None
            if hasattr(e, "args") and len(e.args) >= 1:
                code = e.args[0]
            transient = (
                code in RPC_REJECTED_CODES
                or "rejected" in str(e).lower()
                or "rpc" in str(e).lower()
            )
            if transient:
                pythoncom.PumpWaitingMessages()
                time.sleep(delay * (1.3 ** i))
                pythoncom.PumpWaitingMessages()
                continue
            raise
    raise last_exc


def add_doc(word):
    return retry(word.Documents.Add)


def get_section(wdoc, idx):
    return retry(lambda: wdoc.Sections(idx))

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_section_break_y_handoff.json")

WD_LINE_SPACE_SINGLE = 0
WD_SECTION_CONTINUOUS = 0
WD_SECTION_NEW_PAGE = 2


def setup_section1_basic(wdoc, top_margin=72, bottom_margin=72,
                         left_margin=72, right_margin=72,
                         font="Calibri", size=11):
    sec = get_section(wdoc, 1)
    ps = sec.PageSetup
    ps.LeftMargin = left_margin
    ps.RightMargin = right_margin
    ps.TopMargin = top_margin
    ps.BottomMargin = bottom_margin
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    return sec


def add_continuous_break(wdoc):
    """Append a continuous section break to the end of the document.
    Returns the new section's index (Sections.Count after insert)."""
    end = wdoc.Content.End - 1
    rng = wdoc.Range(end, end)
    rng.InsertBreak(3)  # wdSectionBreakContinuous=3
    return wdoc.Sections.Count


def case_continuous_basic(word):
    """Section 1 (N body paras) → continuous break → Section 2 (M body paras),
    same column count, same margins. Verify gap between last s1 and first s2."""
    results = []
    for n in [1, 2, 5]:
        for body_size in [11, 14]:
            wdoc = add_doc(word)
            sec1 = setup_section1_basic(wdoc, font="Calibri", size=body_size)
            # Section 1 body
            wdoc.Content.Text = "\r".join(f"S1_B{i+1}" for i in range(n))
            for i in range(1, wdoc.Paragraphs.Count + 1):
                p = wdoc.Paragraphs(i)
                p.Range.Font.Name = "Calibri"
                p.Range.Font.Size = body_size
                p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0

            # Insert continuous section break
            end = wdoc.Content.End - 1
            wdoc.Range(end, end).InsertBreak(3)  # wdSectionBreakContinuous

            # Append section 2 paragraphs
            end = wdoc.Content.End - 1
            rng = wdoc.Range(end, end)
            rng.InsertAfter("S2_B1\rS2_B2\rS2_B3")
            for i in range(1, wdoc.Paragraphs.Count + 1):
                p = wdoc.Paragraphs(i)
                p.Range.Font.Name = "Calibri"
                p.Range.Font.Size = body_size
                p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0

            wdoc.Repaginate()
            time.sleep(0.05)

            paras = []
            for i in range(1, wdoc.Paragraphs.Count + 1):
                p = wdoc.Paragraphs(i)
                paras.append({
                    "i": i,
                    "y": round(p.Range.Information(6), 4),
                    "x": round(p.Range.Information(5), 4),
                    "page": p.Range.Information(3),
                    "text": p.Range.Text.strip()[:20],
                })

            rec = {
                "case": "continuous_basic",
                "n_section1_paras": n,
                "body_size": body_size,
                "section_count": wdoc.Sections.Count,
                "paragraphs": paras,
            }
            results.append(rec)
            wdoc.Close(False)
    return results


def case_continuous_column_change(word):
    """Section 1 (1col) → continuous break → Section 2 (2col)."""
    results = []
    for n_s1 in [1, 5]:
        for s2_cols in [2, 3]:
            wdoc = add_doc(word)
            sec1 = setup_section1_basic(wdoc)
            # Section 1 body
            if n_s1 > 0:
                wdoc.Content.Text = "\r".join(f"S1_B{i+1}" for i in range(n_s1))
                for i in range(1, wdoc.Paragraphs.Count + 1):
                    p = wdoc.Paragraphs(i)
                    p.Range.Font.Name = "Calibri"
                    p.Range.Font.Size = 11
                    p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
                    p.Format.SpaceBefore = 0
                    p.Format.SpaceAfter = 0

            # Insert continuous section break
            end = wdoc.Content.End - 1
            wdoc.Range(end, end).InsertBreak(8)

            # Add section 2 paragraphs (enough to fill columns)
            end = wdoc.Content.End - 1
            wdoc.Range(end, end).InsertAfter(
                "\r".join(f"S2_B{i+1}" for i in range(40))
            )
            time.sleep(0.1)
            n_secs = retry(lambda: wdoc.Sections.Count)
            if n_secs < 2:
                print(f"  WARN: only {n_secs} section(s) after InsertBreak; skipping")
                wdoc.Close(False)
                continue
            for i in range(1, wdoc.Paragraphs.Count + 1):
                p = wdoc.Paragraphs(i)
                p.Range.Font.Name = "Calibri"
                p.Range.Font.Size = 11
                p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0

            sec2 = get_section(wdoc, 2)
            sec2.PageSetup.TextColumns.SetCount(s2_cols)
            try:
                sec2.PageSetup.TextColumns.LineBetween = False
            except Exception:
                pass
            # Default spacing in Word is ~36pt (0.5in) between columns

            wdoc.Repaginate()

            # Record paragraphs grouped by section + page + (x,y) so we can
            # see column structure
            paras = []
            for i in range(1, wdoc.Paragraphs.Count + 1):
                p = wdoc.Paragraphs(i)
                paras.append({
                    "i": i,
                    "y": round(p.Range.Information(6), 4),
                    "x": round(p.Range.Information(5), 4),
                    "page": p.Range.Information(3),
                    "text": p.Range.Text.strip()[:20],
                })

            sec2_props = {
                "TextColumns_count": sec2.PageSetup.TextColumns.Count,
                "TextColumns_evenlySpaced": bool(sec2.PageSetup.TextColumns.EvenlySpaced),
            }
            # First column dimensions (Word internal)
            try:
                col1 = sec2.PageSetup.TextColumns(1)
                sec2_props["col1_width"] = round(col1.Width, 4)
                sec2_props["col1_spaceAfter"] = round(col1.SpaceAfter, 4)
            except Exception as e:
                sec2_props["col_dim_error"] = str(e)

            rec = {
                "case": "continuous_column_change",
                "n_section1_paras": n_s1,
                "section2_columns": s2_cols,
                "section2_props": sec2_props,
                "paragraphs": paras,
            }
            results.append(rec)
            wdoc.Close(False)
    return results


def case_continuous_margin_change(word):
    """Section 1 default margins → continuous break → Section 2 different
    leftMargin. Verify section 2 first paragraph X."""
    wdoc = add_doc(word)
    sec1 = setup_section1_basic(wdoc)
    wdoc.Content.Text = "S1_B1\rS1_B2"
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    end = wdoc.Content.End - 1
    wdoc.Range(end, end).InsertBreak(3)  # wdSectionBreakContinuous
    end = wdoc.Content.End - 1
    wdoc.Range(end, end).InsertAfter("S2_B1\rS2_B2\rS2_B3")
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0

    sec2 = get_section(wdoc, 2)
    sec2.PageSetup.LeftMargin = 144  # 2 inch instead of 1
    sec2.PageSetup.RightMargin = 144

    wdoc.Repaginate()

    paras = []
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        paras.append({
            "i": i,
            "y": round(p.Range.Information(6), 4),
            "x": round(p.Range.Information(5), 4),
            "page": p.Range.Information(3),
            "text": p.Range.Text.strip()[:20],
        })

    rec = {
        "case": "continuous_margin_change",
        "section1_left_margin": 72,
        "section2_left_margin": 144,
        "paragraphs": paras,
    }
    wdoc.Close(False)
    return rec


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    # Word needs a moment after first launch before COM calls succeed reliably.
    time.sleep(3.0)
    # Warm up COM with a no-op operation to ensure Word is ready
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass
    results = []
    try:
        print("=== continuous_basic (no column change, varying s1 paras + body size) ===")
        for r in case_continuous_basic(word):
            results.append(r)
            print(f"  n_s1={r['n_section1_paras']} sz={r['body_size']:4}pt: section_count={r['section_count']}")
            for p in r["paragraphs"]:
                print(f"    P{p['i']:2}@page{p['page']} y={p['y']:7.2f} x={p['x']:6.2f} '{p['text']}'")

        print("\n=== continuous_column_change (s1=N paras → s2 with 2/3 cols) ===")
        for r in case_continuous_column_change(word):
            results.append(r)
            sc = r["section2_props"]
            print(f"  n_s1={r['n_section1_paras']} cols={r['section2_columns']} "
                  f"col1_w={sc.get('col1_width')} ev={sc.get('TextColumns_evenlySpaced')}")
            for p in r["paragraphs"][:18]:
                print(f"    P{p['i']:3}@p{p['page']} y={p['y']:7.2f} x={p['x']:6.2f} '{p['text']}'")
            if len(r["paragraphs"]) > 18:
                print(f"    ... ({len(r['paragraphs']) - 18} more)")

        print("\n=== continuous_margin_change ===")
        r = case_continuous_margin_change(word)
        results.append(r)
        for p in r["paragraphs"]:
            print(f"    P{p['i']:2}@p{p['page']} y={p['y']:7.2f} x={p['x']:6.2f} '{p['text']}'")

    finally:
        word.Quit()

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_JSON}")


if __name__ == "__main__":
    main()
