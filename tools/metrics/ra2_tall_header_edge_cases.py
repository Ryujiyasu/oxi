"""
Ra2: Tall-header edge cases (complement to ra2_tall_header_pushdown.py).

Investigates:
  E1. Empty header (no paragraphs / single empty paragraph) — does body_y == tm?
  E2. Footer mirror behavior: tall footer pushes body_y up from below?
  E3. First-page-only header (DifferentFirstPageHeaderFooter): does pushdown
      apply only to page 1? Does odd/even apply correctly?
  E4. Header with explicit lineSpacingRule=exact — does block_h use exact
      line height instead of natural?
"""
import win32com.client
import json
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_tall_header_edge_cases.json")

WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_EXACT = 4
WD_LAYOUT_DEFAULT = 0
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST = 2


def case_E1_empty_header(word):
    """Empty header — body_y should equal tm exactly."""
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    ps.LayoutMode = WD_LAYOUT_DEFAULT

    # Don't add header text — keep default empty header
    wdoc.Content.Text = "B1\rB2"
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    wdoc.Repaginate()
    body_y = round(wdoc.Paragraphs(1).Range.Information(6), 4)
    hdr = sec.Headers(WD_HEADER_FOOTER_PRIMARY)
    hdr_y = round(hdr.Range.Information(6), 4) if hdr.Range.Text.strip() else None
    rec = {
        "case": "E1_empty_header",
        "top_margin": 72, "header_dist": 36,
        "body_y": body_y,
        "header_text_present": bool(hdr.Range.Text.strip()),
        "header_y": hdr_y,
        "predicted": "tm=72 (no overflow)",
        "match": body_y == 72.0,
    }
    wdoc.Close(False)
    return rec


def case_E2_tall_footer(word):
    """Tall footer — does it push the body bottom up from bottomMargin?"""
    results = []
    for n_lines in [1, 2, 3, 4, 5]:
        wdoc = word.Documents.Add()
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.HeaderDistance = 36
        ps.FooterDistance = 36
        ps.LayoutMode = WD_LAYOUT_DEFAULT

        ftr = sec.Footers(WD_HEADER_FOOTER_PRIMARY)
        ftr.Range.Text = "\r".join(f"F{i+1}" for i in range(n_lines))
        for i in range(1, n_lines + 1):
            p = ftr.Range.Paragraphs(i)
            p.Range.Font.Name = "Calibri"
            p.Range.Font.Size = 14  # tall enough to overflow
            p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE

        # Body: many paragraphs to fill page so we hit bottom
        wdoc.Content.Text = "\r".join(f"B{i+1}" for i in range(60))
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            p.Range.Font.Name = "Calibri"
            p.Range.Font.Size = 11
            p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        wdoc.Repaginate()

        # Find last body paragraph on page 1
        last_body_y_p1 = None
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            page = p.Range.Information(3)  # wdActiveEndPageNumber
            y = round(p.Range.Information(6), 4)
            if page == 1:
                last_body_y_p1 = y
            else:
                break

        ftr_paras = []
        for i in range(1, ftr.Range.Paragraphs.Count + 1):
            fp = ftr.Range.Paragraphs(i)
            ftr_paras.append({
                "i": i, "y": round(fp.Range.Information(6), 4),
                "text": fp.Range.Text.strip()[:20],
            })

        rec = {
            "case": "E2_tall_footer",
            "n_lines": n_lines,
            "footer_paragraphs": ftr_paras,
            "ftr_p1_y": ftr_paras[0]["y"],
            "ftr_pN_y": ftr_paras[-1]["y"],
            "page_height": round(ps.PageHeight, 4),
            "bottom_margin": round(ps.BottomMargin, 4),
            "last_body_y_page1": last_body_y_p1,
        }
        results.append(rec)
        wdoc.Close(False)
    return results


def case_E3_first_page_header(word):
    """Different first-page header — primary header tall, first-page header short.
    Verify pushdown only applies on pages where the active header is tall."""
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.HeaderDistance = 36
    ps.FooterDistance = 36
    ps.LayoutMode = WD_LAYOUT_DEFAULT

    # Body first (to avoid COM state interaction with header switching)
    wdoc.Content.Text = "\r".join(f"B{i+1}" for i in range(80))
    ps.DifferentFirstPageHeaderFooter = -1  # True

    # First-page header: 1 short line
    fhdr = sec.Headers(WD_HEADER_FOOTER_FIRST)
    fhdr.Range.Text = "FirstPgHdr"
    fhdr.Range.Font.Name = "Calibri"
    fhdr.Range.Font.Size = 11

    # Primary header: 3 tall lines
    phdr = sec.Headers(WD_HEADER_FOOTER_PRIMARY)
    phdr.Range.Text = "Pri H1\rPri H2\rPri H3"
    for i in range(1, 4):
        p = phdr.Range.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 14
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = "Calibri"
        p.Range.Font.Size = 11
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    wdoc.Repaginate()

    # First body paragraph on page 1, page 2
    first_p1 = None
    first_p2 = None
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        pg = p.Range.Information(3)
        y = round(p.Range.Information(6), 4)
        if pg == 1 and first_p1 is None:
            first_p1 = y
        elif pg == 2 and first_p2 is None:
            first_p2 = y
            break

    rec = {
        "case": "E3_first_page_different_header",
        "first_pg_hdr_text": "FirstPgHdr (1 line, 11pt)",
        "primary_hdr_text": "3 lines, Calibri 14pt",
        "body_y_page1": first_p1,
        "body_y_page2": first_p2,
        "predicted_p1_no_overflow": "= 72.0 (short first-page header)",
        "predicted_p2_overflow": "= 87.5 (tall primary header, same as Phase C hd=36)",
    }
    wdoc.Close(False)
    return rec


def case_E4_exact_line_spacing(word):
    """Header with lineSpacingRule=exact — block_h should use exact value."""
    results = []
    for exact_pt in [12, 18, 24, 36]:
        wdoc = word.Documents.Add()
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.HeaderDistance = 36
        ps.FooterDistance = 36
        ps.LayoutMode = WD_LAYOUT_DEFAULT

        hdr = sec.Headers(WD_HEADER_FOOTER_PRIMARY)
        hdr.Range.Text = "H1\rH2\rH3"
        for i in range(1, 4):
            p = hdr.Range.Paragraphs(i)
            p.Range.Font.Name = "Calibri"
            p.Range.Font.Size = 11
            p.Format.LineSpacingRule = WD_LINE_SPACE_EXACT
            p.Format.LineSpacing = exact_pt

        wdoc.Content.Text = "B1\rB2"
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            p.Range.Font.Name = "Calibri"
            p.Range.Font.Size = 11
            p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
        wdoc.Repaginate()

        hdr_paras = []
        for i in range(1, hdr.Range.Paragraphs.Count + 1):
            p = hdr.Range.Paragraphs(i)
            hdr_paras.append({"i": i, "y": round(p.Range.Information(6), 4)})
        body_y = round(wdoc.Paragraphs(1).Range.Information(6), 4)
        block_obs = body_y - 36
        block_pred = 3 * exact_pt
        rec = {
            "case": "E4_exact_line_spacing",
            "exact_pt": exact_pt,
            "header_paragraphs": hdr_paras,
            "body_y": body_y,
            "block_observed": block_obs,
            "block_predicted_3xExact": block_pred,
            "match": abs(block_obs - block_pred) < 0.6 if body_y > 72.0 else None,
        }
        results.append(rec)
        wdoc.Close(False)
    return results


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        print("=== E1: empty header ===")
        r = case_E1_empty_header(word)
        results.append(r)
        print(f"  body_y={r['body_y']} (predicted tm=72) match={r['match']}")

        print("\n=== E2: tall footer mirror ===")
        for r in case_E2_tall_footer(word):
            results.append(r)
            ph = r["page_height"]
            print(f"  n={r['n_lines']} ftr_p1_y={r['ftr_p1_y']:.2f} pN_y={r['ftr_pN_y']:.2f} "
                  f"last_body_p1_y={r['last_body_y_page1']} "
                  f"(page_h={ph} bm={r['bottom_margin']})")

        print("\n=== E3: different first-page header ===")
        r = case_E3_first_page_header(word)
        results.append(r)
        print(f"  page1 first body_y={r['body_y_page1']} (short hdr → predicted 72.0)")
        print(f"  page2 first body_y={r['body_y_page2']} (tall hdr → predicted 87.5)")

        print("\n=== E4: header lineSpacingRule=exact ===")
        for r in case_E4_exact_line_spacing(word):
            results.append(r)
            ys = [p["y"] for p in r["header_paragraphs"]]
            print(f"  exact={r['exact_pt']:3}pt: hdr_ys={ys} body_y={r['body_y']} "
                  f"block_obs={r['block_observed']} pred=3x{r['exact_pt']}={r['block_predicted_3xExact']} "
                  f"match={r['match']}")

    finally:
        word.Quit()

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_JSON}")


if __name__ == "__main__":
    main()
