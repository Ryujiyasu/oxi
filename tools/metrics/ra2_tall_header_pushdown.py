"""
Ra2: Tall-header pushdown formula investigation (§8.2 word_layout_spec_ra.md).

Spec §8.2 explicitly flags this as "still TBD":
  > When header content overflows headerDistance and crosses topMargin,
  > body Y is pushed down. Earlier note ("3-line 14pt header → body_y=90pt
  > when topMargin=72") was measured for noGrid; the formula has not been
  > re-verified under the corrected spec and remains a candidate for
  > follow-up Ra2 measurement.

Goal:
  Pin down the formula for body_first_line_y when the header content height
  exceeds (topMargin - headerDistance).

Hypothesis candidates:
  H1: body_y = max(topMargin, headerDistance + sum(header_line_heights))
  H2: body_y = max(topMargin, headerDistance + N_lines * header_first_line_height)
  H3: body_y = max(topMargin, headerDistance + header_block_height + min_gap)
       — where min_gap might be 0 / 1 line / a constant ~?pt
  H4: body_y always == topMargin (no pushdown — old note was a misread)

Sweep design:
  Phase A — line-count threshold sweep (Calibri 11pt, hdrDist=36, tm=72, noGrid):
    Lines 1..7. Find N where body_y > tm starts. Calibri 11pt line_h≈13.5/15.5pt.
  Phase B — font-size threshold sweep (3-line header, hdrDist=36, tm=72, noGrid):
    sizes 8, 10.5, 11, 14, 18, 24 pt. Each gives different line height.
  Phase C — headerDistance sweep (3-line Calibri 14pt header, tm=72, noGrid):
    hdrDist 6, 18, 36, 54, 72.
  Phase D — topMargin sweep (3-line Calibri 14pt header, hdrDist=36, noGrid):
    tm 36, 72, 108, 144.
  Phase E — grid on (3-line Calibri 14pt header, hdrDist=36, tm=72, LineGrid pitch=360tw=18pt):
    Does grid snap apply on top of the pushdown formula?
  Phase F — heterogeneous-line header (8/14/11pt mixed lines, hdrDist=36, tm=72, noGrid):
    Disambiguates H1 (sum of per-line) vs H2 (N * first_line_height).
  Phase G — CJK header font (3-line MS Gothic 14pt, hdrDist=36, tm=72, noGrid):
    Confirms whether CJK 83/64 line-height multiplier participates in pushdown.

For each record we capture:
  hdr_p1_y, hdr_pN_y, hdr_block_height (last_y - first_y + computed_last_line_h)
  body_y, delta_from_topmargin, delta_from_(hdrDist+block_h)
  per-paragraph-y for header

Caveat (per com_information6_caveat):
  Information(6) returns line-box top, NOT glyph top. For verifying our formula
  on body_y this is fine because both header and body are measured the same way.
"""
import win32com.client
import json
import os
import sys

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_tall_header_pushdown.json")

WD_LINE_SPACE_SINGLE = 0
WD_LAYOUT_DEFAULT = 0
WD_LAYOUT_LINEGRID = 2
WD_HEADER_FOOTER_PRIMARY = 1


def make_doc(word, *, top_margin, header_dist, hdr_lines, hdr_font, hdr_size,
             body_font="Calibri", body_size=11, layout_mode=0, line_pitch_tw=0,
             body_para_count=2, hdr_lines_mixed=None):
    """Create a Word doc and return measurement record.

    hdr_lines_mixed: optional list of (font, size) tuples — one per header line.
                    Overrides hdr_lines/hdr_font/hdr_size.
    """
    wdoc = word.Documents.Add()
    sec = wdoc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = top_margin
    ps.BottomMargin = 72
    ps.HeaderDistance = header_dist
    ps.FooterDistance = 36

    if layout_mode == WD_LAYOUT_DEFAULT:
        ps.LayoutMode = WD_LAYOUT_DEFAULT
    else:
        ps.LayoutMode = layout_mode
        if line_pitch_tw > 0:
            ps.LinesPage = int(round(
                (ps.PageHeight - ps.TopMargin - ps.BottomMargin) * 20 / line_pitch_tw
            ))

    # Header
    hdr = sec.Headers(WD_HEADER_FOOTER_PRIMARY)
    if hdr_lines_mixed:
        # Heterogeneous: build one paragraph per (font, size)
        hdr.Range.Text = "\r".join(f"H{i+1}" for i in range(len(hdr_lines_mixed)))
        for i, (f, s) in enumerate(hdr_lines_mixed):
            p = hdr.Range.Paragraphs(i + 1)
            p.Range.Font.Name = f
            try:
                p.Range.Font.NameFarEast = f
            except Exception:
                pass
            p.Range.Font.Size = s
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE
    else:
        hdr.Range.Text = "\r".join(f"H{i+1}" for i in range(hdr_lines))
        for i in range(1, hdr_lines + 1):
            p = hdr.Range.Paragraphs(i)
            p.Range.Font.Name = hdr_font
            try:
                p.Range.Font.NameFarEast = hdr_font
            except Exception:
                pass
            p.Range.Font.Size = hdr_size
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE

    # Body
    wdoc.Content.Text = "\r".join(f"B{i+1}" for i in range(body_para_count))
    for i in range(1, body_para_count + 1):
        p = wdoc.Paragraphs(i)
        p.Range.Font.Name = body_font
        try:
            p.Range.Font.NameFarEast = body_font
        except Exception:
            pass
        p.Range.Font.Size = body_size
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = WD_LINE_SPACE_SINGLE

    wdoc.Repaginate()

    # Measure
    hdr_paras = []
    n_hdr = hdr.Range.Paragraphs.Count
    for i in range(1, n_hdr + 1):
        p = hdr.Range.Paragraphs(i)
        hdr_paras.append({
            "i": i,
            "y": round(p.Range.Information(6), 4),
            "text": p.Range.Text.strip()[:20],
        })

    body_y = round(wdoc.Paragraphs(1).Range.Information(6), 4)
    body2_y = round(wdoc.Paragraphs(2).Range.Information(6), 4) if body_para_count >= 2 else None

    rec = {
        "top_margin": top_margin,
        "header_distance": header_dist,
        "hdr_font": hdr_font,
        "hdr_size": hdr_size,
        "hdr_lines": n_hdr,
        "hdr_lines_mixed": hdr_lines_mixed,
        "body_font": body_font,
        "body_size": body_size,
        "layout_mode": "default" if layout_mode == 0 else f"linegrid_pitch{line_pitch_tw}tw",
        "header_paragraphs": hdr_paras,
        "hdr_p1_y": hdr_paras[0]["y"],
        "hdr_pN_y": hdr_paras[-1]["y"],
        "hdr_block_height_observed": round(hdr_paras[-1]["y"] - hdr_paras[0]["y"], 4),
        "body_y": body_y,
        "body2_y": body2_y,
        "body_line_h_observed": round(body2_y - body_y, 4) if body2_y else None,
        "delta_body_topmargin": round(body_y - top_margin, 4),
    }

    wdoc.Close(False)
    return rec


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    results = []
    try:
        # Phase A — line-count threshold (Calibri 11pt)
        print("=== Phase A: line count sweep, Calibri 11pt, hdrDist=36, tm=72, noGrid ===")
        for n in range(1, 8):
            r = make_doc(word, top_margin=72, header_dist=36,
                         hdr_lines=n, hdr_font="Calibri", hdr_size=11)
            r["phase"] = "A"
            results.append(r)
            print(f"  n={n}: hdr p1_y={r['hdr_p1_y']:.2f} pN_y={r['hdr_pN_y']:.2f} "
                  f"body_y={r['body_y']:.2f} delta_tm={r['delta_body_topmargin']:+.2f}")

        # Phase B — header font size threshold (3-line)
        print("\n=== Phase B: hdr_size sweep, 3-line Calibri, hdrDist=36, tm=72, noGrid ===")
        for sz in [8, 10.5, 11, 14, 18, 24]:
            r = make_doc(word, top_margin=72, header_dist=36,
                         hdr_lines=3, hdr_font="Calibri", hdr_size=sz)
            r["phase"] = "B"
            results.append(r)
            print(f"  size={sz:5}: hdr p1_y={r['hdr_p1_y']:.2f} pN_y={r['hdr_pN_y']:.2f} "
                  f"body_y={r['body_y']:.2f} delta_tm={r['delta_body_topmargin']:+.2f}")

        # Phase C — headerDistance sweep (3-line Calibri 14pt, tm=72, noGrid)
        print("\n=== Phase C: headerDistance sweep, 3-line Calibri 14pt, tm=72, noGrid ===")
        for hd in [6, 18, 36, 54, 72, 90]:
            r = make_doc(word, top_margin=72, header_dist=hd,
                         hdr_lines=3, hdr_font="Calibri", hdr_size=14)
            r["phase"] = "C"
            results.append(r)
            print(f"  hd={hd:4}: hdr p1_y={r['hdr_p1_y']:.2f} pN_y={r['hdr_pN_y']:.2f} "
                  f"body_y={r['body_y']:.2f} delta_tm={r['delta_body_topmargin']:+.2f}")

        # Phase D — topMargin sweep (3-line Calibri 14pt, hdrDist=36, noGrid)
        print("\n=== Phase D: topMargin sweep, 3-line Calibri 14pt, hdrDist=36, noGrid ===")
        for tm in [36, 72, 108, 144]:
            r = make_doc(word, top_margin=tm, header_dist=36,
                         hdr_lines=3, hdr_font="Calibri", hdr_size=14)
            r["phase"] = "D"
            results.append(r)
            print(f"  tm={tm:4}: hdr p1_y={r['hdr_p1_y']:.2f} pN_y={r['hdr_pN_y']:.2f} "
                  f"body_y={r['body_y']:.2f} delta_tm={r['delta_body_topmargin']:+.2f}")

        # Phase E — grid on (3-line Calibri 14pt, hdrDist=36, tm=72, LineGrid pitch=360tw)
        print("\n=== Phase E: LineGrid pitch=360tw (18pt), 3-line Calibri 14pt, hdrDist=36, tm=72 ===")
        for pitch in [360, 480]:
            r = make_doc(word, top_margin=72, header_dist=36,
                         hdr_lines=3, hdr_font="Calibri", hdr_size=14,
                         layout_mode=WD_LAYOUT_LINEGRID, line_pitch_tw=pitch)
            r["phase"] = "E"
            results.append(r)
            print(f"  pitch={pitch}tw: hdr p1_y={r['hdr_p1_y']:.2f} pN_y={r['hdr_pN_y']:.2f} "
                  f"body_y={r['body_y']:.2f} delta_tm={r['delta_body_topmargin']:+.2f}")

        # Phase F — heterogeneous lines (mixed sizes)
        print("\n=== Phase F: heterogeneous-line header, hdrDist=36, tm=72, noGrid ===")
        scenarios = [
            ("8/8/8", [("Calibri", 8), ("Calibri", 8), ("Calibri", 8)]),
            ("14/14/14", [("Calibri", 14), ("Calibri", 14), ("Calibri", 14)]),
            ("8/14/8", [("Calibri", 8), ("Calibri", 14), ("Calibri", 8)]),
            ("14/8/14", [("Calibri", 14), ("Calibri", 8), ("Calibri", 14)]),
            ("11/24/11", [("Calibri", 11), ("Calibri", 24), ("Calibri", 11)]),
        ]
        for name, mix in scenarios:
            r = make_doc(word, top_margin=72, header_dist=36,
                         hdr_lines=len(mix), hdr_font="Calibri", hdr_size=11,
                         hdr_lines_mixed=mix)
            r["phase"] = "F"
            r["scenario"] = name
            results.append(r)
            ph_ys = "/".join(f"{p['y']:.1f}" for p in r["header_paragraphs"])
            print(f"  {name:12s}: hdr ys=[{ph_ys}] body_y={r['body_y']:.2f} "
                  f"delta_tm={r['delta_body_topmargin']:+.2f}")

        # Phase G — CJK header (MS Gothic 14pt, 3 lines, hdrDist=36, tm=72, noGrid)
        print("\n=== Phase G: CJK header sweep, 3-line, hdrDist=36, tm=72, noGrid ===")
        for fnt in ["MS Gothic", "MS Mincho", "Yu Mincho"]:
            for sz in [11, 14]:
                r = make_doc(word, top_margin=72, header_dist=36,
                             hdr_lines=3, hdr_font=fnt, hdr_size=sz)
                r["phase"] = "G"
                results.append(r)
                print(f"  {fnt:12s} {sz:4}pt: hdr p1_y={r['hdr_p1_y']:.2f} pN_y={r['hdr_pN_y']:.2f} "
                      f"body_y={r['body_y']:.2f} delta_tm={r['delta_body_topmargin']:+.2f}")

    finally:
        word.Quit()

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(results)} records to {OUT_JSON}")


if __name__ == "__main__":
    main()
