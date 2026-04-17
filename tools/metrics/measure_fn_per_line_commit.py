"""
Ra: per-line footnote-reference commitment — minimal repro.

Hypothesis (from b837 p4 root cause):
  Word commits footnote-area reservation per LINE, not per paragraph.
  When a paragraph straddles a page boundary with fn refs only on later lines,
  Word places early lines (without fn refs) on the current page using the
  pre-reservation effective content height; fn refs go with the lines that
  contain them.

Minimal repro design:
  1. Fill page with enough body paragraphs that there is just barely room
     for ONE extra text line before the bottom margin.
  2. Place a paragraph with 3-4 lines. Line 1 has no fn ref. Line 2 has a
     fn ref. If per-line commit is correct, Word puts line 1 on the current
     page; lines 2-N and the footnote go on the next page.

Variants (≥3 to satisfy 3-doc rule):
  - V1: fn ref in line 2 of 3-line para
  - V2: fn ref in line 3 of 3-line para (later)
  - V3: fn ref in line 1 of 3-line para (should push entire para — baseline)
  - V4: 2 fn refs split across line 2 and line 3

For each variant, record which lines of the target paragraph land on each
page, the fn's page, and the y position of each line.
"""
import win32com.client, json, os, sys

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

# Fill-line count chosen so that one additional body line fits on page 1
# with standard Calibri 11pt. Will auto-adjust by measuring.
FILL_LINES_PER_PAGE_DEFAULT = 44

# Long string that wraps to multiple lines when placed in body width.
# We measure its actual line count via LineSpacing iteration below.
LONG_CHUNK_A = (
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod "
    "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim "
    "veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex "
)
LONG_CHUNK_B = (
    "ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate "
    "velit esse cillum dolore eu fugiat nulla pariatur. "
)


def build_doc(variant_name, fill_lines, para_segments_with_fn, fill_skip=0):
    """
    para_segments_with_fn: list of (text, needs_fn, fn_text_or_None).
    Each segment appends to a single continuous paragraph.

    fill_skip: number of fill paragraphs to REMOVE to nudge the target para's
        first line closer to the boundary (tunes so splitting is possible).
    """
    wdoc = word.Documents.Add()
    try:
        sec = wdoc.Sections(1)
        ps = sec.PageSetup
        ps.LeftMargin = 72
        ps.RightMargin = 72
        ps.TopMargin = 72
        ps.BottomMargin = 72
        ps.PageHeight = 792  # Letter
        ps.PageWidth = 612

        # Clear default.
        wdoc.Content.Text = ""

        # Fill body with short lines so we control how many land on page 1.
        for i in range(fill_lines - fill_skip):
            if i > 0:
                rng = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
                rng.InsertParagraphAfter()
            p = wdoc.Paragraphs(wdoc.Paragraphs.Count)
            p.Range.Text = f"Fill line {i+1:03d}"
            p.Range.Font.Name = "Calibri"
            p.Range.Font.Size = 11
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.LineSpacingRule = 0  # wdLineSpaceSingle

        # Target paragraph.
        rng_end = wdoc.Range(wdoc.Content.End - 1, wdoc.Content.End - 1)
        rng_end.InsertParagraphAfter()
        target_para_idx = wdoc.Paragraphs.Count  # 1-based
        target = wdoc.Paragraphs(target_para_idx)
        target.Range.Text = ""  # clean start
        target.Range.Font.Name = "Calibri"
        target.Range.Font.Size = 11
        target.Format.SpaceBefore = 0
        target.Format.SpaceAfter = 0
        target.Format.LineSpacingRule = 0

        # Insert segments sequentially. After each text chunk, if needs_fn,
        # attach a footnote at the final character of that chunk.
        seg_fn_records = []
        for seg_i, (text, needs_fn, fn_text) in enumerate(para_segments_with_fn):
            insert_pos = target.Range.End - 1
            ins = wdoc.Range(insert_pos, insert_pos)
            ins.Text = text
            if needs_fn:
                # Attach fn at the last char of this chunk.
                anchor_end = insert_pos + len(text)
                fn_rng = wdoc.Range(anchor_end - 1, anchor_end)
                fn = wdoc.Footnotes.Add(fn_rng, Text=fn_text or f"Footnote text for seg {seg_i+1}.")
                seg_fn_records.append((seg_i, fn))

        wdoc.Repaginate()

        # Measure target paragraph line-by-line by walking its range in 1-char
        # steps and grouping by (page, y_line_approx). Word does not expose a
        # "line iterator" directly via COM; the stable workaround is
        # Information(wdFirstCharacterVerticalPosition=6) per character, then
        # cluster.
        target_after = wdoc.Paragraphs(target_para_idx)
        rng_full = target_after.Range
        start = rng_full.Start
        end = rng_full.End
        per_char = []
        i = start
        while i < end:
            r = wdoc.Range(i, i + 1)
            try:
                y = round(r.Information(6), 4)
                x = round(r.Information(5), 4)
                page = r.Information(3)
            except Exception:
                y, x, page = None, None, None
            ch = r.Text
            per_char.append({"i": i - start, "ch": ch, "y": y, "x": x, "page": page})
            i += 1

        # Cluster by (page, round(y,1)): same line = same page + y within 1.0pt.
        lines = []
        cur = None
        for rec in per_char:
            if rec["y"] is None:
                continue
            key = (rec["page"], round(rec["y"], 1))
            if cur is None or cur["key"] != key:
                cur = {"key": key, "page": rec["page"], "y": rec["y"],
                       "x_min": rec["x"], "x_max": rec["x"],
                       "chars": [rec["ch"]]}
                lines.append(cur)
            else:
                cur["x_min"] = min(cur["x_min"], rec["x"])
                cur["x_max"] = max(cur["x_max"], rec["x"])
                cur["chars"].append(rec["ch"])
        for ln in lines:
            ln["text"] = "".join(ln["chars"]).replace("\r", "").replace("\x07", "")

        # Footnote placements.
        fn_positions = []
        for seg_i, fn in seg_fn_records:
            fr = fn.Range
            fn_positions.append({
                "seg": seg_i,
                "page": fr.Information(3),
                "y_pt": round(fr.Information(6), 4),
                "x_pt": round(fr.Information(5), 4),
                "text": fr.Text.strip()[:40],
            })

        # Last fill paragraph position (end of page 1 body) for context.
        last_fill_y = None
        last_fill_page = None
        # A few fill paragraphs near target.
        fill_ctx = []
        for i in range(max(1, target_para_idx - 3), target_para_idx):
            p = wdoc.Paragraphs(i)
            r = p.Range
            fill_ctx.append({
                "idx": i,
                "y_pt": round(r.Information(6), 4),
                "page": r.Information(3),
                "text": r.Text.strip()[:20],
            })
            last_fill_y = round(r.Information(6), 4)
            last_fill_page = r.Information(3)

        return {
            "variant": variant_name,
            "fill_lines_used": fill_lines - fill_skip,
            "target_lines": [
                {"page": ln["page"], "y": ln["y"], "x_min": ln["x_min"],
                 "text": ln["text"][:80]}
                for ln in lines
            ],
            "footnotes": fn_positions,
            "fill_context_last_few": fill_ctx,
            "last_fill_before_target": {"y": last_fill_y, "page": last_fill_page},
            "page_setup": {
                "page_height": round(ps.PageHeight, 2),
                "top_margin": round(ps.TopMargin, 2),
                "bottom_margin": round(ps.BottomMargin, 2),
            },
        }
    finally:
        wdoc.Close(False)


results = []
try:
    # We want the target para to span page boundary: line 1 on p1, line 2+ on p2.
    # We tune fill_lines so that with no fn refs, ALL target lines fit on page 1.
    # Then with fn refs on later lines, check whether Word keeps line 1 on p1
    # OR pushes entire para to p2.

    # Run a calibration first with no fn refs and longish para.
    target_text_A = LONG_CHUNK_A + LONG_CHUNK_B  # wraps to ~4 lines
    # Sweep fill_skip from 0..12. Record first skip that causes straddle
    # (target lines span both p1 and p2), and the one right before where all on p1.
    sweep = []
    for fs in range(0, 12):
        test = build_doc(
            f"calib_skip_{fs}",
            fill_lines=FILL_LINES_PER_PAGE_DEFAULT,
            para_segments_with_fn=[(target_text_A, False, None)],
            fill_skip=fs,
        )
        pages = sorted({ln["page"] for ln in test["target_lines"]})
        sweep.append((fs, pages, test))
        print(f"  skip={fs}: target pages={pages}")
        if len(pages) > 1:
            break
    calib = sweep[-1][2]
    tuned_skip = sweep[-1][0]
    if len({ln["page"] for ln in calib["target_lines"]}) == 1:
        print("  WARN: no straddle found in sweep; using last")
    results.append(calib)
    print(f"  tuned fill_skip = {tuned_skip}")
    print(f"  calibration target lines:")
    for ln in calib["target_lines"]:
        print(f"    p{ln['page']}  y={ln['y']}  [{ln['text'][:60]}]")

    # Tuned fill_skip found. Now run variants.
    # V1: fn ref in line 2 of 3-line para.
    # Structure the target para so segment boundaries align with expected line breaks.
    # Line 1 should be ~LONG_CHUNK_A words; line 2 onward starts LONG_CHUNK_B.
    # Attach fn in LONG_CHUNK_B so ref lives on line 2+.
    v1 = build_doc(
        "V1_fn_in_later_segment",
        fill_lines=FILL_LINES_PER_PAGE_DEFAULT,
        para_segments_with_fn=[
            (LONG_CHUNK_A, False, None),
            (LONG_CHUNK_B, True, "Footnote anchored mid-para after wrap."),
        ],
        fill_skip=tuned_skip,
    )
    results.append(v1)

    # V3 (baseline): fn ref in FIRST segment (line 1).
    v3 = build_doc(
        "V3_fn_in_first_segment",
        fill_lines=FILL_LINES_PER_PAGE_DEFAULT,
        para_segments_with_fn=[
            (LONG_CHUNK_A, True, "Footnote anchored on line 1."),
            (LONG_CHUNK_B, False, None),
        ],
        fill_skip=tuned_skip,
    )
    results.append(v3)

    # V4: two fn refs, both on later segment.
    v4 = build_doc(
        "V4_two_fn_later",
        fill_lines=FILL_LINES_PER_PAGE_DEFAULT,
        para_segments_with_fn=[
            (LONG_CHUNK_A, False, None),
            (LONG_CHUNK_B, True, "Footnote A later."),
            (LONG_CHUNK_B, True, "Footnote B later."),
        ],
        fill_skip=tuned_skip,
    )
    results.append(v4)

    # Print each variant.
    for v in [v1, v3, v4]:
        print(f"\n=== {v['variant']}  (tuned_skip={tuned_skip}) ===")
        print(f"  last fill before target: {v['last_fill_before_target']}")
        print(f"  target para lines ({len(v['target_lines'])}):")
        for ln in v["target_lines"]:
            print(f"    p{ln['page']}  y={ln['y']}  x_min={ln['x_min']}  [{ln['text'][:70]}]")
        print(f"  footnotes:")
        for fn in v["footnotes"]:
            print(f"    seg={fn['seg']}  p{fn['page']}  y={fn['y_pt']}  [{fn['text']}]")

    # Final analysis.
    print("\n\n========== ANALYSIS ==========")
    def line_pages(v):
        return [ln["page"] for ln in v["target_lines"]]
    def fn_pages(v):
        return [f["page"] for f in v["footnotes"]]

    print(f"V1 (fn on later segment): target line pages = {line_pages(v1)}  fn pages = {fn_pages(v1)}")
    print(f"V3 (fn on first  segment): target line pages = {line_pages(v3)}  fn pages = {fn_pages(v3)}")
    print(f"V4 (two fn later segments): target line pages = {line_pages(v4)}  fn pages = {fn_pages(v4)}")

    v1_split = len(set(line_pages(v1))) > 1
    v3_all_same = len(set(line_pages(v3))) == 1
    print(f"\nHypothesis:")
    print(f"  - V1 target lines should straddle pages (1,2,...) → observed split: {v1_split}")
    print(f"  - V3 target should all be on same page (1 or 2 only) → observed: {v3_all_same}")
    print(f"  - V1 fn should land on the same page as its ref line (p2), not pre-reserving p1")
    if v1["footnotes"]:
        fn1_page = v1["footnotes"][0]["page"]
        line_of_fn = None
        for i, ln in enumerate(v1["target_lines"]):
            if LONG_CHUNK_B[:20] in ln["text"]:
                line_of_fn = i + 1
                break
        print(f"    V1 fn on page {fn1_page}; ref expected on line {line_of_fn}")

finally:
    word.Quit()

out = os.path.join(os.path.dirname(__file__), 'output', 'measure_fn_per_line_commit.json')
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out}")
