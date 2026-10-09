// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `LayoutEngine::layout_paragraph` -- moved out of `layout/mod.rs` so that it is its own
//! codegen unit (see tools/metrics/split_layout_mod.py). Behaviour-preserving.

use super::*;

/// `layout_paragraph` as a method of its own type: rustc puts a method's code in the
/// codegen unit of its self type's module, so this (not the file move alone)
/// is what gives the giant its own unit. Deref keeps `self.x` meaning the engine.
pub(super) struct ParagraphLayouter<'a>(pub(super) &'a LayoutEngine);

impl<'a> std::ops::Deref for ParagraphLayouter<'a> {
    type Target = LayoutEngine;
    fn deref(&self) -> &LayoutEngine {
        self.0
    }
}

impl<'a> ParagraphLayouter<'a> {
    pub(super) fn layout_paragraph(
        &self,
        para: &Paragraph,
        mut start_x: f32,
        cursor: &mut LayoutCursor,
        content_width: f32,
        mut content_height: f32,
        mut page_top: f32,
        page: &Page,
        pages: &mut Vec<LayoutPage>,
        current_elements: &mut Vec<LayoutElement>,
        grid_pitch: Option<f32>,
        prev_style_id: Option<&str>,
        prev_contextual_spacing: bool,
        // S931 (2026-07-18): the previous paragraph's numId IF it carried
        // afterAutospacing — the same-list HTML-autospacing suppression gate.
        // Some(numId) only when (prev.after_autospacing && prev.num_id);
        // only the body call site threads a real value.
        prev_autospacing_numid: Option<&str>,
        // S658 (2026-06-24): the immediately-previous paragraph's pBdr (border-merge
        // gate). Word merges consecutive paragraphs with an IDENTICAL pBdr into one
        // box — the top border (and its reserved gap) is drawn only above the FIRST
        // paragraph of the group. Used to skip the top-border vertical reservation
        // when this paragraph continues a merged group. Only the body call site
        // threads a real value; header/footer/footnote/textbox pass None (no merge
        // detection, matching the existing prev_style_id=None simplification).
        prev_para_borders: Option<&ParagraphBorders>,
        // S739 (2026-07-04): the immediately-previous body paragraph has
        // keepNext. Word gives the FOLLOWER of a keepNext paragraph the
        // LENIENT natural page-bottom test for its first line (so the
        // keep-with-next pair can place heading + >=1 follower line at the
        // page bottom); a non-follower paragraph's first line uses the
        // stricter centered-box rule. Only the body call site threads a real
        // value; header/footer/footnote/textbox pass false.
        prev_keep_next: bool,
        #[allow(unused)] in_textbox: bool,
        prev_space_after: f32,
        body_para_index: Option<usize>,
        mut lm2_grid_cells: Option<&mut usize>,
        mut mult_cumul_raw: Option<&mut f32>,
        adjacent_to_empty_run: bool,
        // S603 (2026-06-18): the next sibling block is a TABLE. Used by the
        // typed-grid page-bottom full-cell rule: Word uses the FULL grid cell
        // (no natural_lh leading-hang leniency) for a paragraph's LAST line when
        // the next block is a table, but keeps the leniency when followed by body
        // text. Derived: 6/6 typed-grid leniency-regressing canary lines (db9ca,
        // kojin×4, roudoujoken×2, 34140) are followed by BODY; 3a4f para278 line4
        // (the cap+ぶら下げ compensating-error junction) is followed by a table.
        next_block_is_table: bool,
        // Step 0 (Option B fn reserve fix): per-page bucket of unique fn_ref
        // ids attributed to the lines rendered on each page offset.
        // line_fn_refs_out[0] = ids on current page; [i>0] = ids moved to
        // i-th new page pushed during this paragraph. Does NOT alter
        // behavior; exists so caller can redistribute fn reserves correctly.
        mut line_fn_refs_out: Option<&mut Vec<Vec<u32>>>,
        // R7.53 (2026-05-13): extra content_h available ONLY for the first
        // line's page-break check. Caller passes the paragraph's own footnote
        // reserve delta so the first-line check ignores fns that will move
        // with the paragraph to the next page if line 0 doesn't fit. After
        // line 0 is placed, subsequent lines use the strict content_height
        // (including this para's fns).
        first_line_extra_content_h: f32,
        // S168 (2026-05-22) Phase B-2: per-fn heights for per-line lenient.
        para_fn_heights: &std::collections::HashMap<u32, f32>,
        // S637 (2026-06-21): multi-column column-flow. When a line overflows the
        // current column and a NEXT column is available on the same page, flow
        // into it (shift start_x to that column's x, reset cursor to the column
        // top) instead of pushing a new page. `num_columns` is >1 ONLY on the
        // heterogeneous multi-column path (kyotei36spec's continuous 2-col
        // 記載心得 sections); it is 1 everywhere else, so the new branch never
        // fires and behavior is byte-identical for the whole 1-col corpus. The
        // final column the paragraph ended in is returned (3rd tuple element) so
        // the caller can keep its column state in sync. col_x_positions holds the
        // per-column left-x and widths; a paragraph can enter a differently sized column.
        num_columns: usize,
        start_column: usize,
        col_x_positions: &[f32],
        col_widths: &[f32],
        // S749 (2026-07-05): the vertical top of the current multi-column BAND
        // on the paragraph's first page. A continuous multi-col section that
        // starts MID-PAGE flows its columns from the section boundary — a
        // column advance returns the cursor HERE, not to page_top (col2 was
        // gaining the whole upper page: probexcont2col -1x12). After an
        // internal page push the band continues at page_top (the local
        // band_top resets). Non-band callers pass page_top (byte-identical).
        col_band_top: f32,
        // S691 (2026-06-29): this paragraph is laid out in a HEADER or FOOTER.
        // Scopes the large-CJK 83/64 baseline placement (the header/footer
        // snapToGrid=0 context Word places at 83/64 but Oxi's renderer drew at
        // raw); the body/textbox/cell large-CJK placement is already calibrated
        // (S455/S457/S662) and must NOT get the extra shift.
        is_header_footer: bool,
        // S726 (2026-07-03): the BODY's page bottom is FOOTER-CONSTRAINED —
        // footer_reserved (footer_distance + footer content height) exceeds the
        // bottom margin, so the space below the body bottom is occupied by
        // FOOTER TEXT, not empty margin. The page-bottom ink/natural leniency
        // (Day-33/S576 family) lets a line's line-spacing LEADING overhang the
        // bottom — legitimate over an empty margin, but Word does NOT let it
        // overhang into a footer area (probeftrtall: Word breaks the 703.4-COM
        // line whose full box crosses the footer-shrunk bottom 716.6 by 2.4pt
        // while ink fits by 2.0 — full-box at footer boundaries). false for
        // header/footer/footnote/textbox contexts (their bottoms differ).
        footer_tight: bool,
        // S755 (2026-07-06): per-page header/footer geometry (titlePg /
        // evenAndOddHeaders). Some ONLY when the first/even/odd variants
        // actually DIFFER in height (the whole corpus is None). On an
        // internal page push, page_top/content_height switch to the NEW
        // page's variant. Only the body call site threads a real value.
        s755_geom: Option<&S755Geom>,
        // S758 (2026-07-06): wrapSquare side-wrap band the paragraph STARTS
        // inside: (band_bottom_abs, width_reduction_pt). The paragraph breaks
        // at (wrap_width − reduction); when the cursor exits the band the
        // placement loop REBREAKS the remaining lines' fragments at the full
        // width (the imgfloat truth: the band cuts MID-paragraph — para 6 has
        // 2 narrow + 2 full lines). None everywhere except the body call site;
        // v1 = floating IMAGES only (all corpus wrapSquare anchors are
        // textboxes → corpus-inert by construction).
        s758_band: Option<(f32, f32, f32)>,
        // S-TWOSEG: (left strip x, left width, right strip x, right width)
        // when the float leaves usable room on both sides.
        s758_two_seg: Option<(f32, f32, f32, f32)>,
        // S835 (2026-07-14): the page's CURRENT footnote reserve is non-zero —
        // the content bottom this paragraph tests against is the FOOTNOTE-AREA
        // top (a SOFT boundary: Word grants fs/16 of line-box relief there;
        // see the s835_fn_relief derivation at the natural break test). Only
        // the body call site threads a real value; header/footer/footnote/
        // textbox/frame callers pass false (their bottoms are not fn areas).
        fn_boundary: FootnoteBoundary,
        // S900 (2026-07-17): the page's committed fn reserve from EARLIER
        // paragraphs (footnote_reserve_current), in pt. Needed to compute the
        // ABSOLUTE margin bottom (= effective bottom + reserve + this para's
        // delta) for the note-deferral test. Only the body call site threads
        // a real value.
        fn_reserve_above: f32,
        // S900: OUT — note ids this paragraph DEFERS to the next page's
        // footnote area (the anchor line stays; Word rolls notes that cannot
        // START in the page's remaining area forward: 81e80 p2 = notes 2..15
        // fill to the margin line, L9 stays with 15 placed and 16/17/18
        // render at the TOP of Word p3's area). Only the body call site
        // passes Some.
        mut fn_deferred_out: Option<&mut Vec<u32>>,
        // S903 (2026-07-17): the NEXT paragraph's pBdr — the bottom-border
        // advance (space + bw/2) applies only to the LAST paragraph of a
        // merged identical-pBdr group (Word's interior boundaries reserve
        // NOTHING: 0008ea8f form boxes render pure 2×line pitch 20.64 while
        // Oxi added +0.25/para). The S658 top-side merge gate's bottom
        // sibling. Only the body call site threads a real value.
        s903_next_borders: Option<&ParagraphBorders>,
        // S916 (2026-07-18): SPLIT a multi-line keepNext paragraph at
        // lines.len()-2 (keep n-2 head lines here, move a 2-line tail + the
        // follower to the next page) instead of the block-level whole-move.
        // Set true by the keepNext lookahead's Case-B (para fits, push driven
        // by the follower) when the para estimate is >=4 lines; layout re-gates
        // on the REAL lines.len() >= 4 (estimate <= real, so estimate>=4 =>
        // real>=4 — the double-gate cannot over-fire). Only the body call site
        // threads a real value; all other callers pass false.
        s916_tail_split: bool,
        body_wrap_bands: Option<(&[(usize, f32, f32, f32, f32, bool, BodyWrapPolicy)], usize)>,
    ) -> (Vec<LayoutElement>, f32, usize) {
        let fn_boundary_active = fn_boundary.active;

        // S1497: the band of a paragraph-relative wrapTopAndBottom float hosted
        // here starts at the block's entry cursor + posOffset (the S734
        // reservation point), whatever spacing is applied below.
        let s1497_entry_y = cursor.cursor_y;
        let s1497_band: Option<(f32, f32)> = if body_para_index.is_some() {
            S1497_BAND.with(|c| c.get())
                .map(|(off, h)| (cursor.cursor_y + off, cursor.cursor_y + off + h))
        } else { None };
        // S673v (2026-06-26): an EMPTY paragraph whose ¶ MARK is hidden
        // (`<w:pPr><w:rPr><w:vanish/></w:rPr>`) COLLAPSES to 0 height — Word does
        // not display/print the hidden mark, so the para contributes nothing (no
        // line, no spacing). The corpus idiom = an invisible separator paragraph
        // before a `<w:tbl>` (3a4f/model/tokyoshugyo each carry exactly one). Word
        // render-truth (single COM render): hidden-¶ TOP→BOTTOM = no-para 14.4
        // (collapsed); Oxi reserved ~25.9 (a partial line). Skip the para entirely
        // (no cursor advance, pass prev_space_after through). Opt-out OXI_S673V_DISABLE.
        if std::env::var("OXI_S673V_DISABLE").is_err()
            && para.runs.iter().all(|r| r.text.is_empty())
            && para.style.ppr_rpr.as_ref().map_or(false, |r| r.vanish)
        {
            return (Vec::new(), prev_space_after, start_column);
        }
        // S749: page-push detection for the column band top — once this
        // paragraph pushes a page, the multi-col band continues at page_top.
        let s749_pages_at_entry = pages.len();
        // Re-evaluate full-column float exclusion after an internal column
        // transition. The outer paragraph prepass only saw the previous column.
        let column_flow_top = |top: f32, x: f32, page_count: usize| {
            let Some((bands, entry_page)) = body_wrap_bands else { return top; };
            let (_, _, advance) = self.body_paragraph_wrap_bands(
                para, page, bands, entry_page + page_count - s749_pages_at_entry,
                top, x, content_width,
            );
            top + advance
        };
        // S730 (2026-07-03): an EMPTY paragraph that carries a CONTINUOUS
        // section-break mark renders at ZERO height in Word (probexmargins
        // COM: the break para's y equals the previous para's last-line row;
        // Oxi's normal empty-para line drifted everything below +18pt).
        // Same skip shape as S673v. Opt-out OXI_S730_DISABLE.
        if std::env::var("OXI_S730_DISABLE").is_err()
            && !S1501_KEEP.with(|c| c.get())
            && para.style.continuous_section_break
            && para.runs.iter().all(|r| r.text.is_empty())
        {
            // S1073: as in the S945 arm below — the box is skipped but the
            // paragraph's own space-after is what the next paragraph (or the
            // next section's first paragraph) collapses against.
            if std::env::var("OXI_DBG1073").is_ok() {
                eprintln!("[S1073-SKIP730] sa={:?} prev={:.3}", para.style.space_after, prev_space_after);
            }
            let sa_out = if !self.doc_body_has_real_cjk
                && std::env::var("OXI_S1073_DISABLE").is_err()
                && std::env::var("OXI_S816_DISABLE").is_err()
            {
                para.style.space_after.unwrap_or(prev_space_after)
            } else {
                prev_space_after
            };
            return (Vec::new(), sa_out, start_column);
        }
        // S945 (2026-07-19, opt-out OXI_S945_DISABLE): an EMPTY paragraph that
        // ENDS a section (in-body sectPr, any type) renders at zero height —
        // the section page break follows immediately, so Word never gives it a
        // page of its own (NDIS technical__0043bfe0 wp41/42: a 16pt empty
        // section-final para overflowed p41's bottom and manufactured a
        // phantom page before the nextPage section). Continuous breaks are
        // already handled by S730 above; this is the non-continuous sibling.
        if std::env::var("OXI_S945_DISABLE").is_err()
            && !S1501_KEEP.with(|c| c.get())
            && !(para.style.page_break_after
                && (std::env::var("OXI_SECTION_EXPLICIT_BREAKS").is_ok()
                    || (para.style.continuous_section_break
                        && std::env::var("OXI_S1454_DISABLE").is_err())))
            && para.style.page_section_break
            && para.runs.iter().all(|r| r.text.is_empty())
        {
            // S1073: the box is skipped, but the paragraph's own space-after is
            // still what the NEXT section's first paragraph collapses against.
            // uk_local_spending's Annex II boundary is exactly this shape — an
            // empty Heading2 carrying the sectPr with after=120tw — and Word's
            // excess there is 12 - 6 = 6pt, which only works if the skipped
            // paragraph's 6pt is carried.
            if std::env::var("OXI_DBG1073").is_ok() {
                eprintln!("[S1073-SKIP] sa={:?} prev={:.3}", para.style.space_after, prev_space_after);
            }
            let sa_out = if !self.doc_body_has_real_cjk
                && std::env::var("OXI_S1073_DISABLE").is_err()
                && std::env::var("OXI_S816_DISABLE").is_err()
            {
                para.style.space_after.unwrap_or(prev_space_after)
            } else {
                prev_space_after
            };
            return (Vec::new(), sa_out, start_column);
        }
        if let Some(v) = line_fn_refs_out.as_deref_mut() {
            if v.is_empty() {
                v.push(Vec::new());
            }
        }
        // S637: current column within a multi-column section (see signature doc).
        let mut cur_col = start_column;
        let mut elements = Vec::new();
        // S467 (2026-05-31, env-gated OFF, OXI_S467_VSNAP): match Word's vertical
        // layout model on the VISUAL track — advance visual_y by the EXACT (un-rounded)
        // raw line height and snap each emitted line's top to the 0.75pt (96-DPI pixel)
        // grid. cursor_y (page-break) keeps the current rounded model → Phase-1 safe by
        // construction (LayoutCursor decoupling, mod.rs:1439). COM (S467, 5 repros):
        // Word snaps line tops to the absolute 0.75pt grid using the exact cumulative
        // (line+spacing) position; Oxi's 10tw line-round + exact-spacing model is the
        // wrong granularity/phase, causing the gen2 list-boundary drift.
        let s467_vsnap = std::env::var("OXI_S467_VSNAP").is_ok();
        // S618 (2026-06-19) — MULTI-LEVER joint search (the "壁を多レバー同時で崩す" task):
        // tried combining the body cumulative-snap (VSNAP) with a TITLE_EXACT lever
        // (advance the single-spacing title's visual_y by its 83/64 raw via advance_split in
        // the use_cumulative=false else branch). RESULT: gen2 A/B = 0 bytes changed — the
        // lever had NO effect even though the title line h=33.0 (floor) < natural 33.6 should
        // satisfy the gate. So the title's SUBSEQUENT-line phase is NOT advanced at that
        // else branch — the title para (a single line followed by a pBdr `shading` element
        // at y=109, the S467 "title pBdr −0.75" note, mod.rs:5539) routes its vertical
        // advance through the pBdr/border + after-spacing sub-system, NOT the plain
        // cursor.advance(line_height). ⇒ the title component of the −1.26 deficit lives in
        // that pBdr/spacing path, a SEPARATE sub-system the line-height lever can't reach.
        // Combined with VSNAP=break-even and S457=optimal, NO multi-lever win was found this
        // session; the deep sub-systems (body 83/64 precision, title pBdr/spacing advance,
        // snap phase) each need separate work. Experimental knobs reverted.
        // S617 (2026-06-19) — cumulative-position-snap fully explored on gen2, NO win.
        // The "縦スタック累積位置スナップ" task. OXI_S467_VSNAP already implements Word's
        // model (visual_y accumulates EXACT line+spacing raw, emit = snap075). gen2 word_png
        // A/B: VSNAP+snap = net −0.0664 vs OFF (worse — snap075 injects ±0.375pt phase
        // noise); a NOSNAP variant (exact-raw, no 0.75pt snap) = ≈break-even with OFF. So
        // neither the snapped nor the exact-raw cumulative beats the current floor/round
        // model. The residual (title-block deficit −1.26 + body slope +0.106/line) is NOT a
        // snap-model issue — it is the 83/64 per-font RAW line-height precision (Oxi 83/64 ≈
        // Word's raw to ~0.006pt but the body's ROUND-to-0.5pt = +0.11/line, and the title
        // FLOOR compensates the glyph-anchor offset S615 exposed) — the deepest most-reverted
        // lever (S510/S612). A clean win needs per-font sub-0.02pt raw precision + the
        // cumulative model TOGETHER, co-gated on word_png pixels.
        let snap075 = |y: f32| -> f32 { (y / 0.75).round() * 0.75 };

        let (effective_spacing, continuous_section_start_spacing) = self.paragraph_spacing_before(
            para, page, grid_pitch, prev_style_id, prev_contextual_spacing,
            prev_autospacing_numid, prev_space_after, body_para_index,
            pages, current_elements, cursor.cursor_y, page_top,
        );
        // Paragraph spacing is an intentional displacement, not grid rounding.
        // Carry it into the ideal twip position so resynchronization cannot
        // mistake a half-line gap for device rounding residue.
        if std::env::var("OXI_GRID_PARAGRAPH_GAPS").is_ok()
            && cursor.lm2_ideal_y > 0.0
        {
            cursor.lm2_ideal_y += effective_spacing * 20.0;
        }
        cursor.advance(effective_spacing);

        // S1181 v2 (2026-08-20, OPT-IN OXI_S1181=1): a NO-TYPE docGrid runs
        // Word's legacy 96dpi engine — every paragraph's FIRST-LINE TOP is
        // PAINTED on a whole device pixel: round(cumulative_exact / 0.75) ×
        // 0.75, half-up on the ABSOLUTE coordinate, while the stream stays
        // EXACT (compounding falsified twice on _pb_gridwalk A-group k=7) AND
        // the page-bottom fit is judged on the EXACT position (_pb_gridfit
        // F240: the FIT→PUSH transition sits exactly on the exact threshold;
        // both rounded-fit variants over-fit x=566/568 and are refuted — v1's
        // cursor-snap regressed nyserda/uklocal/legal_0001482d and cratered
        // technical__002c1ffa through exactly that). So: VISUAL track only —
        // pagination is byte-identical by construction. Restore at exit
        // unless a page break re-synced both tracks. Derived: _pb_gridwalk
        // (110-para walk × pitch 360/326, quantum 0.75 pitch-independent,
        // exact/atLeast/auto all quantized, absolute anchor) + _pb_gridfit.
        // JP/gen2 no-type-grid docs excluded by the real-CJK gate.
        let (s1181_unsnap, s1181_pages0) = if page.doc_grid_no_type
            && !self.doc_body_has_real_cjk
            && !in_textbox
            && body_para_index.is_some()
            && std::env::var("OXI_S1181").is_ok()
        {
            let exact = cursor.visual_y;
            let snapped = (exact / 0.75).round() * 0.75;
            cursor.advance_split(0.0, snapped - exact);
            (exact - snapped, pages.len())
        } else {
            (0.0, usize::MAX)
        };

        // S658 (2026-06-24, pBdr border-merge): reserve the vertical space ABOVE
        // the paragraph for its TOP border (top.space + top.width). The border
        // element itself is drawn at para_top - top.space - top.width (mod.rs:6979)
        // but NOTHING reserved that space, so a boxed paragraph rendered ~5pt too
        // high (perturb_probe.py para_border -5.28 = top.space 4 + top.width 1 +
        // bottom bw/2 residual). The naive S658 attempt reserved this for EVERY box
        // and regressed 3a4f/model 94->95 pages: those docs STACK adjacent boxes
        // (top=6 bottom=6 left=6), and Word MERGES consecutive paragraphs with an
        // IDENTICAL pBdr into one box — the top border + its gap is drawn only above
        // the group's FIRST paragraph; interior boundaries use the "between" border
        // with NO extra gap. So skip the reservation when this paragraph continues a
        // merged group (the immediately-previous paragraph has the same pBdr). A
        // non-bordered paragraph between two boxes breaks the merge (prev != cur) and
        // correctly re-reserves. Default ON, opt-out OXI_S658_DISABLE.
        if std::env::var("OXI_S658_DISABLE").is_err() {
            if let Some(ref borders) = para.style.borders {
                if let Some(ref top) = borders.top {
                    let merges_with_prev = prev_para_borders == Some(borders);
                    if !merges_with_prev {
                        // S1504 (2026-09-20, default ON, opt-out OXI_S1504_DISABLE):
                        // a DOUBLE rule is three strokes wide (2 lines + gap, each
                        // sz/8). grid_probe.py B (typed grid 360, MS Gothic 10.5,
                        // Info6 of the plain paragraphs around): none 18 / single
                        // sz4 sp1 21 / double sz4 sp1 23.25 / single sz8 sp1 22.5 /
                        // single sz4 sp4 27 / top-only 19.5 -> space + full stroke
                        // width on BOTH sides; Oxi reserved double as single
                        // (20.75) and halved the bottom on CJK docs. ca290d's two
                        // bordered headings were 2.3pt short, pushing two
                        // paragraphs across page bottoms.
                        let s1504 = std::env::var_os("OXI_S1504_DISABLE").is_none();
                        let tw = if s1504 && top.style == "double" { top.width * 3.0 } else { top.width };
                        cursor.advance(top.space + tw);
                    }
                }
            }
        }
        // S1134 (2026-08-15): the cursor now stands at this paragraph's content
        // top. An EMPTY bordered paragraph emits no element, and the border draw
        // below falls back to `start_x` — the LEFT MARGIN used as a Y. A footer
        // whose separator rule is exactly that (a text-less paragraph carrying
        // only `pBdr top`) therefore drew its rule near the top of the BODY:
        // technical__002c1ffa's rule lands at 118.75 = 120.5 - space 1 - 0.75
        // against Word's 616.39, on all 368 pages.
        let s1134_content_top = cursor.cursor_y;
        // S1135 (2026-08-15): an `atLeast` line puts its extra leading ABOVE the
        // text, and Word's TOP BORDER comes down with the text rather than
        // staying at the line-box top. Probe _pb_bdratleast_gen.py, 8 arms:
        // atLeast 13pt over 8pt Times New Roman (leading 3.80) moves Word's rule
        // from 82.56 to 86.30 while Oxi holds it at 82.50; atLeast 20pt moves it
        // 10.80. `exact` (text also moves down) and a line MULTIPLIER (leading
        // below the text) leave the rule at the box top in both engines, so the
        // shift is the atLeast leading alone, not the glyph offset. Filled from
        // the first line below.
        let mut s1135_atleast_lead = 0.0f32;

        // Debug: dump per-paragraph cursor_y for Class A FAIL root cause investigation.
        // Gated by env OXI_DUMP_CURSOR_Y. Day 33 part 7 (option B).
        if std::env::var("OXI_DUMP_CURSOR_Y").is_ok() {
            let pi_str = body_para_index
                .map(|v| v.to_string())
                .unwrap_or_else(|| "?".into());
            let n_runs = para.runs.len();
            let txt: String = para
                .runs
                .iter()
                .flat_map(|r| r.text.chars())
                .take(20)
                .collect();
            eprintln!(
                "[CY_DUMP] body_pi={} cursor_y={:.3} space_before={:.3} n_runs={} text={:?}",
                pi_str, cursor.cursor_y, effective_spacing, n_runs, txt
            );
        }

        // When both twip and *Chars values exist, twip is authoritative (pre-computed by Word).
        // Fall back to *Chars × 10.5pt only when twip value is absent.
        let indent_left = para
            .style
            .indent_left
            .or_else(|| self.s1349_left_pt(para, page.grid_char_pitch, page.grid_char_cw_ratio))
            .unwrap_or(0.0);
        let indent_right = para
            .style
            .indent_right
            .or_else(|| para.style.indent_right_chars.map(|c| self.s1349_default_chars_pt(c, para, page.grid_char_pitch, page.grid_char_cw_ratio)))
            .unwrap_or(0.0);
        let first_line_indent_raw = para
            .style
            .indent_first_line
            .or_else(|| para.style.indent_first_line_chars.map(|c| LayoutEngine::s1214_chars_pt(c, para, true, page.grid_char_pitch, page.grid_char_cw_ratio)))
            .unwrap_or(0.0);
        // COM-confirmed (2026-04-25, e3c545 P1 "3．基本的な考え方" + 3a4f + NH_A..F
        // repros): for numbered list paragraphs with hanging indent and tab suffix
        // (default), Word places the marker at `left - hanging` and the first-text
        // character at `left` — the hanging area is consumed by the marker+tab,
        // not used to pull the first line leftward. Treating `first_line_indent`
        // as 0 here prevents the marker and text from overlapping.
        let list_consumes_hanging = para.style.list_marker.is_some()
            && first_line_indent_raw < 0.0
            && matches!(para.style.list_suff.as_deref(), None | Some("tab"));
        let mut first_line_indent = if list_consumes_hanging {
            0.0
        } else {
            first_line_indent_raw
        };

        // 2026-05-08 Bug B (Session 55+ Day 14): leading whitespace absorbs indent.
        // When a paragraph's leading whitespace (ASCII space + CJK fullwidth space)
        // pt > L + FL pt, Word renders text at the page margin, treating the
        // leading whitespace as visual indent.
        //
        // COM-confirmed on bd90b00 pi=24 ('統計センター...' with 60 leading
        // ASCII spaces, L=102tw FL=178tw). Word x=56.5 (page margin) vs
        // Oxi pre-fix x=70.7. After fix Oxi line 1 collapses to 1 line.
        //
        // NARROW trigger (full-context scan over 267 docx, body+table-cell+
        // header+footer+footnote+endnote+textbox): only bd90b00 pi=24 +
        // 3a4f9f pi=1410 match. ZERO PASS doc paragraph matches.
        //
        // Day 13 baseline drift discovery showed Day 12's verify (-0.1365)
        // was drift-induced; after baseline refresh, real Δ is +0.0098 net
        // (bd90b00 p.2 -0.0722→-0.0659 improvement, ed025 p.7 +0.0998 etc).
        let mut indent_absorbed_by_leading_ws = false;
        if indent_left > 0.0 && first_line_indent > 0.0 {
            let para_font_size = self.resolve_font_size(
                para.runs
                    .first()
                    .map(|r| &r.style)
                    .unwrap_or(&RunStyle::default()),
                &para.style,
            );
            let leading_ws_pt: f32 = {
                let mut sum = 0.0_f32;
                'outer: for run in &para.runs {
                    let run_fs = run.style.font_size.unwrap_or(para_font_size);
                    for c in run.text.chars() {
                        match c {
                            ' ' => sum += run_fs * 0.5,
                            '\u{3000}' => sum += run_fs,
                            _ => break 'outer,
                        }
                    }
                }
                sum
            };
            if leading_ws_pt > indent_left + first_line_indent {
                indent_absorbed_by_leading_ws = true;
            }
        }
        let indent_left = if indent_absorbed_by_leading_ws {
            0.0
        } else {
            indent_left
        };
        if indent_absorbed_by_leading_ws {
            first_line_indent = 0.0;
        }
        // COM-confirmed (2026-04-03): charGrid (linesAndChars) ignores paragraph indents
        // for line-break purposes. Text starts at margin and charsLine determines wrapping.
        // data_guideline: indent=12pt but x0=71 (margin), 38ch/line (=charsLine+1 kinsoku).
        // Round 29: when the para has snap_to_grid=false (e.g., footnote text
        // with pStyle "footnote text" / a8), DISABLE charGrid for line wrap
        // even if the page has linesAndChars docGrid. Otherwise the chars get
        // padded to the body's grid pitch and the line wraps ~5 chars early.
        //
        // S342 (2026-05-27) env-gated `OXI_S342_NO_SNAP_GATE=1`: drop the
        // snap_to_grid gate for char-grid (horizontal compression). Per OOXML
        // §17.3.1.32 `snap_to_grid` controls LINE SPACING (vertical), not
        // char pitch. b35123 i=89 has snap_to_grid=false + linesAndChars
        // charSpace=-2714 + Word still compresses chars per the grid (S342
        // direct measurement: avg 8.4375pt/char vs nominal sz=18=9.0pt).
        // Default OFF preserves Round 29 behavior; turn ON to test.
        //
        // S344 (2026-05-27): also pass-through to break_into_lines for per-char
        // fs<default_fs filtering (the actual Word behavior discriminator).
        // S342 SHIP (2026-05-27): default ON. Drops snap_to_grid gate from
        // char-grid (horizontal compression) per OOXML §17.3.1.32. Env-var
        // preserved as opt-OUT.
        let s342_no_snap_gate = std::env::var("OXI_S342_NO_SNAP_GATE")
            .map(|v| v != "0" && v != "false")
            .unwrap_or(true);
        let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT")
            .map(|v| v != "0" && v != "false")
            .unwrap_or(false);
        let snap_pass_through = s342_no_snap_gate || s344_fs_gate;
        let snap_gate_active = !snap_pass_through && !para.style.snap_to_grid;
        let effective_char_pitch = if in_textbox || snap_gate_active {
            None
        } else {
            page.grid_char_pitch
        };
        // 2026-05-05 Track A (Session 55+): COM-measured 8 paragraphs in b837
        // confirmed Word's wrap rule: available = content_w - indent_l - indent_r
        // for both charGrid and non-charGrid (full indent applied, no cell-based
        // tolerance). Combined with fn attribution fix below, b837 pagination
        // score improved 0.9524 → 0.9744 (Phase 1 gate). Other docs unchanged.
        //
        // S1211 (2026-08-24, opt-in `OXI_S1211=1`): under a linesAndChars grid a
        // BODY line may only use a WHOLE number of grid cells -- the remainder
        // (always less than one pitch) is unusable. MEASURED two ways with
        // `tools/metrics/_pb_gridpitch_gen.py` and `_pb_gridfloor_gen.py`:
        //   break  -- charSpace 1966 / 1453 / 532 / none x sizes 9..12, right
        //             indent swept in 0.25pt steps: 18 of 20 arms break exactly
        //             where floor(content/pitch)*pitch - indents predicts, the
        //             other two within 0.03pt over 44-49 characters.
        //   justify - a jc=both paragraph's justified lines END at that width:
        //             charSpace 1966 (pitch 10.98, 38 cells of a 425.2pt measure)
        //             ends 8.5pt short of the right margin and 0.6pt short of the
        //             floored 417.24; charSpace 532 (40 cells, remainder 0.004)
        //             ends at the margin.
        // So the text area really is narrower -- this is not a break-only rule.
        // NOT covered by the measurement: what a CENTERED line centres in, and
        // whether a cell floors (it does NOT -- `_pb_cellpitch_gen.py`, 9 of 9
        // arms break against the cell's own inner width), so only this body
        // site is gated here.
        // The measured domain is charSpace != 0. A `linesAndChars` grid with NO
        // charSpace still yields a pitch here (= the default size), and flooring
        // THAT costs b837808d / harassmanual / parttime a page each (1.0000 ->
        // 0.9577 / 0.9836 / 0.9970) -- all three are outside every arm measured
        // above. Left alone until it is measured on its own.
        let s1211 = std::env::var("OXI_S1211").ok().as_deref() == Some("1");
        // S1211B (2026-08-26, opt-in OXI_S1211B=1): extend the floor to
        // charSpace-LESS linesAndChars grids. The "flooring costs parttime/
        // b837/harassmanual a page" measurement that held this back predates
        // S1233 — parttime's floor pitch was then poisoned to 11.47 by the
        // stub section's charSpace. Word truth on the cs-less shape
        // (_pb_rchars_gen fine sweep, grid 415, run 8pt, jc=both): capacity
        // n = floor(floor(usable/12.0)*12.0 / 8.0) — 63 chars at rc0
        // (42 cells x 12), 64 at rc-80 (43 cells) — the floor is REAL with
        // pitch = the default size.
        let s1211b = std::env::var("OXI_S1211B").ok().as_deref() == Some("1")
            && self.doc_default_sz_declared;
        let grid_has_char_space = match (page.grid_char_pitch, page.grid_char_cw_ratio) {
            (Some(pitch), Some(ratio)) if ratio > 0.0 => (pitch - pitch / ratio).abs() > 0.01,
            _ => false,
        };
        // S1211C (2026-08-26, opt-in OXI_S1211C=1): the COMPLETE floor law —
        // measured across 17 arms (margins 623..923 x run 8..14pt x Normal
        // 16/21/24 x dd present/absent, _pb_rchars/_pb_floorpitch series):
        //   capacity boundary = floor(raw_content/cell) x cell, cell =
        //   docDefaults size (else the default-paragraph-style size, = this
        //   effective_char_pitch), tested against the char's INK — the last
        //   char may overhang the floored edge by its right bearing
        //   (advance - ink ~ 0.12em). The horizontal twin of the S576
        //   page-bottom ink leniency. The "11pt exemption" (F3/F4, b837's
        //   43-char lines vs a 40-char box floor) is exactly this ink slack
        //   crossing a small partial cell; parttime's 63-vs-66 char lines are
        //   the floor itself (+0.108 SSIM when applied).
        // v2 (R8): the boundary test is compat-mode-dependent — settings-less/
        // legacy docs use the strict box floor, but compatibilityMode 15 (the
        // b837 slice bisection: settings with ONLY cm15 flips n 40→41) lets
        // the last char START at ≤ F, i.e. available = floored + one char
        // advance of the paragraph's base size; past the raw width the floor
        // is a no-op (b837: 444+11 > 453.5). R9: the SAME width must reach the
        // body page-fit/keepNext estimates (s1211c_floor_body_width mirror) or
        // Phase-1 breaks (harassmanual/parttime PASS→FAIL on the layout-only
        // floor). Shared logic lives in s1211c_floor_body_width.
        let _ = (s1211, s1211b, grid_has_char_space);
        let grid_content_width = self.s1211c_floor_body_width(
            para,
            content_width,
            effective_char_pitch,
            page.grid_char_cw_ratio,
        );
        let available_width = grid_content_width - indent_left - indent_right;

        // Render list marker if present
        // S517 (2026-06-09): index of the emitted list-marker element so the
        // first-line loop can back-patch its text_y_off to share the body
        // baseline (the marker element is emitted here, before the line loop
        // computes text_y_off). Only set for NON-bullet markers (number markers
        // like ①/(1)); bullets keep their own marker_y_offset tuning untouched.
        let mut s517_marker_el_idx: Option<usize> = None;
        // S776 (2026-07-10, opt-out OXI_S776_DISABLE): suff="nothing" — Word
        // renders the number CONTIGUOUS with the first line's text AT the
        // first-line position (nyserda «3.NON-COLLUSIVE BIDDING…», lvl suff
        // val="nothing": Word/Libra draw number+text together; Oxi placed the
        // marker at left−list_indent with a tab-model gap → every numbered
        // paragraph's left edge differed → the worst nyserda pages 0.53 vs
        // Libra 0.95). The marker element moves to the first-line x and the
        // body's first line starts right after it (effective_first_indent +=
        // marker_width, which also narrows line 0's wrap width like Word).
        let mut s776_marker_extra: f32 = 0.0;
        // S789: hanging-less level — the suffix tab's num stop (see the
        // marker block below).
        let mut s789_stop_extra: f32 = 0.0;
        // S893: over-wide marker pushes the suffix tab to the next default
        // stop (see the marker block below).
        let mut s893_stop_extra: f32 = 0.0;
        // S778 (2026-07-10, opt-out OXI_S778_DISABLE): the numbering LEVEL's
        // w:ind left is the marker SUFFIX-TAB stop, and it SURVIVES a direct
        // w:ind override on the paragraph. nyserda's definition list (level
        // left=720tw=36pt hanging=720, direct ind left=0 firstLine=360):
        // Word places the (empty, numFmt=none) marker at the first-line x
        // (=108) and TABS the text to margin+36 = 126; Oxi started the text
        // at 108 -> line 1 was 18pt WIDER -> fit "paid in cash." where Word
        // wraps -> the per-page line-phase shifts behind nyserda's worst
        // pages (0.54 vs Libra 0.95). extra = level_left - indent_left -
        // first_line (positive part) is 0 for the standard adopted-level
        // case (indent_left == level_left) -> byte-identical there.
        // ★Latin scope (the TABTW/LATINEM discriminator): JP direct-numPr
        // lists (3a4f/model/harassmanual/tokyoshugyo 第X条) do NOT tab to the
        // level left in Word — unscoped, all four flipped PASS→FAIL. The JP
        // suffix-tab model needs its own derivation (recorded).
        // S853 (2026-07-15, opt-out OXI_S853_DISABLE): the level_left tab stop
        // is only correct for a NON-hanging list (nyserda: positive firstLine,
        // marker on the first line, tab to level_left). A HANGING list
        // (list_consumes_hanging: marker hangs left of the text, wrapped lines
        // align at the paragraph's left indent) tabs to the paragraph's OWN
        // indent_left, NOT the level_left — Word respects the direct w:ind.
        // reports__0007f6be: bullet numId=8 level ind left=1440tw=72pt, direct
        // w:ind left=634tw=31.7pt hanging=274; Word tabs the text to 31.7pt
        // (x=103.7), Oxi's S778 tabbed to level 72pt (x=144) → line 1 ~40pt
        // narrower → the long "Reminder MDPH…Register here." wrapped to 2 lines
        // where Word fits 1 → the extra line pushed the bullet off the page 3
        // bottom (+1 page). For the standard adopted-level hanging case
        // (indent_left == level_left) both branches give 0 → byte-identical.
        // ★DISCRIMINATOR (marker position sign): the direct w:ind overrides the
        // level tab stop ONLY when the marker sits at/right of the text margin
        // (indent_left + first_line_indent_raw >= 0, i.e. left >= hanging). When
        // the marker HANGS INTO the left margin (left < hanging, marker_pos < 0)
        // the direct left is a degenerate value and Word keeps tabbing to the
        // level_left. uk_framework numId=20 ilvl=1 direct ind left=153tw=7.65pt
        // hanging=431tw=21.55pt (marker_pos = 7.65-21.55 = -13.9 < 0): Word tabs
        // to level 586tw=29.3pt, NOT to 7.65pt — an unscoped S853 re-wrapped
        // those clauses (SSIM p23 0.88→0.71). The 584tw "normal" ilvl=1 items
        // (marker_pos +7.65 >= 0) shift only 0.1pt (direct 584 ≈ level 586).
        let s853_zero_stop = std::env::var("OXI_S853_DISABLE").is_err()
            && list_consumes_hanging
            && (indent_left + first_line_indent_raw >= 0.0);
        let s778_stop_extra: f32 = if std::env::var("OXI_S778_DISABLE").is_err()
            && !self.doc_body_has_real_cjk
            && para.style.list_suff.as_deref() == Some("tab")
            && !s853_zero_stop
        {
            para.style
                .list_level_left
                .map(|ll| (ll - indent_left - first_line_indent).max(0.0))
                .unwrap_or(0.0)
        } else {
            0.0
        };
        if let Some(ref marker) = para.style.list_marker {
            let default_style = RunStyle::default();
            let marker_style = s1037_marker_style(para).unwrap_or_else(|| {
                para.runs
                    .first()
                    .map(|r| &r.style)
                    .unwrap_or(&default_style)
            });
            let marker_font_size = self.resolve_font_size(marker_style, &para.style);
            // Symbol font bullets (•/●) have large glyphs relative to em-square.
            // No font size adjustment needed — use the paragraph's font size directly.
            let marker_metrics = &*self.metrics_for(marker_style, &para.style);
            if std::env::var("OXI_DBG_MARKER").is_ok() {
                eprintln!("[MARKER] text={:?} style_fam={:?} style_sz={:?} -> resolved fam={:?} fs={:.2} | first_run_fam={:?} ppr_fam={:?} drs={:?}",
                    marker, marker_style.font_family, marker_style.font_size,
                    marker_metrics.family, marker_font_size,
                    para.runs.first().map(|r| r.style.font_family.clone()),
                    para.style.ppr_rpr.as_ref().and_then(|r| r.font_family.clone()),
                    para.style.default_run_style.as_ref().map(|r| r.font_family.clone()));
            }
            // S692 (2026-06-29, SHIPPED default ON, opt-out OXI_MARKERCJK_DISABLE): a numbered-list label that is
            // CJK/full-width (e.g. 「第３４条」, the tokyoshugyo regulation article
            // markers) had its WIDTH computed with metrics_for (the ASCII font, Century),
            // so the full-width digits 「３４」 fell to the proportional ~7.5pt (marker
            // w=36) — but the GDI RENDER paints them in the eastAsia font (MS Mincho,
            // 10.5/char = Word, via font-linking). The under-counted marker shifts the
            // body start ~6pt LEFT → the body line over-fits (fits 「当」 where Word
            // wraps). Resolve the marker width per-char (CJK → metrics_for_char's
            // eastAsia font), matching the render. See [[tokyoshugyo_wrap_not_cellheight]].
            let marker_cjk = std::env::var("OXI_MARKERCJK_DISABLE").is_err();
            let marker_width: f32 = marker
                .chars()
                .map(|c| {
                    // The CJK render paints full-width marker chars (e.g. 「第３４条」's
                    // digits) at the eastAsia monospace width (font_size), NOT the Century
                    // proportional 7.5 that metrics_for resolves at break time. Match the
                    // render so the marker width (→ hanging-indent → body start) is correct.
                    if marker_cjk && crate::font::is_fullwidth(c) {
                        return marker_font_size;
                    }
                    self.registry
                        .char_width_pt_with_fallback(c, marker_font_size, marker_metrics)
                })
                .sum();
            // S1349: a hanging indent given in characters (hangingChars, no twip
            // first line) is the marker's hanging too -- the parser's list_indent
            // could only see the level's twip.
            let s1349_list_indent = if para.style.indent_first_line.is_none() {
                para.style
                    .indent_first_line_chars
                    .filter(|c| *c < 0.0)
                    .map(|c| -LayoutEngine::s1214_chars_pt(c, para, true, page.grid_char_pitch, page.grid_char_cw_ratio))
            } else {
                None
            };
            let list_indent = s1349_list_indent.or(para.style.list_indent).unwrap_or(18.0);
            let mut marker_x = start_x + indent_left - list_indent;
            // S893 (2026-07-17, default ON, opt-out OXI_S893_DISABLE): the
            // list suffix tab can never go BACKWARD. When the MARKER is wider
            // than the hanging indent (marker end overshoots the ind_left
            // text stop), Word tabs the text to the next defaultTabStop
            // multiple past the marker end — Oxi placed the text AT ind_left,
            // OVERLAPPING the number and gaining phantom line-1 room.
            // uk_framework Heading2 «32. [Companies that employ their own
            // staff] Staff²⁵» (numId=20 lvl0 ind left=1080 hanging=360,
            // Humnst777→Calibri 18): marker '32.' = 22.5pt > hanging 18 →
            // marker end rel 58.5; Word text x=171.3 = margin + 72 =
            // ceil(58.5/36)·36 (dts 720tw); Oxi text x=153.2 = ind_left 54,
            // overlapping the marker end 157.7 by 4.5pt AND giving line 1
            // 18pt of extra room — the knife-edge that made the correct S892
            // NBSP width flip wp31 (the S559 pair blocking S892).
            if std::env::var("OXI_S893_DISABLE").is_err()
                && !self.doc_body_has_real_cjk
                && para.style.list_suff.as_deref().unwrap_or("tab") == "tab"
                && list_consumes_hanging
                && s778_stop_extra <= 0.0
                // disjoint from S789 (which requires list_indent == None)
                && para.style.list_indent.is_some()
                && marker_width > list_indent + 0.01
            {
                let marker_end_rel = (indent_left - list_indent) + marker_width;
                let dts = self.default_tab_stop;
                if dts > 0.0 {
                    let stop = ((marker_end_rel / dts).floor() + 1.0) * dts;
                    s893_stop_extra = (stop - indent_left).max(0.0);
                }
            }
            let line_height = self.line_height(
                marker_font_size,
                para.style.line_spacing,
                para.style.line_spacing_rule.as_deref(),
                marker_metrics,
                para.style.snap_to_grid,
                grid_pitch,
            );

            // Determine marker text including suffix
            let suff = para.style.list_suff.as_deref().unwrap_or("tab");
            if matches!(suff, "nothing" | "space") && std::env::var("OXI_S776_DISABLE").is_err() {
                // A non-tab suffix places text after the actual marker advance.
                // Its space consumes line capacity just like the number itself.
                marker_x = start_x + indent_left + first_line_indent;
                s776_marker_extra = marker_width + if suff == "space" {
                    self.registry.char_width_pt_with_fallback(' ', marker_font_size, marker_metrics)
                } else { 0.0 };
            }
            if s778_stop_extra > 0.0 {
                // S778: the marker sits at the paragraph's first-line indent;
                // the suffix tab sends the text to the level-left stop.
                marker_x = start_x + indent_left + first_line_indent;
            }
            // S789 (2026-07-11, opt-out OXI_S789_DISABLE): a numbering level
            // with NO hanging indent (ind left=0 firstLine=0) places its
            // marker AT the first-line position — NOT at the fabricated
            // left−18pt default — and the suffix tab sends the text to the
            // level's own num tab stop. nyserda numId=11 lvl0 ('1.'-'9.'
            // form items, num tab at 288tw): Word marker x=90 (margin), text
            // x=104.4; Oxi put the marker at 72 and text at 90 → line 1 was
            // 18pt wider → '…BECOME EFFECTIVE UNLESS' over-packed (the
            // LRPB-off catalog #2). Latin scope (JP list model untouched).
            if std::env::var("OXI_S789_DISABLE").is_err()
                && !self.doc_body_has_real_cjk
                && suff == "tab"
                && para.style.list_indent.is_none()
                && s778_stop_extra <= 0.0
            {
                if let Some(stop) = para.style.list_tab_stop {
                    marker_x = start_x + indent_left + first_line_indent;
                    s789_stop_extra = (stop - first_line_indent).max(marker_width);
                }
            }
            let marker_text = match suff {
                "space" => format!("{} ", marker),
                "nothing" => marker.clone(),
                // "tab" — marker text alone; tab stop handled by indent_left
                _ => {
                    // For tab suffix: if there's a tab_stop defined, use it to
                    // adjust text start position via indent_left. The marker sits
                    // at marker_x and text starts at indent_left (which should
                    // align with the tab stop).
                    marker.clone()
                }
            };

            // Page break check for marker.
            // S651 (2026-06-24): a 2-cell (sz>=14) numbered chapter heading's MARKER
            // must break to the next page WITH its body. `line_height` above (via
            // com_line_height) can return a 1-cell value (18) while the body's first
            // line snaps to 2 cells (line_height_for_line, 83/64 grid-ceil → 36); when
            // S651 pushes the body's first line to the next page, the marker would
            // otherwise STRAND on this page. Use the grid-snapped 83/64 natural for the
            // marker's break check so the marker moves with its body. Only changes a
            // multi-cell heading at the page bottom (the S651 case); 1-cell markers are
            // unchanged (snapped == line_height).
            let marker_break_h =
                if para.style.snap_to_grid && std::env::var("OXI_S651_DISABLE").is_err() {
                    match grid_pitch {
                        Some(p) if p > 0.0 => {
                            // Use the CJK text metrics (metrics_for_text on the marker text)
                            // to match the BODY's grid-snapped height — marker_metrics (style)
                            // resolves to the ascii/hAnsi font (non-83/64) giving a 1-cell
                            // value, while the body uses the eastAsia 83/64 font (2 cells).
                            let nat_metrics =
                                &*self.metrics_for_text(&marker_text, marker_style, &para.style);
                            let nat = nat_metrics.word_line_height_no_grid(marker_font_size);
                            let snapped = (((nat + p * 0.5) / p) + 0.5).floor().max(1.0) * p;
                            line_height.max(snapped)
                        }
                        _ => line_height,
                    }
                } else {
                    line_height
                };
            // S737 (2026-07-04): the marker's page-break check must not be
            // STRICTER than the body line0's lenient threshold. For a 1-cell
            // marker in a typed grid, marker_break_h = the FULL grid cell (18)
            // while the body line0 uses the Day-33 natural_lh leniency (13.5)
            // — so at a page-bottom cursor in the leniency window the MARKER
            // alone pushed the whole numbered paragraph (problist: pi=35 at
            // cursor 755, 755+18=773>771 pushed; Word keeps the row, body
            // line0 over = −2.5). Use the marker's natural (un-snapped) height
            // for the check; the S651 multi-cell heading case (snapped >
            // 1.5×pitch) keeps its strict box. Opt-out OXI_S737_DISABLE.
            let marker_break_h = if std::env::var("OXI_S737_DISABLE").is_err()
                && para.style.snap_to_grid
                && grid_pitch.map_or(false, |p| p > 0.0 && marker_break_h <= p * 1.5)
            {
                let nat_metrics = &*self.metrics_for_text(&marker_text, marker_style, &para.style);
                nat_metrics
                    .word_line_height_no_grid(marker_font_size)
                    .min(marker_break_h)
            } else {
                marker_break_h
            };
            // S928 (2026-07-18, default ON, opt-out OXI_S928_DISABLE): the
            // no-grid Latin marker must use the SAME page-bottom occupancy
            // threshold as its body line.  S779/S827 derives that threshold as
            // the exact hhea line; the marker pre-check still used the AUTO
            // line-spacing advance and could therefore push the whole list
            // paragraph even when the body line itself fit.  policies f7115
            // pi547: cursor 755.09 + full 15.87 > 769.90 (false push), while
            // cursor + hhea 13.80 = 768.89 fits, as Word does.  This is the
            // no-grid counterpart of S737's typed-grid marker/body alignment.
            let marker_break_h = if std::env::var("OXI_S928_DISABLE").is_err()
                && (page.grid_line_pitch.is_none() || page.doc_grid_no_type)
                && (!self.doc_body_has_real_cjk
                    || (std::env::var("OXI_CJK_MARKER_FIT").is_ok()
                        && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))))
                // Footnote/footer area tops are occupied boundaries, not the
                // plain page margin for which S779's leading overhang was
                // derived.  Keep their established full-box marker check
                // (uklocalspending wp5).
                && !fn_boundary_active
                && !footer_tight
            {
                let nat_metrics = &*self.metrics_for_text(&marker_text, marker_style, &para.style);
                let mut marker_natural = if self.doc_body_has_real_cjk
                    && std::env::var("OXI_CJK_MARKER_FIT").is_ok()
                    && nat_metrics.is_cjk_83_64_font()
                {
                    LayoutEngine::s1367_cjk_box(nat_metrics, marker_font_size)
                } else {
                    nat_metrics.natural_line_height_hhea(marker_font_size)
                };
                // A Symbol bullet contributes ascent, while descent comes from
                // the text (the same component rule as S820b below). Counting
                // Symbol's own descent here can reject a line the body accepts.
                if marker.contains('\u{F0B7}')
                    && !matches!(para.style.line_spacing_rule.as_deref(), Some("exact") | Some("atLeast"))
                    && std::env::var("OXI_SYMBOL_FIT_DESCENT_DISABLE").is_err()
                {
                    let text_descent = para.runs.iter()
                        .filter(|r| !r.text.trim().is_empty())
                        .map(|r| {
                            let fs = self.resolve_font_size(&r.style, &para.style);
                            self.metrics_for_text(&r.text, &r.style, &para.style).win_descent * fs
                        })
                        .reduce(f32::max);
                    if let Some(descent) = text_descent {
                        marker_natural -= (nat_metrics.win_descent * marker_font_size - descent).max(0.0);
                    }
                }
                marker_natural.min(marker_break_h)
            } else {
                marker_break_h
            };
            if cursor.cursor_y + marker_break_h > page_top + content_height {
                let marker_columns = num_columns > 1
                    && (std::env::var("OXI_MARKER_COLUMN_FLOW").is_ok()
                || std::env::var("OXI_S1473_DISABLE").is_err());
                let old_origin = (start_x, cursor.cursor_y);
                if marker_columns && cur_col + 1 < num_columns {
                    cur_col += 1;
                    start_x = col_x_positions[cur_col];
                    cursor.set(column_flow_top(if pages.len() == s749_pages_at_entry { col_band_top } else { page_top }, start_x, pages.len()));
                } else {
                dbg_page_push(pages.len(), 0);
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                if let Some(g) = s755_geom {
                    page_top = g.top(pages.len() + 1);
                    content_height = g.ch(pages.len() + 1);
                }
                cursor.set(page_top);
                if marker_columns {
                    cur_col = 0;
                    start_x = col_x_positions.first().copied().unwrap_or(start_x);
                }
                }
                if marker_columns {
                    let dx = start_x - old_origin.0;
                    let dy = cursor.cursor_y - old_origin.1;
                    marker_x += dx;
                    for element in &mut elements {
                        element.x += dx;
                        element.y += dy;
                        if let LayoutContent::TableBorder { x1, y1, x2, y2, .. } = &mut element.content {
                            *x1 += dx; *x2 += dx; *y1 += dy; *y2 += dy;
                        }
                    }
                }
            }

            // Bullet markers are scaled up (2x) so adjust Y to align with text center
            let marker_y_offset = if marker.contains('\u{2022}') || marker.contains('\u{25CF}') {
                -marker_font_size * 0.15 // shift up slightly
            } else {
                0.0
            };
            // Resolve marker font from the paragraph's first-run style, matching
            // the cell renderer (mod.rs:~4780). Without this the GDI renderer
            // falls back to its default font and halfwidth markers like "(1)"
            // render narrower than Word (user-reported on e3c545 p.1 "(1)
            // 公開するデータの設計" — Word 14px vs Oxi 10px marker width).
            let marker_font_family = if marker_text.contains('\u{F0B7}') {
                // S491: a raw Symbol PUA bullet (kept by map_symbol_bullets under
                // OXI_S491_SYMBOL_BULLET) must render in the Symbol font — the
                // numbering level's rFonts is Symbol, not the paragraph's CJK font.
                Some("Symbol".to_string())
            } else if marker_text.contains('\u{F06E}') {
                // Wingdings-bullet: a raw Wingdings PUA bullet (kept by map_symbol_bullets)
                // renders in the Wingdings font (0x6E = ■), not the body font.
                Some("Wingdings".to_string())
            } else {
                self.resolve_font_family_for_text(&marker_text, marker_style, &para.style)
                    .map(|s| s.to_string())
            };
            let marker_bold = self.resolve_bold(marker_style, &para.style);
            let marker_color = self
                .resolve_color(marker_style, &para.style)
                .map(|s| s.to_string());
            let marker_base_y = if s467_vsnap {
                snap075(cursor.visual_y)
            } else {
                cursor.visual_y
            };
            // S517: remember this marker element's index so the first body line
            // can set its text_y_off to match (the marker shares the body
            // baseline — Word-confirmed dy=0 on b837 ①②③). Scoped to non-bullet
            // markers (marker_y_offset==0) so bullet placement is unchanged.
            if marker_y_offset == 0.0 {
                s517_marker_el_idx = Some(elements.len());
            }
            elements.push(LayoutElement::new(
                marker_x,
                marker_base_y + marker_y_offset,
                marker_width,
                line_height,
                LayoutContent::Text {
                    text: marker_text,
                    font_size: marker_font_size,
                    font_family: marker_font_family,
                    bold: marker_bold,
                    italic: marker_style.italic,
                    underline: marker_style.underline,
                    underline_style: marker_style.underline_style.clone(),
                    strikethrough: marker_style.strikethrough,
                    double_strikethrough: marker_style.double_strikethrough,
                    color: marker_color,
                    highlight: marker_style.highlight.clone(),
                    field_type: None,
                    character_spacing: 0.0,
                    text_scale: 100.0,
                    is_vertical: false,
                    effects: TextEffects::default(),
                },
            ));
        }

        // Collect all text fragments with their styles, field types, and source indices
        // S673vi (2026-06-26): an inline HIDDEN run (`<w:vanish/>` on the run) is NOT
        // displayed/printed by Word — its text reserves 0 width. Word render-truth
        // (minimal repro 前[隠し vanish]後): Word line width 20.16 (= 前後, hidden run
        // gone); Oxi rendered all 4 chars (42.0). Skip vanish runs from the fragment
        // list (filter AFTER enumerate so the run index `i` is preserved for the
        // surviving runs). Greenfield: 0 corpus docs carry an inline vanish run
        // (the 3 vanish docs use it only on the ¶ mark, S673v) → byte-identical.
        // Opt-out OXI_S673VI_DISABLE.
        let s673vi = std::env::var("OXI_S673VI_DISABLE").is_err();
        // S677 (2026-06-27): w:caps / w:smallCaps text transform. Word renders a
        // w:caps run in UPPERCASE, and a w:smallCaps run with originally-lowercase
        // letters as SMALL uppercase (size × 0.8 — measured from the Word PDF:
        // 15.96/20.04) while originally-uppercase letters stay full size. The
        // transform changes both glyphs AND width, so it is applied to the input
        // fragments (break_into_lines copies them into LineFragments → both wrap and
        // render use the transformed text/size). Owned segments the fragments borrow
        // from; a smallCaps run splits into per-case-class segments. Gate-safe: only
        // caps/smallCaps runs are transformed and the corpus's caps/smallCaps live
        // ENTIRELY in UNUSED built-in styles (Subtle/Intense Reference, Book Title,
        // Calendar2 — used=0) → byte-identical. Opt-out OXI_S677_DISABLE.
        // S677b (2026-07-09): resolve caps from the RUN or the paragraph STYLE
        // (para.style.default_run_style, the merged pStyle char props — like
        // resolve_bold). nyserda's Exhibit headings are pStyle=Heading3 whose
        // style rPr sets <w:smallCaps/> with NO direct run rPr; the old
        // run-only check missed them (rendered mixed-case, not SMALL CAPS).
        let para_sc = para
            .style
            .default_run_style
            .as_ref()
            .map_or(false, |d| d.small_caps);
        let para_ac = para
            .style
            .default_run_style
            .as_ref()
            .map_or(false, |d| d.all_caps);
        let caps_active = std::env::var("OXI_S677_DISABLE").is_err()
            && (para_sc
                || para_ac
                || para
                    .runs
                    .iter()
                    .any(|r| r.style.all_caps || r.style.small_caps));
        let preserve_case_source = caps_active && std::env::var("OXI_SOURCE_TEXT_IDENTITY").is_ok();
        let case_sources: Vec<Option<CaseSourceMap>> = if preserve_case_source {
            para.runs.iter().map(|r| {
                (r.style.small_caps || r.style.all_caps || para_sc || para_ac)
                    .then(|| CaseSourceMap::uppercase(&r.text))
            }).collect()
        } else { Vec::new() };
        // Segments retain their position within the transformed run.
        let caps_segments: Vec<(String, RunStyle, Option<FieldType>, usize, usize)> = if caps_active {
            let mut segs = Vec::new();
            for (i, r) in para.runs.iter().enumerate() {
                if s673vi && r.style.vanish {
                    continue;
                }
                if r.style.small_caps || para_sc {
                    let small = self.resolve_font_size(&r.style, &para.style) * 0.8;
                    let mut cur = String::new();
                    let mut rendered_start = 0usize;
                    let mut cur_lower: Option<bool> = None;
                    for ch in r.text.chars() {
                        let is_l = ch.is_lowercase();
                        if cur_lower.is_some() && cur_lower != Some(is_l) && !cur.is_empty() {
                            let mut st = r.style.clone();
                            if cur_lower == Some(true) {
                                st.font_size = Some(small);
                            }
                            let text = std::mem::take(&mut cur);
                            let count = text.chars().count();
                            segs.push((text, st, r.field_type, i, rendered_start));
                            rendered_start += count;
                        }
                        cur_lower = Some(is_l);
                        for u in ch.to_uppercase() {
                            cur.push(u);
                        }
                    }
                    if !cur.is_empty() {
                        let mut st = r.style.clone();
                        if cur_lower == Some(true) {
                            st.font_size = Some(small);
                        }
                        segs.push((cur, st, r.field_type, i, rendered_start));
                    }
                } else if r.style.all_caps || para_ac {
                    segs.push((r.text.to_uppercase(), r.style.clone(), r.field_type, i, 0));
                } else {
                    segs.push((r.text.clone(), r.style.clone(), r.field_type, i, 0));
                }
            }
            segs
        } else {
            Vec::new()
        };
        let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = if caps_active {
            caps_segments
                .iter()
                .map(|(t, st, ft, ri, offset)| (t.as_str(), st, *ft, *ri, if preserve_case_source { *offset } else { 0 }))
                .collect()
        } else {
            para.runs
                .iter()
                .enumerate()
                .filter(|(_, r)| !(s673vi && r.style.vanish))
                .map(|(i, r)| (r.text.as_str(), &r.style, r.field_type, i, 0usize))
                .collect()
        };

        // Resolve font size for line breaking. An empty paragraph has no
        // glyphs whose run formatting can set its size; its paragraph mark
        // uses the paragraph/default inheritance chain instead.
        let default_style = RunStyle::default();
        let para_font_size = self.resolve_font_size(
            para.runs
                .first()
                .filter(|_| para.runs.iter().any(|r| !r.text.is_empty()))
                .map(|r| &r.style)
                .unwrap_or(&default_style),
            &para.style,
        );

        // Round 7: pre-compute ruby paragraph-tail expansion once.
        // Greenfield-dormant: 0/177 baseline docs use w:ruby, so this is
        // 0.0 for all baseline paragraphs. Used at last-line cursor advance
        // and gates ruby-annotation emission below.
        let ruby_para_expansion_pt = self.s1396_ruby_expansion(para, para_font_size);

        // Round 7.7: ruby atomic-wrap budget (conservative).
        // When a run has ruby_w > base_w (V2 case "とくてい" 22pt over
        // "特定" 21pt = 1pt overhang), the inline footprint of that run
        // is field_w = max(base_w, ruby_w), not base_w alone. Because
        // break_into_lines tracks fragment widths per char (not per Run)
        // and refactoring it to thread per-Run extra width is invasive,
        // we instead reserve the total overhang from available_width
        // up-front. This over-reserves slightly on multi-run paragraphs
        // where only one run has overhang, but never under-reserves —
        // ensuring atomic wrap correctness without touching the wrap
        // loop. Greenfield-dormant: total_overhang = 0 when no run has
        // ruby (or when ruby_w ≤ base_w, the common case).
        let ruby_total_overhang_pt: f32 = para
            .runs
            .iter()
            .filter_map(|run| run.ruby.as_ref().map(|r| (run, r)))
            .map(|(run, ruby_ir)| {
                let base_pt = run.style.font_size.unwrap_or(para_font_size);
                let hps_pt = ruby_ir
                    .hps_halfpt
                    .map(|h| h as f32 / 2.0)
                    .unwrap_or(base_pt / 2.0);
                let ruby_metrics = &*self.metrics_for_text(&ruby_ir.text, &run.style, &para.style);
                let base_metrics = &*self.metrics_for_text(&run.text, &run.style, &para.style);
                let ruby_w: f32 = ruby_ir
                    .text
                    .chars()
                    .map(|c| {
                        self.registry
                            .char_width_pt_with_fallback(c, hps_pt, ruby_metrics)
                    })
                    .sum();
                let base_w: f32 = run
                    .text
                    .chars()
                    .map(|c| {
                        self.registry
                            .char_width_pt_with_fallback(c, base_pt, base_metrics)
                    })
                    .sum();
                (ruby_w - base_w).max(0.0)
            })
            .sum();

        // COM-confirmed (d77a): firstLineIndent reduces first line WIDTH but does
        // NOT shift start position. Text starts at margin, line is shorter.
        //
        // S109d fix (2026-05-19): the d77a-derived `effective=0` was zeroing out
        // NEGATIVE first_indent (hanging) too, which loses the line-1 wrap
        // credit. COM-confirmed on hanging+charGrid v2 repros (H1v2/H4v2/
        // H9v2/H10v2/H3v2/4a36b62 para32): Word extends line 1 budget by
        // -first_line_indent for hanging paragraphs. Now we only zero out
        // POSITIVE first_indent (the d77a case); negative (hanging) keeps
        // the raw value so break_into_lines credits the hanging extension.
        // S168 (2026-05-22) Phase B-2 holistic bundle — breakthrough discovery.
        // S164 round 4 (per-line fn tracking alone) → b837 -0.4336 cascade.
        // S164 round 6 (first_indent wrap respect alone) → -0.0054 cascade.
        // BUT combined as a bundle: cascade COMPENSATES → +0.0526 b837 gain,
        // +0.0058 mean IoU strict increase, Phase 1 53/55 unchanged.
        // Mechanism: first_indent fix makes paragraphs wrap one more line
        // (i=50 "地方公共団体..." went 1→2 lines matching Word), AND per-line
        // fn tracking fits paragraph 39 line 2 on page 2 (matching Word).
        // The two boundaries (p2→p3 and p3→p4) compensate's cascading
        // shifts: per-line fn pushes p3 up 1 line, first_indent's extra wrap
        // on p4 i=50 absorbs the shift. Net: pages 3-7 align with Word.
        // S241 (2026-05-23): removed OXI_LEGACY_NO_B2_BUNDLE legacy
        // env-var fallback during hardening pass. S168 Phase B-2 bundle
        // is the canonical path.
        let effective_first_indent = first_line_indent
            + s776_marker_extra
            + s778_stop_extra
            + s789_stop_extra
            + s893_stop_extra;
        // S342: mirror the snap_to_grid gate change for cw_ratio (see effective_char_pitch comment).
        let effective_cw_ratio = if in_textbox || snap_gate_active {
            None
        } else {
            page.grid_char_cw_ratio
        };
        // S1314 (2026-09-05, default ON, opt-out OXI_S1314_DISABLE): the ruby field's
        // extra width is carried by the base run's character spacing (see
        // resolve_ruby_spread_runs), so nothing is reserved off every line.
        let ruby_total_overhang_pt = if std::env::var("OXI_S1314_DISABLE").is_err() { 0.0 } else { ruby_total_overhang_pt };
        let raw_wrap_width = available_width - ruby_total_overhang_pt;
        let wrap_width = if std::env::var("OXI_DEGENERATE_CJK_SPACES").is_ok() { raw_wrap_width } else { raw_wrap_width.max(0.0) };
        // S758: the paragraph starts inside a wrapSquare band — break at the
        // narrowed width; s758_wrap_full is the rebreak target at band exit.
        let s758_wrap_full = wrap_width;
        let s758_lane_minimum = body_wrap_bands.map_or(30.0, |(bands, entry_page)| {
            bands.iter().filter(|b| b.0 == entry_page && b.2 > cursor.cursor_y
                && b.3 < start_x + content_width && b.4 > start_x)
                .map(|b| b.6.minimum_lane_width).fold(30.0, f32::min)
        });
        // Apply the character-grid boundary to the space left by the float.
        // Subtracting its geometry from a previously rounded column would
        // charge the column's discarded fraction a second time.
        let floor_wrap_reduction = |red: f32| -> f32 {
            // S1457 (2026-09-17, default ON, opt-out OXI_S1457_DISABLE):
            // promotes the OXI_GRID_WRAP_WIDTH checkpoint, whose formula the
            // 0ea3ec86 measurement confirms. A side-wrap band must be taken
            // from the RAW column and the remaining lane re-floored to whole
            // grid cells; subtracting it from the already-floored column
            // charges the column's discarded fraction twice. p16 right column
            // (cell 11.5, column 235.6 -> floored 230.0, band reduction 139.4):
            // the old order handed the breaker 90.635, one cell short of the
            // eight that fit, so Word's nine 8-character lines became nine
            // 7-character ones and the column lost a line. Re-floored:
            // floor(96.28 / 11.5) = 8 cells = 92.0 and the document goes
            // 0.9975 -> 1.0.
            if red > 0.0
                && (std::env::var("OXI_GRID_WRAP_WIDTH").is_ok()
                    || std::env::var("OXI_S1457_DISABLE").is_err())
            {
                let free = self.s1211c_floor_body_width(
                    para, content_width - red, effective_char_pitch,
                    page.grid_char_cw_ratio);
                (grid_content_width - free).max(0.0)
            } else { red }
        };
        // S-TWOSEG: the pair supplies its own two widths, so the single-segment
        // narrowing and shift are dropped. The band BOTTOM is kept, because the
        // rebreak at the band exit is what returns the paragraph to full width
        // once it clears the float.
        let marker_left_region = (std::env::var("OXI_MARKER_COLUMN_FLOW").is_ok()
                || std::env::var("OXI_S1473_DISABLE").is_err())
            && (cur_col != start_column || pages.len() > s749_pages_at_entry);
        let s758_two_seg = if marker_left_region { None } else { s758_two_seg };
        let s758_band = if marker_left_region {
            None
        } else if s758_two_seg.is_some() {
            s758_band.map(|(bot, _, _)| (bot, 0.0, 0.0))
        } else {
            s758_band
        };
        let wrap_width = if let Some((_, red, _)) = s758_band {
            (wrap_width - floor_wrap_reduction(red)).max(s758_lane_minimum)
        } else {
            wrap_width
        };
        // S476: this is the MAIN BODY flow (s476_body=true) → S475/S476 yakumono
        // capacity may apply (the demand break). Aux/estimate/cell calls pass false.
        if let Ok(needle) = std::env::var("OXI_DBGFLUSH") {
            if !needle.is_empty() && para.runs.iter().any(|r| r.text.contains(&needle)) {
                eprintln!("[DBGWRAP] wrap_width={:.2} first_line_indent={:.2} eff_first={:.2} s776={:.2} s778={:.2} s789={:.2} avail_width_param={:.2}",
                    wrap_width, first_line_indent, effective_first_indent,
                    s776_marker_extra, s778_stop_extra, s789_stop_extra, available_width);
            }
        }
        let para_has_lrpb = para.runs.iter().any(|r| r.has_last_rendered_page_break);
        // ★S1026-REPLAY diagnostic (OXI_DBG_REPLAY, default OFF): tag the two body
        // break_into_lines calls (first pass + S721 retry) with this paragraph's
        // index so the trace joins to the first-divergence dataset. Behaviour-neutral.
        S1026_REPLAY_PARA.with(|c| c.set(body_para_index));
        // S1636: the lane's left edge relative to the paragraph's (0 at full width).
        self.s1636_lane_shift.set(s758_band.map_or(0.0, |b| b.2));
        let mut lines = self.break_into_lines(
            &fragments,
            wrap_width,
            effective_first_indent,
            &para.style,
            effective_char_pitch,
            effective_cw_ratio,
            page.doc_grid_lines_and_chars,
            true,
            matches!(para.alignment, Alignment::Justify | Alignment::Distribute)
                || (matches!(para.alignment, Alignment::Center)
                    && std::env::var("OXI_S1028_CT_DISABLE").is_err()
                    && self.compat_mode == 14
                    && self.compat_mode_explicit
                    && !self.doc_body_has_real_cjk),
            page.doc_grid_no_type,
            para_has_lrpb,
            caps_active,
            false,
        );
        self.s1636_lane_shift.set(0.0);
        // S-TWOSEG: replace the full-width break with one row per pair of strips.
        if let Some((_, seg1_w, _, seg2_w)) = s758_two_seg {
            let two = self.break_two_segment_lines(
                &fragments,
                seg1_w,
                seg2_w,
                effective_first_indent,
                &para.style,
                effective_char_pitch,
                effective_cw_ratio,
                page.doc_grid_lines_and_chars,
                true,
                matches!(para.alignment, Alignment::Justify | Alignment::Distribute),
                page.doc_grid_no_type,
                para_has_lrpb,
                caps_active,
            );
            if !two.is_empty() {
                lines = two;
            }
        }
        let s1026_replay_first_nlines = lines.len();
        // S721 body arm (2026-07-03, default ON, opt-out OXI_S721_DISABLE):
        // PARAGRAPH-TAIL ORPHAN ELIMINATION via a two-pass re-break. Word accepts
        // ABOVE-normal 約物 compression (~3.9/約物 vs the s590 break cap 1.5) when
        // it removes a short (≤2-glyph) final line. Render-truth tokyoshugyo p36 ③:
        // Word compresses the L1 、 by 3.78 (measured 6.72 render) to pull 数 up,
        // which lets L2 hold the rest (with the 。 hang) → 2 lines; Oxi's cap-1.5
        // break left 数 down → an ん。 orphan L3. Mechanism: if the first pass ends
        // with a ≤2-glyph final line, re-break the whole paragraph with the orphan
        // caps (thread-local flag read by the cap computation) and accept the
        // re-break ONLY when it saves a line (ikujidetail る。: demand ~11 > 3.9 →
        // count unchanged → keep the first pass, matching Word's oidashi there;
        // nedo 子: 1-glyph tail, fits at open 3.4 → saved → matches Word).
        let mut lines = {
            let last_glyphs: usize = lines
                .last()
                .map(|l| l.fragments.iter().map(|f| f.text.chars().count()).sum())
                .unwrap_or(0);
            if std::env::var("OXI_S721_DISABLE").is_err()
                && lines.len() >= 2
                && last_glyphs > 0 && last_glyphs <= 2
                && self.compress_punctuation
                // LEGACY scope (compat<15): the orphan-cap evidence (③ 3.78, ⑧ 3.9)
                // is all compat-11; nedocontract (compat 15, Word break cap ~3.4 per
                // S639b) DECLINES the same-size compression on its wi=51 short-tail
                // para — the retry flipped it PASS→FAIL until this gate.
                && self.compat_mode < 15
            {
                S721_ORPHAN_RETRY.with(|f| f.set(true));
                self.s1636_lane_shift.set(s758_band.map_or(0.0, |b| b.2));
                let retry = self.break_into_lines(
                    &fragments,
                    wrap_width,
                    effective_first_indent,
                    &para.style,
                    effective_char_pitch,
                    effective_cw_ratio,
                    page.doc_grid_lines_and_chars,
                    true,
                    matches!(para.alignment, Alignment::Justify | Alignment::Distribute)
                        || (matches!(para.alignment, Alignment::Center)
                            && std::env::var("OXI_S1028_CT_DISABLE").is_err()
                            && self.compat_mode == 14
                            && self.compat_mode_explicit
                            && !self.doc_body_has_real_cjk),
                    page.doc_grid_no_type,
                    para_has_lrpb,
                    caps_active,
                    false,
                );
                S721_ORPHAN_RETRY.with(|f| f.set(false));
                self.s1636_lane_shift.set(0.0);
                if retry.len() < lines.len() {
                    retry
                } else {
                    lines
                }
            } else {
                lines
            }
        };
        // ★S1026-REPLAY: emit the paragraph END marker (used pass + final nlines) so
        // the analysis unit can select the pass that produced the final render
        // (used_pass=1 iff the S721 retry saved a line). Then clear the para tag so
        // subsequent (non-body) break_into_lines calls never emit.
        if body_para_index.is_some() && std::env::var("OXI_DBG_REPLAY").is_ok() {
            let used_pass: u8 = if lines.len() < s1026_replay_first_nlines {
                1
            } else {
                0
            };
            eprintln!(
                "[S1026-REPLAY-END] para={} used_pass={} nlines={}",
                body_para_index.unwrap(),
                used_pass,
                lines.len()
            );
        }
        S1026_REPLAY_PARA.with(|c| c.set(None));

        // S168 Phase B-2 holistic bundle (b): per-line fn cumul delta.
        // S834 (2026-07-14, opt-out OXI_S834_DISABLE): the SEPARATOR allocation
        // is part of first_line_extra_content_h (delta_if_current includes
        // footnote_sep_alloc for the page's first fn) but was NEVER entered in
        // this per-line committed map — so the REF LINE's lenient extra stayed
        // at +sep_region (~36-49pt) and every anchor line could eat into the
        // separator/notice region → the fn area under-fit → footnote bodies
        // silently DROPPED (uklocal fn4: anchor last line at 664.8 vs the
        // reserved eff 632.6; FN_PLACE fit=0 gap 33.0 < 38.5). Word render
        // truth (fnr_Z sweep): the anchor line must co-locate with the FULL fn
        // machinery — anchor at 645.3 fits, at ~648 the ANCHOR moves to the
        // next page with its fn; no leniency at/after the ref line. FIX: fold
        // the sep part (first_line_extra − Σ para fn heights, clamped ≥0) into
        // the cumulative when the para's FIRST ref commits.
        // S900: per-line ids first committed by that line (doc order) — the
        // deferral test walks them to decide which notes still START on the
        // page. Filled only on the s834 path (Latin fn docs).
        // Plan the text rows before deriving any metrics or footnote arrays.
        // A word that needs an emergency character break waits until the
        // wrapping object has ended; preceding rows stay beside the object.
        // A picture can begin partway through its own anchor paragraph. Plan
        // each remaining row against the registry, including entry into a band.
        // The existing in-band planner is retained for paragraphs starting beside
        // an object; the new entry path uses the same geometry and word-fit rules.
        let future_wrap_band = if s758_band.is_none()
            && std::env::var_os("OXI_FUTURE_WRAP_DISABLE").is_none()
        {
            body_wrap_bands.and_then(|(bands, entry_page)| bands.iter()
                .filter(|b| (b.0 == entry_page && b.1 > cursor.cursor_y
                    || b.0 > entry_page)
                    && b.1 < page_top + content_height && b.2 > b.1
                    && b.3 < start_x + content_width - 6.0 && b.4 > start_x + 6.0)
                .map(|b| (b.2, 0.0, 0.0))
                .max_by(|a,b| a.0.total_cmp(&b.0)))
        } else { None };
        let mut word_fit_widths = Vec::new();
        let mut word_fit_floors: Vec<Option<f32>> = Vec::new();
        let mut word_fit_segments: Vec<Option<(f32,f32,f32,f32)>> = Vec::new();
        let mut word_fit_columns: Vec<usize> = Vec::new();
        if std::env::var("OXI_DBG_WF").is_ok() {
            eprintln!("[WF-GATE] cjk={} bpi={:?} two_seg={:?} clean={} band={:?} text={:?}",
                self.doc_body_has_real_cjk, body_para_index, s758_two_seg,
                fragments.iter().all(|f| !f.0.chars().any(|c| matches!(c, '\t' | '\u{FFFC}'))),
                s758_band,
                para.runs.iter().map(|r| r.text.as_str()).collect::<String>()
                    .chars().take(22).collect::<String>());
        }
        if (std::env::var("OXI_WRAP_WORD_FIT").is_ok()
                || std::env::var("OXI_S1472_DISABLE").is_err())
            && !self.doc_body_has_real_cjk && body_para_index.is_some()
            && fragments.iter().all(|f| !f.0.chars().any(|c| matches!(c, '\t' | '\u{FFFC}')))
        {
            if let Some((bottom, reduction, shift)) = s758_band.or(future_wrap_band) {
                // Requery current page/column bands even when this paragraph
                // starts beside a float. Its continuation can have other lanes.
                let dynamic_entry = body_wrap_bands.is_some();
                let mut row_para = para.clone();
                row_para.style.space_before = Some(0.0);
                let mut remaining: Vec<_> = fragments.iter().map(|f|
                    (f.0.to_owned(), f.1.clone(), f.2.clone(), f.3, f.4)).collect();
                let mut planned = Vec::new();
                let mut widths = Vec::new();
                let mut floors = Vec::new();
                let mut segments = Vec::new();
                let mut columns = Vec::new();
                let mut y = cursor.cursor_y;
                let mut column = start_column;
                let mut active = y < bottom - 0.5;
                let mut pending_floor = None;
                let mut valid = true;
                while !remaining.is_empty() {
                    let mut row_bottom = bottom;
                    let mut row_two = None;
                    let (red, sh) = if dynamic_entry {
                        let (bands, entry_page) = body_wrap_bands.unwrap();
                        let row_column = column % num_columns.max(1);
                        let row_page = entry_page + column / num_columns.max(1);
                        let row_x = col_x_positions.get(row_column).copied().unwrap_or(start_x);
                        let row_content_width = col_widths.get(row_column).copied().unwrap_or(content_width);
                        row_para.runs = remaining.iter().filter_map(|f| {
                            let mut run = para.runs.get(f.3)?.clone();
                            run.text = f.0.clone();
                            Some(run)
                        }).collect();
                        if !planned.is_empty() { row_para.style.indent_first_line = Some(0.0); }
                        let (lane, two, advance) = self.body_paragraph_wrap_bands(
                            &row_para, page, bands, row_page, y, row_x, row_content_width);
                        row_two = two;
                        if advance > 0.0 {
                            y += advance;
                            pending_floor = Some(y);
                            continue;
                        }
                        active = lane.is_some();
                        if let Some((bot, red, sh)) = lane {
                            row_bottom = bot;
                            if row_two.is_some() { (0.0,0.0) } else { (red,sh) }
                        } else { (0.0, 0.0) }
                    } else {
                        (if active { reduction } else { 0.0 }, if active { shift } else { 0.0 })
                    };
                    let row_content_width = col_widths.get(column % num_columns.max(1))
                        .copied().unwrap_or(content_width);
                    let row_full = (self.s1211c_floor_body_width(
                        para, row_content_width, effective_char_pitch, page.grid_char_cw_ratio)
                        - indent_left - indent_right - ruby_total_overhang_pt).max(0.0);
                    let row_reduction = if active {
                        if red > 0.0 && (std::env::var("OXI_GRID_WRAP_WIDTH").is_ok()
                            || std::env::var("OXI_S1457_DISABLE").is_err()) {
                            let full = self.s1211c_floor_body_width(
                                para, row_content_width, effective_char_pitch, page.grid_char_cw_ratio);
                            let free = self.s1211c_floor_body_width(
                                para, row_content_width - red, effective_char_pitch, page.grid_char_cw_ratio);
                            (full - free).max(0.0)
                        } else { red }
                    } else { 0.0 };
                    let width = if active { (row_full - row_reduction).max(s758_lane_minimum) } else { row_full };
                    let refs: Vec<_> = remaining.iter().map(|f|
                        (f.0.as_str(), &f.1, f.2.clone(), f.3, f.4)).collect();
                    let first_indent=if planned.is_empty() { effective_first_indent } else { 0.0 };
                    self.s1636_lane_shift.set(sh);
                    let mut broken = if let Some((_,left_width,_,right_width))=row_two {
                        self.break_two_segment_lines(&refs,left_width,right_width,first_indent,
                            &para.style,effective_char_pitch,effective_cw_ratio,
                            page.doc_grid_lines_and_chars,true,
                            matches!(para.alignment,Alignment::Justify|Alignment::Distribute),
                            page.doc_grid_no_type,para_has_lrpb,caps_active)
                    } else {
                        self.break_into_lines(&refs, width,first_indent,
                            &para.style, effective_char_pitch, effective_cw_ratio,
                            page.doc_grid_lines_and_chars, true,
                            matches!(para.alignment, Alignment::Justify | Alignment::Distribute),
                            page.doc_grid_no_type, para_has_lrpb, caps_active, false)
                    };
                    self.s1636_lane_shift.set(0.0);
                    let Some(first) = broken.first() else { valid = false; break; };
                    let natural = self.natural_line_height_for_line(first, &para.style, para_font_size);
                    // Fit the same declared line box that advances this plan.
                    // In particular, a nonempty exact-height row before a
                    // column/page control cannot fit using smaller glyph ink.
                    let row_height=self.line_height_for_line(first,&para.style,para_font_size,
                        para.style.snap_to_grid,grid_pitch,page.doc_grid_no_type);
                    if !first.fragments.is_empty() && y + row_height > page_top + content_height && y > col_band_top + 0.01 {
                        column += 1;
                        y = if column < num_columns { col_band_top } else { page_top };
                        active = false;
                        pending_floor = None;
                        continue;
                    }
                    // Floating tables permit emergency word breaks in their usable side lanes.
                    // Shape wrapping retains its word-fit deferral below the object.
                    let table_lane = body_wrap_bands.is_some_and(|(bands, entry_page)|
                        bands.iter().any(|b| b.0 == entry_page && b.6.break_long_words
                            && y + natural > b.1 && y < b.2 - 0.5
                            && b.3 < start_x + content_width && b.4 > start_x));
                    if active && first.emergency_word_break && !table_lane {
                        y = y.max(row_bottom);
                        active = false;
                        pending_floor = Some(row_bottom);
                        continue;
                    }
                    let next = broken.get(1).and_then(Line::source_start);
                    let line = broken.remove(0);
                    y += self.line_height_for_line(&line, &para.style, para_font_size,
                        para.style.snap_to_grid, grid_pitch, page.doc_grid_no_type);
                    let explicit_page = line.break_type == LineBreakType::PageBreak;
                    let explicit_column = line.break_type == LineBreakType::ColumnBreak;
                    planned.push(line);
                    widths.push((s758_wrap_full - row_full + row_reduction, sh));
                    floors.push(pending_floor.take());
                    segments.push(row_two);
                    columns.push(column);
                    if dynamic_entry && (explicit_page || explicit_column) {
                        column += if explicit_page { num_columns.max(1) - column % num_columns.max(1) } else { 1 };
                        y = if column < num_columns { col_band_top } else { page_top };
                        active = false;
                        pending_floor = None;
                    }
                    if y >= bottom - 0.5 { active = false; }
                    if let Some((ri, co)) = next {
                        let Some(index) = remaining.iter().position(|f| f.3 == ri && f.4 <= co
                            && co <= f.4 + f.0.chars().count()) else { valid = false; break; };
                        let count = co - remaining[index].4;
                        if index == 0 && count == 0 { valid = false; break; }
                        remaining.drain(..index);
                        remaining[0].0 = remaining[0].0.chars().skip(count).collect();
                        remaining[0].4 = co;
                    } else { remaining.clear(); }
                }
                if std::env::var("OXI_DBG_WF").is_ok() {
                    eprintln!("[WF-PLAN] valid={} n={} floors={:?} text={:?}",
                        valid, planned.len(), floors,
                        para.runs.iter().map(|r| r.text.as_str()).collect::<String>()
                            .chars().take(22).collect::<String>());
                }
                if valid && !planned.is_empty() {
                    lines = planned;
                    word_fit_widths = widths;
                    word_fit_floors = floors;
                    word_fit_segments = segments;
                    word_fit_columns = columns;
                }
            }
        }
        // Finish explicit column reflow before deriving row metrics and note
        // arrays. The retained prefix includes its source control; only the
        // remaining source is broken against the newly entered column width.
        // Word controls independently vary both widths and a soft/column break.
        if num_columns > 1 && word_fit_columns.is_empty()
            && col_widths.iter().any(|w| (*w - content_width).abs() > 0.1)
            && lines.iter().any(|line| line.break_type == LineBreakType::ColumnBreak)
        {
            let mut remaining_lines = std::mem::take(&mut lines);
            let mut column = start_column;
            let mut planned = Vec::new();
            let mut widths = Vec::new();
            let mut columns = Vec::new();
            let mut current_red = s758_band.map_or(0.0, |b| b.1);
            let mut current_shift = s758_band.map_or(0.0, |b| b.2);
            while !remaining_lines.is_empty() {
                let boundary = remaining_lines.iter().position(|line|
                    matches!(line.break_type, LineBreakType::ColumnBreak | LineBreakType::PageBreak));
                let count = boundary.map_or(remaining_lines.len(), |i| i+1);
                let break_type = remaining_lines[count-1].break_type;
                for line in remaining_lines.drain(..count) {
                    planned.push(line);
                    widths.push((current_red, current_shift));
                    columns.push(column);
                }
                if remaining_lines.is_empty() { break; }
                column += if break_type == LineBreakType::PageBreak {
                    num_columns - column % num_columns
                } else { 1 };
                let new_content_width = col_widths.get(column % num_columns)
                    .copied().unwrap_or(content_width);
                let new_wrap_width = (self.s1211c_floor_body_width(
                    para, new_content_width, effective_char_pitch, page.grid_char_cw_ratio)
                    - indent_left - indent_right - ruby_total_overhang_pt).max(0.0);
                let fragments: Vec<_> = remaining_lines.iter().flat_map(Line::source_fragments).collect();
                let refs: Vec<_> = fragments.iter().map(|(text,style,field,run,offset)|
                    (text.as_str(), style, field.clone(), *run, *offset)).collect();
                self.s1636_lane_shift.set(0.0);
                remaining_lines = self.break_into_lines(&refs, new_wrap_width, 0.0,
                    &para.style, effective_char_pitch, effective_cw_ratio,
                    page.doc_grid_lines_and_chars, true,
                    matches!(para.alignment, Alignment::Justify | Alignment::Distribute),
                    page.doc_grid_no_type, para_has_lrpb, caps_active, false);
                current_red = s758_wrap_full - new_wrap_width;
                current_shift = 0.0;
            }
            lines = planned;
            word_fit_widths = widths;
            word_fit_columns = columns;
        }
        let mut line_own_fn_ids: Vec<Vec<u32>> = vec![Vec::new(); lines.len()];
        // S900: note ids DEFERRED to the next page's area (excluded from this
        // page's per-page bucket; handed to the caller via fn_deferred_out).
        let mut s900_deferred_ids: Vec<u32> = Vec::new();
        let committed_fn_delta_at_line: Vec<f32> = if !para_fn_heights.is_empty() {
            // ★BUNDLED with S833 (opt-in OXI_S833=1, default OFF byte-identical):
            // S834 alone flips uklocalspending's LRPB-mode gate {-1:1} (an S559
            // compensation); the pair ships default-ON together with the Latin
            // LRPB-drop once the uklocal natural residual {+1:7} closes.
            let s834 = std::env::var("OXI_S833_DISABLE").is_err() && !self.doc_body_has_real_cjk;
            let sep_part: f32 = if s834 {
                (first_line_extra_content_h - para_fn_heights.values().sum::<f32>()).max(0.0)
            } else {
                0.0
            };
            let mut out = Vec::with_capacity(lines.len());
            let mut cumulative = 0.0_f32;
            let mut seen: Vec<u32> = Vec::new();
            for (li, line) in lines.iter().enumerate() {
                // S834(b): the same word-merge blind spot S826 fixed in the
                // per-line attribution — a ref run merged into a word fragment
                // («organisations.» + "3" → one fragment keeping the first
                // run_index) is invisible to the raw fragment scan. Sweep the
                // full run RANGE this line covers (fragments are emitted in
                // run order: [this line's first fragment run, next line's
                // first fragment run)). Bundled behind OXI_S833; the legacy
                // path keeps the raw per-fragment scan byte-identically.
                let lo = line.fragments.iter().map(|f| f.run_index).min();
                let hi = if s834 {
                    lines
                        .get(li + 1)
                        .and_then(|nl| nl.fragments.iter().map(|f| f.run_index).min())
                        .unwrap_or(para.runs.len())
                } else {
                    // legacy: only runs that appear as fragment run_index
                    // (reproduced by sweeping each fragment's own run only)
                    0
                };
                if !s834 {
                    for f in &line.fragments {
                        if let Some(run) = para.runs.get(f.run_index) {
                            if let Some(id) = run.footnote_ref {
                                if !seen.contains(&id) {
                                    seen.push(id);
                                    if let Some(&h) = para_fn_heights.get(&id) {
                                        cumulative += h;
                                    }
                                }
                            }
                        }
                    }
                    out.push(cumulative);
                    continue;
                }
                if let Some(lo) = lo {
                    for r in lo..hi.max(lo) {
                        if let Some(id) = para.runs.get(r).and_then(|x| x.footnote_ref) {
                            if !seen.contains(&id) {
                                if seen.is_empty() {
                                    cumulative += sep_part;
                                }
                                seen.push(id);
                                if let Some(&h) = para_fn_heights.get(&id) {
                                    cumulative += h;
                                }
                                if let Some(v) = line_own_fn_ids.get_mut(li) {
                                    v.push(id);
                                }
                            }
                        }
                    }
                }
                out.push(cumulative);
            }
            out
        } else {
            vec![0.0; lines.len()]
        };

        // Widow/orphan control: pre-compute line heights for lookahead
        // ★ohnoikuji −1×3 finding (2026-06-18, S606 ATTEMPTED+REVERTED): the −1 paras
        // are pushed by a4-style (name "header") snapToGrid=0 list items 「（１）（２）（３）」
        // (MS Mincho 10.5, type=lines) that Word renders at the GRID pitch 18.0pt but
        // Oxi renders at NATURAL 13.5pt (line_height_inner passes None grid_pitch when
        // snap_to_grid=false) → Oxi over-fills the page by 3×4.5pt → the next para −1.
        // FIX (snapToGrid=0 body line height = grid pitch) was FALSIFIED on the gate:
        // it fixed ohnoikuji but REGRESSED 9a8e (1pg→2pg, 1.0→0.74)/bd90b00/roudoujoken
        // (−2 net) — Word does NOT grid-snap ALL body snapToGrid=0 paras. A
        // lines.len()==1 gate fixed roudoujoken (multi-line sg0) but 9a8e/bd90b00 STILL
        // regressed (single-line sg0 paras Word natural-izes). The discriminator (Word
        // grids ohnoikuji's SHORT EMBEDDED a4 sg0 run between gridded paras but
        // natural-izes 9a8e's large sg0 BLOCK) is the snapToGrid line-height precision
        // wall — deferred. line_height_inner:8218 is the natural-vs-grid site.
        // Reflow a gridded text paragraph against the region in which each
        // line will be placed. All line metrics are calculated below from the
        // resulting lines, rather than retained from a previous wrapping width.
        let mut region_line_widths: Vec<(f32, f32)> = word_fit_widths;
        if region_line_widths.is_empty() && std::env::var("OXI_GRID_REGION_REFLOW").is_ok()
            && num_columns > 1 && body_para_index.is_some()
            && grid_pitch.is_some_and(|p| p > 0.0) && para.style.snap_to_grid
            && !matches!(para.style.line_spacing_rule.as_deref(), Some("exact" | "atLeast"))
            && para.style.line_spacing.map_or(true, |v| (v - 1.0).abs() < 0.01)
            && !caps_active && s758_two_seg.is_none()
            && fragments.iter().all(|f| f.2.is_none()
                && !f.0.chars().any(|c| matches!(c, '\n' | '\r' | '\t' | '\u{FFFC}')))
        {
            let bands: Vec<_> = para.shapes.iter().filter_map(|shape| {
                if shape.wrap_type != Some(crate::ir::WrapType::Square) { return None; }
                let pos = shape.position.as_ref()?;
                let (rx, rw) = match pos.h_relative.as_deref() {
                    Some("page") => (0.0, page.size.width),
                    Some("margin") => (page.margin.left, page.size.width - page.margin.left - page.margin.right),
                    _ => return None,
                };
                let (ry, rh) = match pos.v_relative.as_deref() {
                    Some("page") => (0.0, page.size.height),
                    Some("margin") => (page.margin.top, page.size.height - page.margin.top - page.margin.bottom),
                    _ => return None,
                };
                let x = match pos.h_align.as_deref() {
                    Some("right") => rx + rw - shape.width,
                    Some("center") => rx + (rw - shape.width) * 0.5,
                    Some("left") => rx,
                    _ => rx + pos.x,
                };
                let y = match pos.v_align.as_deref() {
                    Some("top") => ry,
                    Some("bottom") => ry + rh - shape.height,
                    Some("center") => ry + (rh - shape.height) * 0.5,
                    _ => ry + pos.y,
                };
                Some((0usize, y, y + shape.height,
                    x - pos.dist_l.unwrap_or(9.0), x + shape.width + pos.dist_r.unwrap_or(9.0), false, BodyWrapPolicy::OBJECT))
            }).collect();
            if !bands.is_empty() {
                let mut remaining: Vec<_> = fragments.iter().map(|f|
                    (f.0.to_owned(), f.1.clone(), f.2.clone(), f.3, f.4)).collect();
                let mut planned = Vec::new();
                let mut widths = Vec::new();
                let mut y = cursor.cursor_y;
                let mut col = start_column;
                let mut valid = true;
                while !remaining.is_empty() {
                    let natural = self.natural_line_height_for_line(&lines[0], &para.style, para_font_size);
                    if y + natural > page_top + content_height && col + 1 < num_columns {
                        col += 1;
                        y = col_band_top;
                    }
                    if y + natural > page_top + content_height { valid = false; break; }
                    let (band, two, advance) = self.body_paragraph_wrap_bands(
                        para, page, &bands, 0, y, col_x_positions[col], content_width);
                    if two.is_some() || advance > 0.0 { valid = false; break; }
                    let (_, red, shift) = band.unwrap_or((0.0, 0.0, 0.0));
                    let refs: Vec<_> = remaining.iter().map(|f|
                        (f.0.as_str(), &f.1, f.2.clone(), f.3, f.4)).collect();
                    self.s1636_lane_shift.set(shift);
                    let mut broken = self.break_into_lines(&refs, (s758_wrap_full - floor_wrap_reduction(red)).max(s758_lane_minimum),
                        if planned.is_empty() { effective_first_indent } else { 0.0 },
                        &para.style, effective_char_pitch, effective_cw_ratio,
                        page.doc_grid_lines_and_chars, true,
                        matches!(para.alignment, Alignment::Justify | Alignment::Distribute),
                        page.doc_grid_no_type, para_has_lrpb, false, false);
                    if broken.is_empty() { valid = false; break; }
                    let next = broken.get(1).and_then(Line::source_start);
                    let line = broken.remove(0);
                    y += self.line_height_for_line(&line, &para.style, para_font_size,
                        para.style.snap_to_grid, grid_pitch, page.doc_grid_no_type);
                    planned.push(line);
                    widths.push((red, shift));
                    if let Some((ri, co)) = next {
                        let Some(index) = remaining.iter().position(|f| f.3 == ri && f.4 <= co
                            && co <= f.4 + f.0.chars().count()) else { valid = false; break; };
                        let count = co - remaining[index].4;
                        if index == 0 && count == 0 { valid = false; break; }
                        remaining.drain(..index);
                        remaining[0].0 = remaining[0].0.chars().skip(count).collect();
                        remaining[0].4 = co;
                    } else { remaining.clear(); }
                }
                if valid && !planned.is_empty() {
                    lines = planned;
                    region_line_widths = widths;
                }
            }
        }
        let s779_latin = (page.grid_line_pitch.is_none() || page.doc_grid_no_type)
            && !self.doc_body_has_real_cjk
            && std::env::var("OXI_S779_DISABLE").is_err();
        let header_inline_geometry = is_header_footer
            && (std::env::var_os("OXI_HEADER_INLINE_OBJECTS").is_some()
                || (para.style.line_spacing_rule.as_deref() == Some("exact")
                    && para.runs.iter().any(|r| r.style.inline_object_image.is_some())));
        let header_exact_inline = header_inline_geometry
            && para.style.line_spacing_rule.as_deref() == Some("exact");
        // Initial breaking and later band-exit reflow must use the same line
        // boxes, including font unions, inline objects, occupied ink and leading.
        let line_boxes = |lines: &[Line]| {
        let mut line_heights: Vec<f32> = lines
            .iter()
            .map(|line| {
                self.line_height_for_line(
                    line,
                    &para.style,
                    para_font_size,
                    para.style.snap_to_grid,
                    grid_pitch,
                    page.doc_grid_no_type,
                )
            })
            .collect();
        // Day 33 part 65 (2026-05-12): natural line heights (ascent+descent only)
        // for page-break threshold. Word allows grid-snap LEADING to extend
        // into bottom margin; only the text-occupying zone (ascent+descent)
        // must fit within content area. db9ca18 i=37 confirmed via COM:
        // line at y=758.25, grid line_h=18, line bottom=776.25 (5.25pt past
        // pgBot=771) — Word fits, while Oxi (using full line_h for break
        // check) rejected.
        let mut natural_line_heights: Vec<f32> = lines
            .iter()
            .map(|line| self.natural_line_height_for_line_inner(
                line, &para.style, para_font_size,
                (grid_pitch.is_none() || page.doc_grid_no_type)
                    && std::env::var("OXI_CJK_EXACT_BODY_CAPACITY").is_ok(),
            ))
            .collect();
        // S576 (2026-06-15): glyph-ink line heights (typo_sum*fs ≈ em) for the
        // page-bottom break-fit. The natural_line_heights above are the SPACING
        // box (win*83/64 = 1.297*em for CJK), ~3.2pt larger than the real glyph
        // ink — that over-count rejected page-bottom lines Word fits (their grid
        // leading hangs into the margin). See break_threshold below.
        // S779 (2026-07-11, opt-out OXI_S779_DISABLE): the DERIVED LM0-Latin
        // page-bottom rule (controlled sweep _pb_latin_gen.py, TNR 12pt: keep
        // iff baseline + win descent <= content_bottom, flip within 0.1pt) =
        // threshold win_ascent+win_descent (13.29 @12pt), NOT the em-ink
        // leniency (~11.3, ~2pt too lenient: nyserda p10 fit a line at base
        // 719.3 -> base+desc = 721.9 > 720 where Word widow-pushes the whole
        // 3-line paragraph -> the page-phase shifts behind its worst pages).
        // Scope: TRUE no-docGrid docs (page.grid_line_pitch None — gen2's
        // no-type docGrid keeps its S571-family rules) + pure-Latin
        // (!doc_body_has_real_cjk — JP no-grid docs keep their calibration).
        // Scope includes NO-TYPE docGrid docs (nyserda linePitch=299 no-type):
        // the probe re-run WITH that docGrid flips at the same cbot window —
        // the rule is grid-independent for non-snapping (no-type/LM0) Latin.
        // S827 (2026-07-13, opt-out OXI_S827_DISABLE): the S779 floor was a
        // MIS-TRANSLATION of the derived rule. The derivation said "keep iff
        // baseline + win_descent <= content_bottom"; from the line TOP the
        // baseline sits at hhea_ascent + lineGap, so the top-relative
        // threshold = hhea asc+desc+lineGap = the FULL hhea line (TNR 12pt
        // 13.799), NOT win_asc+win_desc (13.289). Re-derived directly
        // (_pb_latinbot_gen: no-grid Letter, TNR 12pt singles, 2tw bottom
        // sweep): Word's capacity flips EXACTLY at line_top + 13.799
        // (cbot 720.5 pushes / 720.6 keeps line47@706.746; the win model
        // predicts keeps at 720.1-720.5 — violated). Settings-absent and
        // compat-14 variants flip at the SAME point (compat is NOT a
        // discriminator here). The [13.29, 13.80) window let Oxi keep lines
        // Word pushes = the nyserda natural-flow over-pack class.
        let mut s779_win_heights: Vec<f32> = if s779_latin {
            let s827_hhea = std::env::var("OXI_S827_DISABLE").is_err();
            lines
                .iter()
                .map(|line| {
                    let mut mx: f32 = 0.0;
                    for f in &line.fragments {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                        let h = if s827_hhea {
                            m.natural_line_height_hhea(fs)
                        } else {
                            (m.win_ascent + m.win_descent) * fs
                        };
                        if h > mx {
                            mx = h;
                        }
                    }
                    mx
                })
                .collect()
        } else {
            Vec::new()
        };
        let mut ink_line_heights: Vec<f32> = lines
            .iter()
            .map(|line| self.ink_line_height_for_line(line, &para.style, para_font_size))
            .collect();

        // S1116 (2026-08-14, default ON, opt-out OXI_S1116_DISABLE): the
        // S851/S875/S1095 object line
        // target, for the cumulative raw basis below. The growth lands in
        // line_heights[] but a SINGLE-line paragraph's cursor advances by
        // raw_spaced_tw (the documented S612z three-sites trap) — S773 folds the
        // vector-group target and S795 the bullet-marker target, but the inline
        // PICTURE / w:object target had no sibling, so a `[text][picture]`
        // paragraph advanced by its text line alone.
        // A fixed line in a header or footer clips an inline drawing;
        // retaining that drawing as a run must not turn fixed spacing into
        // a minimum height. Use the shared fixed-line baseline and clipping
        // model for retained inline images under the default rules as well.
        // Keep the text-only line box before inline objects enlarge it.
        // Visual group placement below composes the same object with this box.
        let text_only_line_heights = line_heights.clone();
        let mut s1116_line0_target: f32 = 0.0;
        let mut story_image_leading = vec![0.0f32; lines.len()];
        // S851 (2026-07-14, opt-out OXI_S851_DISABLE): an inline w:object
        // form-field image (routed as a run-level inline object,
        // RunStyle.inline_object_image) contributes its box HEIGHT to its host
        // line — the FFFC width fragment (break_into_lines) reserves x but not
        // y. Grow every height array on the affected line so the object's box
        // (~18pt) reserves vertical space (render advance + page-break). Fires
        // ONLY when a fragment carries inline_object_image (EN form docs;
        // corpus JP has 0 inline OLE-less w:objects → byte-identical).
        if std::env::var("OXI_S851_DISABLE").is_err() {
            for (li, line) in lines.iter().enumerate() {
                let obj_h = line
                    .fragments
                    .iter()
                    .filter_map(|f| {
                        if f.style.inline_object_image.is_some()
                            || f.style.hr_rule.is_some()
                            || f.style.inline_math.is_some()
                        {
                            f.style.inline_object_extent.map(|(_, oh)| oh)
                        } else {
                            None
                        }
                    })
                    .fold(0.0f32, f32::max);
                if obj_h > 0.0 && !header_exact_inline {
                    // The object BOTTOM sits on the text baseline (emit), so the
                    // line height = object (all above baseline) + the line's text
                    // DESCENT (below baseline). Word's PA-form member-info gap
                    // 23.2 = line 20.2 (=18 + Arial-10 descent 2.2) + after 3.0;
                    // reserving only obj_h (18) under-counts ~2.2pt/line → d-1.
                    let descent = line
                        .fragments
                        .iter()
                        .filter(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                        .map(|f| {
                            let fs = f.style.font_size.unwrap_or(para_font_size);
                            self.metrics_for_text(&f.text, &f.style, &para.style)
                                .win_descent
                                * fs
                        })
                        .fold(0.0f32, f32::max);
                    // S875 (2026-07-16, default ON, opt-out OXI_S875_DISABLE):
                    // the line rule's EXTRA LEADING on an
                    // object line — the
                    // 2-arm model DERIVED by _pb_objline_gen.py (36 configs,
                    // obj {12,18,24} × line {240,276,360} × mixed/solo ×
                    // Arial/Calibri, Word COM):
                    //   MIXED: H = max(normal_line, obj + win_desc + extra)
                    //   SOLO:  H = obj + extra          (NO descent, NO clamp:
                    //          c12240s = 12.0 < the 13.43 Calibri text line)
                    // where extra = the AUTO line-rule multiple's addition
                    // (240→0, 276→+15%, 360→+50%). The sweep DECISIVELY
                    // rejects two rivals: solo=obj+raw-desc (REPORT2's read —
                    // the 240 solo row is obj EXACTLY) and desc×factor (360
                    // predicts 21.5 vs observed 24.0); the Calibri series pins
                    // WIN descent. The old obj+desc (v1) = the 240 degenerate
                    // form. ★The "sweep vs real doc" contradiction RESOLVED
                    // (the investigation's Word box histogram over all 40
                    // solo  lines): the doc is TRI-MODAL — 12 lines are
                    // o:hr RULE paragraphs (Word box ~8.25; the S852 basis
                    // 13.8 was +5.5/HR over), 17 are true 18pt object lines
                    // at obj+1.875 = EXACTLY this extra-leading model (no
                    // FORMTEXT discriminator exists), 10 carry space-before
                    // composites. The HR overage was partially COMPENSATING
                    // the missing object extra inside each Section — S875
                    // alone exposed it ({+1:2}); S875 + the S879 HR box fix
                    // ship together, with hr_rule lines EXCLUDED from the
                    // extra (a rule paragraph has no text leading to scale).
                    let hr_line = line.fragments.iter().any(|f| f.style.hr_rule.is_some());
                    let extra = if std::env::var("OXI_S875_DISABLE").is_err() && !hr_line {
                        // line_spacing under the auto rule is stored as the
                        // MULTIPLE itself (1.15, not 13.8pt — the [S875] trace:
                        // drug rows ls=Some(1.15), member rows ls=Some(1.0)).
                        let factor = if matches!(
                            para.style.line_spacing_rule.as_deref(),
                            None | Some("auto")
                        ) {
                            para.style.line_spacing.map(|l| l.max(1.0)).unwrap_or(1.0)
                        } else {
                            1.0
                        };
                        if factor > 1.0 {
                            let visible_natural = if header_inline_geometry {
                                line.fragments.iter()
                                    .filter(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                                    .map(|f| {
                                        let fs = f.style.font_size.unwrap_or(para_font_size);
                                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                                        if m.is_cjk_83_64_font() {
                                            self.line_height_inner(fs, Some(1.0), Some("auto"), m, false, None, false)
                                        } else {
                                            m.natural_line_height_hhea(fs)
                                        }
                                    })
                                    .fold(0.0_f32, f32::max)
                            } else { 0.0 };
                            if visible_natural > 0.0 {
                                visible_natural * (factor - 1.0)
                            } else {
                                line_heights[li] * (1.0 - 1.0 / factor)
                            }
                        } else {
                            0.0
                        }
                    } else {
                        0.0
                    };
                    if is_header_footer { story_image_leading[li] = extra; }
                    if std::env::var("OXI_DBG_S875").is_ok() {
                        let txt: String = line
                            .fragments
                            .iter()
                            .map(|f| f.text.as_str())
                            .collect::<String>()
                            .chars()
                            .take(20)
                            .collect();
                        eprintln!("[S875] li={} obj_h={:.1} desc={:.2} extra={:.2} lsr={:?} ls={:?} lh={:.2} txt={:?}",
                            li, obj_h, descent, extra, para.style.line_spacing_rule,
                            para.style.line_spacing, line_heights[li], txt);
                    }
                    // S1095 (2026-08-07, opt-out OXI_S1095_DISABLE): an inline
                    // object run that carries `w:position` is RAISED/LOWERED, so
                    // it does not sit with its bottom on the baseline. Word
                    // composes the line from the two sides independently:
                    //     ascent  = max(text_ascent,  obj_h - lower)
                    //     descent = max(text_descent, lower)
                    // (lower = -position, clamped at 0). Word render-truth on
                    // policies__0016b30b's `Construct an [x̄] control chart …`
                    // (obj 17.0, position -6 = 3pt lowered, TNR 12):
                    //   entry pitch 2.3.3→2.3.4  28.60 vs a plain 25.80 = +2.80
                    //       = max(11.20, 17.0-3.0) - 11.20   ✓ (ascent side)
                    //   exit advance line1→'limits:' 14.20 vs 13.799 = +0.404
                    //       = max(2.596, 3.0) - 2.596        ✓ (descent side)
                    //   line total = 14.0 + 3.0 = 17.0 = obj_h
                    // This UNIFIES the two rules already in the tree: with
                    // lower = 0 it is algebraically the old `obj_h + descent`
                    // (S875's MIXED, and its SOLO stays obj_h since a
                    // text-less line has descent = 0), and it reproduces
                    // S1066b's cell finding (`max(lh, obj_h)`) for a
                    // position=-6 object. Only a positioned object changes.
                    // Latin scope: the model was derived on a no-type-docGrid
                    // Latin document. The only CJK docs that carry a positioned
                    // inline object (3a4f / model, `w:position=-22` on a manual
                    // fraction) sit in a TYPED docGrid whose lines snap to whole
                    // grid cells and whose paragraph spacing does not decompose
                    // into the same ascent/descent sum (model p29 measures an
                    // entry pitch of +26.79 and an exit of +27.24 against a
                    // 29.25pt object) — that stack stays on its calibration.
                    let target = if !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S1095_DISABLE").is_err()
                    {
                        let lower = -line
                            .fragments
                            .iter()
                            .filter(|f| {
                                f.style.inline_object_image.is_some()
                                    || f.style.hr_rule.is_some()
                                    || f.style.inline_math.is_some()
                            })
                            .filter_map(|f| f.style.position)
                            .fold(0.0f32, f32::min);
                        // S1252: an inline maths box's own DESCENT is the
                        // `lower` of this composition — Word's `2π/3` arm grows
                        // the line +2.28 above and +2.76 below the plain box,
                        // i.e. the two sides compose independently.
                        let math_lower = line
                            .fragments
                            .iter()
                            .filter_map(|f| {
                                f.style.inline_math.as_ref().map(|mb| {
                                    let mfs = f.style.font_size.unwrap_or(para_font_size);
                                    crate::layout::math::inline_math_ink(mb, mfs).2
                                })
                            })
                            .fold(0.0f32, f32::max);
                        let lower = lower.max(0.0).max(math_lower);
                        // Picture positions belong to the object box, not its placeholder font.
                        let picture_only = line.fragments.iter().any(|f| f.style.inline_object_image.is_some())
                            && !line.fragments.iter().any(|f| f.style.inline_math.is_some() || f.style.hr_rule.is_some());
                        // Compose a story picture with the unspaced text box.
                        // The line's additional leading is added once below,
                        // after combining the text and picture on each side
                        // of their shared baseline.
                        let text_box = if is_header_footer && picture_only {
                            (line_heights[li] - extra).max(0.0)
                        } else {
                            line_heights[li]
                        };
                        let text_asc = if descent > 0.0 {
                            (text_box - descent).max(0.0)
                        } else {
                            0.0
                        };
                        let object_ascent = if picture_only {
                            line.fragments.iter().filter(|f| f.style.inline_object_image.is_some())
                                .filter_map(|f| f.style.inline_object_extent.map(|(_, h)|
                                    (h + f.style.position.unwrap_or(0.0)).max(0.0)))
                                .fold(0.0_f32, f32::max)
                        } else { obj_h - lower };
                        object_ascent.max(text_asc) + lower.max(descent) + extra
                    } else {
                        obj_h + descent + extra
                    };
                    let effect_bottom = if !is_header_footer && !self.doc_body_has_real_cjk
                        && descent == 0.0
                        && std::env::var("OXI_BODY_IMAGE_EFFECT_EXTENT_DISABLE").is_err() {
                        line.fragments.iter().filter_map(|f| f.style.inline_object_image.as_ref())
                            .map(|im| im.effect_extent_b.max(0.0) + im.effect_extent_t.max(0.0)).fold(0.0_f32, f32::max)
                    } else { 0.0 };
                    let target = target + effect_bottom;
                    // Grid leading changes line advance, not the occupied object box.
                    let natural_target = target;
                    // S1418 (2026-09-15, default ON, opt-out OXI_INLINE_IMAGE_GRID_DISABLE):
                    // the checkpoint's opt-in promoted. In a typed docGrid a line
                    // holding an inline picture takes whole cells: technical__90f5b9d4
                    // (4 ideographic spaces + a 244.07pt screenshot, effectExtent t 1.5
                    // b 1.95) spans 252 = 14 x 18 in Word (125.25 -> 147.0 -> 395.25 ->
                    // 413.25), Oxi 244.18; 98ef3583 is the same class. Both pass
                    // without the saved page-break markers once the line is quantized.
                    // S1596 (2026-09-29, default ON, opt-out OXI_S1596_DISABLE): an
                    // inline oMath line takes whole grid cells the same way.
                    // `_pb_cjkmath_gen.py` (educational__20d9968b slice, `lines`
                    // grid 18pt): marker+test pairs span 36 for あ+a/b, x^2, x_i
                    // (1 cell) and 54 for b/sqrt(a^2+b^2) and (a/b)/(c/d) (2 cells);
                    // Oxi left the math line at its natural 18.68.
                    let s1596_math = std::env::var_os("OXI_S1596_DISABLE").is_none()
                        && line.fragments.iter().any(|f| f.style.inline_math.is_some());
                    let target = if std::env::var_os("OXI_INLINE_IMAGE_GRID_DISABLE").is_none()
                        && !page.doc_grid_no_type && para.style.snap_to_grid
                        && (line.fragments.iter().any(|f| f.style.inline_object_image.is_some()) || s1596_math)
                    {
                        // Raised/lowered objects contribute to opposite sides
                        // of the baseline. Quantize that occupied box, rather
                        // than treating a lowered object's full height as ascent.
                        let object_ascent = line.fragments.iter()
                            .filter(|f| f.style.inline_object_image.is_some()
                                || f.style.hr_rule.is_some() || f.style.inline_math.is_some())
                            .filter_map(|f| f.style.inline_object_extent.map(|(_, h)|
                                (h + f.style.position.unwrap_or(0.0)).max(0.0)))
                            .fold(0.0_f32, f32::max);
                        let object_descent = line.fragments.iter()
                            .filter(|f| f.style.inline_object_image.is_some()
                                || f.style.hr_rule.is_some() || f.style.inline_math.is_some())
                            .map(|f| (-f.style.position.unwrap_or(0.0)).max(0.0))
                            .fold(descent, f32::max);
                        let occupied = object_ascent + object_descent + extra + effect_bottom;
                        // S1611 (2026-09-29, default ON, opt-out OXI_S1611_DISABLE): a
                        // maths line counts grid cells from the PER-GLYPH ink of the
                        // maths (`math::inline_math_ink_extent`: fraction shifts +
                        // numerator/denominator ink, radical depth = radicand depth,
                        // cramped superscripts, upright `m:sty p` glyphs), not its
                        // layout box, which gives every letter 0.7em. Blocks share the
                        // baseline: tallest top + deepest bottom + 1.37 <= n x pitch.
                        // `_pb_cjkmath_gen.py` 22 fraction arms + blind-G JA
                        // educational__20d9968b: one cell up to 16.60 (1/x, a/b,
                        // a/sqrt(a^2+b^2)), two from 16.66 (A/b, b/.., a^2/.., a/g) ->
                        // margin in (1.34, 1.40].
                        let s1611 = std::env::var_os("OXI_S1611_DISABLE").is_none()
                            && s1596_math
                            && !line.fragments.iter().any(|f| f.style.inline_object_image.is_some()
                                || f.style.hr_rule.is_some());
                        if s1611 {
                            let m: f32 = std::env::var("OXI_S1611_M").ok()
                                .and_then(|v| v.parse().ok()).unwrap_or(1.37);
                            // Blocks on one line share the baseline: the line
                            // needs the tallest top plus the deepest bottom.
                            let (ink_t, ink_b) = line.fragments.iter()
                                .filter_map(|f| f.style.inline_math.as_ref().map(|mb| {
                                    let mfs = f.style.font_size.unwrap_or(para_font_size);
                                    let (t, b) = crate::layout::math::inline_math_ink_extent(mb, mfs);
                                    if std::env::var_os("OXI_DBG_S1611").is_some() {
                                        eprintln!("[S1611-INK] fs={:.2} top={:.2} bot={:.2}", mfs, t, b);
                                    }
                                    (t, b)
                                }))
                                .fold((0.0f32, 0.0f32), |(a, d), (t, b)| (a.max(t), d.max(b)));
                            let ink = ink_t + ink_b;
                            let font_boxes=line.fragments.iter().filter_map(|f|f.style.inline_math.as_ref().and_then(|mb| {
                                crate::layout::math::inline_math_typographic_extent(mb,f.style.font_size.unwrap_or(para_font_size))
                            })).fold(None,|boxes:Option<(f32,f32)>,(a,d)|Some(boxes.map_or((a,d),|(ba,bd)|(ba.max(a),bd.max(d)))));
                            let occ=if let Some((a,d))=font_boxes {
                                let (ha,hd)=line.fragments.iter().filter(|f|f.style.inline_math.is_none()
                                    && f.style.inline_object_image.is_none() && f.style.hr_rule.is_none())
                                    .map(|f|self.metrics_for_text(&f.text,&f.style,&para.style)
                                        .design_font_box_pt(f.style.font_size.unwrap_or(para_font_size),true))
                                    .fold((0.0_f32,0.0_f32),|(a,d),(fa,fd)|(a.max(fa),d.max(fd)));
                                a.max(ink_t).max(ha)+d.max(ink_b).max(hd)
                            }else {(ink+m).max(1.0)};
                            // Replacing the math object's estimated grid box
                            // must retain the typographic minimum of its host
                            // text. Measure that text with the ordinary line
                            // policy, excluding the math placeholder itself.
                            let text_line = Line {
                                fragments: line.fragments.iter()
                                    .filter(|f| f.style.inline_math.is_none()).cloned().collect(),
                                empty_break_style: line.empty_break_style.clone(),
                                whitespace_paragraph: line.whitespace_paragraph,
                                ..Line::default()
                            };
                            let text_floor = if text_line.fragments.is_empty() { 0.0 } else {
                                self.line_height_for_line(&text_line, &para.style, para_font_size,
                                    para.style.snap_to_grid, grid_pitch, page.doc_grid_no_type)
                            };
                            grid_pitch.filter(|p| *p > 0.1)
                                .map_or(target, |pitch| ((occ / pitch).ceil().max(1.0) * pitch).max(text_floor))
                        } else {
                        grid_pitch.filter(|p| *p > 0.1)
                            .map_or(target, |pitch| ((occupied / pitch).ceil() * pitch).max(target))
                        }
                    } else { target };
                    if li == 0 && target > line_heights[0] {
                        s1116_line0_target = target;
                    }
                    // S1611: the ink count REPLACES the earlier grid snap of the
                    // maths layout box (which had already taken two cells).
                    let s1611_set = std::env::var_os("OXI_S1611_DISABLE").is_none()
                        && std::env::var_os("OXI_S1596_DISABLE").is_none()
                        && std::env::var_os("OXI_INLINE_IMAGE_GRID_DISABLE").is_none()
                        && !page.doc_grid_no_type && para.style.snap_to_grid
                        && grid_pitch.map_or(false, |p| p > 0.1)
                        && line.fragments.iter().any(|f| f.style.inline_math.is_some())
                        && !line.fragments.iter().any(|f| f.style.inline_object_image.is_some()
                            || f.style.hr_rule.is_some());
                    if s1611_set {
                        if std::env::var_os("OXI_DBG_S1611").is_some() {
                            eprintln!("[S1611] line {} {:.2} -> {:.2}", li, line_heights[li], target);
                        }
                        line_heights[li] = target;
                    } else if target > line_heights[li] {
                        line_heights[li] = target;
                    }
                    if natural_target > natural_line_heights[li] {
                        natural_line_heights[li] = natural_target;
                    }
                    if natural_target > ink_line_heights[li] {
                        ink_line_heights[li] = natural_target;
                    }
                    if li < s779_win_heights.len() && natural_target > s779_win_heights[li] {
                        s779_win_heights[li] = natural_target;
                    }
                }
            }
        }

            (line_heights, natural_line_heights, s779_win_heights,
                ink_line_heights, text_only_line_heights, s1116_line0_target, story_image_leading)
        };
        let (mut line_heights, mut natural_line_heights, mut s779_win_heights,
            mut ink_line_heights, mut text_only_line_heights, s1116_line0_target,
            mut story_image_leading) = line_boxes(&lines);

        // S689 (2026-06-29, SHIPPED default ON, opt-out OXI_S689_DISABLE): a list
        // paragraph whose numbering marker is a SYMBOL-FONT bullet (\u{F0B7}) renders
        // its line ~0.67pt TALLER in Word/LibreOffice (Symbol natural ratio 1.2251 >
        // Cambria 1.1724; at 11pt×1.15 = 15.49 vs 14.83). Oxi omits the marker glyph's
        // font from the line-height max → the bullet line is too short → cumulative
        // drift (the gen2 family's BIGGEST residual; Libra — matching the word_png at
        // ~0.98 — proved it FIXABLE, reframing the deferred "render-anchor wall"; the
        // body 11pt is already correct via S671). Bump line 0 (the bullet line) by the
        // Symbol/body natural-ratio so line_height = marker_natural × factor. SCOPED to
        // the S671 no-type-docGrid Latin path (line_height = natural_hhea × factor, so
        // the ratio scale transfers the factor cleanly; s671_fine routes the advance to
        // line_heights[0]). ★GATE: corpus SSIM A/B net +1.4326 (page-sum), 33 improved
        // (the whole gen2 Latin family — gen2_064 0.9167→0.9733/+0.0566, gen2_041
        // +0.109, gen2_067 +0.088 …) / 1 negligible (test_lists −0.0057). ALL page
        // counts STABLE (N/N) → no pagination shift; only 34 synthetic gen2/gen/test
        // docs byte-change → every Phase-1 regulation doc BYTE-IDENTICAL → Phase-1
        // preserved by construction (S689 fires only on no-type-docGrid + Latin-line0
        // + Symbol bullet, which the candidate scan confirms is gen2/gen/test ONLY;
        // CJK gen2 excluded by all_latin). lib 142/0/6. Tools: _bugfind_rank.py,
        // Libra-PDF per-line baselines. See [[gen2_vertical_drift]].
        // S785 note (2026-07-11): the NO-GRID (LM0) Symbol-marker line height
        // (nyserda bullets: Word 14.64 = Symbol 1.2251×12 device-snapped) is
        // handled by the REGISTRY 'Symbol' metrics entry (font_metrics_compact
        // .json, added 2026-07-11) via the normal marker-metrics path — an
        // explicit S689-style bump here proved redundant (identical output).
        // S821 shared: the marker line-0 growth target (S689 ratio model /
        // S795 component model, max) — attributed to the paragraph ENTRY
        // (cursor) instead of the line's own advance when S821 is on.
        let mut s821_growth_target: f32 = 0.0;
        let s821_entry_attribution = std::env::var("OXI_S821_DISABLE").is_err();
        if std::env::var("OXI_S689_DISABLE").is_err()
            && page.doc_grid_no_type
            && !lines.is_empty()
            && !lines[0].fragments.is_empty()
            && para
                .style
                .list_marker
                .as_deref()
                .map_or(false, |m| m.contains('\u{F0B7}'))
        {
            // body natural-hhea of line 0 (max over fragments), and confirm all-Latin
            let mut body_nat: f32 = 0.0;
            let mut body_desc: f32 = 0.0;
            let mut body_box_asc: f32 = 0.0;
            let mut all_latin = true;
            for f in &lines[0].fragments {
                let fs = f.style.font_size.unwrap_or(para_font_size);
                let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                if m.is_cjk_83_64_font() {
                    all_latin = false;
                }
                let n = m.natural_line_height_hhea(fs);
                if n > body_nat {
                    body_nat = n;
                }
                let d = m.win_descent * fs;
                if d > body_desc {
                    body_desc = d;
                }
                let a = m.win_ascent * fs;
                // The box's TOP is the text ascent plus its external leading
                // (hhea natural − win sum), which is where a taller marker
                // starts to overflow.
                let boxa = a + (n - (m.win_ascent + m.win_descent) * fs).max(0.0);
                if boxa > body_box_asc {
                    body_box_asc = boxa;
                }
            }
            // S821b (2026-07-13): the marker line target is the COMPONENT
            // model (Symbol asc + TEXT desc, S820b — differential probe
            // 13.386 = (2059+434)/2048×11), not the win-SUM ratio (1.2251
            // → 13.476, +0.086/bullet = the uklocal p6/p7 residual slope).
            // S1263 (2026-08-30, default ON, opt-out OXI_S1263_DISABLE): the
            // Symbol marker's ascent is measured at the MARKER's own declared
            // size, not the paragraph's. `w:lvl` carries its own `<w:sz>` and
            // the parser already resolves it into `list_marker_size` (S1037
            // uses it for the marker's width); only this height rule still read
            // the body size.
            // WITNESS `reports__0018715b4769984f` (EN Phase-1 FAIL 0.9560, whose
            // only feature is footnotes): its recommendation list is `numId=36`
            // -> a Symbol U+F0B7 bullet declared at `w:sz=20` (10pt) inside
            // 11pt Calibri body text. Word sets consecutive items 22.44pt apart
            // -- exactly `11 x 1.2207 x 1.0792 + 8 = 22.49`, i.e. the plain body
            // line plus the docDefaults 8pt after, with NO marker growth. Oxi
            // set them 23.07 apart because it measured a 10pt marker as if it
            // were 11pt:
            //     11pt: 1.00537 x 11    = 11.06 > body box ascent 10.68 -> +0.38
            //     10pt: 1.00537 x 10    = 10.05 < 10.68                 -> +0.00
            // The overflow is the whole point of the rule, so feeding it the
            // wrong size does not merely scale it -- it invents one.
            let s1263_marker_fs = if std::env::var("OXI_S1263_DISABLE").is_err() {
                para.style.list_marker_size.unwrap_or(para_font_size)
            } else {
                para_font_size
            };
            const SYM_ASC_R: f32 = 2059.0 / 2048.0;
            const SYMBOL_RATIO: f32 = 1.2251; // legacy win-sum model (A/B)
            let marker_nat = if s821_entry_attribution {
                SYM_ASC_R * s1263_marker_fs + body_desc
            } else {
                SYMBOL_RATIO * para_font_size
            };
            if all_latin && body_nat > 0.0 && marker_nat > body_nat {
                let scale = marker_nat / body_nat;
                // S947 (2026-07-19, opt-out OXI_S947_DISABLE): the scale model
                // applies only to body-proportional lines (rule None/auto,
                // line = nat x factor — the gen2 derivation domain). An
                // atLeast FLOOR line (NDIS bullets: line=300 = 15.0 >
                // marker component 14.01) absorbs the marker — multiplying
                // the floor by the component ratio grew every bullet +0.88
                // (Word inter-item 20.04 = plain, measured). atLeast compares
                // the raw component against the floor; exact never grows.
                let s947 = std::env::var("OXI_S947_DISABLE").is_err();
                // S1112 (2026-08-13, SHIPPED default-ON, opt-out
                // OXI_S1112_DISABLE, with S1091 + S1074 + S1113 + S1114 — see
                // the bundle note at the S1091 site): the marker's
                // growth is the ASCENT OVERFLOW, ADDED to the multiplied text
                // line — it is NOT itself multiplied by the line factor. Word
                // truth (_pb_bullet_gen.py, 8 numbering arms x 20 items, span
                // precision ±0.04): Arial 12 + full-size Symbol bullet at
                // line=276 measures 16.658 = 15.869 + (12.064 − 11.256), where
                // the old ratio/component model multiplies the whole marker
                // natural and gives 16.799 (+0.14/bullet — the policies__000f7115
                // page-26 excess that cancelled its empty-paragraph deficit and
                // held S1091/S1074 opt-in). The two models COINCIDE at factor
                // 1.0, which is where S820b/S821b were derived, so the earlier
                // probes stay satisfied: uklocalspending Arial 11 + Symbol =
                // 12.649 + 0.741 = 13.390 vs the recorded differential 13.386.
                let s1112 = std::env::var("OXI_S1112_DISABLE").is_err();
                let marker_overflow =
                    (SYM_ASC_R * s1263_marker_fs - body_box_asc).max(0.0);
                let s689_target = match para.style.line_spacing_rule.as_deref() {
                    Some("exact") if s947 => line_heights[0],
                    Some("atLeast") if s947 => line_heights[0].max(marker_nat),
                    _ if s1112 => line_heights[0] + marker_overflow,
                    _ => line_heights[0] * scale,
                };
                if s821_entry_attribution {
                    s821_growth_target = s821_growth_target.max(s689_target);
                } else {
                    line_heights[0] *= scale;
                    natural_line_heights[0] *= scale;
                    ink_line_heights[0] *= scale;
                }
            }
        }

        // S795 (2026-07-12, opt-out OXI_S795_DISABLE): NO-GRID (LM0) Symbol-bullet
        // marker line height — Word's line model is a per-COMPONENT GDI max:
        // line = max(tmAscent) + max(tmDescent) + max(tmExternalLeading) over the
        // text fonts AND the marker font (tm* = OS/2 win metrics; ext = hhea
        // natural − win sum). A Symbol marker (winAsc 2059/2048 = 1.0054em) grows
        // the line's ASCENT while the text font keeps its larger DESCENT (Calibri
        // 550 > Symbol 450) — the sum-max model (S689's flat 1.2251 ratio) cannot
        // express this: Calibri+Symbol = (2059+550)/2048 = 1.2739em, NOT 1.2251.
        // DERIVED from a controlled probe (fw_probe2, Word PDF baselines):
        // plain Calibri 13.44 / Calibri+Symbol 14.02 (= 1.2739×11.04 device-
        // rounded) / Humnst-sub+Symbol 13.51 (= 1.2251 — the sub font's desc <
        // 450); ukframework's real list pitch 13.92-14.04 matches, and its
        // line=259 bullets measure 15.12 = component-max × 1.0792 → the growth
        // applies to LINE 0 ONLY and composes MULTIPLICATIVELY with the auto
        // line-spacing factor. Scope: grid_pitch None (covers BOTH pure no-grid
        // AND no-type docGrid — ukframework carries the default
        // <w:docGrid linePitch=360>, so the no-type case must be included;
        // S689's sum-ratio bump runs first and this rule supersedes it only
        // when the component-max exceeds it, i.e. text desc > Symbol desc —
        // for Cambria/TNR the two are degenerate within 0.03pt, probe-
        // confirmed), Latin doc (!doc_body_has_real_cjk → JP corpus
        // byte-identical by construction), auto line rule only.
        let mut s795_line0_target: f32 = 0.0;
        // S821: the marker growth advanced at the paragraph ENTRY (cursor),
        // keeping the line's own advance plain.
        let mut s795_entry_extra: f32 = 0.0;
        if std::env::var("OXI_DBG795").is_ok() && para.style.list_marker.is_some() {
            eprintln!(
                "[DBG795] marker={:?} gp_none={} no_type={} cjk={} rule={:?} ls={:?} nlines={}",
                para.style
                    .list_marker
                    .as_deref()
                    .map(|m| m.chars().take(3).collect::<String>()),
                grid_pitch.is_none(),
                page.doc_grid_no_type,
                self.doc_body_has_real_cjk,
                para.style.line_spacing_rule,
                para.style.line_spacing,
                lines.len()
            );
        }
        if std::env::var("OXI_S795_DISABLE").is_err()
            && grid_pitch.is_none()
            && !self.doc_body_has_real_cjk
            && !lines.is_empty()
            && !lines[0].fragments.is_empty()
            && !matches!(
                para.style.line_spacing_rule.as_deref(),
                Some("exact")
            )
            && (para
                .style
                .list_marker
                .as_deref()
                .map_or(false, |m| m.contains('\u{F0B7}'))
                || (std::env::var("OXI_S1037_DISABLE").is_err()
                    && s1037_marker_style(para).is_some()))
        {
            let mut asc: f32 = 0.0;
            let mut desc: f32 = 0.0;
            // S820 (2026-07-13): the external-leading term comes from the
            // TALLEST (max win-sum) component, NOT the max ext across
            // components — uklocalspending bullets (Arial 11 text ext 67/2048
            // + full-size Symbol marker ext 0): Word pitch 13.44 =
            // (2059+450)/2048×11 with NO ext (the Symbol marker is tallest);
            // max(ext) gave 13.84 = +0.4/bullet → ~+6pt over a 16-bullet page
            // → the wp6/7/15 page-bottom pushes. Degenerate for ukframework
            // (Calibri tallest AND ext-max → identical); fw_probe2's
            // Calibri+full-Symbol 14.02 also matches (argmax=Symbol → ext 0 →
            // 13.95 ≈ 14.02 device) where max(ext) gave 14.26.
            let mut best_sum: f32 = 0.0;
            let mut ext: f32 = 0.0;
            let mut marker_growth_ok = true;
            for f in &lines[0].fragments {
                let fs = f.style.font_size.unwrap_or(para_font_size);
                let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                asc = asc.max(m.win_ascent * fs);
                desc = desc.max(m.win_descent * fs);
                let win_sum = (m.win_ascent + m.win_descent) * fs;
                let fext = (m.natural_line_height_hhea(fs) - win_sum).max(0.0);
                if win_sum > best_sum {
                    best_sum = win_sum;
                    ext = fext;
                }
            }
            // Microsoft Symbol (symbol.ttf): win 2059/450, upm 2048, lineGap 0.
            const SYM_ASC: f32 = 2059.0 / 2048.0;
            const SYM_DESC: f32 = 450.0 / 2048.0;
            // S801b: the marker glyph is sized by the LEVEL rPr's w:sz when
            // declared (ukframework bullet levels = Symbol + sz=20: the 10pt
            // marker's ascent 10.05 < Calibri@11's 10.47 → NO line growth,
            // Word pitch 13.44; sizing it at para_font_size wrongly grew the
            // line to 14.01).
            let marker_fs = para.style.list_marker_size.unwrap_or(para_font_size);
            // S820b (2026-07-13): the Symbol marker contributes its ASCENT
            // only — the line's descent stays the TEXT fonts' max, and the
            // external leading drops when the marker drives the height.
            // Differential probe (pbb_ 4/8/16/24 bullets, TARGET ink deltas —
            // ink offsets cancel): Arial 11 + full-size Symbol advance =
            // 13.380/13.399/13.380 ≈ 13.386 = (2059 + 434)/2048×11 = Symbol
            // asc + ARIAL desc (Symbol's own winDesc 450 NOT used; win-sum
            // max gave 13.476, max(ext) 13.84 = the old wp6/7/15 drift).
            // fw_probe2 Calibri+Symbol 14.02 = 11.10 + Calibri desc 2.90 ✓.
            // S1037 (2026-07-29): a NON-Symbol marker takes its ascent from
            // its own RESOLVED style (level rPr > direct paragraph-mark rPr >
            // paragraph style) - the source the 22-arm Word probe identified.
            // Same ascent-only topology as S820b/S821: a tall marker grows the
            // pitch INTO the line (dUp +10.44 for Courier-24 over Arial-10) and
            // leaves the outgoing pitch untouched (dDown +0.00), for every
            // suffix (tab/space/nothing) and both hanging/firstLine.
            let is_symbol_marker = para
                .style
                .list_marker
                .as_deref()
                .map_or(false, |m| m.contains('\u{F0B7}'));
            let marker_asc = if is_symbol_marker {
                SYM_ASC * marker_fs
            } else {
                match s1037_marker_style(para) {
                    Some(ms) => {
                        let mfs = ms.font_size.unwrap_or(marker_fs);
                        // S1614 (2026-09-30, default ON, opt-out OXI_S1614_DISABLE):
                        // a NUMBERING MARKER in Arial Unicode MS keeps AUM's own
                        // ascent (1.332em) even with the face absent -- S1564's
                        // Arial box is for text RUNS. `_pb_ckl_heading_gen.py`
                        // (blind-G EN policies__0097fbf2, ChecklistLevel1 headings,
                        // marker AUM hint=eastAsia, Calibri-Bold 10 text): Word row
                        // pitch 17.25 with the AUM marker, 12.75 with no numbering
                        // and 12.75 with the level font set to Arial; a 12pt marker
                        // 19.5. The PDF draws the digit in Calibri-Bold, but the line
                        // grows by AUM's ascent overflow (10pt +3.8, 12pt +6.46).
                        let aum_marker = std::env::var_os("OXI_S1614_DISABLE").is_none()
                            && !self.doc_body_has_real_cjk
                            && self.resolve_font_family(ms, &para.style) == Some("Arial Unicode MS");
                        let m = if aum_marker {
                            self.registry.get_with_style(
                                "Arial Unicode MS Latin",
                                self.resolve_bold(ms, &para.style),
                                self.resolve_italic(ms, &para.style),
                            )
                        } else {
                            self.metrics_for(ms, &para.style)
                        };
                        m.win_ascent * mfs
                    }
                    None => 0.0,
                }
            };
            let _ = SYM_DESC;
            let _ = best_sum;
            // The TEXT fonts' own external leading, kept for the S1112 box top
            // (S820 drops `ext` from the legacy target when the marker wins).
            let text_ext = ext;
            if marker_asc > asc {
                ext = 0.0;
            }
            // Self-limiting: a marker no taller than the body contributes
            // nothing, so an ordinary list whose marker shares the body font
            // stays byte-identical.
            if !is_symbol_marker && marker_asc <= asc {
                marker_growth_ok = false;
            }
            let target_nat = asc.max(marker_asc) + desc + ext;
            let factor = para.style.line_spacing.unwrap_or(1.0);
            let factor = if factor > 0.0 { factor } else { 1.0 };
            // S1112: the marker overflows the TEXT box's top (ascent + the
            // text's own external leading) and that overflow is added to the
            // multiplied text line rather than multiplied with it. `ext` is
            // zeroed above when the marker drives the height (S820), so the
            // text's leading is recovered from `text_ext` for the box top.
            let target = if para.style.line_spacing_rule.as_deref() == Some("atLeast") {
                line_heights[0] + LayoutEngine::minimum_marker_entry_overflow(
                    (marker_asc - asc - text_ext).max(0.0),
                    asc + desc + text_ext, line_heights[0], Some("atLeast"))
            } else if std::env::var("OXI_S1112_DISABLE").is_err() {
                (asc + desc + text_ext) * factor + (marker_asc - asc - text_ext).max(0.0)
            } else {
                target_nat * factor
            };
            if target > line_heights[0] && marker_growth_ok {
                // S821 (2026-07-13, opt-out OXI_S821_DISABLE): the marker
                // growth belongs to the pitch INTO the bullet line, not out
                // of it — junction probe matrix (pk_, real F0B7 markers):
                // entry plain→bullet 13.44 / interior 13.44 / EXIT
                // bullet→plain 12.60 = PLAIN (marker irrelevant); spacing
                // adds on top (kB12 24.6 = 12.6+12, kEB12 25.44 = 13.44+12).
                // Word model: pitch = prev_TEXT_desc + next_max_asc(incl.
                // marker) + next_ext(tallest). Attribution: advance the
                // CURSOR by the growth at the paragraph ENTRY and keep the
                // line's own advance PLAIN (the old line_heights[0] bump put
                // the growth on the EXIT → +0.74 after every list followed
                // by spacing — uklocal p6/p7 page-bottom pushes). The pbb_
                // "exit=13.704" that suggested otherwise was the '•' GLYPH
                // ink (1.1 above the text ink) — measure bullet junctions
                // text-ink to text-ink.
                if s821_entry_attribution {
                    s821_growth_target = s821_growth_target.max(target);
                } else {
                    line_heights[0] = target;
                    if natural_line_heights[0] < target {
                        natural_line_heights[0] = target;
                    }
                    if ink_line_heights[0] < target {
                        ink_line_heights[0] = target;
                    }
                    s795_line0_target = target;
                }
            }
        }
        // Compare marker and text box tops above their shared baseline.
        // East Asian natural boxes distribute additional leading on both sides;
        // a Latin box already has its external leading above the glyph ascent.
        // Only the marker's overflow grows the line; its descent is not text.
        // Typed grids quantize the expanded box and exact rules remain fixed.
        if self.doc_body_has_real_cjk && !lines.is_empty()
            && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
            && (grid_pitch.is_none() || !page.doc_grid_no_type)
        {
            if let Some(marker_style) = s1037_marker_style(para) {
                let marker_fs = self.resolve_font_size(marker_style, &para.style);
                let marker_metrics = &*self.metrics_for(marker_style, &para.style);
                let box_top = |m: &FontMetrics, fs: f32| {
                    let win = (m.win_ascent + m.win_descent) * fs;
                    let leading = if m.is_cjk_83_64_font() {
                        (LayoutEngine::s1367_cjk_box(m, fs) - win).max(0.0) * 0.5
                    } else {
                        (m.natural_line_height_hhea(fs) - win).max(0.0)
                    };
                    m.win_ascent * fs + leading
                };
                let body_ascent = lines[0].fragments.iter().filter(|f| !f.text.trim().is_empty())
                    .map(|f| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        box_top(&self.metrics_for_text(&f.text, &f.style, &para.style), fs)
                    }).fold(0.0f32, f32::max);
                let body_descent = lines[0].fragments.iter().filter(|f| !f.text.trim().is_empty())
                    .map(|f| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                        let total = if m.is_cjk_83_64_font() {
                            LayoutEngine::s1367_cjk_box(m, fs)
                        } else { m.natural_line_height_hhea(fs) };
                        (total - box_top(m, fs)).max(0.0)
                    }).fold(0.0f32, f32::max);
                let extra = (box_top(marker_metrics, marker_fs) - body_ascent).max(0.0);
                if body_ascent > 0.0 && extra > 0.0 {
                    // Quantize the union of actual boxes, before device rounding.
                    natural_line_heights[0] = (natural_line_heights[0] + extra)
                        .max(body_ascent + body_descent + extra);
                    let factor = para.style.line_spacing.unwrap_or(1.0);
                    let target = if para.style.snap_to_grid {
                        grid_pitch.filter(|pitch| *pitch > 0.0).map(|pitch| {
                            let cells = ((natural_line_heights[0] * 20.0).round() / (pitch * 20.0)).ceil().max(1.0);
                            pitch * cells.max(factor)
                        }).unwrap_or(line_heights[0] + extra * factor)
                    } else {
                        line_heights[0] + extra * factor
                    };
                    line_heights[0] = line_heights[0].max(target);
                    s795_line0_target = s795_line0_target.max(target);
                }
            }
        }
        if s821_growth_target > line_heights.first().copied().unwrap_or(0.0) {
            s795_entry_extra = s821_growth_target - line_heights[0];
            cursor.advance(s795_entry_extra);
        }
        let _ = s795_entry_extra;

        // S773: the line-0 target also feeds the single-LM0 cumulative advance
        // basis (raw_spaced_tw) below — a SINGLE-line paragraph's cursor
        // advance uses that basis, not line_heights[0] (the documented
        // S612z three-sites trap; without it the title-para bump was a no-op
        // once S774 made the title one line).
        let mut s773_line0_target: f32 = 0.0;
        // S837 (2026-07-14): the S773 bump fired on line 0 — its cy, for the
        // baseline-alignment override at emit (0 = not fired).
        let mut s837_fired_cy: f32 = 0.0;
        // S773 (2026-07-10, opt-out OXI_S773_DISABLE): an INLINE visual drawing
        // (wp:inline wpg/wps vector group without textbox text — hmrc's crown
        // logo, checkbox rows, black rules) grows its host LINE like an inline
        // object. COM probe (wpg_probe.docx, hmrc's actual drawings injected
        // into 12pt hosts): text line = cy + ~0.25×fs (cy16.5→19.5, cy60→63.0);
        // an object-only paragraph's line = cy EXACTLY (16.5/30.0 exact — the
        // pinned S537 image-only rule). Without this the crown para was −17.7
        // and the box rows −5..−8 each → hmrc's body ran ~50pt above Word by
        // the Employee Statement section. Identified by the S535 synthetic-
        // inline signature (position (0,0) column/paragraph + WrapType::None +
        // no text blocks; textbox-text drawings get the S741 placeholder and
        // must not double-count). Corpus scan: non-pic non-txbx wp:inline
        // drawings exist ONLY in uk_hmrc_checklist → byte-identical elsewhere.
        if std::env::var("OXI_S773_DISABLE").is_err() && !lines.is_empty() {
            let s773_cy: f32 = body_para_index.map_or(0.0, |bi| {
                page.text_boxes
                    .iter()
                    .filter(|tb| {
                        tb.anchor_block_index == bi
                            && (std::env::var("OXI_EXPLICIT_INLINE_BOX_DISABLE").is_ok() || tb.inline)
                            && tb.blocks.is_empty()
                            && matches!(tb.wrap_type, Some(crate::ir::WrapType::None))
                            && tb.position.as_ref().map_or(false, |p| {
                                p.x == 0.0
                                    && p.y == 0.0
                                    && p.h_relative.as_deref() == Some("column")
                                    && p.v_relative.as_deref() == Some("paragraph")
                            })
                    })
                    .map(|tb| tb.height)
                    .fold(0.0f32, f32::max)
            });
            if s773_cy > 0.0 {
                // S839: the U+FFFC object fragment is NOT text — an
                // object-only paragraph keeps the "line = cy EXACTLY" rule.
                let has_text = lines[0]
                    .fragments
                    .iter()
                    .any(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty());
                // S839b (2026-07-14): the mixed line = cy + the TEXT line's
                // own BELOW-BASELINE extent (line0 − win_asc×text_fs) — the
                // baseline sits at box+cy (S837) and the text's descent
                // portion (incl any lineRule-auto multiplier extra) hangs
                // below it. hmrc title (Calibri 18, line=259 auto): Word
                // line 65.78 = cy 59.46 + (23.5 − 0.952×18) ✓; the old flat
                // "+0.25×para_font_size" used the FIRST run's size (11 →
                // +2.75) and left the title 3.7pt short — the doc-wide
                // upward drift's origin. The S773 probe's "+0.25×fs" is the
                // m=1.0 case of the same formula (12pt host: 14.4 − 11.42 ≈
                // 3.0 = 0.25×12).
                let target = if has_text {
                    let max_asc = lines[0]
                        .fragments
                        .iter()
                        .filter(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                        .map(|f| {
                            let fs = f.style.font_size.unwrap_or(para_font_size);
                            let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                            m.win_ascent * fs
                        })
                        .fold(0.0f32, f32::max);
                    s773_cy + (text_only_line_heights[0] - max_asc).max(0.0)
                } else {
                    s773_cy
                };
                if std::env::var("OXI_DBG773").is_ok() {
                    eprintln!(
                        "[S773] blk={:?} cy={:.1} lines={} line0={:.1} target={:.1} fired={}",
                        body_para_index,
                        s773_cy,
                        lines.len(),
                        line_heights[0],
                        target,
                        target > line_heights[0]
                    );
                }
                if target > line_heights[0] {
                    line_heights[0] = target;
                    if natural_line_heights[0] < target {
                        natural_line_heights[0] = target;
                    }
                    if ink_line_heights[0] < target {
                        ink_line_heights[0] = target;
                    }
                    s837_fired_cy = s773_cy;
                }
                if lines.len() == 1 {
                    s773_line0_target = target;
                }
            }
        }
        // S839: ordered page.text_boxes indices of this paragraph's INLINE
        // visual vector groups (the S535 signature + vector_shapes). Consumed
        // one per U+FFFC fragment in the emit loop (run order == page order).
        let s839_tbs: Vec<usize> = if std::env::var("OXI_S839_DISABLE").is_err() {
            body_para_index
                .map(|bi| {
                    page.text_boxes
                        .iter()
                        .enumerate()
                        .filter(|(_, tb)| {
                            tb.anchor_block_index == bi
                                && tb.blocks.is_empty()
                                && !tb.vector_shapes.is_empty()
                                && matches!(tb.wrap_type, Some(crate::ir::WrapType::None))
                                && tb.position.as_ref().map_or(false, |p| {
                                    p.x == 0.0
                                        && p.y == 0.0
                                        && p.h_relative.as_deref() == Some("column")
                                        && p.v_relative.as_deref() == Some("paragraph")
                                })
                        })
                        .map(|(i, _)| i)
                        .collect()
                })
                .unwrap_or_default()
        } else {
            Vec::new()
        };
        let mut s839_next: usize = 0;

        // COM-confirmed (2026-04-05, test_widow): Multiple spacing uses cumulative ceil
        // for intra-paragraph Y positions. Last line uses per-line ceil for paragraph gap.
        // COM-confirmed (2026-04-08, 683ffcab86e2): SINGLE spacing also benefits from
        // cumulative round in LM=0, but ONLY when raw > per-line round (i.e., cumulative
        // gives MORE advance than per-line), which preserves page-break decisions.
        // When raw < per-line round (e.g., Meiryo 10.5pt: raw=20.43 vs round=20.5),
        // cumulative would tighten content and shift page breaks (LOD_Handbook lost a
        // page in bdd9321 → reverted in cb35baa). Gating by sign keeps both gains.
        let is_multiple_spacing = match (
            para.style.line_spacing_rule.as_deref(),
            para.style.line_spacing,
        ) {
            (Some("exact"), _) | (Some("atLeast"), _) => false,
            (_, Some(f)) if (f - 1.0).abs() > 0.001 => true,
            _ => false,
        };
        let is_single_lm0 = !is_multiple_spacing
            && grid_pitch.is_none()
            && match (
                para.style.line_spacing_rule.as_deref(),
                para.style.line_spacing,
            ) {
                (Some("exact"), _) | (Some("atLeast"), _) => false,
                _ => true,
            };
        // S1111: a `<w:br w:type="page"/>`-only stub takes its font's RAW natural
        // height (see line_height_for_line_inner) -- keep it off the cumulative
        // multiple-spacing basis, which would re-apply the multiplier and the
        // 0.5pt cumulative round on top (the probe's 14pt stub: 18.000 instead of
        // the measured 16.099).
        let s1111_stub = lines.first().is_some_and(|l| l.fragments.is_empty())
            && para.style.page_break_after
            && !self.doc_body_has_real_cjk
            && matches!(
                para.style.line_spacing_rule.as_deref(),
                None | Some("auto")
            )
            && std::env::var("OXI_S1111_DISABLE").is_err();
        let use_cumulative_basis = (is_multiple_spacing || is_single_lm0) && !s1111_stub;
        // REPORT_administrative__0010e437 (2026-07-26, opt-out OXI_NTMULT_DISABLE):
        // a NO-TYPE docGrid Latin AUTO-MULTIPLE-spacing paragraph decides its
        // page-bottom keep/push with the S576 INK box, NOT the win/hhea spacing
        // box that S779 (normal break) and the natural widow arm apply. The two
        // consumers share no threshold, so BOTH must switch to ink — fixing only
        // one leaves the 3-line target either 1 line on p11 (widow off) or all 3
        // on p11 (S779 off). Derived by _pb_notypemult_gen.py (Word PDF, Calibri
        // + TNR × L 240..480 × cb sweep): Word flips at the line's fitz-bbox
        // (ink) bottom at EVERY factor 1.06-2.0× (d_ink 0.1-0.7pt vs d_win
        // 1.6-16.6pt), both fonts, both widow=0 (split) and widow=1 (whole-move).
        // The ink model is thus UNIFORM across L (the report's `<1.15` was an
        // unverified S695-boundary borrow, refuted here); OXI_NTMULT_ALL removes
        // the <1.15 canary scope to apply the full physics. Default keeps <1.15
        // (excludes uklocal/nyserda's 276×many + nyserda's 360/480×13 = frozen
        // PASS canaries with no PASS benefit from the change). Scope: no-type
        // docGrid (true no-grid keeps S827; typed grid keeps its snap rules) +
        // Latin (CJK keeps S608 MS-Gothic natural) + multiple (singles keep S779).
        // S1079 (2026-08-06, opt-out OXI_S1079_DISABLE): the threshold above is
        // the NATURAL (unmultiplied hhea) line, NOT the ink box. S1009 read its
        // flip off the fitz SPAN BBOX, whose height is the font's own
        // ascender+descender (1.2207em for Calibri) = exactly the natural line —
        // but the code then used `ink_lh`, a DIFFERENT quantity (the typo box,
        // 1.0em), leaving the threshold 2.4-2.7pt too lenient.
        // _pb_inkleniency_gen.py (Word PDF, 240 arms) pins it directly: with the
        // target box-top swept in 0.3pt steps the flip window is
        //   Calibri 11 x1.0792  (13.391, 13.691] ∋ natural 13.428   (box 14.491, ink 11.0)
        //   Calibri 12 x2.0     (14.397, 14.697] ∋ natural 14.648   (box 29.297, ink 12.0)
        //   TNR     12 x1.5     (13.598, 13.898] ∋ natural 13.799   (box 20.698, ink 12.0)
        // and neither box nor ink is inside ANY window. Round 1 of the same probe
        // also showed a wrapped continuation line, a <w:br/> continuation line, a
        // fresh 1-line paragraph and a fresh paragraph after an 8pt gap all
        // flipping at the SAME box-top, so the line class is not a discriminator.
        // An untyped docGrid does not change the Latin page-bottom capacity.
        // Use the hhea line as for a document with no grid; the older path
        // below substitutes the rounded spacing height for that capacity.
        let no_type_multiple_ink = std::env::var("OXI_NATURAL_CAPACITY_DISABLE").is_ok()
            && page.doc_grid_no_type
            && is_multiple_spacing
            && !self.doc_body_has_real_cjk
            && std::env::var("OXI_NTMULT_DISABLE").is_err()
            && (std::env::var("OXI_NTMULT_ALL").is_ok()
                || para.style.line_spacing.unwrap_or(1.0) < 1.15);
        let s1079_natural = std::env::var("OXI_S1079_DISABLE").is_err();
        let raw_spaced_tw: f32 = if use_cumulative_basis && !lines.is_empty() {
            let first_line = &lines[0];
            let base = {
                let mut ma: f32 = 0.0;
                let mut md: f32 = 0.0;
                let mut has_latin = false;
                if first_line.fragments.is_empty() {
                    // Match line_height_for_line_inner: empty paragraphs use
                    // pPr/rPr font + para_mark metrics (not doc_default).
                    let font_size = para
                        .style
                        .ppr_rpr
                        .as_ref()
                        .and_then(|r| r.font_size)
                        .unwrap_or(para_font_size);
                    let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                    // S707: no-grid empty-para line height is governed by the ASCII font.
                    // S949: Latin custom-pitch no-type grids included (see the
                    // line_height_for_line_inner empty branch).
                    // S1305: the estimate half of the same empty-paragraph rule.
                    // Wiring only the render half would leave the page-fit
                    // estimate disagreeing with what is drawn.
                    let s949_ascii = grid_pitch.is_none()
                        || (!para.style.snap_to_grid
                            && std::env::var("OXI_S1305_DISABLE").is_err()
                            && std::env::var("OXI_S1306_DISABLE").is_err())
                        || (page.doc_grid_no_type
                            && !self.doc_body_has_real_cjk
                            && std::env::var("OXI_S949_DISABLE").is_err());
                    let m = &*self.metrics_for_para_mark_g(&rpr_ref, &para.style, s949_ascii);
                    ma = m.word_ascent_pt(font_size);
                    md = m.word_descent_pt(font_size);
                } else {
                    // S1045: skip an OVERSIZED EDGE whitespace-only fragment — it does
                    // not drive the height (Word A/B: baselines invariant at +0.000pt).
                    // forms__002a64445e58ed78 para 11: 75 leading spaces at 9pt over 8pt
                    // visible text made this basis 15.524 instead of 13.799 (+1.725).
                    let s1045 = self.s1045_height_drivers(&first_line.fragments, para_font_size);
                    for (fi, frag) in first_line.fragments.iter().enumerate() {
                        if LayoutEngine::s1045_skip(s1045, fi, frag, para_font_size) {
                            continue;
                        }
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&frag.text, &frag.style, &para.style);
                        let border_pad = LayoutEngine::run_border_height_pad(&frag.style);
                        ma = ma.max(m.word_ascent_pt(fs) + border_pad);
                        md = md.max(m.word_descent_pt(fs) + border_pad);
                        if !frag.text.chars().all(|c| kinsoku::is_cjk(c)) {
                            has_latin = true;
                        }
                    }
                    // COM-confirmed (2026-04-07): Latin text on a line causes Word to also
                    // consider the ASCII font's CJK 83/64 height for the base.
                    // S1045: take the first fragment that DRIVES the height, so a skipped
                    // leading space run does not stand in for the visible text here.
                    if has_latin {
                        if let Some(frag) = first_line
                            .fragments
                            .iter()
                            .enumerate()
                            .find(|(fi, f)| !LayoutEngine::s1045_skip(s1045, *fi, f, para_font_size))
                            .map(|(_, f)| f)
                        {
                            let fs = frag.style.font_size.unwrap_or(para_font_size);
                            let latin_m = &*self.metrics_for(&frag.style, &para.style);
                            if latin_m.is_cjk_83_64_font() {
                                let la = latin_m.word_ascent_pt(fs);
                                let ld = latin_m.word_descent_pt(fs);
                                if la > ma {
                                    ma = la;
                                }
                                if ld > md {
                                    md = ld;
                                }
                            }
                        }
                    }
                }
                // For LayoutMode=0, use the no-grid formula (matches line_height_for_line_inner)
                // For Multiple spacing cumulative round, use RAW win_sum*fontSize (no floor)
                // so that cumulative ceil(j*raw_tw/10)*10 matches Word's sub-twip precision.
                // COM-confirmed (2026-04-09, test_widow Cambria 11pt 1.15x):
                //   raw = win_sum/upm * fontSize = 12.896pt, NOT floor'd 12.5pt
                //   raw * 1.15 = 14.830pt = 296.6tw
                //   cumulative ceil gives 15.0, 15.0, 14.5... matching Word exactly.
                let run_base = ma + md;
                if grid_pitch.is_none() {
                    let mut no_grid_max: f32 = 0.0;
                    let mut no_grid_raw_max: f32 = 0.0;
                    // S1363, second site: the cursor advance of a no-grid
                    // paragraph comes from THIS basis, not from
                    // `line_height_for_line_inner`, and it folded ascent and
                    // descent separately from the pixel-rounded components --
                    // the same pair of errors. Word (`_pb_line_pitch.py`, span
                    // over 30 lines): MS Mincho 10pt x1.15 = 14.920 = 12.971 x
                    // 1.15; a Latin line carrying kanji numerals takes the
                    // Mincho box too (Calibri 11 + numerals: 14.27 = 11 x 83/64,
                    // x1.15 = 16.41), where the separate fold gave 15.25 / 17.53.
                    // The tallest fragment's exact box, nothing mixed.
                    let s1363_site2 = std::env::var("OXI_S1363_DISABLE").is_err();
                    let s1363_box_of = |m: &FontMetrics, fs: f32| -> f32 {
                        if m.is_cjk_83_64_font() {
                            LayoutEngine::s1367_cjk_box(m, fs)
                        } else {
                            m.natural_line_height_hhea(fs)
                        }
                    };
                    let mut s1363_box_max: f32 = 0.0;
                    // Empty paragraphs: compute no_grid from para mark font
                    // (matching line_height_for_line_inner's empty-para logic).
                    if first_line.fragments.is_empty() {
                        let font_size = para
                            .style
                            .ppr_rpr
                            .as_ref()
                            .and_then(|r| r.font_size)
                            .unwrap_or(para_font_size);
                        let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        // S707: this is the no_grid block (grid_pitch.is_none()); the
                        // empty-para line height is governed by the ASCII font.
                        let m = &*self.metrics_for_para_mark_g(
                            &rpr_ref,
                            &para.style,
                            grid_pitch.is_none(),
                        );
                        no_grid_max = m.word_line_height_no_grid(font_size);
                        no_grid_raw_max = (m.win_ascent + m.win_descent) * font_size;
                        if s1363_site2 {
                            s1363_box_max = s1363_box_of(m, font_size);
                        }
                        // S612z-empty (2026-06-26): empty Zen Old Mincho para → embedded
                        // ¶ box 17.376@12pt (1.448em), not the body fallback 15.56 (rt.pdf
                        // empty-spacer = 17.40). This is the cumulative-LM0 advance basis
                        // (the operative cursor height for aiguideline's LM0 single-spacing).
                        // Opt-out OXI_S612ZE_DISABLE. See line_height_for_line_inner site.
                        if std::env::var("OXI_S612ZE_DISABLE").is_err()
                            && m.family == "Zen Old Mincho"
                        {
                            no_grid_max = no_grid_max.max(font_size * 1448.0 / 1000.0);
                            no_grid_raw_max = no_grid_raw_max.max(font_size * 1448.0 / 1000.0);
                            // S1363: the embedded face's own box is the box.
                            s1363_box_max = s1363_box_max.max(font_size * 1448.0 / 1000.0);
                        }
                    }
                    // S805: exact hhea natural (asc+desc+lineGap, no GDI px
                    // quantization) — Word's true-LM0 Latin per-line height
                    // (Arial 11: 12.649 vs the px-rounded run_base 12.75;
                    // fn_probe PDF para pitch 24.6 = 12.649 device-floored).
                    // Same value/model as S671's no-type-grid Latin.
                    let mut s805_hhea_max: f32 = 0.0;
                    // S876 (2026-07-16, default ON, opt-out OXI_S876_DISABLE):
                    // a Latin-doc EMPTY paragraph takes the SAME hhea basis as
                    // its text lines — the fragment loop below never runs for
                    // an empty para, so it fell to the GDI/com-table
                    // word_line_height_no_grid (Calibri 10.5: 12.0 vs hhea
                    // 12.817). correspondence__000a79e2: 5 empties before
                    // "EXPLORE/CONTEMPLATE:" at Word pitch 12.75-12.81 (the
                    // page arithmetic pins ~12.81 = hhea) vs Oxi 12.00 →
                    // −4.03pt → the knife-edge para stayed on the wrong page.
                    // The S815(cell)/S862(header) family: hhea-exact reaching
                    // the last un-covered empty-para path. Latin scope; the JP
                    // empty-para calibration (S583/S195/S707) is untouched.
                    // S902 (2026-07-17, opt-out OXI_S902_DISABLE): a line whose
                    // runs are ALL WHITESPACE does not take its height from
                    // them — the ¶ MARK's rPr governs, exactly like the empty
                    // para (the CELLPAIR "whitespace-only runs excluded from
                    // the cell line height" rule, body sibling). 0008ea8f p3:
                    // a spacer para = one 12pt SPACE run + an 8pt mark renders
                    // ~9.2 in Word (mark) vs Oxi 13.8 (the space run) — the
                    // +4.6 jump before SECTION TWO that pushed the checkbox
                    // table off p3.
                    let s902_all_ws = !first_line.fragments.is_empty()
                        && first_line
                            .fragments
                            .iter()
                            .all(|f| LayoutEngine::mark_spacing_only_text(&f.text))
                        && (!self.doc_body_has_real_cjk
                            || (first_line.whitespace_paragraph
                                && std::env::var("OXI_CJK_WHITESPACE_MARK").is_ok()))
                        && std::env::var("OXI_S902_DISABLE").is_err();
                    let s1080_mark_fs = std::env::var("OXI_S1080_DISABLE").is_err();
                    if (first_line.fragments.is_empty() || s902_all_ws)
                        && (!self.doc_body_has_real_cjk
                            || (s902_all_ws && std::env::var("OXI_CJK_WHITESPACE_MARK").is_ok()))
                        && std::env::var("OXI_S876_DISABLE").is_err()
                    {
                        // S1080: when S902 fires, the size fallback must come
                        // from the ¶ mark's own inheritance chain, not from
                        // `para_font_size` (= the FIRST RUN's size, i.e. the
                        // whitespace run S902 exists to exclude).
                        let font_size = para
                            .style
                            .ppr_rpr
                            .as_ref()
                            .and_then(|r| r.font_size)
                            .unwrap_or(if s902_all_ws && s1080_mark_fs {
                                self.resolve_font_size(&RunStyle::default(), &para.style)
                            } else {
                                para_font_size
                            });
                        let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        let m = &*self.metrics_for_para_mark_g(&rpr_ref, &para.style, true);
                        if !m.is_cjk_83_64_font() {
                            s805_hhea_max = m.natural_line_height_hhea(font_size);
                        }
                        if s1363_site2 && s902_all_ws {
                            // Whitespace-only runs do not drive the height; the
                            // mark's box replaces whatever the fragment loop
                            // would have folded (it is skipped there anyway).
                            s1363_box_max = s1363_box_of(m, font_size);
                        }
                    }
                    for frag in &first_line.fragments {
                        if s902_all_ws {
                            continue;
                        }
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&frag.text, &frag.style, &para.style);
                        // S1119: pick the per-CHAR face FIRST. The fallback face
                        // REPLACES the run font for the chars it covers and can be
                        // SHORTER (Calibri black-square: Word 20.438 = Courier New,
                        // below Calibri's own 21.973), so this has to happen before
                        // the run font's own value is folded into s805_hhea_max —
                        // folding first makes the rule unable to lower anything,
                        // which is exactly why the Calibri arms stayed put.
                        let s1119_on = !self.doc_body_has_real_cjk
                            && std::env::var("OXI_S1119_DISABLE").is_err()
                            && !frag.text.is_empty();
                        let s1119_faces: Option<Vec<crate::font::FontMetricsRef<'_>>> = if s1119_on
                            && frag
                                .text
                                .chars()
                                .any(|c| self.registry.symbol_fallback_face(c, &m).is_some())
                        {
                            Some(
                                frag.text
                                    .chars()
                                    .map(|c| self.registry.symbol_fallback_face(c, &m).unwrap_or_else(|| m.into()))
                                    .collect(),
                            )
                        } else {
                            None
                        };
                        let mut h = match &s1119_faces {
                            Some(fs_list) => fs_list
                                .iter()
                                .map(|f| f.word_line_height_no_grid(fs))
                                .fold(0.0f32, f32::max),
                            None => m.word_line_height_no_grid(fs),
                        };
                        let border_extra = 2.0 * LayoutEngine::run_border_height_pad(&frag.style);
                        h += border_extra;
                        match &s1119_faces {
                            Some(fs_list) => {
                                let hh = fs_list
                                    .iter()
                                    .filter(|f| !f.is_cjk_83_64_font())
                                    .map(|f| f.natural_line_height_hhea(fs))
                                    .fold(0.0f32, f32::max) + border_extra;
                                if hh > s805_hhea_max {
                                    s805_hhea_max = hh;
                                }
                            }
                            None => {
                                if !m.is_cjk_83_64_font() {
                                    let hh = m.natural_line_height_hhea(fs) + border_extra;
                                    if hh > s805_hhea_max {
                                        s805_hhea_max = hh;
                                    }
                                }
                            }
                        }
                        // Raw (un-floored) height for Multiple spacing cumulative base
                        let mut raw = match &s1119_faces {
                            Some(fs_list) => fs_list
                                .iter()
                                .map(|f| (f.win_ascent + f.win_descent) * fs)
                                .fold(0.0f32, f32::max),
                            None => (m.win_ascent + m.win_descent) * fs,
                        };
                        // S612z-circle (2026-06-23): embedded Zen Old Mincho renders CIRCLED
                        // NUMBERS (①-⑳) in its deep-win-descent box (1.448em = 17.376@12pt)
                        // while kanji fall back to MS Mincho (15.56). DERIVED from the Word
                        // PDF (aiguideline p2 circled lines 17.40 / kanji 15.56). Scoped to
                        // family "Zen Old Mincho" = aiguideline ONLY (canary-safe). The
                        // cumulative single-LM0 basis uses no_grid_max → bump it for the
                        // circled fragment. Opt-out OXI_S612ZC_DISABLE.
                        if std::env::var("OXI_S612ZC_DISABLE").is_err()
                            && m.family == "Zen Old Mincho"
                            && frag
                                .text
                                .chars()
                                .any(|c| matches!(c as u32, 0x2460..=0x2473))
                        {
                            h = h.max(fs * 1448.0 / 1000.0);
                        }
                        // S1119 (see the chain in `symbol_fallback_face`): the
                        // THIRD site. `no_grid_max` / `no_grid_raw_max` /
                        // `s805_hhea_max` are recomputed here from the fragments,
                        // discarding the ma/md fold above — the same
                        // recomputed-downstream shape that made two earlier
                        // attempts at S1118 byte-exact no-ops. Fold the fallback
                        // face in HERE or the rule cannot reach a no-grid Latin
                        // paragraph's advance at all.
                        if h > no_grid_max {
                            no_grid_max = h;
                        }
                        if raw > no_grid_raw_max {
                            no_grid_raw_max = raw;
                        }
                        if s1363_site2 {
                            let mut b = match &s1119_faces {
                                Some(fs_list) => fs_list
                                    .iter()
                                    .map(|f| s1363_box_of(f, fs))
                                    .fold(0.0f32, f32::max),
                                None => s1363_box_of(m, fs),
                            };
                            // The S612ZC circled-number box of the embedded
                            // face (aiguideline): the basis takes it like `h`.
                            if std::env::var("OXI_S612ZC_DISABLE").is_err()
                                && m.family == "Zen Old Mincho"
                                && frag.text.chars().any(|c| matches!(c as u32, 0x2460..=0x2473))
                            {
                                b = b.max(fs * 1448.0 / 1000.0);
                            }
                            b += border_extra;
                            if b > s1363_box_max {
                                s1363_box_max = b;
                            }
                        }
                    }
                    // OXI_ASCDESC_MIX: with no CJK 83/64 face on the line the box is
                    // the largest ascent(+gap) over the largest descent of its faces.
                    if std::env::var_os("OXI_ASCDESC_MIX_DISABLE").is_none() && !s902_all_ws
                        && !super::ascdesc_mix_excluded(&first_line.fragments)
                        && !first_line.fragments.is_empty()
                    {
                        let ms: Vec<(crate::font::FontMetricsRef<'_>, f32)> = first_line.fragments.iter()
                            .filter(|f| !f.text.trim().is_empty())
                            .map(|f| (self.metrics_for_text(&f.text, &f.style, &para.style),
                                f.style.font_size.unwrap_or(para_font_size)))
                            .collect();
                        if !ms.is_empty() && !ms.iter().any(|(m, _)| m.is_cjk_83_64_font()) {
                            let mix = super::ascdesc_mix_height(ms.iter().map(|(m, fs)| (&**m, *fs)));
                            if mix > s805_hhea_max { s805_hhea_max = mix; }
                            if s1363_site2 && mix > s1363_box_max { s1363_box_max = mix; }
                        }
                    }
                    if has_latin {
                        if let Some(frag) = first_line.fragments.first() {
                            let fs = frag.style.font_size.unwrap_or(para_font_size);
                            let latin_m = &*self.metrics_for(&frag.style, &para.style);
                            if latin_m.is_cjk_83_64_font() {
                                let h = latin_m.word_line_height_no_grid(fs);
                                if h > no_grid_max {
                                    no_grid_max = h;
                                }
                                let raw = (latin_m.win_ascent + latin_m.win_descent) * fs;
                                if raw > no_grid_raw_max {
                                    no_grid_raw_max = raw;
                                }
                            }
                        }
                    }
                    if s1363_site2 && s1363_box_max > 0.0 {
                        s1363_box_max.max(self.cjk_baseline_union(first_line, &para.style, para_font_size))
                    } else if is_multiple_spacing
                        && (first_line.fragments.is_empty() || s902_all_ws)
                        && s805_hhea_max > 0.0
                        && std::env::var("OXI_S910_DISABLE").is_err()
                    {
                        // S910 default-ON (2026-08-20, second audit): the
                        // 2026-07-17 hold was ukframework PASS→FAIL {-1:13};
                        // that coupling has dissolved (frozen real_en 6/6
                        // PASS 1.0 A/B-identical, EN 248 changed=0, JP golden
                        // changed=0, ssim_ab 238 = 0 changed bytes). Word truth
                        // re-derived across Arial/Calibri/TNR × 10-14pt ×
                        // f 1.15/1.5 (_pb_crfig empty-vs-text sweep): the empty
                        // line ALWAYS equals the text line, hhea × factor.
                        // (arm ordered BEFORE the raw-max arm: the EMPTY branch
                        // above also sets no_grid_raw_max from the mark's win
                        // metrics, which would shadow this fix at 13.5.)
                        // S910 (2026-07-17, opt-out OXI_S910_DISABLE): a Latin
                        // EMPTY (or whitespace-only) paragraph under a LINE
                        // MULTIPLE takes the hhea basis like its text siblings —
                        // the S876 arm below is !is_multiple-gated, so an empty
                        // ×1.15 fell to the win-based run_base (f7115: Default
                        // style Arial 12, line=276 → Oxi 13.4×1.15→15.5 vs Word
                        // hhea 13.799×1.15=15.87; text lines are correct via
                        // s671_fine). ~10 spacer empties/page × −0.43 = the
                        // −4.5pt/page systematic under-fill → the wp5 knife-edge.
                        // JP inert by construction: s805_hhea_max is only set
                        // for empties under the S876 !cjk gate. The factor is
                        // applied downstream (basis semantics unchanged).
                        s805_hhea_max
                    } else if is_multiple_spacing
                        && !first_line.fragments.is_empty()
                        && !s902_all_ws
                        && s805_hhea_max > 0.0
                        && !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S1049_DISABLE").is_err()
                    {
                        // S1049 (2026-07-31): the LM0 MULTIPLE-spacing cumulative basis
                        // is the exact hhea natural, like its S805 single-spacing sibling
                        // two arms below — NOT the GDI/win componentwise `run_base`.
                        // Word's pitch is `hhea_natural(font, size) × (w:line / 240)`
                        // accumulated exactly and device-snapped only at render.
                        // DERIVED (6-arm Word probe, no docGrid, line=259 auto, the same
                        // effective spacing given via docDefaults AND via direct pPr):
                        //   TNR 12     Word 14.8959  hhea×f 14.8912  Oxi 14.5588
                        //   Calibri 11 Word 14.4935  hhea×f 14.4908  Oxi 14.5588
                        //   Cambria 12 Word 15.1853  hhea×f 15.1821  Oxi 15.3824
                        // hhea explains all three within ±0.005pt; Oxi's error REVERSES
                        // sign by font (TNR short, Calibri/Cambria long) so no constant
                        // or ratio can fix it. docDefaults and direct pPr give byte-equal
                        // Word pitches ⇒ the spacing SOURCE is not a discriminator.
                        // EMPTY / whitespace-only lines are deliberately excluded: their
                        // multiple-spacing basis is S910's (held opt-in — default-ON there
                        // regressed ukframework), so they keep the raw-max arm below.
                        s805_hhea_max
                    } else if is_multiple_spacing && no_grid_raw_max > 0.0 {
                        run_base.max(no_grid_raw_max)
                    } else if !is_multiple_spacing
                        && s805_hhea_max > 0.0
                        && !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S805_DISABLE").is_err()
                    {
                        // S805 basis: exact hhea natural for the Latin-doc LM0
                        // single-spacing cumulative (see advance-site comment).
                        s805_hhea_max
                    } else {
                        run_base.max(no_grid_max)
                    }
                } else {
                    run_base
                }
            };
            if std::env::var("OXI_DBG805").is_ok() {
                eprintln!("[DBG805] base={:.4} run_base(ma+md)={:.4}", base, {
                    let mut ma: f32 = 0.0;
                    let mut md: f32 = 0.0;
                    for frag in &lines[0].fragments {
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&frag.text, &frag.style, &para.style);
                        if m.word_ascent_pt(fs) > ma {
                            ma = m.word_ascent_pt(fs);
                        }
                        if m.word_descent_pt(fs) > md {
                            md = m.word_descent_pt(fs);
                        }
                    }
                    ma + md
                });
            }
            let s1306_mult = para.style.line_spacing.unwrap_or(1.0);
            let raw = base * s1306_mult * 20.0;
            // S584 (2026-06-16): a TYPED docGrid line (body OR cell) is never
            // shorter than 1 grid cell, even with a COMPRESSING auto multiplier
            // (line<240). The BODY multiple-spacing path uses this cumulative
            // raw-twip model (bypassing line_height_for_line's grid snap), so
            // the floor must be applied to raw_spaced_tw here; the CELL path has
            // the mirror clamp in line_height_inner. COM-confirmed (mult_grid
            // repro, MS Mincho 10.5pt linePitch=360): line=204 (0.85x) AND
            // line=240 (1.0x) both render 18.0pt (=1 cell = 360tw); Oxi's
            // un-clamped 0.85*natural gave 11.5pt. (The actual corpus win is
            // tokyoshugyo's パワハラ list — 11 line=204 paras — but those live in
            // a TABLE cell, fixed by the line_height_inner clamp; this body site
            // is a corpus no-op since the only typed-grid body auto-mult is
            // 3a4f/model's line=360 empty which already exceeds 1 cell. Kept for
            // body correctness, validated by the repro.) Scope: snap_to_grid
            // only (a snap_to_grid=false para uses its natural height), typed
            // grid only (!doc_grid_no_type — no-type uses device-snapped
            // natural). mult>=1.25 (raw>pitch*20) is a no-op. The exact
            // fractional-cell formula does NOT generalize across font sizes
            // (14pt is flat 29.25), so only the universal "line >= 1 grid cell"
            // floor is applied. Opt-out OXI_S584_DISABLE.
            if let Some(pitch) = grid_pitch {
                if para.style.snap_to_grid
                    && !page.doc_grid_no_type
                    && pitch > 0.0
                    && std::env::var("OXI_S584_DISABLE").is_err()
                {
                    // S1306 (2026-09-04, default ON, opt-out `OXI_S1306_DISABLE`): in a TYPED grid
                    // the multiplier COMPETES with the cell count instead of
                    // multiplying the natural height -- S1185's composition,
                    // derived for vertical columns, is the horizontal law too.
                    // `_pb_cjkmult_gen.py`, one paragraph wrapping 13 times so the
                    // 0.75pt device step is spent over 12 gaps, Word PDF:
                    //   pt    mult  pitch    Word     pitch x max(cells, mult)
                    //   14.0  1.0   14.60    29.204   2 x 14.6   (cells = 2)
                    //   14.0  1.5   14.60    29.204   2 x 14.6   <- not natural x1.5
                    //   14.0  2.0   14.60    29.204   2 x 14.6   <- nor x2
                    //   10.5  1.5   14.60    21.912   1.5 x 14.6 (cells = 1)
                    //   14.0  1.5   18.00    36.006   2 x 18
                    //   10.5  1.5   18.00    27.006   1.5 x 18
                    // The note above ("Word does not snap a multiple-spaced line")
                    // holds only where the multiplier is at or above the cell
                    // count -- every arm a cells=1 document can offer. S584's floor
                    // is the same law truncated to "at least ONE cell".
                    // ★This is the site that decides it: `effective_lh` takes
                    // `raw_spaced_tw` for a multiple-spaced line and ignores
                    // `line_height`, so the snap inside `line_height_for_line_inner`
                    // never reaches the cursor.
                    let s1306 = std::env::var("OXI_S1306_DISABLE").is_err();
                    if s1306 && (s1306_mult - 1.0).abs() > 0.001 {
                        let p_tw = (pitch * 20.0).round();
                        let cells = ((base * 20.0).round() / p_tw).ceil().max(1.0);
                        p_tw * cells.max(s1306_mult)
                    } else {
                        raw.max(pitch * 20.0)
                    }
                } else {
                    raw
                }
            } else {
                raw
            }
        } else {
            0.0
        };
        // S773: a single-line paragraph hosting an inline visual drawing
        // advances by the object line (cy[+desc]) — fold the target into the
        // cumulative basis (single-LM0/multiple both route through it).
        let raw_spaced_tw: f32 = if s773_line0_target > 0.0 && raw_spaced_tw > 0.0 {
            raw_spaced_tw.max(s773_line0_target * 20.0)
        } else {
            raw_spaced_tw
        };
        // S795: a SINGLE-line Symbol-bullet paragraph advances by the marker-
        // grown line0 (the S612z three-sites trap — the cumulative basis, not
        // line_heights[0], drives the cursor for single-LM0 paragraphs). The
        // target already includes the line-spacing factor; the basis is the
        // pre-factor raw, so fold the factored value only for lines.len()==1
        // where the whole-paragraph advance IS line0.
        let raw_spaced_tw: f32 =
            if s795_line0_target > 0.0 && raw_spaced_tw > 0.0 && lines.len() == 1 {
                raw_spaced_tw.max(s795_line0_target * 20.0)
            } else {
                raw_spaced_tw
            };
        // S1116 (2026-08-14, default ON, opt-out OXI_S1116_DISABLE): fold the
        // inline-object line-0 target (S851/S875/S1095) into the cumulative
        // basis, exactly as S773 does for a vector group and S795 for a Symbol
        // bullet. Word truth (`_pb_brimg_gen.py`, 54 arms, Calibri 11 natural
        // 13.4277): a line holding BOTH text and an inline picture is
        // `max(natural, img_h + text_descent)` — 18pt picture → 20.906,
        // 60pt → 62.906 — while Oxi advanced 13.428 for EVERY picture height
        // because the basis is rebuilt from the first line's FONT metrics and
        // never consults line_heights. SCOPED to lines.len() == 1 (S795's
        // scope): on a multi-line paragraph a grown line already diverges from
        // line_heights[0], so the equal-height `use_cumulative` test fails and
        // the per-line `cursor.advance(line_height)` branch takes over and is
        // correct there (the [txt][br][img] arms measure 31.500 vs Word 31.406
        // today).
        // ★REACHABLE PATH = the S805 LM0 one (a document with NO <w:docGrid>
        // element at all). A doc that HAS a no-type docGrid takes s671_fine,
        // whose `cursor.advance(line_height)` already carries the growth —
        // educational__0003daa8's 48pt-picture line advances 68.7→119.2 = 50.5
        // correctly today. 306 of 350 corpus docs carry a docGrid, and the
        // intersection {no docGrid} × {single-line para whose picture exceeds
        // its own text line} is EMPTY over all 719 docs, so this ships
        // byte-identical: EN 248 213→213 / JP 96 92→92 / SSIM 238 bases 0
        // changed bytes / the 15 text+picture docs all element-identical.
        // Where the block does fire on a real doc the object is SMALL
        // (reference__0042471c: obj 11.2 → target 13.42 against a text basis
        // already ≥ that), so the max() is a no-op. Latent-correctness fix.
        let raw_spaced_tw: f32 = if std::env::var("OXI_S1116_DISABLE").is_err()
            && s1116_line0_target > 0.0
            && raw_spaced_tw > 0.0
            && lines.len() == 1
        {
            raw_spaced_tw.max(s1116_line0_target * 20.0)
        } else {
            raw_spaced_tw
        };
        // LM2 (charGrid) and LM0 single spacing: carry cumul_line_idx across paragraphs.
        // COM-confirmed (2026-04-12, 0e7a p2): Word maintains cumulative line index
        // across paragraph boundaries for LM0 single spacing, producing continuous
        // 11.5/11.5/12.0 round pattern instead of resetting at each paragraph.
        // Multiple spacing (1.15x etc): reset per paragraph (cumul uses raw base).
        let carry_cumul = is_single_lm0 || grid_pitch.is_some();
        let mut cumul_line_idx: usize = if carry_cumul {
            lm2_grid_cells.as_deref().copied().unwrap_or(0)
        } else {
            0
        };

        // S671 (2026-06-25): a NO-TYPE docGrid NON-CJK paragraph advances by its
        // EXACT per-line hhea-natural height (line_height_for_line_inner) instead of
        // the LM0 cumulative-round-to-0.5pt model (the CJK single-spacing device-snap,
        // S629). The 0.5pt cumulative round mis-tracks Word's per-line multiple-spacing
        // heights (±0.25pt/line phase noise); Word accumulates the EXACT line height and
        // device-snaps only the RENDERED baseline. test_line_heights mean |O−W|
        // 0.266→0.032. CJK paras keep the cumulative model (their S629 device-snap is a
        // separate wall). Scope flag from the first line (must be all non-CJK).
        // S1091 (2026-08-07, opt-out OXI_S1091_DISABLE): an EMPTY paragraph has
        // no fragments, so S671's `!lines[0].fragments.is_empty()` guard (there
        // only so the all()-over-fragments CJK test is not vacuously true) threw
        // it back onto the 10tw cumulative round — every empty Latin paragraph
        // advanced by a 0.5pt-quantised height while its text siblings advanced
        // by the exact one.  correspondence__000f9471: Oxi advanced 15.50 where
        // line_height_for_line_inner had already computed 15.442 (Calibri 11
        // hhea x 1.15) and Word steps 25.50/25.50/.../24.75 = 15.42 average;
        // 0.08pt x 17 empties = the 1.4pt drift that put its last spacer over
        // the page bottom.  Decide the CJK test from the ¶ MARK instead — the
        // same font the height itself was measured from (S583/S707/S876/S989).
        // ★SHIPPED default-ON 2026-08-13 (opt-out OXI_S1091_DISABLE) as the
        // {S1112, S1091, S1074, S1113, S1114} BUNDLE.  They are Word-correct
        // and only work together: S1112 removes the marker line's +0.14pt/
        // bullet, which is exactly what cancelled this rule's per-empty deficit
        // in policies__000f7115 (its p26 carries 5 empties at -0.37 against 18
        // Symbol-bullet lines at +0.11).
        // ★The two exposures that held the bundle opt-in are RESOLVED by
        // deriving what each of them actually was:
        //   ukframework 1.0000 -> 0.9723 {-1:13} was this rule's shorter empty
        //     slipping under S736's 2.5pt page-bottom tolerance.  That tolerance
        //     is not a constant at all (S1113): the fit test runs on the height
        //     BEFORE the lineRule=auto multiplier, so the leniency is exactly
        //     the leading the multiplier added -- zero here.  With S1113 the
        //     JP corpus is 92 -> 92 with NO doc's score moving.
        //   reference__0042471c 1.0000 -> 0.9692, which S1113 then exposed, was
        //     a +13.55pt page-1 cursor gain in a `<w:br/>` + inline-picture
        //     paragraph that fell between the S537 and S1034 routes (S1114).
        // ★GATES (5 flags together): EN frozen sets 211 -> 212 (FAIL->PASS
        // reports__00377a16 + educational__002354115a), JP corpus 92 -> 92 with
        // 0 docs changed, SSIM sentinel net +0.2731 over 238 bases (29 docs
        // improved, 3 worse by <= 0.0019), lib tests 152/0.
        // ★KNOWN COST, root attributed elsewhere: reports__0020157f
        // 1.0000 -> 0.8065 -- its p1 br-page stub overflows the content bottom
        // by 0.36pt (cursor 718.00 + 11.50 vs limit 729.14) and its own break
        // leaves a BLANK page; Word fits the stub at 715.50, i.e. Oxi's cursor
        // is +2.5 low there from the 65-border form table above it (a
        // cell-height root, not this rule).
        // Its own target correspondence__000f9471 also stays 0.9524 -- measured
        // 2026-08-13, that doc is NOT an empty-paragraph case at all: Word
        // pushes the following TEXT paragraph, and Oxi sits ~27pt high there.
        let s1091_empty_fine = lines.len() == 1
            && lines[0].fragments.is_empty()
            && !self.doc_body_has_real_cjk
            && std::env::var("OXI_S1091_DISABLE").is_err()
            && {
                let rpr = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                !self
                    .metrics_for_para_mark_g(&rpr, &para.style, true)
                    .is_cjk_83_64_font()
            };
        let s671_fine = !lines.is_empty()
            && page.doc_grid_no_type
            && std::env::var("OXI_S671_DISABLE").is_err()
            && (s1091_empty_fine
                || (!lines[0].fragments.is_empty()
                    && lines[0].fragments.iter().all(|f| {
                        !self
                            .metrics_for_text(&f.text, &f.style, &para.style)
                            .is_cjk_83_64_font()
                    })));

        // S842 (2026-07-14, opt-out OXI_S842_DISABLE): a PAGE-anchored
        // wrapTopAndBottom float band pushes this paragraph's lines below it.
        // hmrc's p2 top rule (Group 3426: positionV page 19.75, 2pt): Word
        // starts the anchor (p2-first) para at band bottom 21.75, not the
        // 15.75 margin — the block-loop arm alone missed it because the
        // p1→p2 transition happens INSIDE this paragraph's line flow.
        // The band is SPATIAL, not anchor-scoped: Word pushes ANY line
        // intersecting it on the float's page — hmrc's p2-FIRST empty para
        // (one block BEFORE the anchor) is pushed to 21.75 too, which is
        // what places the heading at 43.15 = Word 43.5. Forward-pass
        // approximation: blocks within 2 of the anchor share the band.
        let s842_band: Option<(f32, f32)> = if std::env::var("OXI_S842_DISABLE").is_err() {
            body_para_index.and_then(|bi| {
                page.text_boxes
                    .iter()
                    .find(|tb| {
                        bi <= tb.anchor_block_index + 2
                            && tb.anchor_block_index <= bi + 2
                            && matches!(tb.wrap_type, Some(crate::ir::WrapType::TopAndBottom))
                            && tb
                                .position
                                .as_ref()
                                .map_or(false, |tp| tp.v_relative.as_deref() == Some("page"))
                    })
                    .map(|tb| {
                        let tp = tb.position.as_ref().unwrap();
                        (tp.y, tp.y + tb.height)
                    })
                    .or_else(|| {
                        // Fixed-coordinate images participate in the same local
                        // top/bottom exclusion as shapes. A margin-relative
                        // image at y=2 with height=60 moves the first line from
                        // 72 to 134; a page-relative image at y=74 does likewise.
                        // The band can intersect the paragraph before its anchor,
                        // so it must be considered during line flow as well.
                        page.floating_images.iter().find_map(|img| {
                            if bi > img.anchor_block_index + 2
                                || img.anchor_block_index > bi + 2
                                || img.wrap_type != Some(crate::ir::WrapType::TopAndBottom)
                            {
                                return None;
                            }
                            let pos = img.position.as_ref()?;
                            let top = match pos.v_relative.as_deref() {
                                Some("page") => pos.y,
                                Some("margin") => page.margin.top + pos.y,
                                _ => return None,
                            };
                            Some((top, top + img.height))
                        })
                    })
            })
        } else {
            None
        };
        if std::env::var("OXI_DBG842").is_ok() {
            if s842_band.is_some() {
                eprintln!(
                    "[S842] blk={:?} band={:?} cursor={:.2} lines={}",
                    body_para_index,
                    s842_band,
                    cursor.cursor_y,
                    lines.len()
                );
            }
        }
        // S842 applier: run after every internal page push — the p1->p2
        // transition happens inside this paragraph's line flow, so the
        // loop-head check alone sees the pre-push cursor.
        let s842_apply = |c: &mut LayoutCursor| {
            if let Some((bt, bb)) = s842_band {
                if c.cursor_y >= bt - 12.0 && c.cursor_y < bb {
                    c.set(bb);
                }
            }
        };
        // S758 refactor: index-based iteration so the side-wrap rebreak can
        // splice the remaining lines at a float-band exit (byte-identical
        // when no rebreak fires — same order, same borrows).
        let mut line_idx = 0usize;
        // A region-aware plan already covers the entire remaining source,
        // including the float-band exit and page/column transitions. Rebreaking
        // it again changes the lines without rebuilding their cached geometry.
        let mut s758_rebroken = s758_band.is_none() || !region_line_widths.is_empty()
            || !word_fit_columns.is_empty();
        let s758_entry_pages = pages.len();
        while line_idx < lines.len() {
            if std::env::var("OXI_DBG_WF").is_ok() && !word_fit_floors.is_empty() {
                eprintln!("[WF-EMIT] li={} y={:.2} pages={} entry={} col={} start_col={} floor={:?}",
                    line_idx, cursor.cursor_y, pages.len(), s758_entry_pages, cur_col,
                    start_column, word_fit_floors.get(line_idx));
            }
            let actual_flow_column=(pages.len()-s758_entry_pages)*num_columns.max(1)+cur_col;
            if word_fit_columns.get(line_idx).copied().map_or(
                pages.len()==s758_entry_pages && cur_col==start_column,
                |column|column==actual_flow_column) {
                if let Some(Some(bottom)) = word_fit_floors.get(line_idx) {
                    if cursor.cursor_y < *bottom { cursor.set(*bottom); }
                }
            }
            // S842: a line starting inside (or whose box would overlap) the
            // page-anchored band moves below it.
            if let Some((bt, bb)) = s842_band {
                if cursor.cursor_y >= bt - 12.0 && cursor.cursor_y < bb {
                    cursor.set(bb);
                }
            }
            // S758: the cursor exited the band (or a page push left it behind)
            // → rebreak the remaining lines' fragments at the full width.
            if !s758_rebroken {
                if let Some((band_bot, _, _)) = s758_band {
                    if (cursor.cursor_y >= band_bot - 0.5 || pages.len() > s758_entry_pages)
                        && line_idx < lines.len()
                    {
                        s758_rebroken = true;
                        let rem: Vec<(String, RunStyle, Option<FieldType>, usize, usize)> = lines
                            [line_idx..]
                            .iter()
.flat_map(Line::source_fragments)
                            .collect();
                        let refs: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = rem
                            .iter()
                            .map(|(t, st, ft, ri, co)| (t.as_str(), st, ft.clone(), *ri, *co))
                            .collect();
                        self.s1636_lane_shift.set(0.0);
                        let nl = self.break_into_lines(
                            &refs,
                            s758_wrap_full,
                            0.0,
                            &para.style,
                            effective_char_pitch,
                            effective_cw_ratio,
                            page.doc_grid_lines_and_chars,
                            true,
                            matches!(para.alignment, Alignment::Justify | Alignment::Distribute),
                            page.doc_grid_no_type,
                            para_has_lrpb,
                            caps_active,
                            false,
                        );
                        let (mut advance, mut natural, mut latin_fit, mut ink,
                            text_only, _, leading) = line_boxes(&nl);
                        // A paragraph's first source row also carries marker/group
                        // contributions applied after the common text/object fold.
                        // Keep those contributions if reflow starts at source row zero.
                        if line_idx == 0 && !nl.is_empty() {
                            advance[0] = advance[0].max(line_heights[0]);
                            natural[0] = natural[0].max(natural_line_heights[0]);
                            ink[0] = ink[0].max(ink_line_heights[0]);
                            if !latin_fit.is_empty() {
                                latin_fit[0] = latin_fit[0].max(s779_win_heights[0]);
                            }
                        }
                        line_heights.truncate(line_idx);
                        line_heights.extend(advance);
                        natural_line_heights.truncate(line_idx);
                        natural_line_heights.extend(natural);
                        s779_win_heights.truncate(line_idx);
                        s779_win_heights.extend(latin_fit);
                        ink_line_heights.truncate(line_idx);
                        ink_line_heights.extend(ink);
                        text_only_line_heights.truncate(line_idx);
                        text_only_line_heights.extend(text_only);
                        story_image_leading.truncate(line_idx);
                        story_image_leading.extend(leading);
                        lines.truncate(line_idx);
                        lines.extend(nl);
                        if line_idx >= lines.len() {
                            break;
                        }
                    }
                }
            }
            let line = &lines[line_idx];
            // S1335: set when this line is a bare column break that the natural
            // overflow (S637) already carried into the next column.
            let mut s1335_break_consumed = false;
            let _first_style = line
                .fragments
                .first()
                .map(|f| &f.style)
                .unwrap_or(&default_style);
            let line_height = line_heights[line_idx];
            // S1497: the paragraph-relative wrapTopAndBottom band.
            if let Some((bt, bb)) = s1497_band {
                if pages.len() == s758_entry_pages {
                    if line_idx == 0 {
                        // The FIRST line's box begins at the block's entry (its
                        // space-before belongs to it): a band that starts inside
                        // [entry, entry + before + line] cuts it, and Word moves the
                        // whole paragraph -- space-before included -- below the band
                        // (S1089's technical__002c6778: the 0.1pt rule at posOffset 1.25
                        // sits inside the 7.3pt before; text at 145.83 = band bottom +
                        // 7.3). Testing the line against the band alone missed it.
                        if bt >= s1497_entry_y - 0.01
                            && bt < s1497_entry_y + effective_spacing.max(0.0) + line_height - 0.01
                        {
                            cursor.set(bb + effective_spacing.max(0.0));
                        }
                    } else if cursor.cursor_y + line_height > bt + 0.01 && cursor.cursor_y < bb {
                        cursor.set(bb);
                    }
                }
            }
            // Full-width text frame below the anchor (see TEXT_FRAME_EXCLUSION).
            if body_para_index.is_some() && pages.len() == s758_entry_pages {
                if let (Some((pg, top, bottom)), Some((_, cur))) =
                    (TEXT_FRAME_EXCLUSION.with(|c| c.get()), body_wrap_bands)
                {
                    if pg == cur && cursor.cursor_y + line_height > top + 0.01 && cursor.cursor_y < bottom {
                        cursor.set(bottom);
                    }
                }
            }
            // Fixed top margins allow header text to overlap the body, but
            // wrapping header drawings exclude each intersecting body line.
            if body_para_index.is_some() {
                let bands = self.negative_header_wrap_bands(page, pages.len(),
                    s755_geom.map_or(false, |g| g.first_even));
                while let Some((_, bottom)) = bands.iter().find(|(top, bottom)|
                    cursor.cursor_y + line_height > *top && cursor.cursor_y < *bottom) {
                    cursor.set(*bottom);
                }
            }
            // BODYLINE instrument: reliable per-LINE body text + char count + cursor_y
            // (the GDI --dump-layout emits per-RUN for body, hiding line breaks). Use to
            // localize per-line body over-fit (Oxi chars/line vs Word PDF). OXI_DUMP_BODYLINE.
            if std::env::var("OXI_DUMP_BODYLINE").is_ok() {
                let lt: String = line.fragments.iter().map(|f| f.text.as_str()).collect();
                let nc = lt.chars().count();
                let pi = body_para_index
                    .map(|v| v.to_string())
                    .unwrap_or_else(|| "?".into());
                eprintln!(
                    "[BODYLINE] pi={} li={} cy={:.1} lh={:.1} nc={} «{}»",
                    pi,
                    line_idx,
                    cursor.cursor_y,
                    line_height,
                    nc,
                    lt.chars().take(44).collect::<String>()
                );
            }

            // Page break check with widow/orphan control
            // TextBox content: no page breaks, no widow/orphan. Overflow is clipped.
            let effective_lh = if is_multiple_spacing && raw_spaced_tw > 0.0 && !s671_fine {
                let old_pos = mult_cumul_raw.as_deref().copied().unwrap_or(0.0);
                let new_pos = old_pos + raw_spaced_tw;
                let cn = (new_pos / 10.0).round() as i32 * 10;
                let cc = (old_pos / 10.0).round() as i32 * 10;
                (cn - cc) as f32 / 20.0
            } else {
                // S671: a NON-CJK no-type-grid line uses the INDEPENDENT per-line
                // height (line_height_for_line_inner = snap_0.12(hhea natural ×
                // factor)) — Word computes multiple-spacing line heights per-line,
                // NOT via the cumulative-round model (which is the LM0 single-spacing
                // device-snap, S629). line_height already carries the S671 value.
                line_height
            };
            // Day 33 part 65 (2026-05-12): use natural_lh (ascent+descent) for
            // break threshold; the grid-snap LEADING (line_h − natural_lh) is
            // allowed to extend into bottom margin. Cursor still advances by
            // full line_h. COM-confirmed via db9ca18 i=37 (+5.25pt overflow
            // accepted by Word).
            let natural_lh = natural_line_heights
                .get(line_idx)
                .copied()
                .unwrap_or(effective_lh);
            // S548b (2026-06-12, opt-out OXI_S548B_DISABLE): the Day-33
            // leniency is the INK-BOTTOM rule and does NOT apply to
            // lineRule=exact. For exact lines the text sits at the BOTTOM of
            // the box (S495 bottom-align) — there is no spare leading below
            // the ink — so the FULL box height must fit above the bottom
            // margin. 3a4f p43→p44: ① para (line=350 exact, 17.5pt) at
            // cursor 741.5, content bottom 756.85: natural_lh 13.6 fit it
            // (755.1) where Word pushes (741.25+17.5=758.75 > 756.85) — one
            // of the 5 delta=-1 boundary paras behind the Phase-1 sole FAIL.
            // Auto/grid lines keep the Day-33 natural_lh leniency (db9ca's
            // +5.25 leading acceptance is the auto-grid case: ink at top,
            // leading below).
            let s548b_exact_full = para.style.line_spacing_rule.as_deref() == Some("exact")
                && std::env::var("OXI_S548B_DISABLE").is_err();
            // S562 SHIP (2026-06-14, default ON, opt-out OXI_S562B_DISABLE): the
            // Day-33 natural_lh leniency does NOT apply to EMPTY paragraphs.
            // roudoujoken's −1: a trailing empty para (i=147) at the page-2 bottom
            // ends ~1.7pt past the bottom margin; Oxi's natural_lh leniency kept it
            // on p2, but Word pushes it to p3 (the empty para's full grid box must
            // fit). That keeps Oxi's page 3 starting ~15.7pt higher (no empty atop
            // p3) → ８.「休暇」 fit p3 where Word has it on p4. db9ca (the leniency's
            // COM source) is a CONTENT line (ink at top, leading below) — empties
            // have no ink to anchor, so Word uses their full box. Discriminator =
            // empty para. GATE: Phase-1 55/57 → 56/57 (roudoujoken FAIL→PASS), 0
            // PASS→FAIL, mean 0.9980; only kyotei (multi-col residual) still fails.
            // S1375 (see the S739 site below): an empty line whose NEXT block is
            // an empty section-end paragraph is judged by its natural height,
            // not by its full grid box -- Word keeps it down to `natural`
            // remaining (8pt: 10.35 kept / 9.85 pushed; full box 16.25 would
            // push at 16). forms__00830ac053a2c57a p6.
            let s1375_before_section_end = std::env::var("OXI_S1375_DISABLE").is_err()
                && body_para_index
                    .and_then(|bi| page.blocks.get(bi + 1))
                    .map_or(false, |b| matches!(b, Block::Paragraph(n)
                        if n.style.page_section_break
                            && n.runs.iter().all(|r| r.text.is_empty())));
            // S1572 (2026-09-26, default ON, opt-out OXI_S1572_DISABLE): an empty
            // line whose snapped box spans two or more grid cells is judged like
            // text by the centred box, not by the full box. Probe gridbottom_fit
            // (28 arms, Meiryo on a lines grid, empty and text identical): a
            // 2-row box may overhang the body bottom by 6.9 (kept) but not 9.3.
            // administrative__108cf8: Word keeps an empty Meiryo line whose 31.6
            // box ends 6.1 past the bottom; Oxi pushed it and cascaded p2/p3.
            // Single-cell empties keep S562b (roudoujoken, and the four Latin
            // documents that regressed when S562b was dropped wholesale, S1557).
            let s1572_multicell = std::env::var_os("OXI_S1572_DISABLE").is_none()
                && !page.doc_grid_no_type
                && para.style.snap_to_grid
                && grid_pitch.is_some_and(|p| p > 0.0 && effective_lh > p * 1.5);
            let s562b_empty_full = std::env::var("OXI_S562B_DISABLE").is_err()
                && para.runs.iter().all(|r| r.text.is_empty())
                && !s1375_before_section_end
                && !s1572_multicell;
            // S576 (2026-06-15, default ON, opt-out OXI_S576_DISABLE): the
            // page-bottom break-fit measures the GLYPH INK (≈ em), not the
            // line-SPACING box. natural_lh is win_sum*83/64 = 1.297*em for CJK
            // (MS Mincho 11pt → 14.25), ~3.2pt larger than the real ink; that
            // over-count rejected page-bottom lines Word fits (their grid
            // leading hangs into the margin). PDF gold-standard ikujidetail p9
            // "３ 請求…": ink bbox h=11.04 ≈ em=11.0, fits its 14.3 grid box;
            // Oxi at 14.25 rejected → +1 cascade on word pages 9/12-16 (12
            // paras). ink_lh = typo_sum*fs (= em for MS/Yu Mincho/Gothic).
            // Exact lines (S548b: text bottom-aligned, no spare leading) and
            // empty paras (S562b: no ink to anchor) keep the FULL box.
            // SCOPE = no-type docGrid ONLY. A TYPED docGrid (w:type=lines /
            // linesAndChars) grid-SNAPS each line to a whole cell, so the
            // page-bottom occupant is the full grid cell, not the glyph ink —
            // applying ink-leniency to typed grids let Oxi fit a line Word
            // breaks (ikujikaigo + model each picked up −1×3, PASS→FAIL). A
            // no-type docGrid uses the natural device-snapped advance (S571b),
            // so its leading genuinely overhangs the margin like LM0.
            let ink_lh = if std::env::var("OXI_S576_DISABLE").is_ok() || !page.doc_grid_no_type {
                // OXI_OFFSLOT_INK (opt-in experiment, ROWBOX2 companion): an
                // OFF-GRID paragraph (the S592 class — proportional CJK font
                // in a linesAndChars grid; its lines do NOT snap to slots)
                // uses the glyph-INK page-bottom threshold (like no-type
                // grids / LM0); grid-snapped monospace bodies keep natural +
                // the S739 centered floor. Specimen: kojin pi=52 last line
                // under ROWBOX2 (Word-correct +2pt table heights): Word
                // keeps it at natural-over +1.65 — ink (~em 11.8) fits where
                // natural (13.5) pushes. ★A first cut scoped by CURSOR PHASE
                // (off-slot within the page) also fixed kojin but flipped
                // tokyoshugyo (alone) + kyotei36spec (with ROWBOX2): the
                // phase is Oxi-geometry-relative (the S739 fragility). The
                // S592 font-class scope is docx-derivable: only kojin +
                // parttime (HGPGothicM) match in the corpus.
                // ★SHIPPED default ON 2026-07-07 (opt-out OXI_OFFSLOT_INK_DISABLE)
                // as part of the ROWBOX2 bundle.
                let offgrid_ink = std::env::var("OXI_OFFSLOT_INK_DISABLE").is_err()
                    && !page.doc_grid_no_type
                    && page.doc_grid_lines_and_chars
                    && para
                        .runs
                        .iter()
                        .find(|r| {
                            r.text
                                .chars()
                                .any(|c| !c.is_whitespace() && c != '\u{3000}')
                        })
                        .and_then(|r| self.metrics_for_cjk(&r.style, &para.style))
                        .map_or(false, |m| {
                            matches!(
                                m.family.as_str(),
                                "MS PGothic" | "MS PMincho" | "HGPGothicM"
                            )
                        });
                if offgrid_ink {
                    ink_line_heights
                        .get(line_idx)
                        .copied()
                        .unwrap_or(natural_lh)
                        .min(natural_lh)
                } else {
                    natural_lh
                }
            } else if std::env::var_os("OXI_CJK_TEXT_BOTTOM_BOX_DISABLE").is_none() /* S1409 */
                && lines[line_idx].fragments.iter().any(|fragment| {
                    fragment.text.chars().any(kinsoku::is_cjk_ideograph_or_kana)
                })
            {
                natural_lh
            } else {
                ink_line_heights
                    .get(line_idx)
                    .copied()
                    .unwrap_or(natural_lh)
                    .min(natural_lh)
            };
            // S582 (2026-06-15) FALSIFIED the "S576 ink-leniency is ~1.75pt too
            // loose" hypothesis for ikujidetail's +1×2: an OXI_INK_MARGIN sweep
            // showed margin 0 (= ink=em, current) is OPTIMAL; +1.0/+1.5 unchanged,
            // +1.75 WORSE (+1×5), +2.0..box +1×12 (the S571b state). So the
            // page-bottom threshold is correct; the +1×2 are the doc-wide para-spill
            // break-POINT cascade (which line of a wrapped para lands at the bottom),
            // reset by the real pi=149/263 LRPBs — not a threshold calibration.
            // S603 (2026-06-18, default ON, opt-out OXI_S603_DISABLE): a TYPED docGrid
            // (w:type=lines / linesAndChars) line normally keeps the Day-33 natural_lh
            // leading-hang leniency at the page bottom (the grid LEADING, cell−natural_lh,
            // is allowed to overhang the bottom margin). EXCEPTION: the LAST line of a
            // paragraph immediately followed by a TABLE block uses the FULL grid cell —
            // Word does NOT let that line's leading hang when a table follows (the
            // table/para junction is measured at full line height). DERIVED from the
            // Phase-1 gate: a blanket "typed grid uses full cell" regressed 6 docs
            // (db9ca/kojin/roudoujoken/34140) whose page-bottom leniency lines are all
            // followed by BODY text → keep the leniency; 3a4f para278 line4 (followed by
            // the box279 table) is the ONLY full-cell case. 3a4f para278: cursor 741.1 +
            // cell 18 = 759.1 > content_bottom 756.85 → break to page 34 (= Word); the
            // natural_lh 13.6 → 754.7 leniency wrongly kept it on page 33, the
            // compensating S559 error behind the cap-3.1+ぶら下げ −1 at para306. Exact
            // (S548b) and empty (S562b) lines already use the full box.
            // S1194 (2026-08-22, opt-in `OXI_S1194`): S603 says "a TYPED docGrid
            // line" and its whole argument is about the GRID LEADING (cell −
            // natural_lh) hanging into the bottom margin — but the condition
            // never tests for a grid. `!doc_grid_no_type` is satisfied by a
            // document with NO docGrid at all, so the rule also fires on no-grid
            // Latin, where there is no leading to hang and the derived capacity
            // is S779/S827's hhea line. 00501ca3 p8 pi=134 «(Added 2002)
            // (Amended 2010)» is the last line before a table in a no-grid
            // Latin doc: S603 forces the 12.000 box against a 720.000 bottom at
            // cursor 708.110 and pushes a line Word keeps (hhea 11.499 → −0.391).
            // SHIPPED default-ON 2026-08-22 (opt-out `OXI_S1194_DISABLE`).
            // GATE: inert everywhere measured — EN 248 PASS 224 → 224 with ZERO
            // docs changed in all five sets; the golden Phase-1 census (98 docs)
            // unchanged; the SSIM sentinel over all 82 pure-Latin golden docs
            // byte-identical. The 156 CJK golden docs are structurally excluded
            // (`s779_latin` needs `!doc_body_has_real_cjk`, and the grid gate
            // only relaxes where there is no docGrid at all). What it buys is
            // S1192: `S1189+S1192` on 00501ca3 goes FAIL 0.9963 → PASS 1.0000.
            let s1194_grid_scope = std::env::var("OXI_S1194_DISABLE").is_ok()
                || page.grid_line_pitch.is_some();
            // S1605 (2026-09-29): S603 RETIRED to opt-in (`OXI_S603=1`). Its one
            // derivation case (3a4f para278) passes without it, and a faithful
            // slice (`_pb_s603_gen.py`, linesAndChars 416, HGPGothicM 12pt, exact
            // spacer swept in 0.5pt) shows Word keeps the last line up to the same
            // spacer (710.0) whether a body paragraph, a table, or a two-line
            // paragraph + table follows; S603 moved the table arms to 707.5 / 708.0.
            // blind-G policies__1f014c0f p20 «※帳票は…» (followed by a table) is
            // the corpus instance. Full batch_gate with it off: 1075 -> 1075.
            let s603_typed_fullbox = next_block_is_table
                && line_idx + 1 == lines.len()
                && std::env::var_os("OXI_S603").is_some()
                && !page.doc_grid_no_type
                && s1194_grid_scope
                && !s548b_exact_full
                && !s562b_empty_full;
            // S605 (2026-06-18, default ON, opt-out OXI_S605_DISABLE): the FIRST line
            // of a TWO-line typed-grid paragraph uses the FULL grid cell (no natural_lh
            // leniency) — but ONLY when line0's cell-bottom OVERFLOWS the content bottom
            // (line0 fits the page bottom only via the leniency hang). Word does not let
            // a 2-line para's first line hang its grid leading into the bottom margin:
            // that would split the para 1+1 (the last line alone at the top of the next
            // page = a WIDOW, and the first alone at the bottom = an ORPHAN). It uses the
            // full cell, so the whole 2-line para moves down. ohnoshugyo pidx=203 (第３３
            // 条…, 2-line, MS Mincho 10.5, type=lines): line0 cursor 743 + cell 18 = 761
            // > content_bottom 756.85 → Word puts the whole para on page 9; Oxi's
            // natural_lh 13.6 → 756.6 leniency wrongly fit it on page 8 (1+1) → the −1.
            // ★DISCRIMINATOR = line0's FULL CELL OVERFLOWS (a leniency over-fit). The
            // BROADER widow_effective "a 2-line para is never split 1+1" was FALSIFIED on
            // the gate (−9, 10 PASS→FAIL incl ikujikaigo/kojin/model): Word DOES split
            // most 2-line paras 1+1 when line0 fits the page bottom by the FULL cell
            // (a legitimate fit). Only the leniency-over-fit line0 is pushed. Single-line
            // paras and continuation/last lines keep the leniency (S603 covers the
            // last-line-before-a-table case separately). Phase-1 75→76 (ohnoshugyo
            // FAIL→PASS), 0 PASS→FAIL.
            // S748 (2026-07-05): S605's full-box is scoped to OFF-SLOT cursors.
            // A controlled 2-line-para bottom sweep (_pb_capacity_gen cap2, 25
            // COM points, 0.1pt fine flip) derived that an ON-SLOT 2-line
            // para's line0 obeys the SAME centered-box rule as S739 — the flip
            // lands EXACTLY in (15.8, 15.9] ∋ (pitch+nat)/2 = 15.809, and Word
            // KEEPS line0 + splits 1+1 when the centered box fits (probeqbrkchars
            // 第20条 at slot 754.75: centered 770.56 ≤ 771 → Word keeps; S605's
            // full box 18 pushed it → the probe's +1×2). The earlier "S605
            // counterexample" reading used the Info6 y (slot + ~2.25 centering
            // offset) as the slot — the corrected arithmetic shows NO exception.
            // ohnoshugyo pidx=203 (S605's source) is OFF-slot (phase ~6pt, mixed
            // heights above) where the centered rule does not apply (S739's
            // on-slot gate) — the full box stays as the off-slot approximation
            // that reproduces Word's push there.
            let s748_on_slot = grid_pitch.map_or(false, |p| {
                if p <= 0.0 {
                    return false;
                }
                let phase = (cursor.cursor_y - page_top).rem_euclid(p);
                phase < 1.0 || phase > p - 1.0
            }) && std::env::var("OXI_S748_DISABLE").is_err();
            // S1417 (2026-09-15, default ON, opt-out OXI_MULTICELL_GRID_FIT_DISABLE):
            // the checkpoint's opt-in promoted. reference__1323b9bc (Meiryo 12 in
            // 2-cell 36pt grid lines, widowControl): the S608 natural look-ahead
            // read the last line at the cell top and split a 2-line paragraph
            // 1+1; Word centres the natural box in its cells (Info(6) sits 5.75
            // below the cell top = (36 - 24.5) / 2) and pushes both lines. Env
            // gates: golden 183 -> 185 (probeqsizes, probexbigrun), ja 188 same.
            let centered_multicell_grid = std::env::var_os("OXI_MULTICELL_GRID_FIT_DISABLE").is_none()
                && !page.doc_grid_no_type
                && para.style.snap_to_grid
                && grid_pitch.is_some_and(|p| p > 0.0 && effective_lh > p * 1.5);
            let s605_line0_2 = !centered_multicell_grid
                && std::env::var("OXI_S605_DISABLE").is_err()
                && line_idx == 0
                && lines.len() == 2
                && !page.doc_grid_no_type
                && !s748_on_slot
                && !s548b_exact_full
                && !s562b_empty_full;
            // S693 (2026-06-29, default ON, opt-out OXI_S693_DISABLE): generalize
            // S605 from 2-line paras to ANY NON-LAST line — a typed-grid page-bottom
            // line with MORE lines following in the same paragraph, whose natural-
            // leniency over is a HAIRLINE (nat_over > OXI_S693_OV, default -1.0; the
            // full box barely overflows the content bottom), breaks at the FULL grid
            // cell (no natural_lh leniency). DERIVED discriminator (BR_DUMP keep/break
            // vs Word PDF): tokyoshugyo «給月給» (pi=444, line 2 of a 5-line para,
            // over=-0.80, Word BREAKS) vs kojin pi=52 (over=-0.35, LAST line, Word
            // KEEPS) and 〔例２〕 (pi=205, over=-1.10, LAST line, Word KEEPS) — over is
            // INVERTED (kojin keeps the tighter -0.35), so the cut is last-vs-non-last
            // (NOT over, refuting OXI_TGINK_K/S651), with an over-hairline gate WITHIN
            // non-last lines (see break_threshold override below). NARROWER than the
            // FALSIFIED S687 ("all continuation lines", idx>=1) which broke the
            // last-line continuations Word keeps (+1 5->30); s693 EXCLUDES last lines.
            // Shipped DEFAULT-ON JOINTLY with S694 (the widow box-split fix): S693
            // pushes «給月給» down (fixing the chapter -1) but EXPOSES the chapter
            // over-tallness (the 精勤手当/賞与 widow-push), which S694 then fixes — the
            // two are a compensating pair (S559 pattern), net tokyoshugyo 0.9855→
            // 0.9874. Corpus-safe: only tokyoshugyo changes (every canary byte-identical).
            // The full grid-cell fit rule requires an actual line grid.
            // A missing docGrid also has doc_grid_no_type=false.
            let s693_nonlast = std::env::var("OXI_S693_DISABLE").is_err()
                && grid_pitch.is_some()
                && line_idx + 1 < lines.len()
                && !page.doc_grid_no_type
                && !s548b_exact_full
                && !s562b_empty_full;
            // ★tokyoshugyo #2 (2026-06-22 s3): typed-grid full-cell break — one of the
            // ~3 stackable #2 components (page-bottom natural-leniency ~11 lines). Forces
            // effective_lh when the cell snaps > natural. Canary-clean on ikujikaigo/
            // model/3a4f/ikujidetail (their page-bottom paras don't hit the snapped-larger
            // gate). OXI_S_TGFULL. See [[tokyoshugyo_wrap_not_cellheight]].
            let s_tgfull = std::env::var("OXI_S_TGFULL").ok().as_deref() == Some("1")
                && !page.doc_grid_no_type
                && effective_lh > natural_lh + 0.1;
            // S651 (2026-06-24, default ON, opt-out OXI_S651_DISABLE): a typed-grid
            // line that snaps to 2+ grid cells (effective_lh > 1.5*pitch — a sz>=14pt
            // chapter heading: natural 18.16 > pitch 18 → cells=2 → effective_lh=36)
            // must fit its FULL box at the page bottom. The Day-33/S576 natural_lh
            // leniency forgives only ONE grid cell's leading (~4pt, cell−ink); a 2-cell
            // heading has a WHOLE empty 2nd cell (~18pt) below the ink, and Word does
            // NOT let that hang into the bottom margin. tokyoshugyo 第４章 at Op22 y730:
            // box 36 → bottom 766 > content_bottom 756.85 → Word pushes to p23 (=Word
            // p23 top), but Oxi's natural_lh 18.16 → 748.16 wrongly fit it on p22 →
            // the WHOLE 第４章 chapter region (一般勤務 …) cascaded to −1. NARROWER than
            // OXI_S_TGFULL (which fires on every 1-cell line where effective_lh>natural_lh
            // = the ~4pt grid leading Word DOES allow to hang → over-corrects the 賃金
            // body). Only multi-cell (sz>=14) headings trigger. Exact (S548b) / empty
            // (S562b) lines already use the full box.
            let s651_multicell_head = !centered_multicell_grid
                && std::env::var("OXI_S651_DISABLE").is_err()
                && !page.doc_grid_no_type
                && grid_pitch.map_or(false, |p| p > 0.0 && effective_lh > p * 1.5);
            // S687 FALSIFIED (2026-06-28): "a mid-paragraph CONTINUATION line uses the
            // FULL grid box (no ink leniency)" — aimed at the tokyoshugyo 賃金-chapter −1
            // origin, p46 «給月給（定額賃金制…» (Oxi cursor 742.55, ink_lh 13.5 → over=-0.8
            // FIT; Word breaks to p47, full box → over=+3.7; cursors aligned ~0.65pt).
            // Forcing full box for ALL continuation lines OVER-CORRECTS: tokyoshugyo −1
            // 23→8 (the 賃金 origin IS continuation-line leniency) BUT +1 5→30 (most
            // continuation lines Word KEEPS via leniency), net 0.9817→0.9761 WORSE. The
            // «給月給» (Word ink overflows by 0.15pt) and the 30 Word-keeps lines INTERLEAVE
            // at the sub-pt ink boundary — the documented "no threshold separates
            // push-vs-keep" wall (S651). Needs Word's exact per-line ink/device-snap, not
            // a continuation-line gate. See [[tokyoshugyo_wrap_not_cellheight]].
            // S736 (2026-07-03): the S562b empty-para full box gets a small
            // KEEP TOLERANCE — Word retains a page-bottom EMPTY paragraph whose
            // full box overhangs the content bottom by a little. Two-sided
            // bounds (S725-style window): probexempty's spacer empty overflows
            // +2.0 and Word KEEPS it (i=27 at y757.0, box 772.8 > 771); the
            // S562b roudoujoken empty is PUSHED — its measured overflow bounds
            // the window above at ~(3.0, 3.5] (the S562-era "~5.75" was an
            // approximation; the TOL sweep flips roudoujoken between 3.0 and
            // 3.5). Verified clean window [2.1, 3.0]; default 2.5. Tune
            // OXI_S736_TOL. Fires only when
            // the empty rule is the SOLE full-box reason (exact/table-next/
            // footer-tight etc. keep the strict box).
            let s736_tol: f32 = std::env::var("OXI_S736_TOL")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(2.5);
            // S1055 (2026-08-02 derived, SHIPPED 2026-08-06 default ON, opt-out
            // OXI_S1055_DISABLE): a br-page STUB (an empty paragraph whose only
            // content was <w:br w:type="page"/>, parsed to page_break_after) is
            // NOT a spacer empty — it gets no S736 keep tolerance and moves to
            // the next page like any other line when its full box overflows.
            // MEASURED on Word (legal__001a2c7f, ExportAsFixedFormat): the stub
            // overflows p11 by +1.805pt, Word moves it to p12 — and because its
            // break then sends the following content one page further, Word's
            // p12 holds header + footer only (0 body lines, verified). Two Word
            // variants pin the mechanism: removing the FOLLOWING paragraph's
            // lastRenderedPageBreak leaves 28 pages with the same blank p12 (the
            // LRPB is inert), while removing the br drops to 27 pages with 16.2
            // at p12 y=132.50 — one 10pt line below the page top, i.e. the same
            // paragraph landing on p12 as a plain empty.
            //
            // ★The three docs that blocked this in 2026-08 (policies__000f7115,
            // educational__00161422, reports__0020157f) were each over-full by
            // ~one line where their stub sits; all three are PASS 1.0000 with
            // the stub rule on since S1057 (the empty FOOTER ¶ mark takes the
            // ASCII font) and S1059 (an over-long token packs to the last
            // fitting character) landed. Re-verified 2026-08-06.
            let s1055_br_stub = std::env::var("OXI_S1055_DISABLE").is_err()
                && !self.doc_body_has_real_cjk
                && para.style.page_break_after
                && para.runs.iter().all(|r| r.text.is_empty());
            let s736_empty_tol = s562b_empty_full
                && !s548b_exact_full
                && !s603_typed_fullbox
                && !s605_line0_2
                && !s_tgfull
                && !s651_multicell_head
                && !footer_tight
                && !s1055_br_stub
                && std::env::var("OXI_S736_DISABLE").is_err();
            // S1041 (2026-07-29, opt-out OXI_S1041_DISABLE): the S736 tolerance
            // must not be clamped to ink_lh on an EMPTY line. An empty paragraph
            // paints no glyphs, so ink_lh there is a font-derived fiction, and
            // for a 11.724pt empty line (ink_lh 10.0) the clamp cut the
            // configured 2.5pt tolerance down to 1.724pt - which is why the
            // OXI_S736_TOL sweep is inert above 2.5 (the clamp dominates).
            // reference__0042471c's p4-terminal empty paragraph overflows its
            // full box by 2.513pt: Word keeps it on p4, the clamp pushed it, and
            // that single break supplied a p5-p8 cascade that ended as a 35pt
            // shift on p7 (two of six pre-figure empties spilling onto the page).
            // Without the clamp the residual is +0.013pt, inside S967's half-twip
            // tolerance, and the paragraph stays where Word puts it.
            // S1113 (2026-08-13, SHIPPED default-ON, opt-out
            // OXI_S1113_DISABLE): S736's "keep
            // tolerance" is not a constant. The page-bottom fit test for an
            // EMPTY paragraph runs on the line height BEFORE the lineRule=auto
            // multiplier is applied, so exactly the leading that multiplier
            // added — and nothing else — may hang past the content bottom.
            // DERIVED with _pb_emptytail_gen.py + _pb_emptytol_gen.py (Word COM;
            // a shim paragraph of EXACT line height at the top of each arm walks
            // the whole stack across the content bottom in 0.25pt steps):
            //   * FIVE successor variants (nothing follows / 1-line / 4-line /
            //     4-line with widowControl, whose successor always moves / a
            //     second empty) flip at the SAME overflow ⇒ the successor plays
            //     no part. That falsifies group-movement, page-terminal and
            //     ink-vs-empty, which the four specimens had suggested.
            //   * 14 combos, monotone throughout, every flip window containing
            //     lh − natural and nothing else:
            //       Calibri 11 x1.000  (-0.23, 0.02] ∋ 0.00
            //       Calibri 11 x1.079  ( 0.98, 1.23] ∋ 1.06
            //       Calibri 11 x1.150  ( 2.01, 2.26] ∋ 2.01
            //       Calibri 11 x1.500  ( 6.51, 6.76] ∋ 6.71
            //       Calibri 22 x1.000  (-0.23, 0.02] ∋ 0.00
            //       Calibri 22 x1.150  ( 4.02, 4.27] ∋ 4.03
            //       Arial   12 x1.150  ( 2.01, 2.26] ∋ 2.07
            //       TNR     12 x1.500  ( 6.77, 7.02] ∋ 6.90
            //     A constant is refuted by 6 of the 8 auto arms; k×size by the
            //     x1.000 arms and by Times.
            //   * atLeast 18pt, exact 18pt and exact 10pt all flip at
            //     (0.00, 0.25] = the FULL box, no leniency at all — the
            //     multiplier is the only source of hangable leading.
            //   * the empty's own w:after is NOT part of the test (after=200
            //     leaves the flip point untouched).
            // Scope: no typed docGrid. The 2.5 constant was fitted to
            // probexempty, whose grid is a different mechanism — its recorded
            // flip (1.8, 2.1] brackets the S739 centered box (pitch+natural)/2
            // = 1.875, not 2.5 — so the JP/grid side keeps the old path and is
            // byte-identical here.
            // atLeast zero bypasses grid snapping when the line is measured.
            // Its fit test therefore uses the full natural box even when the
            // section declares a grid; there is no grid leading to overhang.
            let empty_atleast_natural = para.style.line_spacing_rule.as_deref() == Some("atLeast")
                && para.style.line_spacing == Some(0.0)
                && std::env::var_os("OXI_EMPTY_ATLEAST_NATURAL_DISABLE").is_none();
            let s1113_pre_mult = s736_empty_tol
                && std::env::var("OXI_S1113_DISABLE").is_err()
                && (grid_pitch.is_none() || page.doc_grid_no_type || empty_atleast_natural);
            // Set when the max-chain below actually selects the centered box,
            // so S1154 can withdraw the half-twip tolerance in exactly that case.
            let mut centered_box_is_threshold = false;
            let mut brk_branch = "else";
            let break_threshold = if s1113_pre_mult {
                brk_branch = "s1113_pre_mult";
                // auto carries a factor, atLeast/exact carry a length: only the
                // former inflates the box, so only the former is divided out.
                let factor = match para.style.line_spacing_rule.as_deref() {
                    None | Some("auto") => para.style.line_spacing.unwrap_or(1.0).max(1.0),
                    _ => 1.0,
                };
                (effective_lh / factor).max(0.0)
            } else if s736_empty_tol {
                brk_branch = "s736_empty_tol";
                if std::env::var("OXI_S1041_DISABLE").is_err() {
                    (effective_lh - s736_tol).max(0.0)
                } else {
                    (effective_lh - s736_tol).max(ink_lh)
                }
            } else if s548b_exact_full || s562b_empty_full
                || s603_typed_fullbox || s605_line0_2 || s_tgfull || s651_multicell_head
                // S726: footer-constrained bottom → full box (the leniency's
                // leading overhang would land inside footer text; Word breaks).
                || footer_tight
            {
                brk_branch = if s548b_exact_full { "s548b" }
                    else if s562b_empty_full { "s562b" }
                    else if s603_typed_fullbox { "s603" }
                    else if s605_line0_2 { "s605" }
                    else if s_tgfull { "s_tgfull" }
                    else if s651_multicell_head { "s651" }
                    else { "footer_tight" };
                if footer_tight
                    && !s548b_exact_full
                    && !s603_typed_fullbox
                    && !s605_line0_2
                    && !s_tgfull
                    && !s651_multicell_head
                    && !page.doc_grid_no_type
                    && para.style.snap_to_grid
                    && grid_pitch.is_some_and(|pitch| pitch > 0.0)
                    && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                {
                    // A grid line places half of its added leading below the
                    // natural line box, including above an occupied footer.
                    centered_box_is_threshold = true;
                    effective_lh.min((effective_lh + natural_lh) / 2.0)
                } else {
                    effective_lh
                }
            } else {
                // S688 PROBE/SCAFFOLD (2026-06-28, default 0 = byte-identical, opt-in
                // OXI_TGINK_K=<pt>): the typed-grid page-bottom break threshold (= natural_lh,
                // the win*83/64 SPACING box, 13.617 for MS Mincho 10.5pt) UNDER-estimates the
                // TRUE rendered ink bottom from the cursor by ~1.0pt. MEASURED on tokyoshugyo
                // «る賃金» (p46 mid-para continuation, pi=444): cursor 742.55, Word baseline
                // 755.70 (Oxi baseline IDENTICAL — no cursor drift, confirmed via --dump-glyphs),
                // glyph descent 1.477 → ink bottom 757.18 > content_bottom 756.85 → Word BREAKS;
                // natural_lh 13.617 gives cursor+13.617 = 756.17 ≤ 756.85 → Oxi KEEPS. The
                // baseline sits LOW in the 18pt cell (leading 4.1pt ABOVE the ink), so the real
                // ink bottom from cursor = baseline_offset(13.15) + descent(1.477) = 14.63, NOT
                // 13.617. The principled K ≈ true_ink_bottom − natural_lh − Word_tolerance
                //   = 14.63 − 13.617 − 0.19 ≈ 0.82 (sweep flips at K∈(0.80, 0.85]).
                // ★Canaries db9ca/ikujikaigo/model/3a4f/roudoujoken/34140 ALL stay PASS at K≤2.0
                //   (the typed-grid bump has NO canary risk — supersedes the S576/S687 fear).
                // ★NOT shipped: a uniform K can't fix tokyoshugyo. It is a DOC-WIDE BIDIRECTIONAL
                //   distributed wall (~41 reliable page-top divergences, both Oxi-ahead AND
                //   Oxi-behind), and breaking «る賃金» alone over-corrects the chapter via the
                //   coupled empty-para/section-heading SPACING drift (Oxi's «（）»-heading gap is
                //   +1.9pt mean vs Word; the «（基本給）» section renders ~18.25/increment vs Word
                //   18.10, accumulating ~+0.6/section). K=0.85 nets reliable 41→39 only. The
                //   coupled fix needs the threshold (−1) AND the spacing drift (+1) together.
                //   See [[tokyoshugyo_wrap_not_cellheight]].
                // S690 (2026-06-29) FINDING: the ink-threshold is CONFIRMED NOT PRODUCTIVE
                // for tokyoshugyo AFTER the body fix (S689/S590 default-ON). Post-S690 the
                // reliable page-top is 36 at BOTH K=0 and K=0.85 — the «る賃金» break does
                // NOT reduce real divergences (the body 約物 capacity break already captured
                // that region); the gate {-1:7,1:30} it produces is PURE text-prefix ARTIFACT.
                // AND the kojin coupling is NOT separable by continuation-vs-para-start: a
                // continuation-line scope (line_idx>0) STILL regressed kojin (its flips are
                // continuation CASCADES, not para-starts — the para_idx detection misled).
                // ⇒ the ink-threshold under-estimate is real but a DEAD pagination lever for
                // tokyoshugyo post-body-fix; OXI_TGINK_K kept as a default-0 byte-identical
                // probe only. The p47 賃金-region divergence is a DIFFERENT cause (re-examine).
                let tgink_k = if !page.doc_grid_no_type {
                    std::env::var("OXI_TGINK_K")
                        .ok()
                        .and_then(|v| v.parse::<f32>().ok())
                        .unwrap_or(0.0)
                } else {
                    0.0
                };
                // S714 (2026-07-02, default ON, opt-out OXI_S714_DISABLE): a
                // lineRule=atLeast line demands its EXPLICIT box (the atLeast value)
                // at the page bottom — the ink/natural leniency only forgives GRID
                // slack above the box, not the author's explicit minimum (the
                // exact/empty analogy: S548b / S562b). tokyoshugyo 〔例３〕 heading
                // (atLeast 350=17.5): Word pushes it to p31; the 13.5 font-natural
                // leniency wrongly kept it at the p30 bottom (over −1.85 vs +2.15).
                // Scoped to typed docGrid snap_to_grid paragraphs.
                let atleast_floor = if std::env::var("OXI_S714_DISABLE").is_err()
                    && ((!page.doc_grid_no_type && para.style.snap_to_grid)
                        || (std::env::var("OXI_ATLEAST_BODY_FLOOR_DISABLE").is_err()
                            && !self.doc_body_has_real_cjk))
                    && para.style.line_spacing_rule.as_deref() == Some("atLeast")
                {
                    para.style.line_spacing.unwrap_or(0.0)
                } else {
                    0.0
                };
                // S739 (2026-07-04, default ON, opt-out OXI_S739_DISABLE): a
                // paragraph's FIRST line in a typed docGrid uses the CENTERED-BOX
                // page-bottom threshold — the glyph box is vertically centered in
                // its grid slot, and Word requires the CENTERED box's bottom
                // (slot_top + (pitch + natural)/2) to fit the content area. Only
                // the TOP half of the grid leading may hang past the bottom, not
                // the full leading (the Day-33 natural rule was too lenient for
                // first lines by (pitch−natural)/2 ≈ 1-2.2pt).
                // ★DERIVED from a controlled capacity sweep (_pb_capacity_gen.py:
                // uniform single-line paras, linePitch {312,360} × fs {9,10.5,11,
                // 12}pt × bottom-margin sweep 1300..1520/20 × top-margin sweep —
                // 38 Word COM measurements): every per-page line-capacity flip
                // lands EXACTLY on the centered-box boundary (knife-edge checks:
                // p312/b1400 771.91 vs bottom_eff 771.90 → push ✓; p360 b1420 keep
                // /b1440 push bracket 770.75 ✓; fs sweep 8/8 ✓). Pure-natural and
                // full-box models both violate multiple sweep points.
                // ★EXCEPTION (prev_keep_next): the follower of a keepNext
                // paragraph keeps the LENIENT natural test — Word places heading
                // + ≥1 follower line at the page bottom (probekeepnext p2: body
                // line0 at slot 756.15, centered 771.9 > 771 would push, Word
                // KEEPS; the non-follower probelac line0 at 757.3 is PUSHED).
                // ★CONTINUATION lines are NOT touched (kojin pi=52 nat_over −0.35
                // LAST line Word KEEPS — the corpus leniency keeps live on
                // continuation/last lines; S693's non-last hairline still applies).
                // Fixes probekeepnext (heading nat_over −0.44 → centered +1.2 →
                // push = Word) and probelac (line0 centered → 44 lines/page = Word).
                // Line scope: the paragraph's EDGE lines (first AND last) use the
                // centered-box rule; INTERIOR lines keep the natural leniency
                // (probekeepnext body line1 at slot 756: centered 771.8 > 771
                // would push, Word KEEPS — interior confirmed natural). The
                // m-sweep (3-line paras, y_p2 straddle discriminator) pinned the
                // LAST line to the same centered flip as line0 (p312 keep@14.7 /
                // push@14.5, centered 14.61 inside; natural 13.62 excluded).
                // compat 11/14/15 all measured IDENTICAL (compat is NOT a
                // discriminator here).
                // S1155 (2026-08-16, default ON 2026-08-26 (opt-out `OXI_S1155_DISABLE`), bundle promotion):
                // the edge condition is not one either. Re-running the whole
                // _pb_lastline fine sweep with the test paragraph wrapped to
                // THREE lines -- same cursor, same slack, but the boundary line
                // is now index 1 of 3, i.e. NON-last -- reproduces the 2-line
                // table exactly: keep at boxover 2.950, SPLIT at 3.050, all four
                // phases. So the centered box is the page-bottom threshold for
                // every typed-grid line. That also subsumes S693: its non-last
                // arm gives the full box for natslack < 1.0 (same verdict as
                // centered) and the INK leniency for natslack in (1.0, 2.9875),
                // where this sweep shows Word breaking. On that 3-line probe the
                // default is 0/32 against Word (32 keeps where Word keeps 12);
                // with S1155 it is 32/32.
                // ★NOT default yet — ONE blocker: Phase 1 95 -> 94, tokyoshugyo
                // PASS -> FAIL (wi 1152/1177, +1 page). The new break is at
                // pi=205 line 2, and it is a KNOCK-ON, not a wrong call: Word
                // puts ONLY line 0 of pi=205 on p26 while Oxi puts three, i.e.
                // that page is already 21.63pt high before this rule runs.
                // `_kojin_rowgeom.py scan` shows the whole document zig-zagging
                // (+19 on p4-5, +21 on p18-19, -17 from p21, -35 on p27-28), and
                // its HEAD is p4: an overlong Latin token (an e-gov URL). Word
                // moves the whole token to the next line and only then breaks it
                // at the margin; Oxi starts it on the partly-filled line and then
                // breaks it at a punctuation opportunity, leaving line 2 a
                // quarter empty -- one extra line, and every later page inherits.
                // Fix that first; this rule is derived and waiting.
                let s739_edge = (line_idx == 0 && !prev_keep_next)
                    || (line_idx > 0 && line_idx + 1 == lines.len())
                    || std::env::var("OXI_S1155_DISABLE").is_err();
                // ★ON-SLOT gate — FALSIFIED 2026-08-16 (S1153, removed by
                // default, opt-back-in OXI_S1153_DISABLE). Its only evidence was
                // kojin pi=52 "Word keeps at nat_over −0.35", and that line was
                // sitting 13.89pt below where Word puts it because of the
                // mid-line LRPB S1151 fixed; there is no such Word decision.
                // _pb_lastline_gen.py sweeps slot phase and slack INDEPENDENTLY
                // (the spacer sets the phase, the section's bottom margin sets
                // the slack) and Word flips at the SAME slack for every phase:
                //   phase 0.00 / 4.10 / 8.20 / 12.25 pt, all four
                //   keep at boxover 2.950, SPLIT at boxover 3.050
                // and the centered box predicts (BOX−NAT)/2 = 2.9875, i.e.
                // inside that 0.1pt bracket. So the centered box is the rule at
                // EVERY phase and the gate only kept the leniency alive off
                // slot. 56 coarse + 32 fine arms, Word's own PDF.
                // Gate: _pb_cjk2line 2/6 -> 6/6 and _pb_lastline 29/32 (32/32
                // with S1154), Phase 1 95/96 with zero per-doc change, and all
                // 238 SSIM sentinel documents BYTE-IDENTICAL — no corpus page
                // currently sits off-slot at the centered-box boundary, which
                // is why the gate could survive this long.
                // S1375 (2026-09-13, default ON, opt-out OXI_S1375_DISABLE): the
                // last line before an EMPTY SECTION-END paragraph is kept by the
                // natural (Day-33) test, not the centred one. MEASURED
                // (`tests/fixtures/empty_edge`, linesAndChars 325, 8pt): with a
                // plain MARK following, the probe is pushed once less than
                // ~13.5pt remains (the centred value 13.31); with an empty
                // sectPr paragraph following -- plain, after a table, framed or
                // not -- it is kept down to 10.35pt left and pushed at 9.85
                // (natural 10.38). forms__00830ac053a2c57a's page 6 ends in
                // such a pair: the centred test pushed the empty line to a
                // blank page 7 and the section start to page 8.
                let s739_centered = if std::env::var("OXI_S739_DISABLE").is_err()
                    && s739_edge
                    && !s1375_before_section_end
                    && !page.doc_grid_no_type
                    && para.style.snap_to_grid
                    && grid_pitch.map_or(false, |p| p > 0.0)
                {
                    let pitch = grid_pitch.unwrap_or(0.0);
                    let phase = (cursor.cursor_y - page_top).rem_euclid(pitch);
                    let on_slot = phase < 1.0 || phase > pitch - 1.0;
                    if on_slot || std::env::var("OXI_S1153_DISABLE").is_err() {
                        (effective_lh + natural_lh) / 2.0
                    } else {
                        0.0
                    }
                } else {
                    0.0
                };
                let s779_floor = if s779_latin && !no_type_multiple_ink {
                    s779_win_heights.get(line_idx).copied().unwrap_or(0.0)
                } else if no_type_multiple_ink && s1079_natural {
                    // S1079: the floor is the natural (unmultiplied) line.
                    natural_line_heights.get(line_idx).copied().unwrap_or(0.0)
                } else {
                    0.0
                };
                // S1194 (2026-08-22, opt-in `OXI_S1194`): S827's hhea line is the
                // page-bottom capacity ITSELF, not a floor under the ink box.
                // The derivation it cites (`_pb_latinbot_gen` TNR 12pt, and
                // `_pb_latinbot_cal` Calibri 11pt at 2tw, which separates the
                // full-box model from the baseline+typo_desc one and picks the
                // box) reads "keep iff line_top + hhea <= content_bottom" —
                // an equality, in both directions. Entering it as `.max()`
                // only ever makes Oxi STRICTER, which was the direction nyserda
                // needed; where the ink box is the WIDER of the two the floor is
                // inert and Oxi stays stricter than Word. Times New Roman is
                // that case: word_asc+word_desc = 1.200em against hhea 1.14990em,
                // so at 10pt Oxi reserves 12.000 where Word reserves 11.499 and
                // rejects a last line Word keeps by 0.11pt (00501ca3 p8 pi=134
                // «(Added 2002) (Amended 2010)»: cursor_y 708.110, cbot 720.000).
                // Same shape as S576 on the CJK side — the spacing box is not
                // the capacity.
                let s1194_hhea = std::env::var("OXI_S1194_DISABLE").is_err()
                    && s779_latin
                    && !no_type_multiple_ink;
                let base = if s1194_hhea && s779_floor > 0.0 {
                    s779_floor
                } else {
                    ink_lh + tgink_k
                };
                let v = base
                    .max(atleast_floor)
                    .max(s739_centered)
                    .max(s779_floor)
                    .min(effective_lh);
                centered_box_is_threshold = s739_centered > 0.0
                    && (v - s739_centered).abs() < 1e-6;
                // S1438: a line that carries ruby needs its natural box plus the
                // ruby expansion at the page bottom (proberuby vs its plain twin).
                let s1438_ruby = std::env::var_os("OXI_S1438_DISABLE").is_none()
                    && ruby_para_expansion_pt > 0.0
                    && lines.get(line_idx).map_or(false, |l| {
                        l.fragments.iter().any(|fragment| {
                            para.runs
                                .get(fragment.run_index)
                                .map_or(false, |r| r.ruby.is_some())
                        })
                    });
                let v = if s1438_ruby {
                    let exp_line = if std::env::var_os("OXI_S1641_DISABLE").is_none() {
                        lines.get(line_idx).map_or(ruby_para_expansion_pt,
                            |l| self.s1641_line_ruby_expansion(l, para, para_font_size))
                    } else {
                        ruby_para_expansion_pt
                    };
                    v.max(natural_lh + exp_line)
                } else {
                    v
                };
                if std::env::var("OXI_DBG_PB").is_ok() {
                    eprintln!("[PB-RUBY] ruby_line={} exp={:.2}", s1438_ruby, ruby_para_expansion_pt);
                    let head: String = lines.get(line_idx).map(|l| l.fragments.iter().flat_map(|f| f.text.chars()).take(16).collect()).unwrap_or_default();
                    eprintln!("[PB] cy={:.2} bottom={:.2} eff={:.2} nat={:.2} ink={:.2} c739={:.2} v={:.2} sect_end_next={} bi={:?} «{}»",
                        cursor.cursor_y, page_top + content_height, effective_lh, natural_lh, ink_lh, s739_centered, v, s1375_before_section_end, body_para_index, head);
                }
                v
            };
            // R7.53: first-line lenient check using `first_line_extra_content_h`.
            // S168 Phase B-2 (c): per-line lenient.
            let line_lenient_extra = if !para_fn_heights.is_empty() {
                let committed = committed_fn_delta_at_line
                    .get(line_idx)
                    .copied()
                    .unwrap_or(0.0);
                (first_line_extra_content_h - committed).max(0.0)
            } else if line_idx == 0 {
                first_line_extra_content_h
            } else {
                0.0
            };
            let effective_break_bottom = page_top + content_height + line_lenient_extra;
            // S693 (2026-06-29, default ON, opt-out OXI_S693_DISABLE): a NON-LAST
            // typed-grid line whose natural-leniency over is a HAIRLINE (nat_over >
            // -1.0pt, the full box barely overflows the content bottom) breaks at the
            // FULL grid cell. The over-hairline gate is the S693-commit's pi=106 fix:
            // a comfortable-margin non-last line (over < -1.0) keeps the leniency.
            // Two-part discriminator derived from BR_DUMP keep/break vs Word PDF:
            //   (1) last-vs-non-last: kojin pi=52 (over -0.35, LAST line) Word KEEPS;
            //       tokyoshugyo «給月給» (line 2/5, over -0.80, non-last) Word BREAKS.
            //   (2) WITHIN non-last lines, over still matters: pi=205 «〔例２〕»
            //       (non-last, over -1.10) Word KEEPS, «給月給» (over -0.80) BREAKS —
            //       a non-last line with comfortable margin (over < -1.0) keeps the
            //       leniency; only a hairline non-last line (over in (-1.0, 0)) breaks.
            // OXI_S693_OV overrides the -1.0 threshold for sweeping.
            // S1152 (2026-08-16, opt-in OXI_S1152=1) — FALSIFIED as a blanket
            // rule, kept as the lever the scoped version will reuse. It makes a
            // typed grid use the FULL grid box at the page bottom for EVERY
            // line, last one included.
            // Why it was tried: both observations S693's last-line leniency
            // rests on were taken at drifted positions and are void. kojin
            // pi=52 sat 13.89pt low behind the mid-line LRPB (S1151), and
            // tokyoshugyo pi=205 sits 21.63pt low on p26 — Word puts only its
            // line 0 on that page, so its «over -0.60 KEEP» is not a decision
            // Word ever made. What survives is _pb_cjk2line_gen.py: 6 arms,
            // self-authored, and Word splits 1+1 in EVERY one where Oxi's
            // natural height still fits by 2.675pt. With S1152 the probe goes
            // 2/6 -> 6/6.
            // Why it cannot ship as written: Phase 1 95 -> 90 (34140, db9ca,
            // ohnochingin, roudoujoken, tokyoshugyo all PASS -> FAIL), so the
            // leniency is real for those pages and the probe's regime is
            // narrower than "any typed-grid last line". The probe differs from
            // them in BOTH slot phase (its exact-height spacer leaves the
            // cursor off-slot) and natural/box ratio (10.375/16.35 = 0.63 vs
            // 13.5/18 = 0.75); separating those needs a phase x slack sweep,
            // which moves the bottom margin (slack alone) independently of the
            // spacer (phase and slack together).
            let s1152_full_box = std::env::var("OXI_S1152").ok().as_deref() == Some("1")
                && !page.doc_grid_no_type
                && !s548b_exact_full
                && !s562b_empty_full;
            // S1484 (2026-09-19, default ON, opt-out OXI_S1484_DISABLE): with
            // S1483 a CJK document without a line grid advances by the EXACT
            // line box, so the page-bottom fit must ask for that same box.
            // The S576 ink box / 0.5pt-quantized natural (13.5 for MS Mincho
            // 10.5pt against a 13.617 box) was calibrated on the S571 pitch and
            // now lets a line stay with its box past the content bottom
            // (policies__094c44cd p39 「例示」 over=-0.056 on 13.5, +0.06 on
            // 13.617; Word pushes). `bottomlimit2.py`: Word's limit is exactly
            // "box top + advance <= content bottom" for auto/atLeast lines of
            // every face and size; exact lines and empty paragraphs keep their
            // own laws (S548b / S562b / S1113), and a multiplied auto line keeps
            // the ink leniency (the multiplier's leading may hang, S576's case).
            let s1484_full_box = std::env::var_os("OXI_S1484_DISABLE").is_none()
                && self.doc_body_has_real_cjk
                && (grid_pitch.is_none() || page.doc_grid_no_type)
                && !s548b_exact_full
                && !s562b_empty_full
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && para.style.line_spacing.unwrap_or(1.0) <= 1.0;
            // S1617 (2026-09-30, opt-in OXI_S1617=1) -- FALSIFIED, kept only as a lever.
            // It asked for the FULL line box at a typed-grid page bottom once a table
            // sits on the page. The table probe that seemed to show it
            // (`_pb_gridbottom_tbl_gen.py`) was really showing S1618: Word starts the
            // block after a table 0.75 lower than Oxi did, and with that width restored
            // the ordinary leniency explains every arm. Default ON it pushed five last
            // lines that Word keeps (1245d99e, 13abeaf6, 01c5a769, 03704f36,
            // ohnochingin_02).
            let s1617_table_on_page = std::env::var("OXI_S1617").as_deref() == Ok("1")
                && grid_pitch.is_some()
                && !page.doc_grid_no_type
                && page.grid_char_pitch.is_none()
                && !s548b_exact_full
                && !s562b_empty_full
                && current_elements.iter().any(|e| e.cell_row_index.is_some());
            let break_threshold = if s1152_full_box || s1484_full_box || s1617_table_on_page {
                effective_lh
            } else if s693_nonlast {
                let nat_over = cursor.cursor_y + natural_lh - effective_break_bottom;
                let ov_thr = std::env::var("OXI_S693_OV")
                    .ok()
                    .and_then(|v| v.parse::<f32>().ok())
                    .unwrap_or(-1.0);
                if nat_over > ov_thr {
                    effective_lh
                } else {
                    break_threshold
                }
            } else {
                break_threshold
            };
            // S832 (2026-07-13, opt-out OXI_S832_DISABLE): a TRAILING-BR EMPTY
            // line (the S684 line — a paragraph-final <w:br/> renders an empty
            // line) at the PAGE BOTTOM stays on the current page. Word starts
            // the next page with the FOLLOWING paragraph (nyserda p54:
            // «…for Reports.¶<br>» then the Option-2 heading at p55 y=72.5 =
            // the page top; Oxi pushed the empty br line → +26pt at the p55
            // top → the wp55 ×2 tail in BOTH LRPB modes). Latin scope; JP
            // byte-identical by construction.
            let s832_trailing_empty = line_idx > 0
                && line_idx + 1 == lines.len()
                && line.fragments.iter().all(|f| f.text.trim().is_empty())
                && !self.doc_body_has_real_cjk
                && std::env::var("OXI_S832_DISABLE").is_err();
            // S835 (2026-07-14, default ON, opt-out OXI_S835_DISABLE): the
            // FOOTNOTE-AREA top is a SOFT boundary — Word lets the body line's
            // box enter the separator paragraph's leading by 1/16 em (fs/16)
            // before pushing. DERIVED on uk_framework wp26 (Calibri 11): the
            // area chain is locked by the sep-rule baseline (rule center 722.98
            // = area_top 714.82 + asc 8.25; keep-all-afters + Normal-hhea sep),
            // and the body flip L = 0.71 ± 0.05 (W8 spacer ladder ×2 + W9
            // bottom-margin ladder keep→push 1144→1146) = fs/16 = 0.6875@11pt.
            // The ceil-0.75 device-slot alternative was REJECTED by W9 (predicts
            // push at 1136; Word keeps through 1144). The PLAIN page bottom
            // keeps the FULL hhea box (S827; Calibri probe lbc_1240→1242 =
            // box-exact — the relief is fn-boundary-ONLY). Scope: Latin +
            // fn-ref-carrying para (the derived case; a no-own-ref para above
            // a committed area is not yet covered).
            let s835_boundary_is_fn = fn_boundary_active
                || committed_fn_delta_at_line
                    .get(line_idx)
                    .copied()
                    .unwrap_or(0.0)
                    > 0.0;
            // (S889 footer-boundary relief ATTEMPTED + FALSIFIED 2026-07-17:
            // the _pb_ftext "+0.64 uniform bias" that suggested it was a
            // probe ARITHMETIC BUG — the model constant was 697.25 where
            // 841.9−144 = 697.9; with the correct constant the blank-footer
            // flip sits at stack 0 EXACTLY (no relief) and the hhea stack
            // model needs no companion. A footer-part relief also regressed
            // uklocalspending 1.0→0.59 / ukframework 1.0→0.97, whose
            // S806-S835 fine probes close at ±0.05 with NO relief term.)
            let exact_fn_separator = self.fn_special_declared
                && std::env::var("OXI_EXACT_FN_SEPARATOR_DISABLE").is_err()
                && page.footnotes.iter().find(|n| n.number == u32::MAX)
                    .and_then(|n| n.blocks.iter().find_map(|b| match b {
                        Block::Paragraph(p) => Some(p),
                        _ => None,
                    }))
                    .is_some_and(|p| p.style.line_spacing_rule.as_deref() == Some("exact"));
            // A typed line grid allocates complete body slots above notes.
            // Its capacity is independent of the ink's centered position or
            // the body's exact/atLeast/auto rule. Natural leading relief does
            // not apply to this slot boundary.
            let typed_grid_fn_boundary = s835_boundary_is_fn
                && grid_pitch.is_some() && !page.doc_grid_no_type;
            // Word's modern compatibility mode reserves the complete advance
            // of a body paragraph following an already committed note area.
            // Modes 12 and 14 retain the natural leading: a fresh 24-arm
            // Word comparison changes only compatibility mode and grid presence.
            // Reference-bearing paragraphs retain their own marker/box rule.
            let modern_committed_note_slot = s835_boundary_is_fn
                && fn_boundary.automatic_numbering
                && para_fn_heights.is_empty()
                && self.compat_mode >= 15 && self.compat_mode_explicit
                && !self.doc_body_has_real_cjk
                && std::env::var_os("OXI_MODERN_FN_CAPACITY_DISABLE").is_none();
            let s835_fn_relief = if s835_boundary_is_fn
                && !modern_committed_note_slot
                // Untyped stories retain their natural last-line leading.
                // Typed grids allocate complete slots independently of ink.
                && !typed_grid_fn_boundary
                && !exact_fn_separator
                && !self.doc_body_has_real_cjk
                && std::env::var("OXI_S835_DISABLE").is_err()
            {
                para_font_size / 16.0
            } else {
                0.0
            };
            // S967 (2026-07-21, opt-out OXI_S967_DISABLE): OOXML coordinates are
            // twips, so a page-bottom overshoot below HALF A TWIP is not a real
            // overshoot. S926 already grants exactly this tolerance to the orphan
            // lookahead (12120); the natural comparison kept a raw float `>`, so
            // policies__00148f8d wi=906 broke on over=+0.013pt (0.26 twip) where
            // Word keeps the line. Census over all 200 EN corpus documents:
            // only THREE natural breaks fall in 0 < over < 0.025 (this one, plus
            // legal__0010437a pi=415 and technical__002c6778 pi=121, both in
            // already-failing documents) — so the blast radius is three lines,
            // and round-to-twip and half-twip tolerance agree at all of them.
            // Same constant as S926: no second magic number.
            // S1154 (2026-08-16, default ON, opt-out OXI_S1154_DISABLE): the
            // centered box carries NO such tolerance. Its own value is a
            // quarter-twip quantity ((327+207.5)/2 = 267.25tw for the probe
            // face), so a half-twip slop swallows the whole flip: on
            // _pb_lastline's 32 fine arms Word splits at over +0.0125 and the
            // tolerance made Oxi keep 3 of them. Withdrawing it there takes the
            // probe 29/32 -> 32/32; S967's own derivation (policies__00148f8d
            // wi=906, Word KEEPS at +0.013) is a Latin natural-threshold case,
            // which this leaves alone. Gate: Phase 1 95/96 unchanged and all
            // 238 SSIM sentinel documents byte-identical, same as S1153.
            // Private Word-measured candidate: automatic multiples must fit their
            // natural final line exactly. The atLeast half-twip rule remains separate.
            let multiple_exact_fit = std::env::var("OXI_MULTIPLE_FIT_NO_TOL_DISABLE").is_err()
                && !self.doc_body_has_real_cjk && page.grid_line_pitch.is_none()
                && !s835_boundary_is_fn
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && para.style.line_spacing.is_some_and(|f| f > 1.0);
            // Automatic Latin line boxes must fit at both the line and orphan checks.
            let automatic_exact_fit = std::env::var("OXI_AUTO_LINE_FIT_DISABLE").is_err()
                && !self.doc_body_has_real_cjk && page.grid_line_pitch.is_none()
                && !s835_boundary_is_fn
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"));
            // Natural CJK body capacity and exact advance share the same box.
            // Do not add the atLeast half-twip allowance to this auto box.
            let cjk_exact_fit = std::env::var("OXI_CJK_EXACT_BODY_CAPACITY").is_ok()
                && (grid_pitch.is_none() || page.doc_grid_no_type)
                && !s835_boundary_is_fn
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && para.style.line_spacing.unwrap_or(1.0) <= 1.0
                && (break_threshold - natural_lh).abs() < 1e-6
                && lines[line_idx].fragments.iter().any(|fragment| {
                    fragment.text.chars().any(kinsoku::is_cjk_ideograph_or_kana)
                });
            // Exact twip inputs still accumulate binary f32 subtraction noise.
            // Absorb only a couple of coordinate ULPs, not a physical half-twip
            // allowance: the measured automatic separator differs from an
            // exact separator by much less than one twip and must still break.
            let fn_coordinate_roundoff = 2.0 * f32::EPSILON * effective_break_bottom.abs();
            let s967_tol = if typed_grid_fn_boundary {
                fn_coordinate_roundoff
            } else if std::env::var("OXI_S967_DISABLE").is_ok()
                || cjk_exact_fit
                || automatic_exact_fit
                || multiple_exact_fit
                || (centered_box_is_threshold
                    && std::env::var("OXI_S1154_DISABLE").is_err())
            {
                0.0
            } else {
                0.025
            };
            // A reference line reserves its note together with its multiplied
            // text box. Apply the same height and reservation to look-ahead.
            let footnote_fit_height = |idx: usize, threshold: f32| -> f32 {
                let above_notes = committed_fn_delta_at_line.get(idx).copied().unwrap_or(0.0) > 0.0;
                // Grid slots use the full advance. Without a typed grid,
                // retain the natural last-line threshold, including its
                // actual run metrics rather than imposing an extra body slot.
                if (fn_boundary_active || above_notes)
                    && (typed_grid_fn_boundary || modern_committed_note_slot) {
                    return line_heights.get(idx).copied().unwrap_or(threshold);
                }
                if std::env::var("OXI_FOOTNOTE_REF_FIT_DISABLE").is_ok()
                    || !above_notes
                    || self.doc_body_has_real_cjk
                    // A linePitch without a grid type is an untyped story:
                    // use the natural font box and note reservation, as for noGrid.
                    || (page.grid_line_pitch.is_some() && !page.doc_grid_no_type)
                {
                    return threshold;
                }
                let factor = match para.style.line_spacing_rule.as_deref() {
                    None | Some("auto") => para.style.line_spacing.unwrap_or(1.0).max(1.0),
                    _ => 1.0,
                };
                let hhea = s779_win_heights.get(idx).copied().unwrap_or(threshold);
                (hhea * factor).max(threshold)
            };
            let break_threshold = if typed_grid_fn_boundary {
                effective_lh
            } else {
                footnote_fit_height(line_idx, break_threshold)
            };
            let footnote_fit_bottom = |idx: usize| -> f32 {
                let bottom = page_top + content_height;
                if std::env::var("OXI_FOOTNOTE_REF_FIT_DISABLE").is_ok()
                    || self.doc_body_has_real_cjk
                    // A linePitch without a grid type is an untyped story:
                    // use the natural font box and note reservation, as for noGrid.
                    || (page.grid_line_pitch.is_some() && !page.doc_grid_no_type)
                {
                    return bottom;
                }
                let extra = if !para_fn_heights.is_empty() {
                    let committed = committed_fn_delta_at_line.get(idx).copied().unwrap_or(0.0);
                    (first_line_extra_content_h - committed).max(0.0)
                } else if idx == 0 {
                    first_line_extra_content_h
                } else {
                    0.0
                };
                bottom + extra
            };
            // S1248 (default ON, opt-out OXI_S1248_DISABLE): on a page whose foot is spoken
            // for by notes, a paragraph's TRAILING SPACE has to fit above the note
            // area as well as its last line. Measured both ways on the real
            // document (`_pb_18715_gap`, sweeping the paragraph's own w:after with
            // its position fixed: Word keeps at 0 and pushes at 0.8pt, exactly
            // where line_bottom + after crosses Oxi's own content bottom) and on a
            // self-authored repro. At a PLAIN page bottom the same sweep does NOT
            // move the flip (`_pb_botafter` with no notes: after 0 / 8 / 16pt all
            // flip at the same spacer), so the term is footnote-scoped -- there the
            // trailing space simply falls off the page.
            // Like the roll rule (S1244) this is compatibilityMode >= 15 only: the
            // same repro with no settings.xml flips at one spacer for after 0, 8
            // and 16pt, where at 15 the flip moves by exactly the after value.
            let s1248_modern = self.compat_mode >= 15 && self.compat_mode_explicit;
            let s1248_para_after = if std::env::var("OXI_S1248_DISABLE").is_err()
                && ((s835_boundary_is_fn && s1248_modern)
                    || s758_band.is_some_and(|(bottom, _, _)| cursor.cursor_y < bottom))
                && !self.doc_body_has_real_cjk
            {
                para.style.space_after.unwrap_or(0.0)
            } else {
                0.0
            };
            // The term belongs to the paragraph's LAST line, so the per-line test
            // takes it at line_idx == len-1 and the widow look-ahead (which asks
            // about line_idx+1) at len-2.
            let s1248_after = if line_idx + 1 == lines.len() {
                s1248_para_after
            } else {
                0.0
            };
            let s1248_next_after = if line_idx + 2 == lines.len() {
                s1248_para_after
            } else {
                0.0
            };
            // Moving an unavoidable first line from an empty page only creates
            // an empty page. Keep it here; subsequent paragraphs still paginate.
            let oversized_first_line = line_idx == 0
                && elements.is_empty()
                && current_elements.is_empty()
                && (cursor.cursor_y - page_top).abs() < 0.025
                && effective_lh > effective_break_bottom - page_top + 0.025;
            // A source boundary without line content ends the current flow
            // region. It does not first need room for an empty painted row.
            let empty_hard_control=line.fragments.iter().all(|fragment|fragment.text.is_empty())
                && matches!(line.break_type,LineBreakType::PageBreak|LineBreakType::ColumnBreak);
            let mut natural_needs_page_break = if in_textbox || s832_trailing_empty || oversized_first_line || empty_hard_control {
                false
            } else {
                cursor.cursor_y + break_threshold + s1248_after - s835_fn_relief
                    > effective_break_bottom + s967_tol
            };
            // Opt-in trace of the natural-break arithmetic. Every term of this
            // one comparison decides which page a line lands on, and reading
            // them off a running document is otherwise a matter of guessing
            // which of `break_threshold` / `s1248_after` / `effective_break_bottom`
            // is the odd one out.
            if std::env::var("OXI_DBG_BREAK").is_ok() {
                eprintln!(
                    "[BREAK] line={}/{} y={:.2} thr={:.2} after={:.2} relief={:.2} bottom={:.2} tol={:.2} -> {} lh={:.2} text={:?}",
                    line_idx,
                    lines.len(),
                    cursor.cursor_y,
                    break_threshold,
                    s1248_after,
                    s835_fn_relief,
                    effective_break_bottom,
                    s967_tol,
                    natural_needs_page_break,
                    line_height,
                    lines[line_idx]
                        .fragments
                        .iter()
                        .map(|f| f.text.as_str())
                        .collect::<String>()
                        .chars()
                        .take(18)
                        .collect::<String>()
                );
            }
            // S916 (2026-07-18, opt-out OXI_S916_DISABLE): force the split at
            // lines.len()-2 for a multi-line keepNext paragraph (the keepNext
            // lookahead requested s916_tail_split). This keeps n-2 head lines on
            // the current page and lets the existing mid-paragraph natural-break
            // handler move the 2-line tail (and, being the next block, the
            // follower) to the next page — matching Word, which SPLITS a
            // multi-line keepNext body paragraph 3+2 (legal pi=2128, its own
            // saved mid-paragraph LRPB) rather than whole-moving it. Re-gated on
            // the REAL lines.len() (>= 4 => len-2 keeps >= 2 head lines, so no
            // orphan). line_idx > 0 by construction (len-2 >= 2). Fires exactly
            // once (line_idx reaches len-2 on the current page — the para fits,
            // so there is no competing natural break before it).
            if s916_tail_split
                && lines.len() >= 4
                && line_idx == lines.len() - 2
                && std::env::var("OXI_S916_DISABLE").is_err()
            {
                natural_needs_page_break = true;
            }
            // S900 (2026-07-17, default ON, opt-out OXI_S900_DISABLE): Word
            // DEFERS a line's footnote NOTES (not the line) when they cannot
            // START in the page's remaining area. 81e80 p2: notes 2..15 fill
            // the area to the margin line (710.9 + 9.24 = 720.1), L9 (refs
            // 15..18) STAYS as the page's last body line with note 15 placed;
            // notes 16/17/18 render at the TOP of Word p3's area (measured).
            // Reconciles fnr_Z ("the anchor must co-locate, else it moves"):
            // there the SINGLE note could not start below the line (empty
            // area, sep+note past the margin) → the pair moves; deferral
            // applies only when at least one note places OR earlier notes
            // already fill the area (moving the line would free nothing).
            // v1 limitation: later lines of the same paragraph keep the full
            // committed map (conservative). Latin scope.
            // ★v3 default-ON (opt-out OXI_S900_DISABLE): cutoff closed by
            // (a) the S900b separator est fix (default-para-style resolution,
            // 12.649→13.8 for 81e80) + (b) the last-line AFTER-SPACE term +
            // (c) no epsilon. 81e80 Word arithmetic closes EXACT: fill
            // 570.25 + sep 13.8 + notes 2..14 → room 15.85 → note 15 places,
            // 16/17/18 roll = Word's p2/p3 split. The controlled probe
            // (_pb_fnarea) reproduced the roll (R07 stays, note 7 rolls at
            // room 3.7 < 9.2) and bracketed the keep-test sep to (13.5,15.5]
            // ∋ one Normal line = the same model.
            // S1244: rolling is a LEGACY-compat behaviour. The same document
            // rolls with compatibilityMode <= 14 (or none) and stops rolling
            // at 15: at 15 Word reserves the line's own notes in full and
            // moves the LINE instead, so a note never lands on a later page
            // than its reference. 81e80, which S900 was derived from, is
            // compatibilityMode 12; the docs where the roll produced a wrong
            // break are 15.
            let s900_legacy_compat = (self.compat_mode <= 14 || !self.compat_mode_explicit)
                || std::env::var("OXI_S1244_DISABLE").is_ok();
            if natural_needs_page_break
                && !self.doc_body_has_real_cjk
                && !para_fn_heights.is_empty()
                && s900_legacy_compat
                && std::env::var("OXI_S900_DISABLE").is_err()
            {
                let own = line_own_fn_ids.get(line_idx).cloned().unwrap_or_default();
                if !own.is_empty() {
                    let committed_prev = if line_idx > 0 {
                        committed_fn_delta_at_line
                            .get(line_idx - 1)
                            .copied()
                            .unwrap_or(0.0)
                    } else {
                        0.0
                    };
                    // Absolute margin bottom = effective bottom (reserve-shrunk)
                    // + earlier-para reserve + this para's full delta.
                    let absolute_bottom = effective_break_bottom - line_lenient_extra
                        + fn_reserve_above
                        + first_line_extra_content_h;
                    let prior_fill = fn_reserve_above + committed_prev;
                    let sep_needed = if prior_fill <= 0.0 {
                        (first_line_extra_content_h - para_fn_heights.values().sum::<f32>())
                            .max(0.0)
                    } else {
                        0.0
                    };
                    // v2 (probe _pb_fnarea + 81e80 closure): the placement cap
                    // boundary includes the LAST line's paragraph after-space
                    // (81e80 L9: 556.5 + after 13.75 + sep 13.8 + notes 2..14
                    // → room 15.85 → note 15 places, 16/17/18 roll = Word
                    // EXACT; without the after term 16/17 also placed = the
                    // measured {-1:3}). Approximation: autospacing 13.75 /
                    // explicit space_after; mid-para ref lines get 0.
                    let s900_after_term = if line_idx + 1 == lines.len() {
                        if para.style.after_autospacing
                            && std::env::var("OXI_S675_DISABLE").is_err()
                        {
                            if std::env::var("OXI_S907_DISABLE").is_err()
                                || (!self.doc_body_has_real_cjk
                                    && std::env::var("OXI_S901_DISABLE").is_err())
                            {
                                14.0
                            } else {
                                13.75
                            }
                        } else {
                            para.style.space_after.unwrap_or(0.0)
                        }
                    } else {
                        0.0
                    };
                    let line_bottom = cursor.cursor_y + break_threshold + s900_after_term;
                    let mut fill = line_bottom + prior_fill + sep_needed;
                    let mut placed = 0.0_f32;
                    let mut deferred: Vec<u32> = Vec::new();
                    for id in &own {
                        let h = para_fn_heights.get(id).copied().unwrap_or(0.0);
                        // No epsilon: the 81e80 cutoff (note 16 over by 0.35)
                        // sits inside a ±1 slack — Word's boundary is the
                        // margin line itself.
                        if deferred.is_empty() && fill + h <= absolute_bottom {
                            placed += h;
                            fill += h;
                        } else {
                            deferred.push(*id);
                        }
                    }
                    if !deferred.is_empty() && (placed > 0.0 || prior_fill > 0.0) {
                        let new_committed = committed_prev + sep_needed + placed;
                        let new_lenient = (first_line_extra_content_h - new_committed).max(0.0);
                        let new_bottom = effective_break_bottom - line_lenient_extra + new_lenient;
                        if cursor.cursor_y + break_threshold - s835_fn_relief <= new_bottom {
                            natural_needs_page_break = false;
                            if std::env::var("OXI_DBG900").is_ok() {
                                eprintln!("[S900] li={} defer={:?} placed={:.1} prior={:.1} abs_bot={:.1}",
                                    line_idx, deferred, placed, prior_fill, absolute_bottom);
                            }
                            s900_deferred_ids.extend(deferred);
                        }
                    }
                }
            }
            // S391 (2026-05-27): per-LINE LRPB respect. When THIS line is the
            // first to contain a run R that has has_last_rendered_page_break
            // (char_offset==0 for run R's fragment on this line), AND this is
            // not the paragraph's first line, force a mid-paragraph page break
            // before this line. Word honors the LRPB position even mid-paragraph
            // (b837 pi=71: run 1 has LRPB; in Word run 1 starts at top of
            // page 6; in Oxi run 1 starts at line 3 of page 5 because Oxi's
            // natural per-line break sees room remaining). More surgical than
            // the R7.45-rejected "force whole paragraph". Env-gated.
            // S395 SHIP (2026-05-27): per-LINE LRPB respect with doc-level
            // LRPB count threshold. DEFAULT ON.
            //
            // History: R7.45 (2026-05-13) ignored LRPB on non-first run citing
            // 34140 w_i=535 cascade concern when the WHOLE paragraph is moved
            // to next page. S391 (2026-05-27) implements per-LINE LRPB respect
            // (only the line containing the LRPB-bearing run's first char
            // moves to next page, not the whole para) — strictly more
            // surgical. S394 (2026-05-27) adds doc-level LRPB count threshold
            // to discriminate clean current LRPB hints (b837=6, d77a=11) from
            // stale-LRPB-saturated docs (3a4f=82 had 38 non-first-run LRPBs
            // that catastrophically cascaded under blanket-enable).
            //
            // Corpus impact (threshold=30):
            //   b837808d  0.7398 -> 0.9407  (+0.2009)  ← largest single-doc
            //   d77a58    0.8992 -> 0.9119  (+0.0127)         gain ever found
            //   ed025     0.9198 -> 0.9179  (-0.0019)         small
            //   3a4f      0.7916 -> 0.7919  (+0.0003)  ← threshold filtered
            //   corpus: 0.9603 -> 0.9641 (+0.0038), Phase 1 53/55 PRESERVED.
            //
            // S397 (2026-05-28) FALSIFIED: "skip per-line LRPB when LRPB-bearing
            // run text is short" hypothesis. Aimed at b837 page 7 +18.50pt
            // step (pi=89 LRPB on 1-char "の" particle, suspected stale/artifact
            // vs pi=71 LRPB on full sentence run, clean). At OXI_S397_LRPB_MIN_LEN=4:
            // b837 IoU 0.9535 -> 0.9776 (+0.0241, page 7 step fixed) BUT
            // b837 transitions Phase 1 PASS -> FAIL (53/55 -> 52/55 sentinel
            // regression). All L in {4,5,6,8} hit identical Phase 1 52/55.
            // No safe L. The pagination depends on per-line LRPB firing for
            // ALL b837 LRPBs (including short-run ones) — partial-firing
            // breaks Phase 1 alignment. Per CLAUDE.md no-EXCEPTION-stacking,
            // the spec needs re-derivation from richer input space (not a
            // per-run text-length filter).
            //
            // Opt-out:
            //   OXI_S391_PER_LINE_LRPB=0  -> disable per-line LRPB respect
            //   OXI_S394_LRPB_MAX=<N>     -> override threshold (default 30)
            let s391_on = std::env::var("OXI_S391_PER_LINE_LRPB")
                .map(|v| v != "0" && v != "false")
                .unwrap_or(true)
                // S811: distrusted saved LRPBs (metric-incompatible font
                // substitution) skip the per-line respect too.
                && !self.doc_lrpb_distrust
                && !(std::env::var("OXI_NATURAL_LRPB_DISABLE").is_err() && self.lrpb_count_distrust.get())
                // S897 (2026-07-17, default ON, opt-out OXI_S897_DISABLE;
                // ships WITH S895+S898 — separately each is a PASS-doc
                // trade): LATIN docs
                // retire the per-line LRPB respect too — the S836
                // block-level drop completed. The whole frozen EN 6 is PASS
                // 1.0000 with per-line OFF (measured), i.e. the Latin natural
                // flow is fully LRPB-independent after S827-S835/S884-S894;
                // what remains of per-line respect only re-plays STALE saved
                // breaks (legal__00081e80: the saved mark sits one line EARLY
                // vs fresh Word — Word keeps «plaintiffs, in the issues…» on
                // p1 at box 696.5 ≤ cbot ~696.8, the mark forced it to p2 =
                // the doc-wide +13.3 chain; with S895's correct autospacing
                // the doc reads 0.9746 vs 0.83 with the mark). JP keeps the
                // full LRPB model (b837/d77a/3a4f load-bearing).
                && (self.doc_body_has_real_cjk
                    || std::env::var("OXI_S897_DISABLE").is_ok())
                // S1491: the per-line respect retires for CJK bodies as well
                // (see the block-level site); opt-in OXI_S1491_LEGACY_LRPB.
                && std::env::var_os("OXI_S1491_LEGACY_LRPB").is_some();
            let s391_lrpb_break = if line_idx > 0 && !in_textbox && s391_on {
                let lrpb_frag = |f: &LineFragment| {
                    f.char_offset == 0
                        && para
                            .runs
                            .get(f.run_index)
                            .map(|r| r.has_last_rendered_page_break)
                            .unwrap_or(false)
                };
                // S1151 (2026-08-16, default ON, opt-out OXI_S1151_DISABLE):
                // the mark records where Word STARTED a page, so Word's break
                // is a LINE BOUNDARY. When our wrap puts the mark mid-line the
                // two boundaries around it are both candidates, and breaking
                // BEFORE the line unconditionally (the pre-S1151 rule) throws
                // the whole head of that line onto the next page. Snap to the
                // NEARER boundary by advance width instead.
                //   kojin  pi=20: 44 chars before the mark, 1 after -> break
                //     AFTER. Word (its own PDF) keeps both lines of pi=20 on
                //     p1; the old rule broke before, and every page from 2 on
                //     ran 13.89pt low (the p3 table inherits the offset -- its
                //     own row pitch is already exact at 38.40).
                //   b837  pi=89: 34 chars before, 3 after -> break AFTER, and
                //     Word likewise keeps BOTH of those lines on p6 (the old
                //     rule kept only one).
                //   3a4f  pi=42/113/190/223/785/904/933: 1-3 chars before,
                //     23-40 after -> break BEFORE, unchanged.
                // Marks that open their line (the common case: 6/6 b837, 7/7
                // d77a, 13/20 3a4f, 4/5 kojin) are unaffected either way.
                // Gate: Phase 1 95/96 with ZERO per-doc change; SSIM sentinel
                // 1 of 238 changed bytes (b837 +0.0701 over 7 pages, nothing
                // regressed); kojin (no cached reference, scored against its
                // own Word PDF) 0.8431 -> 0.8581, p2 +0.1464 / p3 +0.1843.
                // Requiring the mark to OPEN the line instead was tried first
                // and is WRONG: it drops b837 pi=89 entirely (Phase 1
                // PASS -> FAIL), because there the nearer boundary is the
                // following one, not "no boundary at all".
                let mark_in_head = |ln: &Line| -> Option<bool> {
                    let i = ln.fragments.iter().position(lrpb_frag)?;
                    let before: f32 = ln.fragments[..i].iter().map(|f| f.width).sum();
                    let after: f32 = ln.fragments[i..].iter().map(|f| f.width).sum();
                    Some(before <= after)
                };
                let s1151_on = std::env::var("OXI_S1151_DISABLE").is_err();
                let has_lrpb_here = if s1151_on {
                    mark_in_head(line) == Some(true)
                        || (line_idx >= 1 && mark_in_head(&lines[line_idx - 1]) == Some(false))
                } else {
                    line.fragments.iter().any(lrpb_frag)
                };
                if std::env::var("OXI_DUMP_LRPB").is_ok() {
                    if let Some(i) = line.fragments.iter().position(lrpb_frag) {
                        let before: usize =
                            line.fragments[..i].iter().map(|f| f.text.chars().count()).sum();
                        let after: usize =
                            line.fragments[i..].iter().map(|f| f.text.chars().count()).sum();
                        let head: String = line.fragments[i].text.chars().take(12).collect();
                        eprintln!(
                            "[LRPB] pi={} line={}/{} cursor_y={:.2} chars_before={} chars_after={} head={} mark_text={:?}",
                            body_para_index.map(|v| v.to_string()).unwrap_or_else(|| "?".into()),
                            line_idx, lines.len(), cursor.cursor_y, before, after,
                            mark_in_head(line) == Some(true), head
                        );
                    }
                }
                let s394_max = std::env::var("OXI_S394_LRPB_MAX")
                    .ok()
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(30);
                // S563 SHIP (2026-06-14, default ON, opt-out OXI_S563_DISABLE): only
                // respect a lastRenderedPageBreak when the current page is substantially
                // full (cursor past content_height/2). A LRPB that fires near the page
                // TOP is a STALE hint (Word re-rendered and the break moved) — respecting
                // it forces a premature mid-paragraph break leaving the page nearly empty.
                // ikujikaigo: 1 LRPB in pi=60 fires at cursor ~66 (p4 ~8% full) →
                // premature → 108 paras pushed +1 (0.3455 → 0.9758 with this gate). b837's
                // LRPBs fire near the page BOTTOM (real breaks) → still respected. GATE:
                // full corpus 58/62 (ikujikaigo 0.3455→0.9758, 0 baseline PASS→FAIL;
                // b837/d77a/3a4f/ed025 all PASS). total_lrpb_count≤30 (S394) AND
                // page-substantially-full (S563) together discriminate stale LRPBs.
                let s563_full = if std::env::var("OXI_S563_DISABLE").is_ok() {
                    true
                } else {
                    cursor.cursor_y > page_top + content_height * 0.5
                };
                // S577 margin-discriminator FALSIFIED (had the sign BACKWARDS):
                // it respected LARGE-margin LRPBs and ignored small-margin ones.
                // S581 (2026-06-15) inverts it correctly: a STALE LRPB fires FAR
                // from the page bottom (the line plus the NEXT line both fit, i.e.
                // > 1 line of room below); a REAL page-bottom LRPB fires when the
                // line is the LAST that fits (the next line would overflow). The
                // physical test: respect only when `over > -effective_lh`. PDF
                // render-truth: ikujidetail pi=24 (over=-22.95, p1) is a stale LRPB
                // Word IGNORES (continues para 24 two more lines) → a 2-line page-1
                // shift cascading via para-spills to the +1×2 (wi=355/440); pi=149/263
                // (over=-1.75, line at the bottom, next line overflows) are REAL.
                // b837 pi=89 (over=-22.35) is also stale (b837 PASSES without S391).
                // The reals measured: d77a -0.50, 3a4f -0.85, ikujidetail -1.75 — all
                // within 1 line of the bottom.
                // ★REVERTED to DEFAULT-OFF (2026-06-23, git-bisect f5cd4b13): S581 is a
                // pagination NO-OP now (full gate CHANGED 0 with it off — S595 made
                // ikujidetail's LRPBs redundant) but it WRONGLY classified b837 pi=89's
                // REAL page-bottom LRPB as "stale" and ignored it → the p5/p6 line-split
                // diverged from Word → SSIM A/B b837 net **-0.2547** (the ONLY doc affected
                // among 238 word_png). The `over` discriminator cannot separate b837 pi=89
                // (Word HONORS, render-real) from ikujidetail pi=24 (Word ignores) — both
                // over≈-22 — so the gate mis-fires on b837. S581 shipped 2026-06-15 with a
                // FALSE "+0.0000" from the broken ssim_ab tool (fixed 2026-06-18) — see
                // [[ssim_ab_tool_was_broken]]. Default OFF; opt-in OXI_S581=1.
                let s581_stale = if std::env::var("OXI_S581").ok().as_deref() == Some("1") {
                    let over = cursor.cursor_y + break_threshold - effective_break_bottom;
                    over < -effective_lh
                } else {
                    false
                };
                has_lrpb_here && page.total_lrpb_count <= s394_max && s563_full && !s581_stale
            } else {
                false
            };
            // OXI_DBG_S391: one line per per-LINE LRPB break, with the geometry a
            // discriminator would use. S822 failed by picking a threshold before
            // dumping both populations; do not repeat that -- collect first.
            if s391_lrpb_break && std::env::var("OXI_DBG_S391").is_ok() {
                let preview: String = line
                    .fragments
                    .iter()
                    .flat_map(|f| f.text.chars())
                    .take(16)
                    .collect();
                eprintln!(
                    "[S391] pg={} line_idx={} y={:.2} bottom={:.2} slack={:.2} natural={} text={:?}",
                    pages.len() + 1,
                    line_idx,
                    cursor.cursor_y,
                    effective_break_bottom,
                    effective_break_bottom - cursor.cursor_y,
                    natural_needs_page_break,
                    preview
                );
            }
            let needs_page_break = natural_needs_page_break || s391_lrpb_break;
            // OXI_DUMP_BREAK_Y lowers the ALL-lines cutoff: a page whose bottom
            // is eaten by a footnote area breaks well above 700.
            let dump_break_y: f32 = std::env::var("OXI_DUMP_BREAK_Y")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(700.0);
            if std::env::var("OXI_DUMP_BREAK").is_ok()
                && (line_idx == 0
                    || (std::env::var("OXI_DUMP_BREAK_ALL").is_ok()
                        && cursor.cursor_y > dump_break_y))
            {
                let pi_str = body_para_index
                    .map(|v| v.to_string())
                    .unwrap_or_else(|| "?".into());
                let txt: String = para
                    .runs
                    .iter()
                    .flat_map(|r| r.text.chars())
                    .take(15)
                    .collect();
                eprintln!(
                    "[BR_DUMP] pi={} line0 cursor_y={:.3} eff_lh={:.3} nat_lh={:.3} ink_lh={:.3} brk_thr={:.3} eff_bot={:.3} over={:.3} brk={} s779={} s779h={:.3} text={:?}",
                    pi_str, cursor.cursor_y, effective_lh, natural_lh, ink_lh, break_threshold,
                    effective_break_bottom, cursor.cursor_y + break_threshold - effective_break_bottom,
                    needs_page_break, s779_latin,
                    s779_win_heights.get(line_idx).copied().unwrap_or(-1.0), txt
                );
                eprintln!("[BR_DUMP2] branch={} s693_nonlast={}", brk_branch, s693_nonlast);
                // The reliefs that move the bottom. `eff_bot` already carries
                // lenient; printing the parts says WHICH one bought the line.
                eprintln!(
                    "[BR_DUMP3] li={} lenient={:.3} fn_relief={:.3} s967_tol={:.3} fn_above={:.3} first_extra={:.3} content_bot={:.3}",
                    line_idx, line_lenient_extra, s835_fn_relief, s967_tol, fn_reserve_above,
                    first_line_extra_content_h, page_top + content_height
                );
            }

            // Widow/orphan: if this is line 0 (orphan) and there are 2+ lines,
            // check if only 1 line would fit on this page — if so, push the
            // entire paragraph to the next page.
            // S282 (2026-05-25): experimental env-gate OXI_FORCE_WIDOW=1 to
            // apply widow protection regardless of para.style.widow_control.
            // b837 has <w:widowControl w:val="0"/> in Normal style but Word's
            // actual rendering applies widow protection anyway — S281 found
            // Oxi is consistently 1 page ahead of Word starting at pi=20,
            // which is exactly a 1-line orphan that Word pushes to next page.
            //
            // S283 (2026-05-25): refined to only apply force-widow on paragraphs
            // with ≥5 lines. Falsified hypothesis: "force widow on any 2+ line
            // paragraph". d77a sample test showed 4-line paragraph pi=46
            // (text "イは、編集・加工等の二次利用を行った") regressed: Word
            // does NOT widow-protect it, but force_widow=1 pushed it forward,
            // cascading 2 trailing empty paragraphs (pi=47, pi=48) to wrong
            // page. The b837 win came from pi=20 (7 lines); threshold ≥5 keeps
            // that win while leaving shorter paragraphs alone.
            //
            // S284 attempt (2026-05-25): flipped to DEFAULT ON based on
            // page-match improvement. REVERTED in S285: the page-match
            // metric was misleading because Word's `.Range.Information`
            // idx field is NOT document XML order — some paragraphs render
            // OUT of idx order on the same page (e.g. b837 p2 has idx=22
            // at y=160.5 and idx=21 at y=646.5 on the same page). The
            // `idx = pi + 1` mapping inflated page-match counts. Actual
            // Phase 2 IoU gate result with the fix ON:
            //   b837   0.7466 → 0.4790  (-0.268, REGRESSION)
            //   d77a   0.7719 → 0.9123  (+0.140, improvement)
            //   db9ca1 0.9829 → 0.7945  (-0.188, REGRESSION)
            // Net Phase 2 gate fails. Reverted to env-gated OPT-IN with
            // OXI_FORCE_WIDOW=1; keep the ≥5-lines threshold from S283.
            let force_widow = std::env::var("OXI_FORCE_WIDOW").is_ok();
            let widow_effective = para.style.widow_control || (force_widow && lines.len() >= 5);
            // S608 (2026-06-18, default ON, opt-out OXI_S608_DISABLE): the
            // widow/orphan page-fit LOOK-AHEAD measures the paragraph's LAST line by
            // its NATURAL height (ascent+descent), NOT the full multiplied line box.
            // Word lets the last line's line-spacing LEADING hang into the bottom
            // margin and keeps a 2-line para on the page when its last line's
            // natural box fits — test_line_heights Calibri 14pt x2.0 (natural 17.25,
            // full box 34.5) is KEPT; Oxi's full-box look-ahead widow-pushed it → a
            // per-page accumulating offset down the document. SCOPE: only fires when
            // widow_control is ON (Word default); the CJK Phase-1 corpus has
            // widowControl=0 → byte-identical, gate 76/84 unchanged. RESULT (correct
            // same-binary A/B, OXI_S608 OFF vs ON over 173 widowControl=ON cohort
            // pages): net +0.0006; the gen2_* cohort is BYTE-IDENTICAL (Δ=0.0000,
            // never fires); the whole change is test_line_heights — 6 pages = Word
            // (was 6) with page-mismatches 6→0 (pages 3/4/5 +0.018/+0.021/+0.060,
            // p6 +0.003, only p2 −0.003). ★An earlier ink-based variant over-compacted
            // the doc end (MS Gothic 14 x2.0 fit on p5, losing Word's p6, −0.985);
            // and an earlier "net −0.0144 / gen2 regressions" reading was a
            // MEASUREMENT ARTIFACT (compared against a STALE ssim_baseline.json from
            // an older binary). natural (not ink) is the structural keep/push measure.
            let s608 = std::env::var("OXI_S608_DISABLE").is_err();
            // S1096 (2026-08-08, default ON, opt-out OXI_S1096_DISABLE): a
            // 2-line paragraph's ORPHAN look-ahead measures its last line with
            // the FULL box, not S608's natural height, when the line is
            // single/atLeast spaced.  A 2-line paragraph cannot be split under
            // widowControl, so "does the last line fit" is exactly the per-line
            // break question and must use the same height.
            // S608 was derived on ×2.0 paragraphs (Calibri 14 natural 17.25 vs
            // full box 34.5; MS Gothic 14 natural 18.125) — the leading Word
            // lets hang into the margin IS the multiplier's extra leading.  On a
            // single/atLeast line there is no such leading: legal__0019967c's
            // Indenta style is `line=260 atLeast` where full = effective_lh
            // 13.799 (the hhea line the per-line threshold uses) but
            // natural_line_height_for_line returns 13.500, so the orphan
            // look-ahead measured the last line 0.30pt SHORTER than the
            // per-line break test would — an internal inconsistency, not a
            // Word rule.  wp167 «(b) an account of the insurer» (2 lines,
            // cursor 605.166, cbot 632.5): 605.166+13.799+13.500 = 632.465
            // "fits" → Oxi split it 1+1, while the per-line test rejects line 1
            // (632.764 > 632.5) so only line 0 stayed = an orphan Word never
            // leaves.  Word's own PDF puts BOTH lines on p167, and its content
            // bottom is bracketed to [631.66, 632.90) by the deepest body
            // baseline in the document (629.02 + descent) and by this very
            // push — i.e. Oxi's 632.5 is right and only the look-ahead height
            // was wrong.  Latin scope (the CJK corpus has widowControl=0, so
            // this arm never fires there anyway).
            // SCOPE = the ORPHAN arm only.  Applying it to the WIDOW arm
            // (line_idx == len-2 of a 3+-line paragraph) regressed
            // legal__0014c86f: its 5-line «(1) At any time» ends at box bottom
            // 639.37 vs cbot 639.30, and Word keeps all five (its own last
            // baseline 636.58 → bottom 639.22).  That 0.15pt is this document's
            // recorded TOP-MARGIN 10tw rounding (Oxi 119.00 vs Word 118.80),
            // not a look-ahead error — S608's natural height was absorbing it.
            // A splittable paragraph also has a real choice about where to
            // break, which the 2-line case does not.
            let s1096_full_box = !is_multiple_spacing
                && !self.doc_body_has_real_cjk
                && std::env::var("OXI_S1096_DISABLE").is_err();
            // The widow/orphan STRUCTURAL look-ahead (keep a 2-line para together
            // or push it whole) measures the last line by its NATURAL height
            // (ascent+descent) — NOT the full multiplied box, and NOT the glyph ink.
            // DERIVED from test_line_heights' p5/p6 boundary (last line at cursor
            // 702.5, content bottom 720):
            //   MS Gothic 14 x2.0: natural 18.125 → 720.6 > 720 → PUSH (= Word p6)
            //   Calibri  14 x2.0: natural 17.25  → 713.6 < 720 → KEEP (= Word p2)
            // The per-line RENDER break uses ink (S576, leading hangs into the
            // margin); but the structural keep/push uses the natural line box, so
            // a CJK 83/64 last line (natural ≈ 1.297·em) is pushed where ink (= em)
            // would wrongly fit it. (Using ink here over-compacted test_line_heights
            // to 5 pages, losing Word's page 6 — the −0.985 in the first A/B.)
            // S1332 (2026-09-05, default ON, opt-out OXI_S1332_DISABLE): under an
            // EXACT line rule the look-ahead height IS the exact box. DERIVED
            // (_pb_exactorphan_gen.py, COM Information(3/6), 2-line paragraph,
            // exact 28.8, widowControl, band bottom 756.9): start 698.25 keeps
            // both lines (698.25 + 57.6 = 755.85), start 700.25 / 702.25 / 704.25
            // / 706.25 push the whole paragraph although start + 28.8 + the
            // natural 23.4 (BIZ UDPGothic 18) or 13.6 (MS Mincho 10.5 in the same
            // box) would fit -- so the S608 natural rule (derived on x2.0
            // multiples, where the leading is the multiplier's) does not reach
            // exact lines. technical__898a80 p11: 「②画面上では…」 at 704.3 is
            // Word's p12 (today's truth), Oxi kept line 0 on p11.
            let s1332_exact = para.style.line_spacing_rule.as_deref() == Some("exact")
                && std::env::var("OXI_S1332_DISABLE").is_err();
            // S1558 (2026-09-26, default ON, opt-out OXI_S1558_DISABLE): on a typed
            // grid the per-line page-bottom test is S739's centred box,
            // (box + natural) / 2, for every line (S1155). The widow / orphan
            // look-ahead measured the same last line by S608's natural height,
            // 1.25pt shorter on an 18pt pitch, so a 2-line widowControl paragraph
            // "fit" in the look-ahead and was then split 1+1 by the per-line
            // test -- an orphan Word never leaves (the S1096 shape, on a grid).
            // policies__1db396de p43: 「関係機関の窓口へのリーフレット…」 (2 lines,
            // HG丸ｺﾞｼｯｸM-PRO 12, lines 360, widowControl) at 736.9 on a 771.0
            // bottom: 736.9 + 18 + 15.5 = 770.4 fit the look-ahead, 754.9 + 16.75
            // = 771.65 broke the line; the Word PDF puts both lines at the top
            // of p44. Same predicate as s739_centered (S1153 on-slot gate off).
            let s1558_centered = std::env::var_os("OXI_S1558_DISABLE").is_none()
                && std::env::var("OXI_S739_DISABLE").is_err()
                && !s1375_before_section_end
                && !page.doc_grid_no_type
                && para.style.snap_to_grid
                && grid_pitch.map_or(false, |p| p > 0.0);
            let last_line_fit_h = |idx: usize| -> f32 {
                let full = line_heights.get(idx).copied().unwrap_or(0.0);
                // With precise Latin margins, widow and orphan look-ahead
                // can use the same single/atLeast height as the line fit test.
                // The old shorter widow estimate compensated rounded origins.
                if s1332_exact || (s1096_full_box
                    && std::env::var("OXI_EXACT_LATIN_MARGIN_DISABLE").is_err()
                    && std::env::var("OXI_WIDOW_HEIGHT_DISABLE").is_err()) {
                    return full;
                }
                // A grid line spanning multiple slots still centers its natural
                // box. The orphan look-ahead must use the same bottom as painting.
                if centered_multicell_grid {
                    let natural = natural_line_heights.get(idx).copied().unwrap_or(full).min(full);
                    return (full + natural) / 2.0;
                }
                if s1558_centered {
                    // S1558: the same centred box the per-line test will apply.
                    let natural = natural_line_heights.get(idx).copied().unwrap_or(full).min(full);
                    return ((full + natural) / 2.0).max(natural).min(full);
                }
                // Latin automatic multiples use the same unrounded natural
                // capacity as the per-line fit test. The quarter-point spacing
                // height can otherwise reject a last line that fits when drawn.
                if is_multiple_spacing && s779_latin && !no_type_multiple_ink
                    && ruby_para_expansion_pt == 0.0
                    && std::env::var("OXI_S1194_DISABLE").is_err()
                {
                    let capacity = s779_win_heights.get(idx).copied().unwrap_or(0.0);
                    if capacity > 0.0 { return capacity.min(full); }
                }
                if no_type_multiple_ink {
                    // S1079: natural (unmultiplied) line, not the ink box.
                    let src = if s1079_natural {
                        &natural_line_heights
                    } else {
                        &ink_line_heights
                    };
                    return src.get(idx).copied().unwrap_or(full).min(full);
                }
                if !s608 {
                    return full;
                }
                natural_line_heights
                    .get(idx)
                    .copied()
                    .unwrap_or(full)
                    .min(full)
            };
            let mut widow_fit_offset = 0.0;
            let mut widow_fit_limit = page_top + content_height;
            // The current row's explicit boundary already chooses where its
            // continuation goes. Widow/orphan look-ahead must not pair that
            // row with text in the next explicit flow region. An empty boundary
            // itself cannot become a printed orphan.
            let widow_orphan_break = if !in_textbox && widow_effective && lines.len() >= 2
                && !matches!(line.break_type,LineBreakType::PageBreak|LineBreakType::ColumnBreak) {
                if line_idx == 0 && !needs_page_break {
                    // Orphan: check if the next line would overflow — that would leave
                    // only 1 line on this page. Push entire paragraph to next page.
                    // The next line is the LAST line only for a 2-line para; use the
                    // page-bottom fit height for it (S608), full box otherwise.
                    let next_h = if lines.len() == 2 {
                        if s1096_full_box {
                            line_heights.get(1).copied().unwrap_or(0.0)
                        } else {
                            last_line_fit_h(1)
                        }
                    } else if std::env::var("OXI_S1074_DISABLE").is_err()
                        && is_multiple_spacing
                        && !self.doc_body_has_real_cjk
                    {
                        // S1074: an INTERIOR next line is measured with the same
                        // page-bottom threshold the per-line break test uses
                        // (max(ink, the S779/S827 hhea floor), capped at the box)
                        // rather than the full multiplied box.
                        // ★SHIPPED default-ON 2026-08-13 (opt-out
                        // OXI_S1074_DISABLE) with the {S1112, S1091, S1074,
                        // S1113, S1114} bundle — see the gate summary at the
                        // S1091 site.  Word truth
                        // for the two documents that disagreed: reports__00377a16
                        // keeps 2 lines whose BOX bottom is 3.5pt past the content
                        // bottom (COM: 4-line para split 2+2 at y732.00/752.70,
                        // content bottom 769.90) — only the natural-height test
                        // fits it; policies__000f7115's push is then correct too
                        // once S1112/S1091 put its cursor at Word's 741.00
                        // (it was 738.92, and the lenient test wrongly kept).
                        let full = line_heights.get(1).copied().unwrap_or(0.0);
                        let floor = s779_win_heights.get(1).copied().unwrap_or(0.0);
                        ink_line_heights
                            .get(1)
                            .copied()
                            .unwrap_or(full)
                            .max(floor)
                            .min(full)
                    } else {
                        line_heights.get(1).copied().unwrap_or(0.0)
                    };
                    let next_h = footnote_fit_height(1, next_h);
                    // S835: the fn-area boundary softness (fs/16) applies to the
                    // widow/orphan fit test too — framework wp26 «shall obtain»
                    // (2-line, ref 14) was whole-moved HERE at a 0.17pt overflow
                    // that Word's soft boundary absorbs.
                    // S926: compare the orphan look-ahead at OOXML's twip
                    // precision. Raw font-metric floats can exceed an exact
                    // page-bottom tie by a few thousandths of a point and
                    // whole-move an otherwise splittable paragraph (legal
                    // wp102: +0.012pt). Half a twip is the round-to-nearest
                    // tolerance; larger physical overflows still push.
                    let orphan_rounding_tolerance = if typed_grid_fn_boundary {
                        fn_coordinate_roundoff
                    } else if std::env::var("OXI_S926_DISABLE").is_err() && !automatic_exact_fit {
                        0.025
                    } else {
                        0.0
                    };
                    widow_fit_offset = line_height + next_h - s835_fn_relief;
                    widow_fit_limit = footnote_fit_bottom(1) + orphan_rounding_tolerance;
                    cursor.cursor_y + widow_fit_offset > widow_fit_limit
                        && !current_elements.is_empty()
                } else if (line_idx == lines.len() - 2
                    // A column control closes this paragraph fragment. Its
                    // last printed row is protected from becoming a widow
                    // even when the paragraph continues after the control.
                    || lines.get(line_idx+1).is_some_and(|next|
                        matches!(next.break_type,LineBreakType::PageBreak|LineBreakType::ColumnBreak)
                        && next.fragments.iter().any(|fragment|!fragment.text.trim().is_empty())))
                    && !needs_page_break {
                    // Widow: if the last line would overflow to the next page alone,
                    // break BEFORE this line so at least 2 lines go to the next page.
                    // next line (line_idx+1) is the paragraph's LAST line → S608.
                    // S891 (2026-07-17, default ON, opt-out OXI_S891_DISABLE): a
                    // trailing-BR EMPTY last line (the S832 class) demands NO
                    // page room — it neither becomes a widow nor pushes the
                    // paragraph. usnyserda p54 «NYSERDA will format…¶<br>»:
                    // 2 text lines fit (bottom 694.5 ≤ cbot 706.5) but the
                    // phantom empty line's box crossed by 0.6pt → the widow
                    // arm whole-moved the para where Word keeps it (rt.pdf:
                    // both text lines at p54 667.3/680.6, Option-2 at p55
                    // top). Latin scope via the S832 shape.
                    let s891_next_is_trailing_br_empty =
                        lines.get(line_idx + 1).map_or(false, |l| {
                            l.fragments.iter().all(|f| f.text.trim().is_empty())
                        }) && !self.doc_body_has_real_cjk
                            && std::env::var("OXI_S891_DISABLE").is_err();
                    // When a three-row paragraph cannot fit even an empty
                    // page, both endpoint constraints cannot be satisfied. The
                    // first two printed rows stay together; moving the second
                    // row would create an orphan or repeatedly move the whole
                    // oversized paragraph. Word's exact-height controls retain
                    // the first pair and let the final row continue separately.
                    let first_pair_has_priority=lines.len()==3 && line_idx==1
                        && !elements.is_empty()
                        && lines.iter().all(|row|!matches!(row.break_type,
                            LineBreakType::PageBreak|LineBreakType::ColumnBreak))
                        && line_heights.iter().take(2).sum::<f32>()+last_line_fit_h(2)
                            > footnote_fit_bottom(2)-page_top;
                    if s891_next_is_trailing_br_empty || first_pair_has_priority {
                        false
                    } else {
                        let next_h = footnote_fit_height(
                            line_idx + 1, last_line_fit_h(line_idx + 1));
                        // S1248: the look-ahead asks the same question the
                        // per-line test does, so it has to count the trailing
                        // space too — otherwise the last line is rejected there
                        // and kept here, and the paragraph splits 2+1 where
                        // widowControl forbids any split at all.
                        widow_fit_offset = line_height + next_h + s1248_next_after - s835_fn_relief;
                        widow_fit_limit = footnote_fit_bottom(line_idx + 1)
                            + if typed_grid_fn_boundary { fn_coordinate_roundoff } else { 0.0 };
                        cursor.cursor_y + widow_fit_offset > widow_fit_limit
                    }
                } else {
                    false
                }
            } else {
                false
            };
            // Dump line 0 always, plus EVERY actual break — the widow arm fires
            // at line_idx == lines.len()-2, which a line-0-only dump hides on a
            // paragraph of 3+ lines.
            if std::env::var("OXI_DUMP_WIDOW").is_ok() && (line_idx == 0 || widow_orphan_break) {
                let txt: String = para
                    .runs
                    .iter()
                    .flat_map(|r| r.text.chars())
                    .take(15)
                    .collect();
                eprintln!("[WIDOW] line{} lines={} wc={} cursor_y={:.2} lh={:.2} next_h={:.2} limit={:.2} curr_empty={} break={} text={:?}",
                    line_idx,
                    lines.len(), para.style.widow_control, cursor.cursor_y, line_height,
                    line_heights.get(1).copied().unwrap_or(0.0),
                    page_top + content_height, current_elements.is_empty(),
                    widow_orphan_break, txt);
            }

            // S790 (2026-07-11, default ON, opt-out OXI_S790_DISABLE, Latin
            // scope): the WIDOW arm (line_idx == len−2) must SPLIT the
            // paragraph like Word — lines 0..len−3 STAY on the current page
            // and only the last TWO lines move. The shared whole-move handler
            // (correct for the ORPHAN arm) sent the ENTIRE paragraph to the
            // next page (nyserda kick-off para: Word splits 6/2, Oxi moved
            // all 8 → ~90pt page under-fill = LRPB-off catalog #3; the S770
            // note documents the same whole-move on '2. WAGE'). JP keeps the
            // legacy whole-move (widowControl is mostly 0 there; calibrated).
            // line_idx > 1: the split must KEEP ≥2 lines — at line_idx == 1 it
            // would strand line 0 alone = an ORPHAN, and Word whole-pushes
            // instead (nyserda p2 bottom ⎯-item: 3 lines, 2 fit → Word moves
            // all 3; the first split cut 1/2 = catalog #4).
            // S1571 (2026-09-26, default ON, opt-out OXI_S1571_DISABLE): the JP
            // side splits too. reports__393aa91f p4/5 (linesAndChars 400, MS
            // P明朝 10pt, widowControl from Normal): a 7-line paragraph with 6
            // lines of room -- Word keeps 5 on p4 and moves 2 (PDF baselines
            // 654.7..734.6 / 84.6, 104.4); Oxi's legacy whole-move sent all 7
            // to p5 and the doc ran one paragraph late for five pages.
            let s790_widow_split = line_idx > 1
                && (!self.doc_body_has_real_cjk
                    || std::env::var_os("OXI_S1571_DISABLE").is_none())
                && std::env::var("OXI_S790_DISABLE").is_err();
            // S1323 (2026-09-05, default ON, opt-out OXI_S1323_DISABLE): a
            // widow/orphan break inside a multi-column section goes to the
            // NEXT COLUMN, the way the natural overflow (S637) and an explicit
            // column break (S733) already do -- Word treats the column as the
            // page for widow/orphan control. Both arms pushed a whole PAGE:
            // reports__167853 p28 (2-col continuous section, 42 lines, room
            // for 45): the 4-line `widowControl` paragraph after ●自立支援事業
            // hit the widow arm at line 2 with column 1 still empty, the
            // whole-move opened page 29, its two laid lines kept their
            // page-28 y and the S750 balance then sorted them to the END of
            // column 1 (the p29 picture). Word keeps the section on p28
            // (W29 / O30 -> O29).
            // A widow/orphan move must fit its kept lines in the next column.
            // A short continuous-section band may require a fresh page even
            // when another column exists on this page.
            let widow_next_col_top = if pages.len() > s749_pages_at_entry {
                page_top
            } else { col_band_top };
            let widow_carried_height = if s790_widow_split { 0.0 } else {
                line_heights.iter().take(line_idx).sum::<f32>()
            };
            let s1323_col_advance = widow_orphan_break
                && num_columns > 1
                && cur_col + 1 < num_columns
                && std::env::var("OXI_S637_DISABLE").is_err()
                && std::env::var("OXI_S1323_DISABLE").is_err()
                && (widow_next_col_top <= page_top + 0.025
                    || widow_next_col_top + widow_carried_height + widow_fit_offset <= widow_fit_limit);
            if s1323_col_advance {
                let old_x = start_x;
                cur_col += 1;
                start_x = col_x_positions[cur_col];
                let col_top = column_flow_top(if pages.len() > s749_pages_at_entry {
                    page_top
                } else {
                    col_band_top
                }, start_x, pages.len());
                if std::env::var("OXI_DBG_COL").is_ok() {
                    eprintln!(
                        "[COL] S1323 widow/orphan col {}->{} line_idx={} split={} cursor_y={:.1} col_top={:.1}",
                        cur_col - 1, cur_col, line_idx, s790_widow_split, cursor.cursor_y, col_top
                    );
                }
                if s790_widow_split || elements.is_empty() {
                    // Split (or nothing laid yet): the laid lines stay in the
                    // old column; this line starts the new one.
                    cursor.set(col_top);
                } else {
                    // Whole-move: carry the laid lines to the new column's top,
                    // keeping their stack (S770's shape, per column).
                    let laid_h: f32 = line_heights.iter().take(line_idx).sum();
                    let shift = col_top - (cursor.cursor_y - laid_h);
                    let dx = start_x - old_x;
                    for e in elements.iter_mut() {
                        e.x += dx;
                        e.y += shift;
                        if let LayoutContent::TableBorder { x1, x2, y1, y2, .. } = &mut e.content {
                            *x1 += dx;
                            *x2 += dx;
                            *y1 += shift;
                            *y2 += shift;
                        }
                    }
                    cursor.set(col_top + laid_h);
                }
                s842_apply(cursor);
            } else if widow_orphan_break && s790_widow_split {
                // Same shape as the natural mid-paragraph break: keep the laid
                // lines on this page, continue on the fresh one.
                current_elements.extend(std::mem::take(&mut elements));
                dbg_page_push(pages.len(), 0);
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                if let Some(g) = s755_geom {
                    page_top = g.top(pages.len() + 1);
                    content_height = g.ch(pages.len() + 1);
                }
                cursor.set(page_top);
                s842_apply(cursor);
                cur_col = 0;
                if num_columns > 1 && std::env::var("OXI_WIDOW_COLUMN_RESET_DISABLE").is_err() {
                    start_x = col_x_positions[0];
                }
                if let Some(v) = line_fn_refs_out.as_deref_mut() {
                    // earlier lines' refs stay on the old page; open the new bucket
                    v.push(Vec::new());
                }
            } else if widow_orphan_break {
                // Push current page and move entire paragraph so far to next page
                dbg_page_push(pages.len(), 0);
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                if num_columns > 1 && std::env::var("OXI_WIDOW_COLUMN_RESET_DISABLE").is_err() {
                    let dx = col_x_positions[0] - start_x;
                    cur_col = 0;
                    start_x = col_x_positions[0];
                    for e in elements.iter_mut() {
                        e.x += dx;
                        if let LayoutContent::TableBorder { x1, x2, .. } = &mut e.content {
                            *x1 += dx;
                            *x2 += dx;
                        }
                    }
                }
                if let Some(g) = s755_geom {
                    page_top = g.top(pages.len() + 1);
                    content_height = g.ch(pages.len() + 1);
                }
                cursor.set(page_top);
                s842_apply(cursor);
                // S770 (2026-07-09): the widow/orphan break MOVES this paragraph's
                // already-laid lines to the next page but they KEPT their OLD y
                // (bottom of the previous page) while the cursor reset to page_top.
                // For a paragraph that starts near the page bottom and fills to it
                // before the last line triggers the widow (nyserda Exhibit C "2. WAGE"
                // = 13 lines from y518: 11 laid at y518-683, widow at line 11), the 11
                // moved lines stay at y518-683 on the NEW page while the last 2 lay at
                // y72-99 and the cursor ends at y99 — so the FOLLOWING blocks flow from
                // y99 OVER the orphaned y518-683 lines (the p37 "text overlap collapse").
                // Re-position the moved lines to page_top and advance the cursor PAST
                // them so the whole paragraph sits cleanly on the new page and the next
                // block flows after it. Scoped to !doc_body_has_real_cjk (pure-Latin
                // docs) so the JP corpus is byte-identical while this is verified; the
                // bug is Phase-1-INVISIBLE (page index is correct, only the within-page
                // y is wrong) so JP passed 87/87 despite it. See [[english_corpus_bug_mine]].
                // CJK documents too (proposal 2026-10-05): the stale-y line is
                // NOT Phase-1-invisible -- the cursor restarts at page_top
                // beneath it, so the page runs one line early. legal__0f631d
                // p187/188: a 3-line widowControl paragraph with two free rows
                // (widow arm at line 1, an orphan split, so a whole move); Word
                // starts all three lines at the next page top (faithful slice
                // widow3_free1 y 99 / next paragraph 138), Oxi kept line 1 at
                // 669.88 and the next paragraph began at 125.19.
                if std::env::var("OXI_S770_DISABLE").is_err()
                    && !elements.is_empty()
                {
                    let min_y = elements.iter().map(|e| e.y).fold(f32::INFINITY, f32::min);
                    let shift = page_top - min_y;
                    if shift.abs() > 0.01 {
                        for e in elements.iter_mut() {
                            e.y += shift;
                            if let LayoutContent::TableBorder { y1, y2, .. } = &mut e.content {
                                *y1 += shift;
                                *y2 += shift;
                            }
                        }
                        let max_y = elements
                            .iter()
                            .map(|e| e.y)
                            .fold(f32::NEG_INFINITY, f32::max);
                        // S792 (2026-07-11): at an ORPHAN break (line_idx == 0)
                        // no text line has been laid yet — `elements` holds only
                        // same-row OVERLAYS (the list marker). Advancing past
                        // them consumed a phantom row: the  dash item's
                        // marker repositioned to page_top but its text line
                        // landed one line BELOW (72/85.8; Word renders both on
                        // one row). Keep the cursor at page_top so line 0
                        // overlays the marker; the widow case (text lines moved)
                        // keeps the advance-past behavior.
                        if line_idx == 0 && std::env::var("OXI_S792_DISABLE").is_err() {
                            cursor.set(page_top);
                            s842_apply(cursor);
                        } else {
                            cursor.set(max_y + line_height);
                        }
                    }
                }
                // Session 107: half-leading at page top (see mid-para break
                // note for full rationale).
                let rule_w = para.style.line_spacing_rule.as_deref();
                let skip_hl_w = matches!(rule_w, Some("exact") | Some("atLeast"));
                // S388 (2026-05-27): tested disabling continuation half-leading
                // (OXI_S388_NO_CONT_HALFLEADING). FALSIFIED as blanket change:
                // b837808d improves +0.021 but d77a58 catastrophically regresses
                // -0.147 (Phase 1 UNCHANGED 53/55, so the half-leading's current
                // role is Phase-2 VISUAL position, not pagination). Word applies
                // it to d77a but apparently not b837 despite both being CJK
                // small-leading grid docs — discriminator unknown, needs COM.
                if line_idx > 0
                    && !skip_hl_w
                    && grid_pitch.map_or(false, |p| p > 0.0)
                    && para.style.snap_to_grid
                    && !in_textbox
                {
                    let hl = ((effective_lh - natural_lh) / 2.0).max(0.0);
                    let leading = effective_lh - natural_lh;
                    if hl > 0.0 && leading < 3.0 {
                        cursor.advance(hl);
                    }
                }
                // Step 0: widow/orphan moves all earlier lines (if any) of
                // this paragraph to the new page. Re-slot any refs already
                // attributed to page 0 into page 1, then open a new bucket
                // for subsequent lines.
                if let Some(v) = line_fn_refs_out.as_deref_mut() {
                    let carry = v.pop().unwrap_or_default();
                    v.push(Vec::new()); // OLD page — nothing of this para stays
                    v.push(carry); // NEW page — earlier lines' refs move here
                }
            } else if needs_page_break {
                // S637: multi-column — when a line overflows the current column
                // and a NEXT column exists on this page, flow into it (Word's
                // newspaper column fill) instead of pushing a new page. Keep the
                // already-laid lines in `elements` (same page) and shift start_x
                // to the next column; subsequent lines emit at the new column x.
                // Fires only on the heterogeneous multi-col path (kyotei), so the
                // 1-col corpus is byte-identical (num_columns==1 → else branch).
                // Opt-out OXI_S637_DISABLE for SSIM A/B verification.
                if num_columns > 1
                    && cur_col + 1 < num_columns
                    && std::env::var("OXI_S637_DISABLE").is_err()
                    // A continuous section can begin with less than one line
                    // left on the page. Its next column has the same short band;
                    // start a fresh page when the line cannot fit there either.
                    && (pages.len() > s749_pages_at_entry
                        || col_band_top <= page_top + 0.025
                        || col_band_top + break_threshold + s1248_after - s835_fn_relief
                            <= effective_break_bottom + s967_tol)
                {
                    cur_col += 1;
                    start_x = col_x_positions[cur_col];
                    cursor.set(column_flow_top(if pages.len() > s749_pages_at_entry {
                        page_top
                    } else {
                        col_band_top
                    }, start_x, pages.len()));
                    s842_apply(cursor);
                    // S1335 (2026-09-06, default ON, opt-out OXI_S1335_DISABLE): a
                    // bare `<w:br w:type="column"/>` line that did not fit the
                    // column it started in has just been carried into the next
                    // column by this overflow -- that IS the break. Letting the
                    // explicit break fire again below moved the paragraph a second
                    // column on (= a page from the last column) and left the
                    // column empty. reference__0ea3ec86 p15 (2-col): 「掛金の減額」
                    // starts with the break at cursor 772.5 of a 785 band; Word
                    // ends column 1 there and starts column 2 with the text, Oxi
                    // pushed a page and left column 2 blank (the +1 cascade of
                    // 840 paragraphs that was compensating the missing S1294 blank
                    // page).
                    if line.break_type == LineBreakType::ColumnBreak
                        && line.fragments.iter().all(|f| f.text.trim().is_empty())
                        && std::env::var("OXI_S1335_DISABLE").is_err()
                    {
                        s1335_break_consumed = true;
                    }
                } else {
                    // Phantom-blank-page fix (2026-04-23): when an empty paragraph
                    // with page_break_after overflows and would produce an empty
                    // stub alone on a new page, followed by ANOTHER page break,
                    // skip the stub entirely. Push the current page and return —
                    // caller's page_break_after path is a no-op on empty elements,
                    // so the next block renders on the fresh page directly.
                    // d77a p.11 case: block 127 is just <w:br w:type="page"/>.
                    // See project_d77a_phantom_page_11.md.
                    // S1055 (default ON, opt-out OXI_S1055_DISABLE): a LATIN doc's
                    // stub is a normal line — it moves to the next page and its
                    // break then sends the following block one page further, which
                    // is exactly Word's blank page. The skip stays for CJK, where
                    // it was derived (d77a's Word truth is 12 pages, no blank).
                    if para.runs.is_empty()
                        && para.style.page_break_after
                        && (self.doc_body_has_real_cjk
                            || std::env::var("OXI_S1055_DISABLE").is_ok())
                    {
                        current_elements.extend(std::mem::take(&mut elements));
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(current_elements),
                        });
                        if let Some(g) = s755_geom {
                            page_top = g.top(pages.len() + 1);
                            content_height = g.ch(pages.len() + 1);
                        }
                        cursor.set(page_top);
                        s842_apply(cursor);
                        return (Vec::new(), 0.0, 0);
                    }
                    // Mid-paragraph page break: keep already-laid-out lines on current page,
                    // only the overflowing line (and subsequent) go to the next page.
                    current_elements.extend(std::mem::take(&mut elements));
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(current_elements),
                    });
                    if let Some(g) = s755_geom {
                        page_top = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(page_top);
                    s842_apply(cursor);
                    if std::env::var("OXI_DBG_EMITY").is_ok() {
                        eprintln!("[EMITY] midpara push: page_top={:.2} cursor={:.2} line_idx={}", page_top, cursor.cursor_y, line_idx);
                    }
                    // S637: a real page push lands on column 0 of the new page.
                    cur_col = 0;
                    // S1013 (2026-07-26, opt-out OXI_S1013_DISABLE): a NATURAL
                    // page break inside a paragraph reset cur_col=0 but left start_x
                    // at the previous column's x — so a column-0 continuation drew
                    // at column 1's x (reports__0013bcb8: para_idx=18's p2 top
                    // continuation rendered at x=306.4 overlapping para_idx=22).
                    // cur_col and start_x are two views of the same column state;
                    // the explicit page/column-break sibling (S733) already resets
                    // BOTH — this is the missing half on the natural-break path.
                    if num_columns > 1 && std::env::var("OXI_S1013_DISABLE").is_err() {
                        start_x = col_x_positions[0];
                    }
                    // Session 107 (2026-05-18): apply half-leading at page top for
                    // grid-snapped lines that are CONTINUATIONS of a paragraph
                    // spilling across page breaks. Word's continuation first line
                    // sits at topMargin + (line_h - natural_lh)/2, not topMargin.
                    // Without this offset, the continuation's content on the next
                    // page is 1-3pt higher than Word, causing later lines in the
                    // SAME paragraph to fit on the wrong page (d77a p.2: line 5
                    // of pi=25 fits in Oxi within natural_lh tolerance but Word
                    // breaks → Oxi has 5 lines vs Word's 4).
                    //
                    // Restrictions:
                    // - line_idx > 0 (continuation only — new paragraphs starting
                    //   on a fresh page have compensating glyph misalignment via
                    //   text_y_off that visually matches Word without the LBT
                    //   shift; applying it there causes regressions)
                    // - skip exact/atLeast rules (V1/V2/V4 minimal repros confirm
                    //   Word does NOT apply half-leading to those)
                    // - leading < 3pt threshold (CJK 12pt+ at grid 18pt has small
                    //   leading where natural_lh leniency alone cannot match
                    //   Word's break decisions; larger leadings like TNR 10.5pt
                    //   (6.5pt) or Mincho 10.5pt (4.5pt) already have enough
                    //   tolerance, and applying the shift there regresses
                    //   db9ca18 / Mincho-heavy docs without page-break benefit)
                    let rule = para.style.line_spacing_rule.as_deref();
                    let skip_half_leading = matches!(rule, Some("exact") | Some("atLeast"));
                    // S388 (2026-05-27): blanket-disable FALSIFIED (see widow site).
                    // S396 (2026-05-28): LRPB-triggered breaks skip continuation
                    // half-leading. Discriminator: when Word inserts
                    // <w:lastRenderedPageBreak/> mid-paragraph, that LRPB position
                    // IS the line top — no half-leading added on top. When break
                    // comes from natural overflow, Session 107's hl still applies
                    // (d77a et al). Localized via b837 dump (pages 2-6 uniformly
                    // +1.5pt step traced to this advance), validated by
                    // OXI_S396_NO_CONT_HL=1 b837 IoU 0.9407 -> 0.9535 (+0.0128).
                    let s396_default = std::env::var("OXI_S396_LRPB_SKIPS_HL")
                        .map(|v| v != "0" && v != "false")
                        .unwrap_or(true);
                    let s396_skip_cont_hl = (s396_default && s391_lrpb_break)
                        || std::env::var("OXI_S396_NO_CONT_HL").is_ok();
                    if !s396_skip_cont_hl
                        && line_idx > 0
                        && !skip_half_leading
                        && grid_pitch.map_or(false, |p| p > 0.0)
                        && para.style.snap_to_grid
                        && !in_textbox
                    {
                        let hl = ((effective_lh - natural_lh) / 2.0).max(0.0);
                        let leading = effective_lh - natural_lh;
                        if hl > 0.0 && leading < 3.0 {
                            cursor.advance(hl);
                        }
                    }
                    // Step 0: lines [0, line_idx) stay on OLD page (their refs
                    // already accumulated in current bucket); open a fresh
                    // bucket so line_idx and beyond register on the NEW page.
                    if let Some(v) = line_fn_refs_out.as_deref_mut() {
                        v.push(Vec::new());
                    }
                } // S637: end multi-column else (page-push path)
            }

            // A section's first paragraph keeps its excess before-spacing
            // when its first line itself triggers the physical page change.
            if line_idx == 0 && pages.len() > s749_pages_at_entry {
                if let Some(spacing) = continuous_section_start_spacing {
                    if !para.style.before_autospacing {
                        cursor.advance(spacing);
                    }
                }
            }

            // Step 0: record fn refs rendered on this line's final page
            // (after any pre-line page push above). A run with footnote_ref
            // maps its marker to this line if any fragment here references
            // that run_index.
            if let Some(v) = line_fn_refs_out.as_deref_mut() {
                let bucket = v.last_mut().unwrap();
                let mut seen: Vec<usize> = Vec::new();
                for f in &line.fragments {
                    if seen.contains(&f.run_index) {
                        continue;
                    }
                    seen.push(f.run_index);
                    if let Some(run) = para.runs.get(f.run_index) {
                        if let Some(id) = run.footnote_ref {
                            if !bucket.contains(&id) {
                                bucket.push(id);
                            }
                        }
                    }
                }
                // S276 (2026-05-25): fn-ref run adjacent-merge fix. After
                // renumber_note_refs rewrites <w:footnoteReference w:id="N"/>
                // markers to single-digit text ("1","2","3",...), adjacent
                // fn-ref runs collapse into a single LineFragment in the
                // line-break loop (word_run_index is only set at word START,
                // not within a word; consecutive digit runs merge as a single
                // "word"). Result: only the FIRST fn-ref's run_index appears
                // in line.fragments; subsequent fn-refs are invisible to the
                // attribution loop above, silently dropping their reservation
                // and rendering. RA repro (5 fns on one para) reproduces.
                // Fix: walk forward from each captured run_idx through
                // consecutive footnote_ref-bearing runs and append their ids.
                // S277 (2026-05-25): flipped to DEFAULT ON. Baseline scan
                // (267 docs in tools/golden-test/documents/docx/) found ZERO
                // paragraphs with 2+ adjacent fn-ref runs → fix is a no-op on
                // baseline (no Phase 1/2/SSIM risk by construction). RA/RD
                // minimal repros (5 and 10 adjacent fn-refs) confirm the fix
                // renders all fns instead of silently dropping fns 2..N.
                // Opt-out via OXI_FN_REF_RUN_SWEEP_DISABLE=1 retained for
                // diagnostic isolation (S269 part 7 hardening pattern).
                let disable = std::env::var("OXI_FN_REF_RUN_SWEEP_DISABLE").is_ok();
                if !disable {
                    let captured: Vec<usize> = seen.clone();
                    for &captured_idx in &captured {
                        let mut i = captured_idx + 1;
                        while i < para.runs.len() {
                            if let Some(id) = para.runs[i].footnote_ref {
                                if !bucket.contains(&id) {
                                    bucket.push(id);
                                }
                                i += 1;
                            } else {
                                break;
                            }
                        }
                    }
                }
                // S826 (2026-07-13, opt-out OXI_S826_DISABLE): a footnote-ref
                // run whose marker merged into a WORD fragment that BEGAN at an
                // earlier run (rsid-split words: "P"+"ractice"+[fnref]+".")
                // never appears as a fragment run_index, and the S276 forward
                // sweep above breaks at the first intermediate TEXT run — the
                // reservation AND the footnote body were silently dropped
                // (uk_framework fn 7 "Practice7."). General attribution:
                // fragments are emitted in run order, so every run index in
                // [this line's first fragment run, next line's first fragment
                // run) lies ON this line — sweep the whole range for
                // footnote_ref runs.
                if std::env::var("OXI_S826_DISABLE").is_err() {
                    let lo = line.fragments.iter().map(|f| f.run_index).min();
                    let hi = lines
                        .get(line_idx + 1)
                        .and_then(|nl| nl.fragments.iter().map(|f| f.run_index).min())
                        .unwrap_or(para.runs.len());
                    if let Some(lo) = lo {
                        for r in lo..hi.max(lo) {
                            if let Some(id) = para.runs.get(r).and_then(|x| x.footnote_ref) {
                                // S900: a deferred note renders on the NEXT
                                // page's area — keep it out of this page's
                                // bucket (the caller re-attributes it).
                                if s900_deferred_ids.contains(&id) {
                                    continue;
                                }
                                if !bucket.contains(&id) {
                                    bucket.push(id);
                                }
                            }
                        }
                    }
                }
            }

            // S776: effective_first_indent includes the suff="nothing" marker
            // width so line 0's text starts right AFTER the number (the break
            // above already narrowed line 0 by the same amount).
            let extra_indent = if line_idx == 0 {
                effective_first_indent
            } else {
                0.0
            };
            // COM-confirmed 2026-04-17 (measure_hanging_indent_v2.py): first-line
            // indent DOES shift line_x. Word places line 1 at margin+indent_left+
            // first_line_indent, continuation lines at margin+indent_left. Applies
            // to both positive firstLine (e.g. +21pt) and hanging (negative, e.g. -9pt).
            // S758 side-wrap: an in-band line (not yet rebroken at the band
            // exit) shifts to the free segment's start (left float -> text
            // flows right of the box) and justifies within the NARROWED
            // width (previously the paint stretched a narrowed line to the
            // full width -- the documented paint-only flaw).
            let s758_line_shift = if let Some((_, shift)) = region_line_widths.get(line_idx) { *shift } else if !s758_rebroken {
                s758_band.map(|(_, _, sh)| sh).unwrap_or(0.0)
            } else {
                0.0
            };
            let s758_line_red = if let Some((red, _)) = region_line_widths.get(line_idx) { *red } else if !s758_rebroken {
                s758_band.map(|(_, red, _)| red).unwrap_or(0.0)
            } else {
                0.0
            };
            // S-TWOSEG: a two-segment row starts at the left strip's own x, not
            // at the paragraph indent, and jumps to the right strip mid-row (the
            // fragment loop below does the jump at `seg2_at`).
            let row_two_segments=word_fit_segments.get(line_idx).copied()
                .unwrap_or(if !s758_rebroken {s758_two_seg} else {None});
            let line_x = match (row_two_segments, line.seg2_at) {
                (Some((seg1_x, _, _, _)), Some(_)) => seg1_x + extra_indent,
                _ => start_x + indent_left + extra_indent + s758_line_shift,
            };

            // Alignment offset
            let line_text_width: f32 = line.fragments.iter().map(|f| f.width).sum();
            let is_last_line = line_idx == lines.len() - 1;
            // For alignment/justify, use indent-adjusted width.
            // Justify/alignment uses indent-adjusted width. In charGrid mode,
            // break_into_lines may put more chars than fit in the indented area;
            // negative slack triggers punctuation compression to fit.
            let render_width = content_width - indent_left - indent_right - s758_line_red;
            let align_offset = match para.alignment {
                Alignment::Left => 0.0,
                Alignment::Center => {
                    // Word GDI: integer pixel division at 96dpi for center alignment
                    let slack_tw =
                        ((render_width - extra_indent - line_text_width) * 20.0).round() as i32;
                    let center_tw = slack_tw / 2; // integer division (truncate)
                    center_tw as f32 / 20.0
                }
                Alignment::Right => render_width - extra_indent - line_text_width,
                Alignment::Justify => 0.0,
                // Distribute: when justification applies (multi-fragment lines), offset is 0
                // because slack is distributed across fragments. When justification can't
                // apply (single-fragment line), center the content.
                Alignment::Distribute => {
                    if line.fragments.len() > 1 {
                        0.0
                    } else {
                        let slack = render_width - extra_indent - line_text_width;
                        if slack > 0.0 {
                            slack / 2.0
                        } else {
                            0.0
                        }
                    }
                }
            };

            // Justification (matches Word output, priority order):
            // 1. CJK punctuation compression (full-width -> half-width, 50% savings)
            // 2. Word-space expansion (distribute remaining slack at space characters)
            // Latin text: ONLY expand at word spaces, never between characters.
            // CJK text: compress punctuation first, then expand at inter-character gaps.

            let mut frag_width_adjustments: Vec<f32> = vec![0.0; line.fragments.len()];
            let mut frag_spacing_after: Vec<f32> = vec![0.0; line.fragments.len()];
            let mut justify_char_spacing: f32 = 0.0;
            // S1117 (2026-08-14, HELD OPT-IN OXI_S1117 — see the gate result at the end
            // of this comment): a compat>=15 justified line may
            // be OVER-FULL by design — Word 2013+ fits a word whose natural width
            // exceeds the column and then COMPRESSES the inter-word spaces to land
            // exactly on the right margin. Oxi's BREAK already reproduces that
            // (repro `pipeline_data/_pb_spacecomp2/pset_m15_only.docx`: both Word
            // and Oxi carry "forecast" onto line 1) but the render leaves the
            // spaces natural, so the line runs 9.87pt PAST the margin — violating
            // the invariant recorded below ("Word never overshoots the right
            // margin"). The compression mechanism already exists as S994's negative
            // slack path; it was scoped to wpJustification (2 docs) because that
            // was the only known producer of a wider-than-column line. compat>=15
            // is the other, and it is 412 of the 719 corpus documents.
            //   Word truth (`_pb_spacecomp2`, same paragraph, only settings.xml
            //   differing):  mode 15 -> space 1.5637 = 0.724x natural, line ends ON
            //   the margin;  mode 14/12 -> space 3.7051 = 1.715x, "forecast" wraps.
            //   Oxi mode 15 -> break correct, space 2.2020 = natural, +9.87 overrun.
            // Latin scope (!doc_body_has_real_cjk): a CJK line's negative slack is
            // Phase 1's yakumono business and it carries no ASCII spaces to give.
            // ★★SHIPS ONLY AS A PAIR WITH S1118 (2026-08-14). Measured alone this
            // rule LOSES: SSIM sentinel A/B over 238 word_png bases moved one
            // document the wrong way (db9ca18368cd net −0.0014, 0 improved / 1
            // regressed), because it compresses onto a target that is itself 2.20pt
            // short — S1118's trailing-space bug. Paired with S1118 the same
            // document goes net **+0.1349** (1 improved / 0 regressed) and the repro
            // lands every line exactly on the margin. So `OXI_S1117_DISABLE` alone
            // does NOT return the tree to a neutral state: it leaves S1118 widening
            // lines that this rule is no longer reining in. Disable both or neither.
            let s1117_compat15_overfull = self.compat_mode >= 15
                && !self.doc_body_has_real_cjk
                && std::env::var("OXI_S1117_DISABLE").is_err();

            let is_soft_break_line =
                self.do_not_expand_shift_return && line.break_type == LineBreakType::SoftBreak;
            let should_justify = !in_textbox
                && !is_soft_break_line
                && ((para.alignment == Alignment::Justify && !is_last_line)
                    || para.alignment == Alignment::Distribute);
            if should_justify && line.fragments.len() > 1 {
                // S472 (break-agnostic render): when on, the yakumono compression for
                // RENDER is computed here purely from natural widths vs available
                // (Word's demand model), independent of whatever the break decided —
                // replacing the legacy Phase 1 (×0.5) + Stage 2b (restore) dance which
                // was tuned to the old break-time 、=8.0 pre-compression.
                let s472_render = std::env::var("OXI_S472_DEMAND").is_ok()
                    || std::env::var("OXI_S473_LOCOMP").is_ok()
                    // S475 routes render through the demand water-fill ONLY for the
                    // no-char-grid (type=lines) docs it actually re-breaks; on
                    // linesAndChars (b837) S475's break is off, so leave render
                    // untouched (else it cascades b837 pagination 7→8). Default-ON
                    // (opt-out OXI_S475_DISABLE) to match the break gate.
                    || (std::env::var("OXI_S475_DISABLE").is_err()
                        && self.compress_punctuation && self.compat_mode >= 15
                        && (!page.doc_grid_lines_and_chars
                            || std::env::var("OXI_S476_DISABLE").is_err()));
                // charGrid: subtract grid extra from slack. Grid extra widens chars
                // for positioning but is NOT distributable justify space.
                let grid_extra_on_line = if let Some(pitch) = effective_char_pitch {
                    line.fragments
                        .iter()
                        .map(|f| {
                            let fs = f.style.font_size.unwrap_or(para_font_size);
                            f.text
                                .chars()
                                .filter(|&c| crate::font::is_fullwidth(c) && fs < pitch)
                                .count() as f32
                                * (pitch - fs)
                        })
                        .sum::<f32>()
                } else {
                    0.0
                };
                // ★OPEN (2026-08-14): every justified Latin line lands exactly 2.20pt
                // short of the right margin — 553.10 against a 555.30 edge, on L0/L1/L2
                // of `_pb_spacecomp2/pset_m15_only.docx` alike, with different content
                // and different space counts each. 2.20 is exactly Cambria 10pt's
                // natural space (2.2021), and Word puts those same lines ON the margin
                // (last visible char at 554.72..554.97 = margin minus side bearing).
                // The obvious mechanism — `line_text_width` charging the line's
                // trailing space against the slack while the distribution loops below
                // skip the last fragment — was IMPLEMENTED AND MEASURED TWICE (as a
                // whole all-space last fragment, then as the trailing space run inside
                // the last fragment) and was a byte-exact NO-OP both times: the last
                // fragment carries no trailing space at all. So the constant is real
                // but its source is NOT the trailing space. Instrument the three terms
                // (render_width / extra_indent / line_text_width) before guessing a
                // third time.
                let mut slack = render_width - extra_indent - line_text_width - grid_extra_on_line;
                // OXI_DBG_SLACK: the three terms behind the unexplained 2.20pt
                // undershoot, plus the fragment tail so the trailing-space question
                // is answered from data rather than from the source shape.
                if std::env::var("OXI_DBG_SLACK").is_ok() {
                    let nf = line.fragments.len();
                    let tail: Vec<String> = line
                        .fragments
                        .iter()
                        .skip(nf.saturating_sub(3))
                        .map(|f| format!("{:?}/w={:.3}", f.text, f.width))
                        .collect();
                    eprintln!(
                        "[SLACK] li={} render_w={:.3} extra_ind={:.3} text_w={:.3} grid={:.3} slack={:.3} nfrag={} tail={}",
                        line_idx, render_width, extra_indent, line_text_width,
                        grid_extra_on_line, slack, nf, tail.join(" ")
                    );
                }

                // S472 break-agnostic demand compression (replaces Phase 1 + Stage 2b
                // below when on). Reset standalone 、。 to natural, compute the line's
                // overflow vs available, and distribute that compression across them
                // cap-aware (、,，→fontSize/3 floor=8.0pt; 。．→fontSize/2 floor=6.0pt)
                // via even water-filling. This lands each 、。 on Word's demand-driven
                // advance regardless of the break-time width. Under-full lines reset 、
                // to natural and let Phase 2 distribute.
                if s472_render {
                    let nfr = line.fragments.len();
                    let mut comps: Vec<(usize, f32, f32)> = Vec::new(); // (fi, fs, cap)
                    let mut nat_total = 0.0f32;
                    for fi in 0..nfr {
                        let f = &line.fragments[fi];
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let c0 = f.text.chars().next().unwrap_or(' ');
                        let single = f.text.chars().count() == 1;
                        // cap per char type: 、,，→fs/3 (8.0); 。．& closing brackets→fs/2
                        // (6.0). Opening brackets never compress. width>fs*0.6 excludes
                        // already-pair-compressed (6.0) fragments.
                        // [S475 render-distribution lever B TRIED + REVERTED: a uniform
                        // measured-model cap (openers/、/closing → 1.5→10.5, pair → 6.0)
                        // matched d77a L1 punct exactly (10.5) but REGRESSED the other
                        // d77a pages net −0.0329 — Word's render is per-line-VARIABLE,
                        // not uniform ~10.5; the uniform cap shifts kanji positions and
                        // misaligns (S468 lesson). The break is correct; the render
                        // residual is not cleanly fixable by uniform caps.]
                        let cap = if !single {
                            0.0
                        } else {
                            match c0 {
                                '、' | '，' => fs / 3.0,
                                '。' | '．' => fs / 2.0,
                                '」' | '』' | '】' | '〕' | '》' | '〉' | '｝' | '］' | '）' => {
                                    fs / 2.0
                                }
                                // S578 (2026-06-15): ・ (nakaguro) compresses on demand at
                                // RENDER too. The BREAK already budgets ・ (s475_max_compress
                                // handles it) but the render water-fill OMITTED it (cap 0 → ・
                                // stuck at natural 12.0) = a break/render inconsistency (the
                                // exact class S573 flagged). Word compresses ・ demand-driven:
                                // median ~11.5 (light), down to 5.14 on tight lines (d77a) =
                                // cap ≈ fs/2, same class as 。/closing brackets. MEASURED 3-doc
                                // (_cb_yakumono_compare, Word PDF vs Oxi: ・ signed Oxi−Word
                                // = b837 +0.53, d77a +0.94, ikujikaigo +0.77 — uniformly
                                // UNDER-compressed). Opt-out OXI_S578_DISABLE.
                                '・' if std::env::var("OXI_S578_DISABLE").is_err() => fs / 2.0,
                                _ => 0.0,
                            }
                        };
                        if cap > 0.0 && f.width > fs * 0.6 {
                            comps.push((fi, fs, cap));
                            nat_total += fs;
                        } else {
                            nat_total += f.width;
                        }
                    }
                    let nat_slack = render_width - extra_indent - nat_total - grid_extra_on_line;
                    let mut comp_amt = vec![0.0f32; nfr];
                    if nat_slack < 0.0 && !comps.is_empty() {
                        let mut needed = -nat_slack;
                        let mut active: Vec<(usize, f32)> =
                            comps.iter().map(|(fi, _, cap)| (*fi, *cap)).collect();
                        loop {
                            if active.is_empty() || needed <= 0.001 {
                                break;
                            }
                            let share = needed / active.len() as f32;
                            let capped: Vec<(usize, f32)> = active
                                .iter()
                                .cloned()
                                .filter(|(_, cap)| *cap <= share)
                                .collect();
                            if capped.is_empty() {
                                for (fi, _) in &active {
                                    comp_amt[*fi] = share;
                                }
                                break;
                            }
                            for (fi, cap) in &capped {
                                comp_amt[*fi] = *cap;
                                needed -= cap;
                            }
                            active.retain(|(_, cap)| *cap > share);
                        }
                    }
                    // Set adjustments so each standalone 、。 renders at (natural − comp).
                    for (fi, fs, _) in &comps {
                        frag_width_adjustments[*fi] =
                            (fs - comp_amt[*fi]) - line.fragments[*fi].width;
                    }
                }

                // Phase 1: CJK punctuation compression (full-width -> half-width)
                // Only compress when the line overflows (slack < 0).
                // Matches Word output: TextBox content does NOT use punctuation compression.
                // 2026-04-20 fix: Skip chars whose fragment.width is ALREADY smaller than
                // natural (indicates break_into_lines already compressed them — applying
                // Phase 1 again would DOUBLE-compress, crushing 「」 to w=0pt).
                if slack < 0.0 && !in_textbox && !s472_render {
                    for (fi, frag) in line.fragments.iter().enumerate() {
                        for ch in frag.text.chars() {
                            if kinsoku::is_cjk_compressible(ch) {
                                // Opening brackets have large ABC A-offset (glyph on
                                // right side of cell). Compressing advance to 6pt
                                // causes glyph to extend past cell, overwritten by
                                // next char. Keep fullwidth advance for these.
                                let is_opening_bracket = matches!(
                                    ch,
                                    '（' | '「' | '『' | '〔' | '【' | '《' | '〈' | '｛' | '［'
                                );
                                if is_opening_bracket {
                                    continue;
                                }
                                let fs = frag.style.font_size.unwrap_or(para_font_size);
                                let fm = &*self.metrics_for(&frag.style, &para.style);
                                let char_w = self.registry.char_width_pt_with_fallback(ch, fs, fm);
                                // Skip if fragment.width is already below fullwidth
                                // (break_into_lines already applied yakumono compression
                                // 0.5x or 0.583x). Re-applying 0.5x here would
                                // double-compress, crushing 」、 to near-zero.
                                if frag.width + frag_width_adjustments[fi] < char_w * 0.95 {
                                    continue;
                                }
                                let actual = char_w * 0.5;
                                frag_width_adjustments[fi] -= actual;
                                slack += actual; // reclaim freed space
                            }
                        }
                    }
                }

                // 2026-04-20: Recompute slack after Phase 1.
                // Grid-extra handling is branch-dependent:
                //   - If Phase 1 compressed chars (had yakumono etc.): grid_extra is
                //     already reclaimed; use post_phase1_ltw directly (no subtract).
                //   - If NO compression happened (pure CJK line): natural widths
                //     already match render; grid_extra is positioning padding that
                //     shouldn't be distributed → subtract it.
                let phase1_compressed = frag_width_adjustments.iter().any(|a| *a < -0.01);
                let post_phase1_ltw: f32 = line
                    .fragments
                    .iter()
                    .enumerate()
                    .map(|(i, f)| f.width + frag_width_adjustments[i])
                    .sum();
                // S1118 (2026-08-14, opt-in OXI_S1118): a justified line's TRAILING
                // space fragment HANGS past the right margin in Word — it is not part
                // of the content being justified. `post_phase1_ltw` sums EVERY
                // fragment, so that space is charged against the slack while the
                // distribution loops below skip the last fragment (`fi < len-1`) and
                // never give it back. The visible content is therefore justified to
                // (render_width − one space) and every such line lands exactly one
                // space-advance short of the margin.
                //   Word truth (`_pb_spacecomp2/pset_m15_only.pdf`, Cambria 10): the
                //   last VISIBLE char sits at 554.72 / 554.97 / 554.94 against a
                //   555.30 margin (remainder = side bearing) and the trailing space's
                //   own origin is AT 554.91..555.01 — it starts on the margin and
                //   hangs over. Oxi ships those lines at 553.10 = 555.30 − 2.20, and
                //   2.20 is exactly Cambria 10pt's natural space (2.2021).
                // ★This adjustment MUST live here, not at the first `slack` binding:
                // that one is unconditionally RECOMPUTED from post_phase1_ltw a few
                // lines up, which silently discarded two earlier attempts at this rule
                // (both measured byte-exact no-ops, which read as "the trailing space
                // isn't there" — OXI_DBG_SLACK shows it plainly IS: tail=" "/w=2.202).
                let s1118_trailing_hang = if std::env::var("OXI_S1118_DISABLE").is_err()
                    && !self.doc_body_has_real_cjk
                {
                    line.fragments
                        .last()
                        .filter(|f| !f.text.is_empty() && f.text.chars().all(|c| c == ' '))
                        .map(|f| f.width + frag_width_adjustments[line.fragments.len() - 1])
                        .unwrap_or(0.0)
                } else {
                    0.0
                };
                slack = if phase1_compressed {
                    render_width - extra_indent - post_phase1_ltw + s1118_trailing_hang
                } else {
                    render_width - extra_indent - post_phase1_ltw + s1118_trailing_hang
                        - grid_extra_on_line
                };

                // 2026-04-21: Stage 2b — de-compress pre-compressed 、。,．toward
                // natural using positive slack. COM-proven (d77a + R19 + R6):
                // Word variable 、 advance 9.5-12pt correlates with line overflow
                // demand. Oxi's break-time 0.583× (= 7pt) is a wrap-budget knob,
                // but at render Word restores 、 toward natural when slack allows.
                // Safe because: only fires when slack > 0, only de-compresses
                // (never over-extends).
                if slack > 0.5 && !in_textbox && !s472_render {
                    let mut compressed: Vec<(usize, f32, f32)> = Vec::new();
                    for (fi, frag) in line.fragments.iter().enumerate() {
                        for ch in frag.text.chars() {
                            if !matches!(ch, '、' | '。' | '，' | '．') {
                                continue;
                            }
                            let fs = frag.style.font_size.unwrap_or(para_font_size);
                            // 、 is CJK punct — use CJK metrics if available.
                            let fm_cjk = self.metrics_for_cjk(&frag.style, &para.style);
                            let fm = fm_cjk
                                .unwrap_or_else(|| self.metrics_for(&frag.style, &para.style));
                            let natural = self.registry.char_width_pt_with_fallback(ch, fs, &fm);
                            let current = frag.width + frag_width_adjustments[fi];
                            if current < natural * 0.95 {
                                compressed.push((fi, natural, current));
                            }
                        }
                    }
                    if !compressed.is_empty() {
                        let n = compressed.len() as f32;
                        let per_comp_cap = compressed
                            .iter()
                            .map(|(_, nat, cur)| nat - cur)
                            .fold(f32::INFINITY, f32::min);
                        let per_comp = (slack / n).min(per_comp_cap);
                        if per_comp > 0.0 {
                            for (fi, _, _) in &compressed {
                                frag_width_adjustments[*fi] += per_comp;
                                slack -= per_comp;
                            }
                        }
                    }
                }

                // Phase 2: Distribute remaining slack at word spaces (only if slack > 0 after compression)
                if slack > 0.0 {
                    // Count ASCII word spaces only — CJK fullwidth spaces (U+3000) are NOT
                    // word boundaries for justify purposes.
                    let space_count = line
                        .fragments
                        .iter()
                        .enumerate()
                        .filter(|(i, f)| {
                            *i < line.fragments.len() - 1
                                && f.text.chars().all(|c| c == ' ')
                                && !f.text.is_empty()
                        })
                        .count();

                    if space_count > 0 {
                        let per_space = slack / space_count as f32;
                        for (fi, frag) in line.fragments.iter().enumerate() {
                            if fi < line.fragments.len() - 1
                                && frag.text.chars().all(|c| c == ' ')
                                && !frag.text.is_empty()
                            {
                                frag_spacing_after[fi] += per_space;
                            }
                        }
                    } else {
                        // No word spaces: distribute between CJK characters.
                        // Use character_spacing on each fragment so Canvas/PDF renderers
                        // apply per-character gap (not just fragment-level gap).
                        let total_chars: usize =
                            line.fragments.iter().map(|f| f.text.chars().count()).sum();
                        let has_cjk = line
                            .fragments
                            .iter()
                            .any(|f| f.text.chars().any(|c| kinsoku::is_cjk(c)));
                        if has_cjk && total_chars > 1 {
                            let char_gap_count = total_chars - 1;
                            let per_char_gap = slack / char_gap_count as f32;
                            // S627 (2026-06-19) ATTEMPTED + FALSIFIED + REVERTED: DISCRETE
                            // error-diffused justify. Word distributes justify slack as discrete
                            // δ≈0.11pt bumps on ~16% of chars (measured), NOT the uniform
                            // per_char_gap Oxi spreads on every inter-char gap — the +0.039
                            // "X-jitter" (fitz-position metric: WORD-x recovers it). Implemented
                            // here (justified CJK lines are ~1-char-per-fragment, so the expansion
                            // lives in frag_spacing_after) by error-diffusing the cumulative gap to
                            // a δ grid. RESULT: DWrite-gate slightly REGRESSED across all δ
                            // (0e7af −0.0003..−0.0007, 683f −0.0001..−0.0013). ROOT: cumulative-
                            // round error-diffusion produces A discrete pattern but NOT Word's
                            // EXACT bump PLACEMENT (which specific gap gets each δ) — so the
                            // per-char positions still differ from Word, just differently than
                            // uniform → no match. The X-jitter needs Word's exact justify
                            // distribution ALGORITHM (gap selection), not generic error-diffusion.
                            // The deep reverse-engineering wall (convergent with char_budget_wall).
                            // S627 (2026-06-19) ATTEMPTED 3 WAYS + ALL FALSIFIED + REVERTED.
                            // The X-jitter (+0.039): Word distributes justify slack as discrete
                            // δ≈0.11pt bumps (cumulative-round-to-δ staircase, even-distributed,
                            // EXCLUDING 約物 gaps — the [3,7,11,16,20,24,29,33] / [4,4,5] pattern),
                            // vs Oxi's uniform per_char_gap on every gap. Tried: (a) renderer
                            // character_spacing error-diffusion (INERT — justified CJK lines are
                            // ~1-char-per-fragment, expansion lives in frag_spacing_after, not
                            // character_spacing); (b) layout cumulative-round-to-δ (regressed
                            // −0.0003..−0.0013: a +0.02pt base offset from 約物/char-width
                            // precision shifts the round-crossings by ~1 near δ-boundaries); (c)
                            // layout even-distribution Bresenham-round (regressed −0.0005..−0.0011:
                            // includes 約物 gaps which Word EXCLUDES, + phase off). Each produces
                            // A discrete pattern but NOT Word's EXACT per-char positions → no
                            // closer than uniform, slight regression. ⇒ the X-jitter needs Word's
                            // EXACT justify algorithm (even-distribute n=round(slack/δ) bumps over
                            // the NON-約物 expandable gaps, exact phase) AND per-char base-width
                            // precision (<0.02pt). Unified with the Y-jitter: both = cumulative-
                            // device-snap(δ≈0.11pt) limited by base precision (char-width X /
                            // line-height Y). The deep per-font wall; see memory.
                            // Distribute: fragment-boundary gaps via frag_spacing_after,
                            // internal gaps via frag_width_adjustments (for layout width),
                            // AND set justify_char_spacing for renderer to apply letterSpacing.
                            for fi in 0..line.fragments.len() {
                                let frag_chars = line.fragments[fi].text.chars().count();
                                if frag_chars > 1 {
                                    frag_width_adjustments[fi] +=
                                        per_char_gap * (frag_chars - 1) as f32;
                                }
                                if fi < line.fragments.len() - 1 {
                                    frag_spacing_after[fi] += per_char_gap;
                                }
                            }
                            // Store per_char_gap for use in LayoutElement character_spacing
                            justify_char_spacing = per_char_gap;
                        }
                        // Pure Latin with no spaces: do NOT add inter-character spacing
                    }
                } else if slack < 0.0
                    && (self.wp_justification || s1117_compat15_overfull)
                    && std::env::var("OXI_S994_DISABLE").is_err()
                {
                    // S994 render: wpJustification selected content WIDER than the
                    // physical line (the break used the W_actual × (1 + 281/7200)
                    // budget, MS-OE376 §2.1.481). Compress the inter-word ASCII
                    // spacing back to the physical width — symmetric to the positive
                    // expansion above. Scoped to wp_justification (2 docx_corpus/en
                    // docs) → byte-identical elsewhere by construction.
                    let space_count = line
                        .fragments
                        .iter()
                        .enumerate()
                        .filter(|(i, f)| {
                            *i < line.fragments.len() - 1
                                && f.text.chars().all(|c| c == ' ')
                                && !f.text.is_empty()
                        })
                        .count();
                    if space_count > 0 {
                        let per_space = slack / space_count as f32; // negative → compress
                        for (fi, frag) in line.fragments.iter().enumerate() {
                            if fi < line.fragments.len() - 1
                                && frag.text.chars().all(|c| c == ' ')
                                && !frag.text.is_empty()
                            {
                                frag_spacing_after[fi] += per_space;
                            }
                        }
                    }
                }
            }

            // 2-pass wrap Stage 4/5: context-aware 「 leading gap.
            // Only applies to docs with compressPunctuation+compat15 (where Word's
            // measured shifts originate). doNotCompress docs use Oxi's natural
            // positioning (no shift) — tested 0e7a / 683f regression when S5
            // applied unconditionally.
            // 2026-04-21: Stage 4/5 「-leading +6pt removed.
            //
            // The previous implementation added +6pt to frag_spacing_after[fi-1]
            // when fragment fi started with an opening bracket and fi-1 ended
            // with CJK. This was POST-wrap (added after break_into_lines), so
            // wrap-time current_width never accounted for it. Result: lines
            // that fit at wrap time would overshoot the right margin by 6-12pt
            // at render time (1 char beyond margin per +6pt extra).
            //
            // User observation 2026-04-21: Word never overshoots the right
            // margin (strict invariant). Disabling this gap brings Oxi closer
            // to Word's no-overshoot behavior.
            //
            // Full baseline verify: bottom-5 sum 3.2451 → 3.2464 (+0.0013,
            // d77a p9 +0.0012). 53 pages improved (d77a p1 +0.0028, p4 +0.0011,
            // p5 +0.0028, p8 +0.0050; e8caed +0.0047; c7b923 +0.0035 etc).
            // 3 minor regressions outside bottom-5 (b837 p1 -0.0026, d77a p6
            // -0.0019, d77a p12 -0.0015). Net +0.1253.

            let mut x = line_x + align_offset;

            // S672 (2026-06-26): render-x separation for Latin. Oxi LINE-BREAKS at
            // com_tw (GDI/10tw-rounded) word widths — correct for PAGINATION (matches
            // Word's GDI-screenshot wrap, [[latin_text_wrap_compression]]) — but the
            // DWrite renderer draws each word at its TRUE em advance, so positioning
            // word/space fragments at the com_tw cumulative x renders them COMPRESSED
            // (drifting left ~0.5pt/word vs Word, whose word_png screenshot IS true-wide,
            // UPDATE 2). FIX: keep the break at com_tw (x, pagination unchanged) but emit
            // each fragment at a PARALLEL TRUE cumulative x (render_x, sum of un-rounded
            // char_width_em × fs). RENDER-ONLY (element.x is horizontal; pagination
            // depends on line count/heights, not x) → Phase-1-safe. SCOPE: a pure-Latin
            // (no CJK char) LEFT-aligned line only — the renderer-side UPDATE-3 attempt
            // hit an element-granularity/alignment blocker the LAYOUT resolves (it knows
            // alignment, columns, fragment structure). Justify/center/right and CJK/mixed
            // lines keep the com_tw x (justify fills to the margin; CJK is on-grid).
            // Opt-out OXI_S672_DISABLE.
            // TAB/FIELD exclusion: a tab fragment is positioned at an ABSOLUTE
            // tab stop (not a char-cumulative advance), and a field (page number
            // etc.) has special positioning — the true-cumulative render_x would
            // mis-place everything after them. (Corpus A/B: test_tabs −0.0023 without
            // this guard.)
            // S762 (2026-07-08): CURLY QUOTES (U+2018-201D) do not disqualify a
            // line — is_cjk() classifies General Punctuation as CJK, so ONE
            // apostrophe («Contractor’s», nyserda/real English legal text)
            // killed S672 for the whole line and the com-compressed words
            // swallowed their following spaces («executedand deliveredto»).
            // In an otherwise-Latin line the quotes are Latin-context by
            // construction (the LATINQUOTE break side already widths them as
            // Latin), so exempt exactly these four chars here. Default ON,
            // opt-out OXI_S762_DISABLE.
            let s762_quote_ok = std::env::var("OXI_S762_DISABLE").is_err();
            let s672_candidate = std::env::var("OXI_S672_DISABLE").is_err()
                && matches!(para.alignment, Alignment::Left)
                && !line.fragments.is_empty()
                && line.fragments.iter().all(|f| {
                    !f.text.chars().any(|c| {
                        crate::font::is_complex_script(c)
                            || (kinsoku::is_cjk(c)
                                && !(s762_quote_ok
                                    && matches!(c, '\u{2018}' | '\u{2019}' | '\u{201C}' | '\u{201D}')))
                    }) && f.tab_alignment.is_none()
                        && f.field_type.is_none()
                });
            // TRUE-WIDER gate: apply render-x ONLY when the line's true em width
            // EXCEEDS its com_tw (break) width — i.e. Word renders it TRUE-WIDE and
            // Oxi's com_tw is COMPRESSED. MEASURED (gen_report): an 11pt body line is
            // true-wider (com_tw compressed) → render-x EXPANDS to match Word; but a
            // 26pt TITLE line is true-NARROWER (com_tw=443px=Word EXACTLY, true=437px)
            // → render-x would WRONGLY compress it. So shift only toward Word's true-
            // wide render, never introduce compression. This auto-excludes large
            // titles + any line where com_tw is already correct (principled, no
            // arbitrary font-size threshold).
            let (line_comtw, line_true): (f32, f32) = if s672_candidate {
                line.fragments
                    .iter()
                    .enumerate()
                    .map(|(fi, f)| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                        let tw: f32 = f.text.chars().map(|c| m.char_width_em(c) * fs).sum();
                        (f.width + frag_width_adjustments[fi], tw)
                    })
                    .fold((0.0_f32, 0.0_f32), |(a, b), (c, d)| (a + c, b + d))
            } else {
                (0.0, 0.0)
            };
            let s672_latinx = s672_candidate && line_true > line_comtw + 0.01;
            let mut render_x = x;

            // Matches Word output: exact/atLeast line spacing places text at BOTTOM of line box.
            // Extra space goes above text (ascent increased, descent unchanged).
            // Session 76 Mech A fix: pass in_textbox so the function can distinguish
            // body/cell (top-align for exact) from shape (bottom-align).
            // S619/S620 (2026-06-19) — the gen2 residual is a REAL downward vertical DRIFT,
            // NOT "the SSIM optimum" (an earlier wrong conclusion, corrected). word_png 2D
            // per-band best-shift (gen2_005 p1): the top body is aligned (dy=0) but lower
            // bands are 1-2px TOO LOW and recover **+0.03..+0.10 SSIM each** with a 1-2px
            // UP-shift (e.g. rows 1410-1530: 0.910→0.987 @ dy=-2). So content accumulates
            // ~2px (1pt) too low by the page bottom = a real slope, recoverable per-band. It
            // is NOT a UNIFORM offset (top aligned → uniform shift regresses, OXI_GLOBAL_DY
            // +0.5=+1.2 worse) and NOT the body line-height (VSNAP-nosnap == OFF, byte-
            // identical). The drift SOURCE is untraced (NOT body lines/spacing-fold; candidate:
            // title-block/heading height, or a per-line 96dpi snap-phase the body lines don't
            // carry). Plus the TABLE band (rows ~930) is structurally misaligned (no shift
            // helps = horizontal/layout). gen2 p1 absolute SSIM ≈ 0.91. S614 (+0.1424) was a
            // real win but the vertical stack is NOT exhausted — the drift is a real fixable
            // error whose source must be isolated per-element (each block's height vs Word).
            let text_y_off = self.text_y_offset_for_line(
                line,
                &para.style,
                para_font_size,
                line_height,
                grid_pitch,
                in_textbox,
                page.doc_grid_no_type,
            );
            // The grid capacity and the painted baseline use separate boxes.
            // Center the composed host/math placement box on this actual line;
            // retain established placement for other stories and MATH-only runs.
            let text_y_off = if !in_textbox && !is_header_footer
                && !page.doc_grid_no_type && para.style.snap_to_grid
                && grid_pitch.is_some_and(|pitch| pitch > 0.1)
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && !line.fragments.iter().any(|f| f.style.inline_object_image.is_some()
                    || f.style.hr_rule.is_some())
            {
                let math_box = line.fragments.iter().filter_map(|f| {
                    f.style.inline_math.as_ref().and_then(|block|
                        crate::layout::math::inline_math_baseline_extent(block,
                            f.style.font_size.unwrap_or(para_font_size)))
                }).fold(None, |boxes: Option<(f32, f32)>, (a, d)|
                    Some(boxes.map_or((a, d), |(ba, bd)| (ba.max(a), bd.max(d)))));
                if let Some((a, d)) = math_box {
                    let host_box = line.fragments.iter().filter(|f|
                        f.text != "\u{FFFC}" && f.style.inline_math.is_none())
                        .map(|f| self.metrics_for_text(&f.text, &f.style, &para.style)
                            .design_font_box_pt(f.style.font_size.unwrap_or(para_font_size), true))
                        .fold((0.0_f32, 0.0_f32), |(a, d), (fa, fd)| (a.max(fa), d.max(fd)));
                    // Use the same host run as inline-math emission, so both
                    // host glyphs and the formula keep one shared baseline.
                    if let Some(host) = line.fragments.iter()
                        .find(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                    {
                        let size = host.style.font_size.unwrap_or(para_font_size);
                        let metrics = &*self.metrics_for_text(&host.text, &host.style, &para.style);
                        (line_height + a.max(host_box.0) - d.max(host_box.1)) * 0.5
                            - metrics.win_ascent * size + 1.0
                    } else { text_y_off }
                } else { text_y_off }
            } else { text_y_off };
            let text_y_off = if header_exact_inline
                && line.fragments.iter().any(|f| f.style.inline_object_image.is_some())
            {
                // The image run's font does not define the visible text ascent.
                // Keep the fixed-line baseline shared by the text and image.
                line.fragments.iter()
                    .find(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                    .map(|f| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                        0.8 * line_height - m.win_ascent * fs + 1.0
                    })
                    .unwrap_or(text_y_off)
            } else { text_y_off };
            // S837 (2026-07-14, default ON, opt-out OXI_S837_DISABLE): a line
            // hosting an INLINE visual drawing (the S773 bump) shares its
            // BASELINE with the object's BOTTOM — Word render truth (hmrc
            // title: Calibri-Bold 18 baseline 75.24 = crown bottom 75.1;
            // Word line box = cy + text hhea DESCENT — the S773 probe's
            // "+0.25×fs" was Calibri's 512/2048 descent ratio, resolving that
            // coincidence). The old centering ((line_h − text_hhea)/2) drew
            // the title ~23pt too high (the doubled «Starter Checklist» in
            // the 3-panel). Place the glyph baseline at box_top + cy using
            // the S614 DWrite convention (baseline = el.y + text_y_off − 1.0
            // + win_ascent×fs). Fires only when S773 fired (hmrc-only by the
            // S773 corpus scan) → corpus byte-identical by construction.
            let text_y_off = if line_idx == 0
                && s837_fired_cy > 0.0
                && std::env::var("OXI_S837_DISABLE").is_err()
            {
                let mut off = text_y_off;
                if let Some(f) = line
                    .fragments
                    .iter()
                    .find(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                {
                    let fs = f.style.font_size.unwrap_or(para_font_size);
                    let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                    off = s837_fired_cy - m.win_ascent * fs + 1.0;
                }
                off
            } else {
                text_y_off
            };
            // S1135: the first line's atLeast leading, for the top border below.
            // This is the offset Oxi ACTUALLY draws the glyphs at, so the rule
            // keeps its 1.75pt clearance above the visible text even where the
            // leading itself is a hair off Word's (the 11pt arm holds 0.5 against
            // Word's 13 - 12.649 = 0.351). Reading it back beats recomputing the
            // natural height, which disagrees with the line-height code by 0.2pt
            // at 8pt and by 1.2pt on the specimen's own footer.
            if line_idx == 0
                && para.style.line_spacing_rule.as_deref() == Some("atLeast")
                && para.style.borders.as_ref().map_or(false, |b| b.top.is_some())
            {
                s1135_atleast_lead = text_y_off.max(0.0);
            }

            // S517 (2026-06-09): the body list-marker element was emitted before
            // this loop with the default text_y_off=0.0 (never set), so for wide
            // line boxes (e.g. b837 numbered list, line=18pt font=12pt → off=4.0)
            // the marker sat (line−fontcell) ABOVE the body baseline. Word renders
            // the marker ON the body baseline (COM-confirmed dy=+0.00 on b837 p2/p5
            // ①②③). The cell path already sets marker_el.text_y_off=cell_text_y_off
            // (mod.rs ~10086); the body path did not. Back-patch the first line's
            // text_y_off here. RENDER-ONLY (element.y unchanged) → Phase-1 safe.
            // Guarded: only an element with no paragraph_index (= a marker) is
            // touched, so a page break that moved the marker out of `elements`
            // makes this a no-op (current behavior) rather than corrupting a body
            // fragment.
            if line_idx == 0 {
                if let Some(midx) = s517_marker_el_idx.take() {
                    if let Some(mel) = elements.get_mut(midx) {
                        if mel.paragraph_index.is_none()
                            && matches!(mel.content, LayoutContent::Text { .. })
                        {
                            mel.text_y_off = text_y_off;
                            mel.content_fit_height = Some(break_threshold);
                        }
                    }
                }
            }

            // Compute max ascent across all fragments for baseline alignment.
            // All fragments in a line share the same baseline (matches Word output).
            let line_max_ascent: f32 = if line.fragments.is_empty() {
                // COM-confirmed: empty lines use paragraph font (East Asian in CJK docs)
                self.metrics_for_para_mark(&RunStyle::default(), &para.style)
                    .word_ascent_pt(para_font_size)
            } else {
                // S1045: an oversized edge whitespace-only fragment must not lower the
                // visible baseline either (para 18's 11pt trailing spaces pushed line 2's
                // baseline down; Word keeps the 9pt visible pitch, 10.320pt measured).
                let s1045_ma = self.s1045_height_drivers(&line.fragments, para_font_size);
                line.fragments
                    .iter()
                    .enumerate()
                    .filter(|(fi, f)| !LayoutEngine::s1045_skip(s1045_ma, *fi, f, para_font_size))
                    .map(|(_, f)| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let a = self
                            .metrics_for_text(&f.text, &f.style, &para.style)
                            .word_ascent_pt(fs);
                        // S655: a w:position-raised run extends the line ascent so the
                        // baseline drops to contain it (matches the line-height growth).
                        // S656: an emphasis mark above the char (dot/comma/circle) also
                        // extends the ascent.
                        let em_asc = self
                            .emphasis_above_pt(f, fs)
                            .filter(|v| *v > 0.0)
                            .unwrap_or(0.0);
                        if std::env::var("OXI_S655_DISABLE").is_err() {
                            a + f.style.position.map_or(0.0, |p| p.max(0.0)) + em_asc
                        } else {
                            a + em_asc
                        }
                    })
                    .fold(0.0_f32, f32::max)
            };

            // S1358 (2026-09-11, default ON, opt-out OXI_S1358_DISABLE):
            // the same maximum, but
            // in the ascent the RENDERER actually places a baseline at
            // (`baseline_ascent`, S1264/S1265). `line_max_ascent` above is
            // Word's ascent, which governs the line's HEIGHT; the two differ
            // on 6 of 23 faces, so a shift computed from the wrong one moves
            // the glyphs off the baseline it was meant to reach.
            let line_max_render_ascent: f32 = if line.fragments.is_empty() {
                0.0
            } else {
                let s1045_ma = self.s1045_height_drivers(&line.fragments, para_font_size);
                line.fragments
                    .iter()
                    .enumerate()
                    .filter(|(fi, f)| !LayoutEngine::s1045_skip(s1045_ma, *fi, f, para_font_size))
                    .map(|(_, f)| {
                        let base = f.style.font_size.unwrap_or(para_font_size);
                        // The size the fragment is actually SET at: a super- or
                        // subscript is drawn smaller, and anchoring the line to
                        // its unshrunk size would move every other fragment.
                        let fs = match f.style.vertical_align {
                            Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript) => {
                                LayoutEngine::vertical_align_font_size(base)
                            }
                            _ => base,
                        };
                        self.metrics_for_text(&f.text, &f.style, &para.style)
                            .baseline_ascent()
                            * fs
                    })
                    .fold(0.0_f32, f32::max)
            };
            // S1360 (2026-09-11, default ON, opt-out OXI_S1360_DISABLE):
            // a raised run grows the
            // line's ASCENT, and the glyphs have to come down with it.
            //
            // S655 already grows the line — `_pb_exactline_super` gives the
            // second line height 20.648 against the first's 14.648, the 6pt of
            // a `w:position` raise — but the line still hands the renderer the
            // same glyph top, so every baseline on it sits 5.94pt above Word's.
            // The growth belongs above the baseline: drop the line by it, and
            // the raised run's own offset then puts it back where it was.
            // No corpus document reaches this (the A/B changed 0 bytes), so the
            // claim rests on the probe alone, which is what it was measured on.
            let line_ascent_growth: f32 = if std::env::var("OXI_S1360_DISABLE").is_err() {
                let s1045_ma = self.s1045_height_drivers(&line.fragments, para_font_size);
                line.fragments
                    .iter()
                    .enumerate()
                    .filter(|(fi, f)| !LayoutEngine::s1045_skip(s1045_ma, *fi, f, para_font_size))
                    .map(|(_, f)| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        let raised = if std::env::var("OXI_S655_DISABLE").is_err() {
                            f.style.position.map_or(0.0, |p| p.max(0.0))
                        } else {
                            0.0
                        };
                        raised + self.emphasis_above_pt(f, fs).filter(|v| *v > 0.0).unwrap_or(0.0)
                    })
                    .fold(0.0_f32, f32::max)
            } else {
                0.0
            };

            // Mixed picture/text lines share the baseline required by both boxes.
            let body_picture_baseline = if !is_header_footer && !self.doc_body_has_real_cjk
                && (grid_pitch.is_none() || !para.style.snap_to_grid)
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && line.fragments.iter().any(|f| f.style.inline_object_image.is_some())
                && !line.fragments.iter().any(|f| f.style.inline_math.is_some() || f.style.hr_rule.is_some())
            {
                let text_ascent = line.fragments.iter()
                    .filter(|f| f.style.inline_object_image.is_none() && !f.text.trim().is_empty())
                    .map(|f| {
                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                        let leading = (m.ascent + m.descent + m.line_gap - m.win_ascent - m.win_descent).max(0.0);
                        (m.win_ascent + leading) * f.style.font_size.unwrap_or(para_font_size)
                    })
                    .fold(0.0_f32, f32::max);
                let image_ascent = line.fragments.iter().filter(|f| f.style.inline_object_image.is_some())
                    .filter_map(|f| f.style.inline_object_extent.map(|(_, h)|
                        (h + f.style.position.unwrap_or(0.0)).max(0.0)))
                    .fold(0.0_f32, f32::max);
                Some((text_ascent, image_ascent.max(text_ascent + line_ascent_growth)))
            } else { None };
            let line_ascent_growth = body_picture_baseline
                .map_or(line_ascent_growth, |(text, baseline)| baseline - text);

            // Two fragments SET at different sizes. Same-size lines place the
            // same either way, so this keeps the change off every line the
            // probes did not speak about.
            let line_has_mixed_sizes = {
                let set_size = |f: &LineFragment| {
                    let base = f.style.font_size.unwrap_or(para_font_size);
                    match f.style.vertical_align {
                        Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript) => {
                            LayoutEngine::vertical_align_font_size(base)
                        }
                        _ => base,
                    }
                };
                let mut seen: Option<f32> = None;
                line.fragments.iter().any(|f| {
                    if f.text.is_empty() {
                        return false;
                    }
                    let fs = set_size(f);
                    match seen {
                        None => {
                            seen = Some(fs);
                            false
                        }
                        Some(first) => (fs - first).abs() > 0.01,
                    }
                })
            };

            // R-10: track whether any fragment on this line came from a
            // revision-bearing source run; if so we emit one change-bar at
            // the line's left margin after the fragment loop finishes.
            // Paragraph-level revisions (ppr_change, paragraph_mark_revision)
            // count as revisions on every line of the paragraph — they're
            // detected once at paragraph entry rather than per-fragment.
            let mut line_has_revision =
                para.ppr_change.is_some() || para.paragraph_mark_revision.is_some();

            // S705 (2026-06-30): paragraph-level shd (w:pPr/w:shd) fills the
            // whole text column behind the paragraph. Emit a full-line-width
            // background rect (reusing CellShading, a generic fill) per line,
            // BEFORE the text, so it draws underneath. x = left text edge
            // (start_x + indent_left, NOT offset by first_line_indent), w = the
            // indent-adjusted text width, h = the line box. Run-level shd (S704)
            // is handled per-fragment.
            // DEFAULT ON (opt-out OXI_S705_DISABLE): the gray box wraps Oxi's
            // text with Word-matching padding + ~1.5pt horizontal bleed. Was held
            // opt-in (net −0.0014) because the gray EXPOSED a pre-existing ~1.4pt
            // text-Y offset (the empty-para over-count, S559) — now FIXED by S707,
            // so S705 is net +0.0043 on test_misc_props (0.9930→0.9973). Scope =
            // ONLY test_misc_props among word_png docs (the lone applied-pPr-shd
            // doc; ParagraphStyle.shading never reaches run fragments so S704's
            // run path is untouched). Render-only (CellShading post-layout) →
            // pagination byte-identical.
            if std::env::var("OXI_S705_DISABLE").is_err() {
                if let Some(ref shd) = para.style.shading {
                    if !shd.is_empty() && shd != "auto" {
                        // Word bleeds paragraph shading ~1.5pt past the text margins
                        // (measured test_misc_props: gray x extends ~1.4pt left /
                        // ~1.9pt right beyond the content edge).
                        let bleed = 1.5_f32;
                        let bg_x = (start_x + indent_left - bleed).max(0.0);
                        let bg_w =
                            (content_width - indent_left - indent_right + 2.0 * bleed).max(0.0);
                        if bg_w > 0.0 {
                            let color_hex = if shd.starts_with('#') {
                                shd.clone()
                            } else {
                                format!("#{}", shd)
                            };
                            elements.push(LayoutElement::new(
                                bg_x,
                                cursor.visual_y,
                                bg_w,
                                line_height,
                                LayoutContent::CellShading { color: color_hex },
                            ));
                        }
                    }
                }
            }

            // S706 (2026-06-30): accumulate run-border (w:bdr) spans on this line
            // into ONE box per contiguous bordered run (avoids internal verticals
            // between same-run fragments). (x_left, x_right, border). Opt-out
            // OXI_S706_DISABLE.
            let s706_on = std::env::var("OXI_S706_DISABLE").is_err();
            let mut run_bdr_acc: Option<(f32, f32, BorderDef)> = None;

            // S-TWOSEG: where this row crosses the float, and to what x.
            let two_seg_jump: Option<(usize, f32)> = match (row_two_segments, line.seg2_at) {
                (Some((_, _, seg2_x, _)), Some(at)) if at > 0 => {
                    Some((at, seg2_x))
                }
                _ => None,
            };
            let flow_elements_start = elements.len();
            let mixed_auto_baseline = std::env::var_os("OXI_MIXED_CJK_AUTO_BASELINE").is_some()
                && !is_header_footer
                && (grid_pitch.is_none() || !para.style.snap_to_grid)
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && line.fragments.iter().any(|f| {
                    !f.text.trim().is_empty()
                        && self.metrics_for_text(&f.text, &f.style, &para.style).is_cjk_83_64_font()
                })
                && line.fragments.iter().all(|f| {
                    f.style.vertical_align.is_none() && f.style.position.is_none()
                        && !f.style.vert_in_horz && !f.style.ruby_field
                        && f.style.emphasis_mark.is_none() && f.style.run_border.is_none()
                        && f.style.inline_object_image.is_none()
                });
            let body_baseline = if mixed_auto_baseline {
                Some(line.fragments.iter().filter(|f| !f.text.trim().is_empty()).map(|f| {
                    let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                    let size = f.style.font_size.unwrap_or(para_font_size);
                    let ascent = if m.is_cjk_83_64_font() {
                        let natural = (m.win_ascent + m.win_descent) * (83.0 / 64.0);
                        (natural + m.win_ascent - m.win_descent) * 0.5
                    } else {
                        m.baseline_ascent()
                    };
                    ascent * size
                }).fold(0.0_f32, f32::max))
            } else if std::env::var_os("OXI_BODY_CJK_BASELINE").is_some()
                && !is_header_footer
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && para.style.line_spacing.map_or(true, |v| (v - 1.0).abs() < 0.001)
                && line.fragments.iter().any(|f| !f.text.trim().is_empty())
                && line.fragments.iter().all(|f| {
                    f.style.vertical_align.is_none() && f.style.position.is_none()
                        && !f.style.vert_in_horz && !f.style.ruby_field
                        && f.style.emphasis_mark.is_none() && f.style.run_border.is_none()
                        && self.metrics_for_text(&f.text, &f.style, &para.style).is_cjk_83_64_font()
                })
            {
                let (ascent, descent) = line.fragments.iter()
                    .filter(|f| !f.text.trim().is_empty())
                    .fold((0.0_f32, 0.0_f32), |(a, d), f| {
                        let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                        let size = f.style.font_size.unwrap_or(para_font_size);
                        (a.max(m.win_ascent * size), d.max(m.win_descent * size))
                    });
                Some(line_height * 0.5 + (ascent - descent) * 0.5)
            } else {
                None
            };
            let header_baseline = if std::env::var_os("OXI_HEADER_CJK_BASELINE").is_some()
                && is_header_footer
                && !para.style.snap_to_grid
                && matches!(para.style.line_spacing_rule.as_deref(), None | Some("auto"))
                && para.style.line_spacing.map_or(true, |v| (v - 1.0).abs() < 0.001)
                && line.fragments.iter().any(|f| !f.text.trim().is_empty())
                && line.fragments.iter().all(|f| {
                    self.metrics_for_text(&f.text, &f.style, &para.style).is_cjk_83_64_font()
                })
            {
                Some(line.fragments.iter().filter(|f| !f.text.trim().is_empty()).map(|f| {
                    let m = &*self.metrics_for_text(&f.text, &f.style, &para.style);
                    let size = f.style.font_size.unwrap_or(para_font_size);
                    (m.win_ascent + (m.win_ascent + m.win_descent) * (83.0 / 64.0 - 1.0) * 0.5) * size
                }).fold(0.0_f32, f32::max))
            } else {
                None
            };
            let story_inline_baseline = if header_inline_geometry
                && line.fragments.iter().any(|f| f.style.inline_object_image.is_some())
            {
                let visible_descent = line.fragments.iter()
                    .filter(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                    .map(|f| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        self.metrics_for_text(&f.text, &f.style, &para.style).win_descent * fs
                    })
                    .fold(0.0_f32, f32::max);
                Some(if header_exact_inline {
                    0.8 * line_height
                } else {
                    line_height - visible_descent - story_image_leading[line_idx]
                })
            } else { None };
            for (frag_idx, frag) in line.fragments.iter().enumerate() {
                if let Some((at, seg2_x)) = two_seg_jump {
                    if frag_idx == at {
                        x = seg2_x;
                    }
                }
                let base_font_size = frag.style.font_size.unwrap_or(para_font_size);
                // Round 29: superscript/subscript rendering. Word default for
                // <w:vertAlign w:val="superscript"/> and "subscript":
                //   - font size = `vertical_align_font_size` (S1359)
                //   - vertical offset = +/- (base_size * 0.333) from baseline
                //     (negative = up for superscript, positive = down for subscript)
                //
                // The RISE is font-dependent and this constant is not: measured
                // at a 96pt base it is 0.334 of the size in Calibri, 0.354 in
                // Times New Roman, 0.349 in Arial and 0.266 in Georgia
                // (`_pb_superscript_size.py`). 0.333 suits the first three and
                // misses Georgia by 6.4pt at that size. Left alone here — it is
                // its own derivation, and the size was the tenth-sized error.
                let (resolved_font_size, vert_offset) = match frag.style.vertical_align {
                    Some(VerticalAlign::Superscript) => {
                        let fs = LayoutEngine::vertical_align_font_size(base_font_size);
                        // Raise the glyph: smaller font's baseline shifts up
                        // by ~1/3 of the original font size.
                        (fs, -(base_font_size * 0.333))
                    }
                    Some(VerticalAlign::Subscript) => {
                        let fs = LayoutEngine::vertical_align_font_size(base_font_size);
                        (fs, base_font_size * 0.083)
                    }
                    _ => (base_font_size, 0.0),
                };
                // S655 (2026-06-24): apply the w:position glyph shift. RunStyle
                // .position is pt, positive = UP; vert_offset is negative=up so
                // subtract. The line grows to match (two line-height fns +
                // line_max_ascent). Opt-out OXI_S655_DISABLE.
                let vert_offset = if std::env::var("OXI_S655_DISABLE").is_err() {
                    vert_offset - frag.style.position.unwrap_or(0.0)
                } else {
                    vert_offset
                };
                let resolved_bold = self.resolve_bold(&frag.style, &para.style);
                let adjusted_width = frag.width + frag_width_adjustments[frag_idx];

                // Per-fragment baseline alignment: shift fragments with smaller ascent
                // so all share the same baseline (y + frag_ascent = cursor_y + text_y_off + line_max_ascent)
                let frag_metrics = &*self.metrics_for_text(&frag.text, &frag.style, &para.style);
                let frag_ascent = frag_metrics.word_ascent_pt(resolved_font_size);
                // COM-confirmed (2026-04-14, gen2_001): Word does NOT apply
                // per-fragment baseline adjustment for body text. All fragments share
                // the same line-box TOP and the renderer adds each font's own ascent.
                // S691 (2026-06-29): the EXCEPTION — place a LARGE CJK glyph on a
                // snapToGrid=0 (header/footer/LM0) line at the 83/64 baseline (box_top +
                // word_ascent_pt), where Word draws CJK. The renderer draws baseline =
                // box_top + win_ascent×fs (the RAW win ascent), which for an MS Mincho-
                // class (win_sum≈1.0) CJK font is ~0.255×fs (= word_ascent − raw, the
                // 83/64 excess) TOO HIGH — visible on the 16pt albalunaTaidan header
                // title 対談版 (raw baseline 26.48 vs Word 30.56). The body GRID path
                // already adds this via S457/S166 centering; the snapToGrid=0 header does
                // not. Add baseline_adjust = word_ascent_pt − win_ascent×fs to the CJK
                // fragments. SCOPED: snap_to_grid==false, CJK-83/64, win_sum≈1.0
                // (EXCLUDES Yu Mincho, whose renderer ascent ≠ win_ascent so the layout
                // can't predict it — the Latin "6pt" stays put), fs≥14 (the 6pt body &
                // ≤12pt text untouched — they're the calibrated S455/S457 wall), and NOT
                // a no-type docGrid (S614 already places those via text_y_off).
                // Render-only (text_y_off). Opt-out OXI_S691_DISABLE.
                let win_sum = frag_metrics.win_ascent + frag_metrics.win_descent;
                let baseline_adjust = if std::env::var("OXI_S691_DISABLE").is_err()
                    && is_header_footer
                    && !para.style.snap_to_grid
                    && !page.doc_grid_no_type
                    && resolved_font_size >= 14.0
                    && frag_metrics.is_cjk_83_64_font()
                    && (win_sum - 1.0).abs() < 0.05
                {
                    frag_ascent - frag_metrics.win_ascent * resolved_font_size
                } else {
                    0.0
                };
                // S1358 (2026-09-11, default ON, opt-out OXI_S1358_DISABLE):
                // put every fragment of a line on ONE baseline.
                //
                // The note above says Word does not do this for body text, and
                // it is wrong. It was taken on gen2_001, where every fragment
                // of a line is the same size — and at one size, top-aligning
                // and baseline-aligning are the same picture. With two sizes
                // they are not: `_pb_exactline_latin.py EL_MIXED=20` puts a
                // 9pt run and a 20pt run on one line, and Word's PDF gives
                // BOTH spans origin 89.78. This engine draws them 10.36pt
                // apart, because the line hands the renderer one glyph top and
                // the renderer adds each font's own ascent to it.
                //
                // The shift is in `baseline_ascent`, the ascent the renderer
                // places a baseline at — not `word_ascent_pt`, which governs
                // the line's height and differs on 6 of 23 faces.
                //
                // SCOPE, and why it is this narrow. The unscoped form measured
                // net −0.1839 on the corpus (52 documents worse, 20 better),
                // and the losses are Japanese: nedocontract −0.0339,
                // kyotei36spec −0.0281, roudoujoken −0.0105. The CJK vertical
                // stack (S455/S457/S614/S629) is calibrated ON TOP of the
                // top-aligned convention, so moving a CJK fragment off it
                // breaks a compensation rather than fixing an error. What the
                // probes actually measured is a LATIN line carrying two
                // different font SIZES, so that is all this claims: a
                // same-size line is left exactly where it was.
                let baseline_adjust = baseline_adjust + line_ascent_growth
                    + if std::env::var("OXI_S1358_DISABLE").is_err()
                        && line_max_render_ascent > 0.0
                        // S1641 (2026-10-02): CJK lines share one baseline too
                        // (`_pb_mixedsize_gen.py`: Word's 12pt and 14pt glyph tops
                        // differ by their ascents, 7.21 vs 5.46; Oxi aligned tops).
                        && (!self.doc_body_has_real_cjk || std::env::var_os("OXI_S1641_DISABLE").is_none())
                        && line_has_mixed_sizes
                    {
                        (line_max_render_ascent
                            - frag_metrics.baseline_ascent() * resolved_font_size)
                            .max(0.0)
                    } else {
                        0.0
                    };
                let _ = line_max_ascent;

                // Session 75 Phase D (2026-05-17): y is LINE BOX TOP, renderer adds
                // text_y_off + baseline_adjust + vert_offset at draw time. See
                // memory/session71_y_convention_refactor_design.md.
                // S467: snap the emitted line top to the 0.75pt grid (Word's model).
                // S631 (2026-06-20, opt-in OXI_S631) ATTEMPTED + FALSIFIED: extend the S629
                // device-snap to the gen2-family NO-TYPE docGrid path (doc_grid_no_type) by
                // snapping the emit box-top to the 0.12pt grid anchored at page_top. gen2 body
                // is line=240 SINGLE-spacing in a no-type grid (NOT is_single_lm0 — needs no-grid;
                // NOT is_multiple_spacing). RESULT: gen2 family SSIM A/B net +0.0183 (13 docs
                // WORSE, 2 better) = REGRESSION. ROOT (measured g5_oxi_g vs g5_word per-line): the
                // gen2 residual is the run_base DRIFT (title aligned −0.02, bottom −0.7..−0.96 =
                // Oxi too LOW, ~0.09pt/line CJK line-height EXCESS = gen2 memory's "Error A +0.115
                // MS Mincho"), NOT quantization. The emit-snap refines quantization (±0.06pt) but
                // CANNOT fix a systematic 0.8pt cumulative excess, and the page_top-anchored
                // ABSOLUTE snap mis-phases most lines (Word's grid has a non-page_top phase). ⇒
                // the gen2 lever is the per-line CJK run_base PRECISION (line height ~0.09pt too
                // tall vs Word), the Phase-1-critical 83/64 wall — NOT the device-snap. Kept
                // opt-in/default-OFF so it never ships. See [[gen2_vertical_drift]].
                if std::env::var("OXI_DBG_EMITY").is_ok() {
                    let head: String = line.fragments.iter().flat_map(|f| f.text.chars()).take(12).collect();
                    eprintln!("[EMITY] line emit: cursor_y={:.2} visual_y={:.2} line_idx={} {:?}", cursor.cursor_y, cursor.visual_y, line_idx, head);
                }
                let emit_y = if s467_vsnap {
                    snap075(cursor.visual_y)
                } else if page.doc_grid_no_type
                    && !is_single_lm0
                    && std::env::var("OXI_S631").is_ok()
                {
                    let d = std::env::var("OXI_S631_DELTA")
                        .ok()
                        .and_then(|v| v.parse::<f32>().ok())
                        .unwrap_or(0.12);
                    page_top + ((cursor.visual_y - page_top) / d).round() * d
                } else {
                    cursor.visual_y
                };
                // S673L (2026-06-26): tab-leader rendering. A tab stop with a
                // `w:leader` (dot/middleDot/underscore/hyphen) fills the gap to the
                // stop with leader glyphs — Word draws them, Oxi left the gap BLANK
                // (the leader was parsed into TabStop.leader but never rendered;
                // 3a4f/model/tokyoshugyo form fields). Re-derive the leader from the
                // tab fragment's tab_position (no LineFragment field needed); fill
                // the gap with floor(gap/advance) leader chars. Opt-out OXI_S673L_DISABLE.
                let tab_leader_text: Option<String> =
                    if frag.text == TAB_STRING && std::env::var("OXI_S673L_DISABLE").is_err() {
                        frag.tab_position
                            .and_then(|pos| {
                                para.style
                                    .tab_stops
                                    .iter()
                                    .find(|ts| (ts.position - pos).abs() < 0.5)
                                    .and_then(|ts| ts.leader.as_deref())
                            })
                            .and_then(|ld| {
                                let ch = match ld {
                                    "dot" => '.',
                                    "middleDot" => '\u{00B7}',
                                    "underscore" => '_',
                                    "hyphen" => '-',
                                    _ => return None,
                                };
                                let adv = frag_metrics.char_width_em(ch) * resolved_font_size;
                                if adv <= 0.1 || adjusted_width < adv {
                                    return None;
                                }
                                let n = (adjusted_width / adv).floor() as usize;
                                if n == 0 {
                                    return None;
                                }
                                Some(std::iter::repeat(ch).take(n).collect::<String>())
                            })
                    } else {
                        None
                    };
                // S672: emit at the TRUE cumulative x (render_x) for pure-Latin
                // left-aligned lines; else the com_tw cumulative (x).
                let el_x = if s672_latinx { render_x } else { x };
                // S703 (2026-06-30): render a `combine` run (割注 / two-lines-in-one)
                // as warichu — optional brackets + the n chars split into 2 small
                // (~half-size) rows stacked within one line height. Emits 2-4 plain
                // Text elements in place of the single glyph (no renderer change).
                if frag.style.combine && std::env::var("OXI_S703_DISABLE").is_err() {
                    let fs = resolved_font_size;
                    let small = fs * 0.5;
                    let bsz = fs * 0.8;
                    let chars: Vec<char> = frag.text.chars().collect();
                    let half = (chars.len() + 1) / 2;
                    let top: String = chars[..half].iter().collect();
                    let bot: String = chars[half..].iter().collect();
                    let (lb, rb): (&str, &str) = match frag.style.combine_brackets.as_deref() {
                        Some("round") => ("（", "）"),
                        Some("square") => ("〔", "〕"),
                        Some("angle") => ("〈", "〉"),
                        Some("curly") => ("｛", "｝"),
                        _ => ("", ""),
                    };
                    let ff = self
                        .resolve_font_family_for_text(&frag.text, &frag.style, &para.style)
                        .map(|s| s.to_string());
                    let wcolor = self
                        .resolve_color(&frag.style, &para.style)
                        .map(|s| s.to_string());
                    let rbold = resolved_bold;
                    let mut wpush = |elements: &mut Vec<LayoutElement>,
                                     t: String,
                                     wx: f32,
                                     woff: f32,
                                     wsz: f32| {
                        if t.is_empty() {
                            return;
                        }
                        let mut e = LayoutElement::new(
                            wx,
                            emit_y,
                            wsz,
                            line_height,
                            LayoutContent::Text {
                                text: t,
                                font_size: wsz,
                                font_family: ff.clone(),
                                bold: rbold,
                                italic: false,
                                underline: false,
                                underline_style: None,
                                strikethrough: false,
                                double_strikethrough: false,
                                color: wcolor.clone(),
                                highlight: None,
                                field_type: None,
                                character_spacing: 0.0,
                                text_scale: 100.0,
                                is_vertical: false,
                                effects: TextEffects::default(),
                            },
                        );
                        e.text_y_off = woff;
                        if let Some(pi) = body_para_index {
                            e.paragraph_index = Some(pi);
                        }
                        elements.push(e);
                    };
                    let mut cx = el_x;
                    if !lb.is_empty() {
                        wpush(&mut elements, lb.to_string(), cx, text_y_off, bsz);
                        cx += bsz;
                    }
                    wpush(&mut elements, top, cx, text_y_off, small);
                    wpush(&mut elements, bot, cx, text_y_off + small, small);
                    cx += half as f32 * small;
                    if !rb.is_empty() {
                        wpush(&mut elements, rb.to_string(), cx, text_y_off, bsz);
                    }
                    x += adjusted_width + frag_spacing_after[frag_idx];
                    continue;
                }
                // S839: a U+FFFC object fragment draws its vector group's
                // primitives at the fragment position instead of text. The
                // object BOTTOM sits on the line's text baseline (the S837
                // rule: hmrc NI strip bottom 442.87 = label baseline; when
                // S773 fired, baseline = line_top + cy_max — else the normal
                // S614 convention baseline from the line's first text frag).
                if frag.text == "\u{FFFC}" && frag.style.inline_object_extent.is_some() {
                    // S852: an inline horizontal rule (o:hr) draws a full-width
                    // gray line centered in its own reserved line.
                    if let Some((thickness, color)) = frag.style.hr_rule.as_ref() {
                        let (ow, oh) = frag.style.inline_object_extent.unwrap_or((468.0, 13.8));
                        let ry = emit_y + oh * 0.5;
                        let mut e = LayoutElement::new(
                            el_x,
                            ry - thickness * 0.5,
                            ow.max(0.1),
                            thickness.max(0.5),
                            LayoutContent::TableBorder {
                                x1: el_x,
                                y1: ry,
                                x2: el_x + ow,
                                y2: ry,
                                color: Some(color.clone()),
                                width: *thickness,
                                style: None,
                            },
                        );
                        if let Some(pi) = body_para_index {
                            e.paragraph_index = Some(pi);
                        }
                        elements.push(e);
                        x += adjusted_width + frag_spacing_after[frag_idx];
                        continue;
                    }
                    // S1252: a structured inline oMath draws its expression
                    // tree at the fragment position, its own baseline sitting
                    // on the LINE's text baseline (the same anchor S851 uses
                    // for an object box's bottom). `emit_math_block` places the
                    // baseline at `cursor_y + max(ascent, 0.8*fs)`, so hand it
                    // the baseline minus that.
                    if let Some(mb) = frag.style.inline_math.as_ref() {
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let bbox = crate::layout::math::layout_math_block(mb, fs);
                        let baseline = if let Some(tf) = line
                            .fragments
                            .iter()
                            .find(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                        {
                            let tfs = tf.style.font_size.unwrap_or(para_font_size);
                            let m = &*self.metrics_for_text(&tf.text, &tf.style, &para.style);
                            emit_y + text_y_off - 1.0 + m.win_ascent * tfs
                        } else {
                            // The maths WRAPPED alone onto this line, so there is
                            // no text fragment to read the baseline off. It is
                            // still a text baseline, not the line bottom: Word
                            // draws reference__0042471c's `∑Dewan Pengawas…`
                            // with its baseline at 513.05 against the line's own
                            // trailing-space span at 513.41. Falling back to
                            // `emit_y + line_height` dropped it ~2.6pt.
                            let m = &*self.metrics_for_text(" ", &frag.style, &para.style);
                            emit_y + text_y_off - 1.0 + m.win_ascent * fs
                        };
                        let (mut math_elems, _) = crate::layout::math::emit_math_block(
                            mb,
                            el_x,
                            baseline - bbox.ascent.max(fs * 0.8),
                            fs,
                        );
                        if let Some(pi) = body_para_index {
                            for e in math_elems.iter_mut() {
                                e.paragraph_index = Some(pi);
                            }
                        }
                        elements.append(&mut math_elems);
                        x += adjusted_width + frag_spacing_after[frag_idx];
                        continue;
                    }
                    // S851: an inline w:object form-field image draws its bitmap
                    // at the fragment position (box BOTTOM on the text baseline,
                    // like the S839 vector groups) instead of a vector group.
                    if let Some(img) = frag.style.inline_object_image.as_ref() {
                        let (ow, oh) = frag
                            .style
                            .inline_object_extent
                            .unwrap_or((img.width, img.height));
                        let baseline = if let Some((_, offset)) = body_picture_baseline {
                            emit_y + offset - frag.style.position.unwrap_or(0.0)
                        } else if let Some(offset) = story_inline_baseline {
                            emit_y + offset
                        } else if let Some(tf) = line
                            .fragments
                            .iter()
                            .find(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                        {
                            let fs = tf.style.font_size.unwrap_or(para_font_size);
                            let m = &*self.metrics_for_text(&tf.text, &tf.style, &para.style);
                            emit_y + text_y_off - 1.0 + m.win_ascent * fs
                        } else if is_header_footer {
                            // Additional auto leading belongs below an image-only
                            // story line, while its natural baseline stays unchanged.
                            emit_y + line_height - story_image_leading[line_idx]
                        } else {
                            emit_y + line_height - if !self.doc_body_has_real_cjk
                                && std::env::var("OXI_BODY_IMAGE_EFFECT_EXTENT_DISABLE").is_err() {
                                img.effect_extent_b.max(0.0)
                            } else { 0.0 }
                        };
                        // S1040 (2026-07-29, opt-out OXI_S1040_DISABLE): an object
                        // TALLER than the text ascent must not render ABOVE its own
                        // line. S851 grows the line to obj + text descent, so the
                        // object occupies [line_top, baseline] - but the baseline
                        // here is the TEXT's (centred in the grown line), which for
                        // a 340pt screenshot on a 10.5pt line put the image at
                        // y=-65.6, off the top of the page (JA blind
                        // policies__03a9dca2, SSIM 0.928 -> 0.854 after S1034 first
                        // routed such images inline). Clamping the top to the line
                        // top is exactly the S851 height model, and is inert when
                        // the object is shorter than the text ascent (the EN
                        // reference__0042471c 11.25pt icon is unchanged).
                        let oy = if std::env::var("OXI_S1040_DISABLE").is_err() {
                            (baseline - oh).max(emit_y)
                        } else {
                            baseline - oh
                        };
                        // S1238 (2026-08-27): a data-less flow-reservation
                        // placeholder whose wps shape draws a visible frame
                        // (kyotei 労働保険番号 digit boxes: 0.5pt black ln +
                        // lt1 fill) renders that frame at the flowed position.
                        // Real pictures and frameless placeholders unchanged.
                        if img.data.is_empty() && std::env::var("OXI_S1238_DISABLE").is_err() {
                            if let Some((stroke, sw, fill)) = img.placeholder_outline.as_ref() {
                                let mut e = LayoutElement::new(
                                    el_x,
                                    oy,
                                    img.width,
                                    oh,
                                    LayoutContent::BoxRect {
                                        fill: fill.clone(),
                                        stroke_color: Some(stroke.clone()),
                                        stroke_width: *sw,
                                        corner_radius: 0.0,
                                    },
                                );
                                if let Some(pi) = body_para_index {
                                    e.paragraph_index = Some(pi);
                                }
                                elements.push(e);
                                x += adjusted_width + frag_spacing_after[frag_idx];
                                continue;
                            }
                        }
                        let mut image_y = oy;
                        let mut image_h = oh;
                        let mut crop = img.crop.as_ref().map(|c| (c.top, c.right, c.bottom, c.left));
                        if header_exact_inline && oh > 0.0 {
                            // Keep the image scale and baseline, clipping to the fixed line.
                            let original_y = baseline - oh;
                            image_y = original_y.max(emit_y);
                            let bottom = baseline.min(emit_y + line_height);
                            image_h = (bottom - image_y).max(0.0);
                            let (top, right, bottom_crop, left) = crop.unwrap_or((0.0, 0.0, 0.0, 0.0));
                            let visible_fraction = (100.0 - top - bottom_crop).max(0.0);
                            crop = Some((
                                top + (image_y - original_y) / oh * visible_fraction,
                                right,
                                bottom_crop + (baseline - bottom).max(0.0) / oh * visible_fraction,
                                left,
                            ));
                        }
                        let mut e = LayoutElement::new(
                            el_x,
                            image_y,
                            ow,
                            image_h,
                            LayoutContent::Image {
                                data: img.data.clone(),
                                content_type: img.content_type.clone(),
                                crop,
                            },
                        );
                        if let Some(pi) = body_para_index {
                            e.paragraph_index = Some(pi);
                        }
                        elements.push(e);
                        x += adjusted_width + frag_spacing_after[frag_idx];
                        continue;
                    }
                    if let Some(&tbi) = s839_tbs.get(s839_next) {
                        s839_next += 1;
                        let tb = &page.text_boxes[tbi];
                        let baseline = if line_idx == 0 && s837_fired_cy > 0.0 {
                            emit_y + s837_fired_cy
                        } else if let Some(tf) = line
                            .fragments
                            .iter()
                            .find(|f| f.text != "\u{FFFC}" && !f.text.trim().is_empty())
                        {
                            let fs = tf.style.font_size.unwrap_or(para_font_size);
                            let m = &*self.metrics_for_text(&tf.text, &tf.style, &para.style);
                            emit_y + text_y_off - 1.0 + m.win_ascent * fs
                        } else {
                            emit_y + line_height
                        };
                        let oy = baseline - tb.height;
                        for vs in &tb.vector_shapes {
                            let (vx, vy) = (el_x + vs.x, oy + vs.y);
                            if vs.is_line {
                                let mut e = LayoutElement::new(
                                    vx,
                                    vy,
                                    vs.w.max(0.1),
                                    vs.h.max(vs.stroke_width),
                                    LayoutContent::TableBorder {
                                        x1: vx,
                                        y1: vy,
                                        x2: vx + vs.w,
                                        y2: vy + vs.h,
                                        color: vs.stroke.clone(),
                                        width: vs.stroke_width,
                                        style: None,
                                    },
                                );
                                if let Some(pi) = body_para_index {
                                    e.paragraph_index = Some(pi);
                                }
                                elements.push(e);
                            } else if !vs.path.is_empty() {
                                // S1120: curve custGeom -> real outline
                                let mut e = LayoutElement::new(
                                    vx,
                                    vy,
                                    vs.w,
                                    vs.h,
                                    LayoutContent::VectorPath {
                                        segs: vs.path.clone(),
                                        fill: vs.fill.clone(),
                                        stroke_color: vs.stroke.clone(),
                                        stroke_width: vs.stroke_width,
                                    },
                                );
                                if let Some(pi) = body_para_index {
                                    e.paragraph_index = Some(pi);
                                }
                                elements.push(e);
                            } else {
                                let mut e = LayoutElement::new(
                                    vx,
                                    vy,
                                    vs.w,
                                    vs.h,
                                    LayoutContent::BoxRect {
                                        fill: vs.fill.clone(),
                                        stroke_color: vs.stroke.clone(),
                                        stroke_width: vs.stroke_width,
                                        corner_radius: 0.0,
                                    },
                                );
                                if let Some(pi) = body_para_index {
                                    e.paragraph_index = Some(pi);
                                }
                                elements.push(e);
                            }
                        }
                    }
                    x += adjusted_width + frag_spacing_after[frag_idx];
                    continue;
                }
                // S700: a vert_in_horz fragment renders as a stacked is_vertical
                // column (n chars down a 1-em cell).
                let is_vert_frag =
                    frag.style.vert_in_horz && std::env::var("OXI_S700_DISABLE").is_err();
                // S706: extend/start this line's run-border box for a w:bdr fragment.
                if s706_on {
                    if let Some(ref b) = frag.style.run_border {
                        let sp = b.space;
                        let l = el_x - sp;
                        let r = el_x + adjusted_width + sp;
                        run_bdr_acc = Some(match run_bdr_acc.take() {
                            Some((al, _, _)) => (al.min(l), r, b.clone()),
                            None => (l, r, b.clone()),
                        });
                    }
                }
                // S1650: a spread ruby base draws half a spread share right of
                // its field start (the advance is unchanged; see ruby_lead).
                let s1650_lead = if frag.style.ruby_spread { frag.style.ruby_lead } else { 0.0 };
                let mut el = LayoutElement::new(
                    el_x + s1650_lead,
                    emit_y,
                    adjusted_width,
                    line_height,
                    LayoutContent::Text {
                        // S700: Word renders eastAsianLayout w:vert by ROTATING the run
                        // 90° CCW, so the run's chars run BOTTOM→TOP (first char at the
                        // column bottom, last at the top — Word PDF draw-order confirmed).
                        // The is_vertical path stacks upright TOP→BOTTOM, so REVERSE the
                        // text to land each char in Word's vertical position (横 top, 縦
                        // bottom). The glyph ROTATION (vs upright) is a basics-only residual
                        // (no pixel gate; 0/corpus). The reservation (n·fs) is exact.
                        text: if is_vert_frag {
                            frag.text.chars().rev().collect::<String>()
                        } else {
                            tab_leader_text.clone().unwrap_or_else(|| frag.text.clone())
                        },
                        font_size: resolved_font_size,
                        font_family: self
                            .resolve_font_family_for_text(&frag.text, &frag.style, &para.style)
                            .map(|s| s.to_string()),
                        bold: resolved_bold,
                        italic: self.resolve_italic(&frag.style, &para.style),
                        underline: frag.style.underline,
                        underline_style: frag.style.underline_style.clone(),
                        strikethrough: frag.style.strikethrough,
                        double_strikethrough: frag.style.double_strikethrough,
                        color: self
                            .resolve_color(&frag.style, &para.style)
                            .map(|s| s.to_string()),
                        // S704 (2026-06-30): render run-level shading (w:shd) as a
                        // background by reusing the highlight rect (the renderer draws a
                        // hex highlight). Effective colour computed in the parser. A real
                        // highlight wins; else the shading fills the background.
                        highlight: frag
                            .style
                            .highlight
                            .clone()
                            .or_else(|| frag.style.shading.clone()),
                        field_type: frag.field_type,
                        character_spacing: if is_vert_frag {
                            // The column's per-char DOWN advance is exactly fs (the
                            // renderer adds character_spacing to it) → keep it 0.
                            0.0
                        } else if frag.style.fit_text.is_some() || frag.style.ruby_spread {
                            frag.style.character_spacing.unwrap_or(0.0) + justify_char_spacing
                        } else {
                            snap_character_spacing(frag.style.character_spacing.unwrap_or(0.0))
                                + justify_char_spacing
                        },
                        text_scale: frag.style.text_scale.unwrap_or(100.0),
                        is_vertical: is_vert_frag,
                        // S702: faithful Word character effects.
                        effects: TextEffects {
                            shadow: frag.style.shadow,
                            emboss: frag.style.emboss,
                            imprint: frag.style.imprint,
                            outline: frag.style.outline,
                            no_fill: frag.style.no_fill,
                        },
                    },
                );
                // Session 72 Phase A: populate text_y_off (y still includes it).
                // S700: place the vert column from the line-box top (text_y_off −
                // (n-1)/2·fs) so it stacks down through the (centred) char-row — the
                // line grew by (n-1)·fs and the centring pushed the char-row to the
                // middle, so the column's middle char aligns with the body chars.
                if !is_vert_frag {
                    el.content_fit_height = Some(break_threshold);
                }
                el.text_y_off = if is_vert_frag {
                    let n = frag.text.chars().count().max(1);
                    text_y_off - (n.saturating_sub(1)) as f32 * resolved_font_size * 0.5
                } else {
                    text_y_off + baseline_adjust + vert_offset
                };
                if !is_vert_frag {
                    el.baseline_offset = story_inline_baseline.or(body_baseline).or(header_baseline).map(|b| b + vert_offset);
                }
                // S1527: every text fragment keeps its source run (cell text
                // included) so a footnote reference can be found on its line.
                el.run_index = Some(frag.run_index);
                if let Some(pi) = body_para_index {
                    el.paragraph_index = Some(pi);
                    el.char_offset = Some(frag.char_offset);
                    if frag.field_type.is_none() {
                        if let Some(Some(map)) = case_sources.get(frag.run_index) {
                            if let Some((offset, len, text)) = map.fragment(frag.char_offset, &frag.text) {
                                el.char_offset = Some(offset);
                                el.source_char_len = Some(len);
                                el.source_text = Some(text);
                            }
                        }
                    }
                }
                // Round 7: capture base element x/y BEFORE push (move).
                // Used below to position the ruby annotation above the base.
                let base_el_x = el.x - s1650_lead;
                let base_el_y = el.y;
                let base_el_tyo = el.text_y_off;
                elements.push(el);

                // Round 7: emit ruby annotation glyph element above the
                // base text. Only fires on the FIRST fragment of a Run
                // (char_offset == 0) to avoid duplicate emission when the
                // base text spans multiple fragments. Base width is computed
                // from the full run text (not just this fragment) so the
                // annotation centers over all base chars per V2 §18.5.
                // Currently implements `Center` only — other rubyAlign
                // modes default to center; per-mode positioning is a
                // Round 7.5 follow-up.
                // S1628 (2026-10-01, default ON, opt-out OXI_S1628_DISABLE): a ruby
                // whose base spans several runs (educational__09422f63's title:
                // «注意» and «報» carry different w:spacing under one fitText) is
                // ONE annotation, emitted at the group's first run and spread over
                // the WHOLE group's advance including its character spacing. The
                // per-run emission drew «ちゅういほう» twice, and a base width
                // without the fitText spacing squeezed «かふん» over the first
                // glyph where Word spreads it 47.5pt apart (distributeSpace over
                // 142.5pt: (142.5 - 45) / 3 = 32.5 + 15).
                let s1628 = std::env::var_os("OXI_S1628_DISABLE").is_none();
                let s1628_same = |a: &crate::ir::Ruby, b: &crate::ir::Ruby| a.text == b.text && a.base == b.base;
                let s1628_continuation = s1628
                    && frag.run_index > 0
                    && match (para.runs.get(frag.run_index - 1).and_then(|r| r.ruby.as_ref()),
                              para.runs.get(frag.run_index).and_then(|r| r.ruby.as_ref())) {
                        (Some(prev), Some(cur)) => s1628_same(prev, cur),
                        _ => false,
                    };
                if frag.char_offset == 0 && !s1628_continuation {
                    if let Some(run) = para.runs.get(frag.run_index) {
                        if let Some(ref ruby_ir) = run.ruby {
                            let base_pt = frag.style.font_size.unwrap_or(para_font_size);
                            let hps_pt = ruby_ir
                                .hps_halfpt
                                .map(|h| h as f32 / 2.0)
                                .unwrap_or(base_pt / 2.0);
                            // S1642 (2026-10-02): the drawn raise defaults like the laid-out one
                            // (base - 1, 9.0 at 10.5) -- the fixed 9.0 put a 12pt line's
                            // annotation 1.84pt low (`_pb_bodyruby_gen.py` b24_h6_r0:
                            // Word か 8.52 / Oxi 10.36 while the base matched).
                            let hps_raise_pt = ruby_ir
                                .hps_raise_halfpt
                                .map(|h| h as f32 / 2.0)
                                .unwrap_or_else(|| if std::env::var_os("OXI_S1642_DISABLE").is_none() {
                                    ruby::default_hps_raise_pt(base_pt, hps_pt)
                                } else {
                                    ruby::DEFAULT_HPS_RAISE_PT
                                });
                            let ruby_text = ruby_ir.text.as_str();
                            let mut ruby_run_style = frag.style.clone();
                            ruby_run_style.font_size = Some(hps_pt);
                            let ruby_metrics =
                                &*self.metrics_for_text(ruby_text, &ruby_run_style, &para.style);
                            // Round 7.6: precise per-char widths via GDI metrics.
                            // Replaces the previous `chars × font_size` CJK
                            // monospace approximation; matches non-CJK ruby
                            // and proportional fonts correctly.
                            let ruby_char_count = ruby_text.chars().count();
                            let ruby_w: f32 = ruby_text
                                .chars()
                                .map(|c| {
                                    self.registry.char_width_pt_with_fallback(
                                        c,
                                        hps_pt,
                                        ruby_metrics,
                                    )
                                })
                                .sum();
                            let base_metrics =
                                &*self.metrics_for_text(run.text.as_str(), &frag.style, &para.style);
                            let base_w: f32 = run
                                .text
                                .chars()
                                .map(|c| {
                                    self.registry.char_width_pt_with_fallback(
                                        c,
                                        base_pt,
                                        base_metrics,
                                    )
                                })
                                .sum();
                            // S1314: a spread base is as wide as its ruby field.
                            let base_w = if s1628 {
                                // S1628: every run of the group, each with its own
                                // character spacing (fitText or explicit).
                                let mut w = 0.0_f32;
                                let mut gi = frag.run_index;
                                while let Some(gr) = para.runs.get(gi) {
                                    match gr.ruby.as_ref() {
                                        Some(r) if gi == frag.run_index || s1628_same(r, ruby_ir) => {}
                                        _ => break,
                                    }
                                    let gm = &*self.metrics_for_text(gr.text.as_str(), &gr.style, &para.style);
                                    let gfs = gr.style.font_size.unwrap_or(base_pt);
                                    let n = gr.text.chars().count() as f32;
                                    w += gr.text.chars()
                                        .map(|c| self.registry.char_width_pt_with_fallback(c, gfs, gm))
                                        .sum::<f32>()
                                        + gr.style.character_spacing.unwrap_or(0.0) * n;
                                    gi += 1;
                                }
                                w
                            } else if frag.style.ruby_spread {
                                base_w
                                    + frag.style.character_spacing.unwrap_or(0.0)
                                        * run.text.chars().count() as f32
                            } else {
                                base_w
                            };
                            // Round 7.5: rubyAlign positioning per ECMA-376 §17.3.3.26.
                            // ruby_position returns (x_offset_from_base, per_char_spacing).
                            let (ruby_x_offset, ruby_char_spacing) =
                                ruby::ruby_position(base_w, ruby_w, ruby_char_count, ruby_ir.align);
                            let ruby_x = base_el_x + ruby_x_offset;
                            let ruby_ascent = ruby_metrics.word_ascent_pt(hps_pt);
                            let frag_metrics =
                                &*self.metrics_for_text(&frag.text, &frag.style, &para.style);
                            let base_ascent = frag_metrics.word_ascent_pt(base_pt);
                            // S1632 (2026-10-02, default ON, opt-out OXI_S1632_DISABLE):
                            // the annotation's BASELINE sits exactly hpsRaise above the
                            // base glyph's rendered baseline. `_pb_bodyruby_gen.py`: Word's
                            // ruby-to-base baseline gap is 9.96 / 15.0 / 18.0 / 28.94 /
                            // 39.98 for raise 10 / 15 / 18 / 29 / 40; the word-ascent
                            // estimate below put Oxi's 0.7-1.0pt higher at 10.5-14pt and
                            // 1.2pt lower at 30pt. The renderer draws a glyph top at
                            // y + text_y_off - 1 and its baseline win_ascent below that.
                            let ruby_y = if std::env::var_os("OXI_S1632_DISABLE").is_none() {
                                base_el_y + base_el_tyo + frag_metrics.win_ascent * base_pt
                                    - hps_raise_pt - ruby_metrics.win_ascent * hps_pt
                            } else {
                                base_el_y + base_ascent - hps_raise_pt - ruby_ascent
                            };
                            let ruby_color = self
                                .resolve_color(&ruby_run_style, &para.style)
                                .map(|s| s.to_string());
                            let ruby_font_family = self
                                .resolve_font_family_for_text(
                                    ruby_text,
                                    &ruby_run_style,
                                    &para.style,
                                )
                                .map(|s| s.to_string());
                            let mut ruby_el = LayoutElement::new(
                                ruby_x,
                                ruby_y,
                                ruby_w,
                                hps_pt * 1.2,
                                LayoutContent::Text {
                                    text: ruby_text.to_string(),
                                    font_size: hps_pt,
                                    font_family: ruby_font_family,
                                    bold: false,
                                    italic: false,
                                    underline: false,
                                    underline_style: None,
                                    strikethrough: false,
                                    double_strikethrough: false,
                                    color: ruby_color,
                                    highlight: None,
                                    field_type: None,
                                    character_spacing: ruby_char_spacing,
                                    text_scale: 100.0,
                                    is_vertical: false,
                                    effects: TextEffects::default(),
                                },
                            );
                            if let Some(pi) = body_para_index {
                                ruby_el.paragraph_index = Some(pi);
                                ruby_el.run_index = Some(frag.run_index);
                                ruby_el.char_offset = Some(0);
                            }
                            ruby_el.flow_line_offset = ruby_y - base_el_y;
                            // Round 7.5: when char spacing > 0 (distribute*),
                            // the rendered width grows by (chars × extra). The
                            // element's `width` field tracks the visual extent
                            // for hit testing — bump it so the renderer reserves
                            // the full distributed range.
                            if ruby_char_spacing > 0.0 {
                                ruby_el.width = ruby_w + ruby_char_count as f32 * ruby_char_spacing;
                            }
                            elements.push(ruby_el);
                        }
                    }
                }

                // S656 (2026-06-24): emit emphasis marks (圏点, w:em) — a small
                // mark above each base char (below for underDot). Word does not
                // render the run via w:em alone, it ADDS the marks; Oxi parsed
                // emphasis_mark but never drew it. Per-FRAGMENT (not gated on
                // char_offset) so a wrapped run marks each line's chars. The line
                // already grew via emphasis_above_pt. 0/corpus → coverage.
                if std::env::var("OXI_S656_DISABLE").is_err() {
                    if let Some(em) = frag.style.emphasis_mark.as_deref() {
                        if em != "none" && !frag.text.is_empty() {
                            let mark = match em {
                                "circle" => '○',
                                "comma" => '﹅',
                                "underDot" => '●',
                                _ => '●', // "dot"
                            };
                            let base_pt = frag.style.font_size.unwrap_or(para_font_size);
                            let mark_pt = base_pt * 0.5;
                            let base_metrics =
                                &*self.metrics_for_text(&frag.text, &frag.style, &para.style);
                            let mark_str = mark.to_string();
                            let mark_metrics =
                                &*self.metrics_for_text(&mark_str, &frag.style, &para.style);
                            let mark_w = self.registry.char_width_pt_with_fallback(
                                mark,
                                mark_pt,
                                mark_metrics,
                            );
                            let mark_family = self
                                .resolve_font_family_for_text(&mark_str, &frag.style, &para.style)
                                .map(|s| s.to_string());
                            let mark_color = self
                                .resolve_color(&frag.style, &para.style)
                                .map(|s| s.to_string());
                            let below = em == "underDot";
                            // The char top sits at base_el_y + text_y_off + the
                            // emphasis growth (the grown ascent pushes the baseline
                            // down by that much); place the mark in the gap just
                            // above it (or below the char for underDot).
                            let em_above = base_pt * 0.33;
                            let char_top = base_el_y + text_y_off + em_above;
                            let mut cx = base_el_x;
                            for ch in frag.text.chars() {
                                let cw = self.registry.char_width_pt_with_fallback(
                                    ch,
                                    base_pt,
                                    base_metrics,
                                );
                                let mx = cx + (cw - mark_w) / 2.0;
                                let my = if below {
                                    char_top + base_pt * 0.92
                                } else {
                                    char_top - mark_pt
                                };
                                let mut mel = LayoutElement::new(
                                    mx,
                                    my,
                                    mark_w,
                                    mark_pt * 1.2,
                                    LayoutContent::Text {
                                        text: mark_str.clone(),
                                        font_size: mark_pt,
                                        font_family: mark_family.clone(),
                                        bold: false,
                                        italic: false,
                                        underline: false,
                                        underline_style: None,
                                        strikethrough: false,
                                        double_strikethrough: false,
                                        color: mark_color.clone(),
                                        highlight: None,
                                        field_type: None,
                                        character_spacing: 0.0,
                                        text_scale: 100.0,
                                        is_vertical: false,
                                        effects: TextEffects::default(),
                                    },
                                );
                                if let Some(pi) = body_para_index {
                                    mel.paragraph_index = Some(pi);
                                    mel.run_index = Some(frag.run_index);
                                }
                                elements.push(mel);
                                cx += cw;
                            }
                        }
                    }
                }

                // R-10: detect revision-bearing fragment by looking up the
                // source run's `tracked_change` OR `rpr_change`. The pre-pass
                // mutated `style.underline`/`color`/`strikethrough` but
                // preserved both revision pointers, so this lookup is the
                // canonical signal. Word fires a change bar for any revision
                // — insert/delete/move via `tracked_change`, or formatting
                // change via `rpr_change` (R-12).
                if !line_has_revision {
                    if let Some(run) = para.runs.get(frag.run_index) {
                        if run.tracked_change.is_some() || run.rpr_change.is_some() {
                            line_has_revision = true;
                        }
                    }
                }
                x += adjusted_width + frag_spacing_after[frag_idx];
                // S672: advance the true-cumulative render track by the fragment's
                // un-rounded em width (= the DWrite render advance) so the NEXT
                // fragment emits where this word actually ends on screen.
                if s672_latinx {
                    let true_w: f32 = frag
                        .text
                        .chars()
                        .map(|c| frag_metrics.char_width_em(c) * resolved_font_size)
                        .sum();
                    render_x += true_w + frag_spacing_after[frag_idx];
                }
            }

            // S706: flush this line's run-border box (one stroked rect around the
            // bordered run's text, padded by w:space). Drawn after the text so the
            // thin border sits on top at the edges.
            if let Some((l, r, b)) = run_bdr_acc.take() {
                let sp = b.space;
                // w:bdr color="auto" (parsed to None) resolves to black.
                let col = Some(b.color.clone().map_or_else(
                    || "#000000".to_string(),
                    |c| {
                        if c.starts_with('#') {
                            c
                        } else {
                            format!("#{}", c)
                        }
                    },
                ));
                let w = (r - l).max(0.0);
                if w > 0.0 {
                    elements.push(LayoutElement::new(
                        l,
                        cursor.visual_y - sp,
                        w,
                        line_height + 2.0 * sp,
                        LayoutContent::BoxRect {
                            fill: None,
                            stroke_color: col,
                            stroke_width: b.width,
                            corner_radius: 0.0,
                        },
                    ));
                }
            }

            // R-10: emit one margin change-bar per revision-bearing line.
            // Word's default change bar sits ~12pt outside the body's left
            // edge, ~1.5pt thick, dark grey. Independent of author color so
            // multi-author paragraphs still get a single unambiguous bar.
            if line_has_revision {
                let bar_x = (start_x - 12.0).max(0.0);
                let bar_y = cursor.visual_y;
                let bar_h = line_height;
                let bar_w: f32 = 1.5;
                elements.push(LayoutElement::new(
                    bar_x,
                    bar_y,
                    bar_w,
                    bar_h,
                    LayoutContent::BoxRect {
                        fill: Some("#424242".to_string()),
                        stroke_color: None,
                        stroke_width: 0.0,
                        corner_radius: 0.0,
                    },
                ));
            }

            // Empty-line placeholder (Round 10): for empty paragraphs (no
            // fragments on this line), emit a zero-width Text element so
            // the structure dump / hit-testing tools can still see the
            // paragraph_index. This matters especially for §17.2.2 implicit
            // empty body paragraphs (header_page_number_01, footer_complex_01).
            if line.fragments.is_empty() {
                if let Some(pi) = body_para_index {
                    // Session 75 Phase D: y is LINE BOX TOP; renderer adds text_y_off.
                    let mut el = LayoutElement::new(
                        line_x,
                        cursor.cursor_y,
                        0.0,
                        line_height,
                        LayoutContent::Text {
                            text: String::new(),
                            font_size: para_font_size,
                            font_family: None,
                            bold: false,
                            italic: false,
                            underline: false,
                            underline_style: None,
                            strikethrough: false,
                            double_strikethrough: false,
                            color: None,
                            highlight: None,
                            field_type: None,
                            character_spacing: 0.0,
                            text_scale: 100.0,
                            is_vertical: false,
                            effects: TextEffects::default(),
                        },
                    );
                    // Session 72 Phase A: populate text_y_off (y still includes it).
                    el.text_y_off = text_y_off;
                    el.paragraph_index = Some(pi);
                    // An empty boundary row has no painted glyphs, but its
                    // real source control is still an editable caret position.
                    // Preserve only the character present at that source offset;
                    // ordinary empty paragraphs retain an empty source string.
                    if let Some((run,offset,_))=&line.break_source {
                        if let Some(control)=para.runs.get(*run)
                            .and_then(|source|source.text.chars().nth(*offset)) {
                            if matches!(control,'\x0C'|'\x0B') {
                                el.run_index=Some(*run);
                                el.char_offset=Some(*offset);
                                let boundaries_before=para.runs[..*run].iter()
                                    .flat_map(|source|source.text.chars())
                                    .chain(para.runs[*run].text.chars().take(*offset))
                                    .filter(|ch|matches!(ch,'\x0C'|'\x0B'|'\n'|'\r')).count();
                                // Only an earlier source object makes this the
                                // paragraph's content start. Leading controls
                                // retain editable caret indices without becoming
                                // an extra source-text paragraph opener.
                                let earlier=page.floating_images.iter()
                                    .filter(|object|object.anchor_block_index==pi)
                                    .filter_map(|object|object.position.as_ref())
                                    .chain(page.text_boxes.iter()
                                        .filter(|object|object.anchor_block_index==pi)
                                        .filter_map(|object|object.position.as_ref()))
                                    .chain(para.shapes.iter().filter_map(|object|object.position.as_ref()))
                                    .any(|position|position.flow_boundary_offset<=boundaries_before);
                                let prior_text=para.runs[..*run].iter()
                                    .flat_map(|source|source.text.chars())
                                    .chain(para.runs[*run].text.chars().take(*offset))
                                    .any(|ch|!ch.is_whitespace());
                                // Preserve the source attachment without inventing
                                // a glyph-source fragment for an unpainted control.
                                el.source_boundary_attachment = earlier && !prior_text;
                            }
                        }
                    }
                    elements.push(el);
                }
            }

            // Multiple spacing: cumulative ceil for non-last lines when all lines
            // have the same height. When heights vary (mixed fonts), use per-line height.
            // COM-confirmed (2026-04-07): variable-height paragraphs (e.g., mixed CJK+Latin
            // first line, pure CJK subsequent lines) use per-line height, not cumulative.
            // COM-confirmed (2026-04-08): SINGLE spacing also cumulative round in LM=0
            // but only when raw_per_line > rounded_per_line (preserves page breaks).
            let is_last = line_idx == lines.len() - 1;
            // Round 30: linesAndChars Single spacing uses pitch-based cumulative round.
            // Session 161 (2026-05-21): LM2 cell-advance path is skipped for
            // paragraphs whose `<w:snapToGrid w:val="0"/>` opts them out of grid
            // snap. d1e8 has docGrid linesAndChars + snapToGrid=0 paragraphs
            // whose Word line height is NATURAL (12.75pt sz=10.5, 14.25pt sz=11),
            // not grid pitch (14.6pt). Without this gate, LM2 cell-advance
            // forces ~15pt per paragraph → cumulative drift accumulates.
            // Full-baseline verification (env-gated trial, 2026-05-21):
            //   - Phase 1: 53/55 UNCHANGED
            //   - Phase 2: mean IoU 0.9191 → 0.9236 (+0.0045 strict increase)
            //   - 5 improvements (d1e8 +0.259 dominant), 1 regression
            //     (1636d28e -0.0233 on wi=89-93, separate issue)
            // S236+S237 (2026-05-23): removed OXI_LEGACY_LM2_IGNORE_STG
            // legacy env-var fallback during hardening pass.
            let flow_cursor_before = cursor.cursor_y;
            let is_lm2_single = lm2_grid_cells.is_some()
                && page.grid_char_pitch.is_some()
                && grid_pitch.map_or(false, |p| p > 0.0)
                && para.style.snap_to_grid
                && match (
                    para.style.line_spacing_rule.as_deref(),
                    para.style.line_spacing,
                ) {
                    (Some("exact"), _) | (Some("atLeast"), _) => false,
                    (_, Some(f)) if (f - 1.0).abs() > 0.01 => false,
                    _ => true,
                };
            if is_lm2_single {
                // R56b (2026-05-17): hybrid cursor advance. Cell-aligned cursor
                // entries use the absolute formula (drift-free for uniform-paragraph
                // docs like b837 / 1ec1). Mid-cell entries (after irregular line-
                // height paragraph that took non-LM2 path) use cursor-relative
                // with PROPER ceiling — fixes d1e8 wi=31->wi=32 3pt advance bug.
                //
                // R56 original attempt failed because `(X/10+1)*10` formula adds
                // 10tw extra when X is on 10tw boundary; cursor-relative with
                // cur on 10tw boundary frequently triggered this. Proper ceiling
                // `((X+9)/10)*10` correctly returns X when X is on boundary.
                //
                // See [[session69-lm2-unified-refactor-groundwork]].
                let pitch_tw_i = (grid_pitch.unwrap() * 20.0).round() as i32;
                // S1226 (2026-08-26): the docGrid line lattice anchors at the
                // CURRENT page's content top (`page_top`, which follows the
                // ACTIVE section's top margin through S863 geom switches and
                // page pushes), not the document-first section's
                // page.margin.top. kyotei36spec p4 pixel truth: the 裏面 page
                // begins a top=567tw section after a top=964tw section; Word
                // lays the title at 28.35 and the next line at 39.85 = the
                // 567-anchored lattice. The old fixed anchor made
                // offset=(cur-964tw).max(0)=0 → cell-aligned target =
                // 964+1×pitch = 59.7 → every later-section page drifted
                // +19.85pt, ate 3 lines of the 2-col band, and manufactured a
                // 5th page. Byte-identical when page_top == margin.top (the
                // whole single-section corpus). Opt-out OXI_S1226_DISABLE.
                let margin_tw = if std::env::var("OXI_S1226_DISABLE").is_err()
                    && (page_top - page.margin.top).abs() > 0.01
                {
                    (page_top * 20.0).round() as i32
                } else {
                    (page.margin.top * 20.0).round() as i32
                };
                let cells = (line_height * 20.0 / pitch_tw_i as f32).round().max(1.0) as i32;
                let cur_tw = (cursor.cursor_y * 20.0).round() as i32;
                let offset = (cur_tw - margin_tw).max(0);
                let cell_remainder = offset % pitch_tw_i;
                // S494 (2026-06-04, SHIP default-ON; opt-out OXI_S494_DISABLE):
                // Word advances docGrid lines by the EXACT fractional grid pitch
                // (357tw=17.85pt) and snaps the ABSOLUTE position to the 96dpi
                // device pixel (15tw=0.75pt) — NOT the integer-rounded 18.0pt the
                // mid-cell branch produces. The mid-cell cursor-relative branch
                // rounds `cur+357` to 10tw, which (when cur is 10tw-aligned) ALWAYS
                // yields +360 (18.0pt), over-allocating 0.15pt/line. After a non-LM2
                // paragraph (e.g. `line=360 lineRule=exact`) pushes the cursor off
                // the docGrid phase, EVERY following grid line takes mid-cell → the
                // 0.15pt/line drift accumulates (1ec1: +2.5pt by the floating table,
                // screenshot + COM + minimal-repro confirmed: Word empty-para pitch =
                // grid pitch device-snapped to 17.25/18.00, Oxi = flat 18.00).
                // Fix: carry an un-rounded ideal accumulator (cursor.lm2_ideal_y) and
                // device-snap, matching Word for BOTH cell-aligned and mid-cell runs.
                // SCOPE: EMPTY paragraph lines only (line.fragments empty). The minimal
                // repro confirmed empty-para height = grid pitch device-snapped across
                // grid320/357/360/400; CONTENT-para grid advance is a separate, unconfirmed
                // spec (d1e8 grid292 content paras regress under the device-snap — Word
                // does NOT advance them by the full snapped pitch the same way). Scoping
                // to empty lines keeps 1ec1's empty-chain fix without disturbing the
                // tuned content-para mid-cell path (S324-S327).
                // GATE (235-doc RGB-SSIM refresh, empty-only): Phase-1 54/55 UNCHANGED;
                // mean 0.9420->0.9422 (+0.0001); bottom-10 sum +0.0339 STRICTLY UP
                // (1ec1, the WORST doc, 0.6511->0.6861 +0.0349); only d1e8 -0.0068
                // (non-bottom; the device-snap is more correct than the old under-
                // allocating mid-cell flat advance, but d1e8 had a pre-existing
                // downstream too-low drift that the under-allocation compensated —
                // separate follow-up). 1636d28 -0.0010 (noise).
                if std::env::var("OXI_S494_DISABLE").is_err() && line.fragments.is_empty() {
                    let pitch = pitch_tw_i as f32;
                    let cur_f = cur_tw as f32;
                    // Continue the run if the ideal is still in sync with the cursor
                    // (within half a pitch); otherwise (first line / after a non-LM2
                    // paragraph moved the cursor / new page) resync to the cursor.
                    // S1400 (2026-09-14, opt-out OXI_S1400_DISABLE): the ideal
                    // stream re-syncs to the cursor whenever the cursor moved by
                    // more than device-snap noise -- a paragraph's space-before
                    // is a REAL move, not rounding. The half-pitch window let a
                    // 9pt before (180tw < 193tw) be discarded: the line landed
                    // back on the pre-spacing lattice, 2 cells - 9 = 29.7 where
                    // Word gives before + 2 cells = 47.7 (`_pb_emptybefore_grid_gen.py`,
                    // 9 arms: linesAndChars with/without charSpace, lang en-US /
                    // ja-JP, empty / text lines, all 48/48/47.25 by Info6;
                    // the type=lines arm never took this path and matched).
                    // policies__07543a6b9776a1cf: three such empties on the
                    // cover, -28.5pt, two paragraphs pulled up from page 2.
                    let s1400_tol = if std::env::var("OXI_S1400_DISABLE").is_err() { 12.0 } else { pitch * 0.5 };
                    let ideal0 = if cursor.lm2_ideal_y > 0.0
                        && (cursor.lm2_ideal_y - cur_f).abs() < s1400_tol
                    {
                        cursor.lm2_ideal_y
                    } else {
                        cur_f
                    };
                    let ideal1 = ideal0 + cells as f32 * pitch;
                    let target = (ideal1 / 15.0).round() * 15.0; // 0.75pt = 96dpi px
                    cursor.set(target / 20.0);
                    cursor.lm2_ideal_y = ideal1;
                    cumul_line_idx += cells as usize;
                } else {
                    // R56c: distinguish "slightly past cell start" (uniform LM2
                    // after 10tw ceiling, e.g. pitch=292 cur=margin+1*pitch+5tw)
                    // from "truly mid-cell" (after irregular non-LM2 paragraph).
                    // Threshold: <10tw or >pitch-10tw means within ceiling-noise
                    // of cell boundary → use absolute (drift-free).
                    let cell_aligned = cell_remainder < 10 || cell_remainder > pitch_tw_i - 10;
                    let target_tw = if cell_aligned {
                        // Cell-aligned (or near-aligned): pre-R56 absolute formula
                        //
                        // S324 (2026-05-26) — env-gated fix for cell-near-boundary
                        // case. R56c treats both <10tw (after cell start) and
                        // >pitch-10tw (before next cell) as "aligned", but the
                        // formula `(k + cells) * pitch` assumes cur is at cell-k
                        // start. For cur "near NEXT cell start" (cell_rem >
                        // pitch-10), this UNDERSHOOTS by 1 cell — cursor stays
                        // at cur+~0pt instead of advancing one line. d1e8ac8
                        // para 11→12 trace: cur_tw=5780 cell_rem=284 → target
                        // 5790 (+0.5pt) instead of 6080 (+15pt). With S324_FIX,
                        // k is incremented when cell_rem > pitch-10. R56's
                        // "cur slightly past cell start" case (cell_rem<10)
                        // is preserved.
                        // S327 DEFAULT-ON. Env-var preserved as OPT-OUT.
                        let s324_fix = std::env::var("OXI_S324_FIX_CELL_BOUNDARY")
                            .map(|v| v != "0" && v != "false")
                            .unwrap_or(true);
                        let k_raw = offset / pitch_tw_i;
                        let k = if s324_fix && cell_remainder > pitch_tw_i - 10 {
                            k_raw + 1
                        } else {
                            k_raw
                        };
                        let target_n = k + cells;
                        // S325 (2026-05-26): when S324 is on, ALSO change the
                        // cell_aligned padding from always-+10tw to proper
                        // ceiling (matches mid-cell branch). The always-+10tw
                        // padding was the source of the +0.5pt/line cascade
                        // accumulating across paragraphs after S324 corrected
                        // the missing-cell advance.
                        // S327 DEFAULT-ON. Env-var preserved as OPT-OUT.
                        let s325_fix = std::env::var("OXI_S325_PROPER_CEIL")
                            .map(|v| v != "0" && v != "false")
                            .unwrap_or(true);
                        if s325_fix {
                            let raw = margin_tw + target_n * pitch_tw_i;
                            if raw % 10 == 0 {
                                raw
                            } else {
                                (raw / 10 + 1) * 10
                            }
                        } else {
                            ((margin_tw + target_n * pitch_tw_i) / 10 + 1) * 10
                        }
                    } else {
                        // Mid-cell from irregular predecessor: cursor-relative.
                        //
                        // S326 (2026-05-26) env-gated: change CEIL → ROUND-half-up.
                        // Even with proper ceiling, raw=5772 → 5780 (+8tw)
                        // accumulates ~0.5pt/paragraph over many lines.
                        // CLAUDE.md S301 first attempt showed CEIL→ROUND was
                        // catastrophic STANDALONE, but the cascade-broken state
                        // after S324+S325 may make ROUND viable.
                        // S327 DEFAULT-ON. Env-var preserved as OPT-OUT.
                        let s326_round = std::env::var("OXI_S326_MID_CELL_ROUND")
                            .map(|v| v != "0" && v != "false")
                            .unwrap_or(true);
                        // S1319 (2026-09-05, default ON, opt-out OXI_S1319_DISABLE): a
                        // mid-cell line advances the EXACT ideal stream
                        // (`lm2_ideal_y`, the S494 absolute-snap design) and rounds
                        // that, instead of rounding cur + pitch from the already
                        // rounded cursor. correspondence__03ca64d7 (linesAndChars
                        // 298 = 14.9pt): Word's PDF steps every body line by
                        // exactly 14.9 (41 lines, 110.5 -> 721.5); Oxi's cursor
                        // fell off the cell grid after an empty paragraph's
                        // 0.75pt snap and then rounded 298 -> 300tw on EVERY line
                        // (15.0), +4.1pt by line 41 and the last paragraph on a
                        // 2nd page. The same drift showed on the 298-pitch probe
                        // (`_pb_oikomi_default_gen.py`: 15.0 per line) while the
                        // cell-aligned branch and the 324-pitch probe (16.5/16.0
                        // alternating = no compounding) were fine.
                        // Enabled together with natural cached-break validation.
                        let s1319 = std::env::var("OXI_S1319_DISABLE").is_err();
                        let raw = if s1319 {
                            let pitch_f = pitch_tw_i as f32;
                            let cur_f = cur_tw as f32;
                            // S1400: same re-sync window as the empty branch above.
                            let s1400_tol = if std::env::var("OXI_S1400_DISABLE").is_err() { 12.0 } else { pitch_f * 0.5 };
                            let ideal0 = if cursor.lm2_ideal_y > 0.0
                                && (cursor.lm2_ideal_y - cur_f).abs() < s1400_tol
                            {
                                cursor.lm2_ideal_y
                            } else {
                                cur_f
                            };
                            let ideal1 = ideal0 + cells as f32 * pitch_f;
                            cursor.lm2_ideal_y = ideal1;
                            ideal1.round() as i32
                        } else {
                            cur_tw + cells * pitch_tw_i
                        };
                        if s326_round {
                            // ROUND-half-up to 10tw: (x + 5) / 10 * 10
                            ((raw + 5) / 10) * 10
                        } else if raw % 10 == 0 {
                            raw
                        } else {
                            (raw / 10 + 1) * 10
                        }
                    };
                    cursor.set(target_tw as f32 / 20.0);
                    // S1619 (2026-10-01, default ON, opt-out OXI_S1619_DISABLE): the
                    // mid-cell line lands ON the exact ideal stream, not on it rounded
                    // to 10tw. policies__1f014c0f p22 (linesAndChars 416): Word's line
                    // tops minus the exact stream 241.6 + 20.8k are 4.46-4.54 on all
                    // nine body lines (constant within the PDF's 0.12 quantum), where
                    // Oxi stepped 20.9/20.5/21.0; and the double-spaced heading after
                    // line (7) starts at the exact 428.8 in Word, while Oxi carried the
                    // rounded 429.0 into it -- +0.2 for the rest of the page, which
                    // pushed «※帳票は» (bottom slack 18.00 against the 18.15 centred
                    // box) to p23. `_pb_gridbottom_lac_gen.py` (the slice from the
                    // page's heading on, so no rounded predecessor) flips exactly
                    // where Oxi's centred box does.
                    if !cell_aligned
                        && std::env::var("OXI_S1319_DISABLE").is_err()
                        && std::env::var_os("OXI_S1619_DISABLE").is_none()
                        && cursor.lm2_ideal_y > 0.0
                    {
                        cursor.set(cursor.lm2_ideal_y / 20.0);
                    }
                    if cell_aligned && std::env::var("OXI_S1319_DISABLE").is_err() {
                        // S1319: the cell-aligned target is exact from the margin;
                        // hand the unrounded stream to the next mid-cell line.
                        let k_raw = offset / pitch_tw_i;
                        let k = if cell_remainder > pitch_tw_i - 10 { k_raw + 1 } else { k_raw };
                        cursor.lm2_ideal_y = (margin_tw + (k + cells) * pitch_tw_i) as f32;
                    }
                    cumul_line_idx += cells as usize;
                } // end S494 else (legacy cell-aligned / mid-cell branch)
            } else {
                // For single LM=0, gate by direction: only when raw advances MORE than rounded.
                let single_lm0_safe = if is_single_lm0 && raw_spaced_tw > 0.0 {
                    let raw_pt = raw_spaced_tw / 20.0;
                    let rounded_pt = (raw_pt * 2.0).round() / 2.0;
                    raw_pt > rounded_pt
                } else {
                    false
                };
                // S805 (2026-07-12, default ON, opt-out OXI_S805_DISABLE): a LATIN document's LM0
                // single-spacing line advances by the EXACT raw height (Arial
                // hhea 1.1499×11 = 12.649), not the cumulative CEIL-10tw
                // (12.649→13.0 = +0.35pt/para → +9pt/page; fn_probe render-truth:
                // Word body para pitch 24.6 vs Oxi 25.0). The S510 CEIL was a
                // COMPENSATION for the CJK 83/64 raw deficit (−0.021pt/line) —
                // a Latin raw has no deficit, so CEIL purely over-advances. Same
                // exact-accumulate model as S671 (no-type-grid Latin). JP docs
                // (doc_body_has_real_cjk) byte-identical by construction.
                let s805_latin_lm0 = grid_pitch.is_none()
                    && is_single_lm0
                    && raw_spaced_tw > 0.0
                    && !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S805_DISABLE").is_err();
                // LM=0 cumulative ROUND includes LAST line; LM≥1 cumulative CEIL excludes last.
                // Reuse the first-line basis only for equal-height lines. A small
                // real font-height difference must survive the cursor advance.
                let use_cumulative = (is_multiple_spacing || single_lm0_safe || s805_latin_lm0) && raw_spaced_tw > 0.0
                && line_heights.iter().all(|&h| (h - line_heights[0]).abs() < 0.001)
                && (grid_pitch.is_none() || !is_last)
                // S671 (2026-06-25): a NO-TYPE docGrid NON-CJK paragraph advances by
                // its EXACT per-line height (line_height_for_line_inner = the hhea
                // natural × factor = Word's no-type-grid Latin line height), NOT the
                // LM0 cumulative-round-to-0.5pt model (the CJK single-spacing device-
                // snap, S629). The cumulative round mis-tracks Word's per-line multiple-
                // spacing heights (±0.25pt/line); Word accumulates the exact height and
                // device-snaps only at render. Falling to the `else { cursor.advance(
                // line_height) }` branch uses the exact S671 value. test_line_heights
                // mean |O−W| 0.266→0.032.
                && !s671_fine;
                if use_cumulative {
                    let j = cumul_line_idx;
                    let (cn, cc) = if grid_pitch.is_none() && is_single_lm0 {
                        // COM-confirmed (2026-04-16, 0e7a): LM0 single spacing should use
                        // position-based cumul, not index × raw. When paragraphs have
                        // different raws (9pt body in 10.5pt doc), per-paragraph raw
                        // applied over a shared index underestimates positions.
                        // Use mult_cumul_raw (shared position accumulator) with CEIL.
                        let old_pos = mult_cumul_raw.as_deref().copied().unwrap_or(0.0);
                        let new_pos = old_pos + raw_spaced_tw;
                        // S510 (2026-06-08) FALSIFIED+REVERTED: tried FINER quantization (round
                        // to 1tw vs the CEIL-10tw here) to match Word's fine line pitch. It DID
                        // match Word's pitch SET (683f: {13.5,13.6,13.7} vs the 10tw model's
                        // {13.5,14.0}) BUT made the CUMULATIVE WORSE (last line Oxi−Word −1.50 vs
                        // −1.20). ROOT CAUSE REVEALED: Oxi's RAW CJK line height (83/64) is
                        // 13.605 vs Word's 13.626 (−0.021pt/line); the CEIL-10tw was COMPENSATING
                        // that deficit by bumping to 14.0. So the real vertical lever is the CJK
                        // 83/64 raw line-height PRECISION (~0.02pt/line too small vs Word,
                        // accumulating ~1.2pt over a dense page), NOT the cumulative quantization.
                        // That is per-font line-height precision (Phase-1-critical, deeply-tuned
                        // 83/64 model). Kept the CEIL-10tw (it compensates reasonably). See
                        // session509_renderer_justify_snap / session511.
                        let cn = (new_pos / 10.0).ceil() as i32 * 10;
                        let cc = (old_pos / 10.0).ceil() as i32 * 10;
                        (cn, cc)
                    } else if is_multiple_spacing {
                        // COM-confirmed (2026-04-14, mixed font repro): Multiple spacing
                        // uses cumulative raw position model with ROUND. Each paragraph
                        // adds its raw_tw to a shared running total.
                        // S467 NOTE: the cumulative LINE position is rounded to 10tw (0.5pt)
                        // here, while spacing is advanced exact (line 4006). Word instead
                        // snaps the COMBINED (line+spacing) cumulative position to 15tw
                        // (0.75pt = 96-DPI pixel). A granularity-only experiment (round to
                        // 15tw here) was FALSIFIED on the Cambria repro (mean|drift| 0.188->
                        // 0.229, worse) — matching Word needs the spacing folded into the
                        // snapped cumulative, not just a coarser line-round. Pure-body Cambria
                        // is already within +/-0.5 of Word (this 10tw model is fine); the gen2
                        // drift is the title pBdr (-0.75, see mod.rs:5539) + list-style-boundary
                        // rounding-phase mismatches that only the combined-snap model resolves.
                        let old_pos = mult_cumul_raw.as_deref().copied().unwrap_or(0.0);
                        let new_pos = old_pos + raw_spaced_tw;
                        let cn = (new_pos / 10.0).round() as i32 * 10;
                        let cc = (old_pos / 10.0).round() as i32 * 10;
                        (cn, cc)
                    } else {
                        let cn = (((j + 1) as f32 * raw_spaced_tw / 10.0).round() * 10.0) as i32;
                        let cc = ((j as f32 * raw_spaced_tw / 10.0).round() * 10.0) as i32;
                        (cn, cc)
                    };
                    // S626 (2026-06-19) ATTEMPTED + FALSIFIED + REVERTED: LM0-single CJK
                    // glyph baseline FINER-quant (advance visual_y by the CEIL cumul at 1tw
                    // vs the 10tw cursor advance, KEEPING CEIL so it differs from S510's
                    // CEIL→ROUND). Hypothesis: the 0.5pt snap causes ±0.12pt per-line jitter
                    // (+0.058 SSIM cost, fitz-position metric) that finer-quant would remove.
                    // RESULT: DWrite-gate REGRESSED (0e7af mean 0.9082→0.8701 −0.0381, 683f
                    // −0.0097) with the regression GROWING down the document = a cumulative
                    // DRIFT. ROOT (confirms S510): the 0.5pt CEIL-10tw snap lands CLOSER to
                    // Word's baselines than the fine CEIL-1tw, because Word's per-line CJK
                    // height is CONTENT-dependent (11.64/11.66/11.76/11.88 — taller glyphs on
                    // some lines), NOT a uniform raw — so finer-quant of ANY uniform value
                    // (83/64 OR the measured per-size table) drifts from Word's varying
                    // per-line; and CEIL-10tw quantizes away the table's ±0.015pt correction.
                    // ⇒ the Y-jitter is NOT a quantization fix — it needs Word's EXACT
                    // per-LINE content-dependent baseline algorithm (the deep precision wall).
                    // S628 (2026-06-19) ATTEMPTED + FALSIFIED + REVERTED: Y-jitter fix via
                    // the MEASURED per-size table as the visual_y base (advance visual_y by the
                    // exact measured MS Mincho/Gothic line height, cursor_y by 83/64 CEIL-10tw).
                    // RESULT: REGRESSED BADLY (0e7af −0.1208, 683f −0.0090). ROOT: the constant
                    // per-SIZE table (11.640@9pt) IGNORES Word's CONTENT-dependent per-LINE
                    // variation (11.64/11.66/11.76/11.88 — taller glyphs on some lines), which
                    // the 83/64 CEIL model actually CAPTURES via content-aware raw_spaced_tw. So
                    // the table-base drifts WORSE than the jitter on content-varying lines. ⇒
                    // (5th falsified jitter fix) Word's per-line baseline is CONTENT-dependent
                    // (tallest-glyph height per line, device-snapped); NO constant/table/finer-
                    // quant model captures it. The Y-jitter needs Word's EXACT per-content-line
                    // baseline — the deepest precision wall. See memory.
                    // S629 (2026-06-19, opt-in OXI_S629): Y-jitter fix via DEVICE-SNAP to
                    // δ≈0.12pt of the 83/64 cumulative. NEW measurement: Word's per-line
                    // 9pt baselines are 11.64 (dominant) + periodic 11.76, differing by
                    // EXACTLY 0.12pt (diffs from 11.64 = multiples of 0.12) = the device-snap
                    // of a CONSTANT 83/64 height (11.672) to a 0.12pt grid, NOT content-
                    // dependent (correcting the S628 conclusion). S626 failed using δ=1tw
                    // (0.05, too fine) + CEIL; the correct model is ROUND-to-0.12 of the
                    // 83/64 cumulative. visual_y (glyph track) snaps the cumulative to δ;
                    // cursor_y keeps the 0.5pt-CEIL advance (pagination byte-identical →
                    // Phase-1 safe). δ via OXI_S629_DELTA (default 0.12).
                    let s629 = grid_pitch.is_none()
                        && is_single_lm0
                        && std::env::var("OXI_S629_DISABLE").is_err();
                    // S1172 (2026-08-19, default ON, opt-out OXI_S1172_DISABLE):
                    // the S805 exact-accumulate applies to Latin no-grid MULTIPLE
                    // spacing too. The 10tw cumulative ROUND is only mean-
                    // preserving when the raw straddles a 10tw boundary; a raw
                    // NEAR a multiple rounds the same way every line and the
                    // error turns systematic. DERIVED (`_pb_garapitch_gen.py`,
                    // identical-line arms, Word truth read as MEANS because the
                    // PDF quantizes positions to 600dpi ±0.12):
                    //   Word = hhea x factor EXACTLY, all six arms
                    //     (Garamond 11.246~11.25, x1.15 12.943~12.9375;
                    //      TNR/Arial 11.486~11.499, x1.15 13.234~13.2239)
                    //   Oxi TNR x1.15: alternates 13.5/13.0, mean 13.214 - bounded
                    //   Oxi Garamond x1.15: uniform 13.000 vs 12.9375
                    //     (raw 258.75tw always rounds UP to 260) = +0.0625/line,
                    //     +1.44pt over a 23-line page = forms__002fbe2c's p2
                    //     drift, one lost line, one slipped paragraph.
                    // The bbox-top pitches read off the real document first
                    // (11.28/11.59/10.85) were a TRAP: the bbox top is the
                    // tallest glyph, so the pitch varies with line CONTENT --
                    // identical-line probes are the only clean readout.
                    let s1172_latin_mult = grid_pitch.is_none()
                        && is_multiple_spacing
                        && raw_spaced_tw > 0.0
                        && !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S1172_DISABLE").is_err();
                    // S1483 (2026-09-19, default ON, opt-out OXI_S1483_DISABLE): a CJK
                    // document with no line grid advances by the EXACT raw line
                    // height too. The S510 CEIL-10tw cumulative fired only on lines
                    // whose natural height sits ABOVE the half point (14pt MS Mincho
                    // 18.156 -> 18.5) and fell through to the exact advance below it
                    // (11pt 14.266), so a single heading gained +0.35 the render never
                    // drew (ikujidetail p12 [PARA] end_cur 189.3 vs [EMITY] visual
                    // 188.91; Word PDF: empty line + heading advance 32.40 = 14.27 +
                    // 18.13, the natural height). Formerly the sleeping opt-in
                    // OXI_CJK_EXACT_BODY_ADVANCE.
                    if s805_latin_lm0 || s1172_latin_mult || ((grid_pitch.is_none() || page.doc_grid_no_type) && self.doc_body_has_real_cjk && std::env::var("OXI_S1483_DISABLE").is_err()) {
                        // S805: exact accumulate — no 10tw quantization (Word
                        // device-snaps only at render; per-line snap accumulates
                        // error, the S674 lesson).
                        cursor.advance(raw_spaced_tw / 20.0);
                    } else if s467_vsnap && is_multiple_spacing {
                        // visual_y advances by the EXACT raw line height; cursor_y by the
                        // current rounded amount (page-break unchanged). Emit snaps visual_y.
                        cursor.advance_split((cn - cc) as f32 / 20.0, raw_spaced_tw / 20.0);
                    } else if s629 {
                        let d = std::env::var("OXI_S629_DELTA")
                            .ok()
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(0.15);
                        // S629 session-2: RELATIVE (content-top anchored) snap of the 83/64
                        // cumulative to the δ=0.12 (600-DPI px) grid. baseline_grid_fit.py CONFIRMED
                        // this fits Word PERFECTLY for 9/10/11/12/14pt (9pt = 0.0mpt residual). The
                        // 10.5pt HALF-POINT is the lone OUTLIER (83/64×10.5=13.617 = mid-cell on the
                        // 0.12 grid, 124mpt residual at every δ) — snapping it REGRESSED 683f. So
                        // EXCLUDE 10.5pt lines (use the normal advance) → Pareto-safe (9pt docs gain,
                        // 10.5pt docs untouched).
                        // Derive the line's font size from its 83/64 line HEIGHT (robust — no font
                        // resolution needed, works for 683f whose font_family is empty in the dump):
                        // fs = (raw_spaced_tw/20) / (83/64). Half-point sizes (10.5/11.5/…) are the
                        // grid OUTLIERS (83/64×10.5=13.617 is mid-cell on the 0.12 grid).
                        let fs_from_lh = (raw_spaced_tw / 20.0) / (83.0 / 64.0);
                        let is_half_point = (fs_from_lh.fract() - 0.5).abs() < 0.08;
                        if is_half_point && std::env::var("OXI_S632").is_err() {
                            cursor.advance((cn - cc) as f32 / 20.0);
                        } else {
                            let op = mult_cumul_raw.as_deref().copied().unwrap_or(0.0);
                            let np = op + raw_spaced_tw;
                            // S632 (2026-06-20, opt-in OXI_S632) ATTEMPTED + FALSIFIED: 10.5pt (and
                            // other half-points) ARE on the 0.12 grid (683f GAPS fit δ=0.12 to 2mpt,
                            // alternating 13.68/13.56 = device-snap of 83/64×10.5=13.617 which is
                            // MID-cell, 0.057 sub-cell remainder). A standalone-gap analysis matched
                            // Word's per-line pattern 92% at start-phase 0.05. BUT applying φ=raw mod δ
                            // (0.057) in the real code REGRESSED 683f −0.0766 (ssim_ab), and φ=0 (the
                            // S629 no-exclusion variant) also regressed. ROOT: the absolute phase is
                            // mult_cumul_raw(body) mod δ — i.e. the PREAMBLE's accumulated height — and
                            // because 10.5pt is mid-cell, a sub-0.06pt preamble error FLIPS the snap
                            // anti-phase. Oxi's preamble (title/headings at non-10.5pt) accumulates a
                            // height that differs from Word's by the per-line run_base error → wrong
                            // phase. ⇒ the 10.5pt phase is NOT independently fixable; it ⊂ the SAME
                            // per-line CJK run_base PRECISION wall as gen2's drift (S631). 9pt is
                            // robust (near cell-edge); 10.5pt mid-cell is phase-fragile. The S629
                            // exclusion stands. Sweepable OXI_S632_PHI. See [[gen2_vertical_drift]].
                            let phi = if is_half_point {
                                std::env::var("OXI_S632_PHI")
                                    .ok()
                                    .and_then(|v| v.parse::<f32>().ok())
                                    .unwrap_or_else(|| (raw_spaced_tw / 20.0).rem_euclid(d))
                            } else {
                                0.0
                            };
                            let snap = |x: f32| (((x + phi) / d).round() * d) - phi;
                            let old_v = snap(op / 20.0);
                            let new_v = snap(np / 20.0);
                            cursor.advance_split((cn - cc) as f32 / 20.0, new_v - old_v);
                        }
                    } else {
                        cursor.advance((cn - cc) as f32 / 20.0);
                    }
                    // Update cumulative raw position for Multiple spacing AND LM0 single.
                    if is_multiple_spacing || (grid_pitch.is_none() && is_single_lm0) {
                        if let Some(ref mut cr) = mult_cumul_raw {
                            **cr += raw_spaced_tw;
                        }
                    }
                } else {
                    cursor.advance(line_height);
                }
                // Round 7: ruby paragraph-tail expansion (V7 measurement) —
                // when the current paragraph contains any ruby annotation,
                // add the expansion AFTER the last line's cursor advance.
                // Greenfield-dormant on baseline (ruby_para_expansion_pt = 0
                // when no run has ruby). Estimate path is wired in §18.4
                // estimate_para_height; this is the matching render-side
                // wiring so cursor positions match the estimate.
                // S1312: the expansion belongs to every LINE that carries a ruby
                // run, not to the paragraph tail (see the cell site for the
                // derivation). Opt-out restores the last-line-only advance.
                let s1312_on = std::env::var("OXI_S1312_DISABLE").is_err();
                let s1312_line_ruby = s1312_on
                    && lines.get(line_idx).map_or(false, |l| {
                        l.fragments.iter().any(|f| para.runs.get(f.run_index).map_or(false, |r| r.ruby.is_some()))
                    });
                // S1641: the expansion this LINE needs (its tallest run may already
                // reach above the ruby's room)
                let ruby_para_expansion_pt = if std::env::var_os("OXI_S1641_DISABLE").is_none()
                    && ruby_para_expansion_pt > 0.0
                {
                    lines.get(line_idx).map_or(ruby_para_expansion_pt,
                        |l| self.s1641_line_ruby_expansion(l, para, para_font_size))
                } else {
                    ruby_para_expansion_pt
                };
                if ruby_para_expansion_pt > 0.0
                    && (s1312_line_ruby || (!s1312_on && line_idx + 1 == lines.len()))
                {
                    // S654 (coverage, 2026-06-24): in a TYPED docGrid the furigana
                    // makes the ruby line taller, and Word snaps the ruby-AUGMENTED
                    // line UP to whole grid cells — perturb_probe.py: Word 2 cells
                    // (36pt) vs Oxi 1 cell + raw overhang (23.76), Δ −12.24. The base
                    // line already advanced by its snapped height (no-grid path is
                    // already correct, +0.48), so add only the extra cell(s) that the
                    // natural base + furigana overhang needs. 0/corpus docs use ruby
                    // (greenfield-dormant) → byte-identical gate. Opt-out
                    // OXI_S654_DISABLE.
                    let typed_grid = para.style.snap_to_grid
                        && grid_pitch.map_or(false, |p| p > 0.0)
                        && !page.doc_grid_no_type
                        && std::env::var("OXI_S654_DISABLE").is_err();
                    if typed_grid {
                        let pitch = grid_pitch.unwrap();
                        let nat = natural_line_heights
                            .get(line_idx)
                            .copied()
                            .unwrap_or(line_height);
                        let base_snapped =
                            line_heights.get(line_idx).copied().unwrap_or(line_height);
                        // S752 (2026-07-05): a small tolerance on the cell ceil —
                        // the marginal config (probervsweep24 cfg16: hps=5pt
                        // raise=9pt at base 12: augmented 18.02) gets 1 cell in
                        // Word, not 2; the true 2-cell configs clear the boundary
                        // by >= 1.27pt in the sweep, so 0.5 is mid-window.
                        // S1638 (2026-10-02, default ON, opt-out OXI_S1638_DISABLE): on a
                        // typed line grid Word centres the UNION of the ruby's box and the
                        // base's box in the n-row box (n = ceil(union / pitch)); the base
                        // therefore sits (R - R0)/2 + exp/2 lower than the base-only
                        // centring (R = rows with ruby, R0 = rows of the base alone).
                        // `_pb_bodyruby_gen.py` BR_GRID=1, 16 arms, all within 0.1pt:
                        // b32 h16 r30 Word 32.42 / Oxi 28.72 (= 09422f63's 「2月の保健目標」
                        // 3.5pt high), b24 h12 r22 30.02 / 18.57, and b21 h10 r20 takes TWO
                        // rows in Word (the old -0.5 fudge gave one).
                        let s1638 = std::env::var_os("OXI_S1638_DISABLE").is_none();
                        let fudge = if s1638 { 0.01 } else { 0.5 };
                        // S1641: rows from the exact union when it is known
                        let union = if s1638 && std::env::var_os("OXI_S1641_DISABLE").is_none() {
                            lines.get(line_idx).map_or(nat + ruby_para_expansion_pt,
                                |l| self.s1641_line_ruby_union(l, para, para_font_size).1.max(nat))
                        } else {
                            nat + ruby_para_expansion_pt
                        };
                        let augmented_snapped = ((union - fudge) / pitch).ceil() * pitch;
                        if std::env::var_os("OXI_DBG_RUBYGRID").is_some() {
                            eprintln!("[RUBYGRID] nat={:.2} exp={:.2} pitch={:.2} base_rows={:.2} aug_rows={:.2}",
                                nat, ruby_para_expansion_pt, pitch, base_snapped, augmented_snapped);
                        }
                        cursor.advance((augmented_snapped - base_snapped).max(0.0));
                        if s1638 {
                            let shift = (augmented_snapped - base_snapped).max(0.0) * 0.5
                                + ruby_para_expansion_pt * 0.5;
                            for el in elements[flow_elements_start..].iter_mut() {
                                el.y += shift;
                            }
                        }
                    } else {
                        cursor.advance(ruby_para_expansion_pt);
                        // S1631 (2026-10-02, default ON, opt-out OXI_S1631_DISABLE):
                        // off a typed grid the ruby line's extra height sits ABOVE
                        // the base (the annotation's room), not below it.
                        // `_pb_bodyruby_gen.py` (09422f63 host, snapToGrid 0, MS
                        // Mincho): Word's base sits lower than Oxi's by 4.68 / 9.60 /
                        // 11.39 for expansions 4.75 / 9.75 / 11.25 (10.5pt raise
                        // 10 / 15, 14pt raise 18) while the NEXT line already agrees
                        // within 0.3; the ruby moves with its base. educational__
                        // 09422f63's title «花粉注意報» (30pt, raise 29) drew its base
                        // 14.5pt high with the annotation outside the line.
                        if std::env::var_os("OXI_S1631_DISABLE").is_none() {
                            for el in elements[flow_elements_start..].iter_mut() {
                                el.y += ruby_para_expansion_pt;
                            }
                        }
                    }
                }
                // Only advance cumul index when cumulative round is active.
                // COM-confirmed (683f): paragraphs with non-uniform line heights
                // (use_cumulative=false) do NOT advance the cross-paragraph index.
                if use_cumulative {
                    cumul_line_idx += 1;
                }
            } // end else (non-LM2 single)

            if std::env::var("OXI_COLUMN_FLOW_HEIGHT").is_ok() {
                let advance = cursor.cursor_y - flow_cursor_before;
                if advance > 0.0 {
                    for element in elements[flow_elements_start..].iter_mut()
                        .filter(|e| matches!(e.content, LayoutContent::Text { .. }))
                    {
                        // Grid cursor rounding belongs to the absolute line
                        // lattice. A section tail reserves the ideal grid box,
                        // not the phase-dependent jump to a hypothetical next line.
                        element.flow_space_before = if line_idx == 0 { effective_spacing } else { 0.0 };
                        element.flow_line_height = Some(if is_lm2_single {
                            line_height
                        } else { advance });
                    }
                }
            }

            // Handle explicit page/column breaks after this line
            if line.break_type == LineBreakType::PageBreak
                || (line.break_type == LineBreakType::ColumnBreak && !s1335_break_consumed)
            {
                // S733 (2026-07-03): a COLUMN break in a multi-column section
                // advances to the NEXT COLUMN of the same page (Word semantics);
                // only from the LAST column (or a 1-col section, where Word
                // treats it as a page break) does it start a new page. It was
                // handled identically to PageBreak — every <w:br type="column">
                // pushed a whole page (probexcolbrk: 5 column breaks in a 2-col
                // doc → Oxi 8 pages vs Word 6, score 0.375). Same column-flow
                // shape as S637. Opt-out OXI_S733_DISABLE.
                if line.break_type == LineBreakType::ColumnBreak
                    && num_columns > 1
                    && cur_col + 1 < num_columns
                    && std::env::var("OXI_S733_DISABLE").is_err()
                {
                    cur_col += 1;
                    start_x = col_x_positions[cur_col];
                    cursor.set(column_flow_top(if pages.len() > s749_pages_at_entry {
                        page_top
                    } else {
                        col_band_top
                    }, start_x, pages.len()));
                    // S1240 (2026-08-27, default ON, opt-out OXI_S1240_DISABLE):
                    // a PARAGRAPH-FINAL column break's paragraph MARK occupies
                    // one line at the NEW column's top — the mark follows the
                    // break character, so it lands in the next column and the
                    // next paragraph starts one line lower. forms__000cf39c:
                    // Word's col2 = mark 15.75 + 3×15.75 + 3×14.0 = 89.25 above
                    // the inline table (COM: 58.5 + 89.25 = 147.75 exact); Oxi
                    // sat one mark-line (15.75) high. A break with same-
                    // paragraph text after it produces further lines instead —
                    // no extra cost there.
                    if line_idx + 1 == lines.len()
                        && std::env::var("OXI_S1240_DISABLE").is_err()
                    {
                        cursor.advance(line_height);
                    }
                } else {
                    if std::env::var("OXI_DBG_COL").is_ok() {
                        eprintln!("[COL] break-push line={}/{} break={:?} num_columns={} cur_col={}",
                            line_idx, lines.len(), line.break_type, num_columns, cur_col);
                    }
                    // Day 33 part 59 (2026-05-12): the line that CARRIES the break_type
                    // has its text already rendered into `elements` and should stay on
                    // the CURRENT page (text BEFORE the `<w:br w:type="page"/>` belongs
                    // to current page per OOXML semantics). Original code pushed
                    // current_elements first then merged elements → pi=11 text ended up
                    // on the NEW page. Fix: merge elements into the pushed page first.
                    let mut page_elements = std::mem::take(current_elements);
                    page_elements.extend(std::mem::take(&mut elements));
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: page_elements,
                    });
                    if let Some(g) = s755_geom {
                        page_top = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(page_top);
                    // S733: a real page push (page break, or column break from the
                    // last column) lands on column 0 of the new page.
                    cur_col = 0;
                    if num_columns > 1 {
                        start_x = col_x_positions[0];
                    }
                } // S733 end else
            }
            line_idx += 1;
        }

        // COM-confirmed (2026-04-16, 683f p2 + minimal repro): content paragraphs
        // adjacent to a RUN of ≥2 consecutive empty paragraphs get +0.5pt extra advance.
        // Only applies to LM0 no-grid single spacing. Skip if paragraph caused page break.
        //
        // S1173 (2026-08-20, default ON, opt-out OXI_S1173_DISABLE): CJK docs
        // only. The 2026-04-16 derivation doc is 683f -- JP -- and its COM read
        // predates the Info6 quantization lesson (Information(6) is 0.75pt-
        // quantized in Latin docs; a phantom +0.5 is exactly its resolution).
        // Word PDF truth 2026-08-20 falsifies the rule for Latin no-grid TWICE:
        // the bisect3_e2 minimal repro (Garamond 10, [4 content, 2 empty,
        // 4 content]) measures m03->t00 = 33.720 = 3 x 11.24 EXACT and
        // t00->t01 = 11.28 -- no +0.5 on either side -- and forms__002fbe2c's
        // real [empty, empty, numbered heading] spans measure 33.73/33.84
        // (line-box) where Oxi's +0.5 made 34.25. That +0.5 x 2 spans is the
        // document's whole p2 drift: it loses the last line and slips its one
        // FAIL paragraph. 683f (JP) keeps the rule -- today's falsification is
        // Latin-scoped, and 683f sits in the golden 96 to catch any flip.
        // S1481 (2026-09-19, default OFF, opt-in OXI_S1481_LEGACY=1): the JP
        // half of the rule falls too. The 683f derivation read Information(6),
        // which is 0.75pt-quantized (683f_word_paras.json: 56.5 / 70.0 / 83.5 /
        // 97.5 -- every y a multiple of 0.75), so a phantom +0.5 is its
        // resolution. Word PDF truth on a JP no-grid repro (MS Mincho 11pt,
        // [3 content, N empty, 2 content], N = 1 / 2 / 3, with and without a
        // no-type docGrid) measures the span at (N+1) x 14.30 EXACT (-0.04 ..
        // -0.08); Oxi's +0.49 fires only at N = 2. ikujidetail page 1 carries
        // the same shape (3-line body, 2 empties, heading): Word 45.72 = the
        // exact model, Oxi 46.38 -- the +0.5 is the first divergence of the
        // whole document once S571's pitch snap stops masking it.
        if adjacent_to_empty_run
            && is_single_lm0
            && grid_pitch.is_none()
            && (cursor.cursor_y - page_top).abs() > 0.1
            && std::env::var("OXI_S1481_LEGACY").is_ok()
        {
            cursor.advance(0.5);
        }

        let space_after = if para.style.after_autospacing
            && std::env::var("OXI_S675_DISABLE").is_err()
            // The style-derived body after spacing follows the same rule as
            // direct spacing regardless of unrelated CJK text, like before.
        {
            // S675 (2026-06-26): w:afterAutospacing → flat auto-space (see
            // space_before; S901: 14.0 Latin / 13.75 JP-calibrated).
            // S907 (2026-07-17): the JP true value is ALSO 14.0 — PDF probe
            // aspj_nogrid (MS Mincho 10.5): pitch 27.62 − line 13.617 =
            // 14.006; grid360: 18 + 14 = 32.0 EXACT. The 13.75 was the same
            // COM quantization artifact as the Latin one. Opt-out
            // OXI_S907_DISABLE restores 13.75 for CJK docs.
            if !self.doc_body_has_real_cjk && std::env::var("OXI_S901_DISABLE").is_err() {
                14.0
            } else if std::env::var("OXI_S907_DISABLE").is_err() {
                14.0
            } else {
                13.75
            }
        } else if let (Some(al), Some(pitch)) = (para.style.after_lines, grid_pitch) {
            // afterLines: exact value (al/100 * pitch), no grid snap needed.
            al / 100.0 * s1480_lines_unit(page.doc_grid_no_type, pitch)
        } else if let Some(al) = para.style.after_lines.filter(|_| {
            page.grid_line_pitch.is_none() && std::env::var("OXI_S697_DISABLE").is_err()
        }) {
            // S697: no-grid afterLines = (al/100) × 12.0pt fixed (doc has no docGrid;
            // textbox-in-docGrid excluded via page.grid_line_pitch — see space_before).
            al / 100.0 * NO_GRID_LINE_PT
        } else {
            para.style.space_after.unwrap_or(0.0)
        };
        // NOTE: space_after is NOT added to cursor_y here.
        // It will be collapsed with the next paragraph's space_before via max(sa, sb).

        // S674 (2026-06-26, opt-in OXI_S674, default OFF = byte-identical) ATTEMPTED +
        // FALSIFIED default-on — the gen2-Latin para-spacing device-snap residual (the
        // S671-deferred "per-para exact-cumulative" lever). S671 accumulates each line's
        // EXACT hhea-natural height (multi-line-safe); Word renders each para's LAST
        // baseline + after-spacing device-snapped to the 0.12pt (600-DPI px) grid →
        // a systematic ~+0.12pt/para the exact accumulation lacks (Word body single-line
        // para gap 24.95 vs Oxi exact 24.83 = line round_0.12(14.83→14.88)+0.05 + after
        // ceil_0.12(10.0→10.08)+0.08). S674 applies the EXACT per-component snap (round
        // line + ceil after, ADAPTS per font/size — fixing the failed OXI_S671_SADD's
        // const-0.12 + single-line-only scope) ONCE per para to ALL s671 paras (the
        // boundary snap is multi-line-safe — interior lines stay exact) on the CURSOR.
        // ★RESULT (ssim_ab.py OXI_S674, gen2_/gen_/test_ family, 54 changed): net
        // +0.0573 (33 improved / 21 regressed) — HONEST net-positive but NOT shippable:
        // (1) PAGINATION-NEUTRAL everywhere (verified: para counts byte-identical OFF/ON
        // on gen2_054/067, test_widow, test_keepnext — pure render-Y drift +0.08/para);
        // (2) UNDISCRIMINABLE — gen2_064 (helps −0.034) and gen2_054 (hurts +0.119) have
        // BYTE-IDENTICAL Oxi top-of-page positions AND both word_pngs drift 24.95, yet
        // split help/hurt (they diverge deeper by content flow → the SSIM benefit is
        // dominated by the ABSOLUTE title-block phase, not the body drift, with NO
        // docx-derivable rule); (3) regresses the CLEANEST controlled refs (test_widow
        // −0.185 — its word_png matches S671's EXACT accumulation, Word does NOT drift
        // it; test_keepnext −0.059) and the bottom-N (gen2_054 −0.119, gen2_050 −0.065).
        // ⇒ the body drift is REAL but the lever is the COUPLED title-block absolute
        // phase (S614/S618/S620 wall); shipping S674 ALONE = an S559 compensating error
        // the eventual title-block fix must untangle. Kept opt-in for the coupled
        // multi-lever session (S674 body-drift ⊕ title-block phase, word_png-gated). The
        // 5th device-snap attempt on this wall (S629-CJK shipped, S631/S673/S671_SADD
        // falsified). See [[gen2_vertical_drift]].
        if s671_fine
            && std::env::var("OXI_S674").is_ok()
            && !lines.is_empty()
            && (cursor.cursor_y - page_top).abs() > 0.1
        {
            let d = std::env::var("OXI_S674_DELTA")
                .ok()
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.12);
            let last_lh = *line_heights.last().unwrap_or(&0.0);
            let line_snapped = (last_lh / d).round() * d; // line → ROUND
            let after_snapped = if space_after > 0.01 {
                (space_after / d).ceil() * d // after → CEIL
            } else {
                space_after
            };
            let corr = (line_snapped - last_lh) + (after_snapped - space_after);
            cursor.advance(corr);
        }

        // Paragraph borders (e.g., Title style bottom border)
        if let Some(ref borders) = para.style.borders {
            // S1134: `start_x` was the historical fallback — an X in a Y slot.
            // Opt-out OXI_S1134_DISABLE restores it.
            let para_top = elements.first().map(|e| e.y).unwrap_or(
                if std::env::var("OXI_S1134_DISABLE").is_ok() {
                    start_x
                } else {
                    s1134_content_top
                },
            );
            let para_bottom = cursor.cursor_y;
            let border_x = start_x;
            let border_width = content_width;

            // S990A (2026-07-23): an INTERIOR merged boundary (next para has the
            // same effective borders) with a `between` border reserves the
            // between line's vertical extent (thickness + 2×space) — the derived
            // Round-74 rule (measure_between_border.py, 2026-05-03), previously
            // draw-only (S903 reserved ZERO). educational__00252fa WritingLines:
            // sz=4/space=1 → 0.5 + 2×1 = 2.5pt/boundary × 94, which (with the
            // Gill line height S990B) makes the border pitch = Word PDF 34.14.
            // Latin scope (JP 3a4f 6-box stack keeps its calibration) + the
            // between must be a real border. The duplicate overlapping bottom is
            // skipped below; the between draw stays UNCONDITIONAL so a CJK doc's
            // between line is never lost. Opt-out OXI_S990A_DISABLE.
            // S1042 (2026-07-29, opt-out OXI_S1042_DISABLE): Word draws ONE outer
            // box around a run of consecutive paragraphs whose effective pBdr is
            // identical - the interior top/bottom edges are not painted at all
            // (a `between` border, when declared, is what draws them). Oxi has
            // merged the vertical RESERVATION since S658/S903 but kept painting
            // every paragraph's own top and bottom, so a bordered form grew a
            // horizontal rule at every interior boundary. forms__002a64445e58ed78
            // has three such groups (7 + 13 + 9 paragraphs) and 26 spurious rules;
            // dropping them lifts its SSIM 0.576 -> 0.681 (an Oxi-position raster
            // lower bound; the pure no-shading arm reaches 0.694). Reuses the same
            // two facts S658/S903 already thread in - no new discriminator.
            let s1042 = std::env::var("OXI_S1042_DISABLE").is_err();
            let s1042_has_between = borders
                .between
                .as_ref()
                .map_or(false, |b| b.style != "none" && b.style != "nil");
            // interior TOP: the previous paragraph carries the same box
            let s1042_skip_top =
                s1042 && !s1042_has_between && prev_para_borders.map_or(false, |pb| pb == borders);
            // interior BOTTOM: the next paragraph carries the same box
            let s1042_skip_bottom =
                s1042 && !s1042_has_between && s903_next_borders.map_or(false, |nb| nb == borders);
            let s990a_between = std::env::var("OXI_S990A_DISABLE").is_err()
                && !self.doc_body_has_real_cjk
                && s903_next_borders.map_or(false, |nb| nb == borders)
                && borders
                    .between
                    .as_ref()
                    .map_or(false, |b| b.style != "none" && b.style != "nil");

            if let Some(ref bottom) = borders.bottom {
                let s1504 = std::env::var_os("OXI_S1504_DISABLE").is_none();
                let bw = if s1504 && bottom.style == "double" { bottom.width * 3.0 } else { bottom.width };
                let color = bottom.color.clone().unwrap_or_else(|| "000000".to_string());
                let border_y = para_bottom + bottom.space;
                // S990A: the interior `between` element below draws this line;
                // skip the duplicate overlapping bottom shading.
                if !s990a_between && !s1042_skip_bottom {
                    elements.push(LayoutElement::new(
                        border_x,
                        border_y,
                        border_width,
                        bw.max(0.5),
                        LayoutContent::CellShading {
                            color: format!("#{}", color),
                        },
                    ));
                }
                // S467 (2026-05-31, env-gated OFF default; opt-in OXI_S467_PBDR_ENABLE):
                // the FULL border width (space + bw) is the CORRECT advance, not the
                // midpoint (space + bw/2). The old "bw/2" comment cited a gen2_036
                // measurement that was a non-collapsed-start (R30) artifact; re-measured
                // collapsed-start, gen2_036 title gap = 54.0 (not 38.5), gen2_055/056/067
                // = 51.75. A minimal repro (Calibri 26pt single sa=15 + bottom border
                // sz=8=1.0pt space=4) confirms Word reserves the FULL border width below
                // the text before space-after: with-border gap 51.75 = lineBox(31.5)+
                // space(4)+bw(1.0)+sa(15)+grid-snap(0.25). COM-confirmed the fix makes the
                // title gap EXACT/closer for BOTH EN (gen2_055 -0.75->-0.25) and JP
                // (gen2_001 -0.50->+0.00) titles. NOT SHIPPED default-ON: the fix is
                // correct but propagates a +0.5 shift to ALL p1 content below the title,
                // which HELPS docs whose body drifted too-high (EN gen2: +0.02..+0.05) but
                // HURTS docs whose body was already aligned via a compensating CJK/per-doc
                // line-height error (gen2 OFF-vs-ON: 55 up / 26 down, incl. gen2_054 EN
                // -0.054, gen2_001 JP -0.046 — NOT separable by language). Body alignment
                // is per-doc inconsistent (the gen2 drift is fragmented), so the title fix
                // can only ship together with the body line-height fix. Kept gated for
                // when that lands. Default OFF = byte-identical baseline.
                // S1048 (2026-07-30, default ON for LATIN documents, opt-out
                // OXI_S1048_DISABLE; OXI_S467_PBDR_ENABLE still forces it everywhere):
                // the full width IS Word's rule, now measured font- and size-independently
                // on a 76-arm controlled probe (26/11/12/13/14pt × Calibri/TNR × border
                // sz 8/16 × space 0/4 × after 0..30 × follower before 0/24 × shading):
                // Word's paired bottom-border increment is +0.9600 (sz8 space0),
                // +5.0400 (sz8 space4) and +6.0000 (sz16 space4) = **space + FULL width**
                // (half would be 0.500 / 4.500 / 5.000 — refuted at every arm). Over all
                // 76 arms the residual vs Word is MAE 0.305 / max 1.088 with bw/2 and
                // MAE 0.071 / max 0.196 with bw. Paragraph shading contributes 0.0000pt
                // (drawn only, no advance) and the spacing max-collapse slope already
                // matches Word, so the border width was the sole deficit.
                // ★The S467 note above concluded "NOT separable by language" — that was
                // measured against the CACHED word_png refs. Against the FRESH refs
                // (pipeline_data/word_png_new, regenerated 2026-06-27) the NET separates
                // cleanly: LATIN 49 docs net +0.4237 in favour of the full width (32 up /
                // 16 down), CJK 46 docs net −0.2847 against it (18 up / 28 down). Per-DOC
                // it is still mixed inside each language (gen2_054 EN stays down 0.049),
                // so the honest claim is that the NET is language-separable, not every doc.
                // Latin scope also matches the sibling S903 (whose own note keeps the JP
                // 3a4f 6-box stack on its calibration).
                let pbdr_full = std::env::var("OXI_S467_PBDR_ENABLE").is_ok()
                    || (!self.doc_body_has_real_cjk && std::env::var("OXI_S1048_DISABLE").is_err())
                    // S1504: the CJK half-width was a calibration; the typed-grid
                    // probe reads the full stroke (single sz8: 22.5 not 21.5).
                    || std::env::var_os("OXI_S1504_DISABLE").is_none();
                // S903: an INTERIOR paragraph of a merged identical-pBdr group
                // reserves NO bottom overhead (Word: pure line pitch between
                // merged boxes; the space+bw belongs to the group's LAST para).
                // Latin scope — the JP 3a4f 6-box stack keeps its calibration.
                let s903_interior = (!self.doc_body_has_real_cjk
                    || std::env::var_os("OXI_S1439_DISABLE").is_none())
                    && std::env::var("OXI_S903_DISABLE").is_err()
                    && s903_next_borders.map_or(false, |nb| nb == borders);
                if !s903_interior {
                    cursor.set(border_y + if pbdr_full { bw } else { bw / 2.0 });
                }
            }
            if let Some(ref top) = borders.top {
                if !s1042_skip_top {
                    let bw = top.width;
                    let color = top.color.clone().unwrap_or_else(|| "000000".to_string());
                    let s1135 = if std::env::var("OXI_S1135_DISABLE").is_ok() {
                        0.0
                    } else {
                        s1135_atleast_lead
                    };
                    let border_y = para_top + s1135 - top.space - bw;
                    elements.push(LayoutElement::new(
                        border_x,
                        border_y,
                        border_width,
                        bw.max(0.5),
                        LayoutContent::CellShading {
                            color: format!("#{}", color),
                        },
                    ));
                }
            }
            // Between border (horizontal line between consecutive bordered paragraphs)
            if let Some(ref between) = borders.between {
                let bw = between.width;
                let color = between
                    .color
                    .clone()
                    .unwrap_or_else(|| "000000".to_string());
                let border_y = para_bottom + between.space;
                elements.push(LayoutElement::new(
                    border_x,
                    border_y,
                    border_width,
                    bw.max(0.5),
                    LayoutContent::CellShading {
                        color: format!("#{}", color),
                    },
                ));
                // S990A: reserve the between line's extent (thickness + 2×space)
                // at an interior merged boundary. cursor.cursor_y == para_bottom
                // here so this only advances the flow.
                if s990a_between {
                    cursor.set(para_bottom + bw + 2.0 * between.space);
                }
            }
            // Left border
            if let Some(ref left) = borders.left {
                let bw = left.width;
                let color = left.color.clone().unwrap_or_else(|| "000000".to_string());
                let bx = border_x - left.space - bw;
                elements.push(LayoutElement::new(
                    bx,
                    para_top,
                    bw.max(0.5),
                    para_bottom - para_top,
                    LayoutContent::CellShading {
                        color: format!("#{}", color),
                    },
                ));
            }
            // Right border
            if let Some(ref right) = borders.right {
                let bw = right.width;
                let color = right.color.clone().unwrap_or_else(|| "000000".to_string());
                let bx = border_x + border_width + right.space;
                elements.push(LayoutElement::new(
                    bx,
                    para_top,
                    bw.max(0.5),
                    para_bottom - para_top,
                    LayoutContent::CellShading {
                        color: format!("#{}", color),
                    },
                ));
            }
        }

        if let Some(ref mut cells) = lm2_grid_cells {
            **cells = cumul_line_idx;
        }

        // S900: hand the deferred note ids to the caller (next-page area).
        if let Some(out) = fn_deferred_out.as_deref_mut() {
            out.extend(s900_deferred_ids);
        }

        // S1181 v2 exit: re-sync the visual track to the exact stream — but
        // only if the paragraph stayed on its page (a break already re-synced
        // both tracks at the new page top).
        if s1181_unsnap != 0.0 && pages.len() == s1181_pages0 {
            cursor.advance_split(0.0, s1181_unsnap);
        }
        (elements, space_after, cur_col)
    }
}
