// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `LayoutEngine::layout_page` -- moved out of `layout/mod.rs` so that it is its own
//! codegen unit (see tools/metrics/split_layout_mod.py). Behaviour-preserving.

use super::*;

/// `layout_page` as a method of its own type: rustc puts a method's code in the
/// codegen unit of its self type's module, so this (not the file move alone)
/// is what gives the giant its own unit. Deref keeps `self.x` meaning the engine.
pub(super) struct PageLayouter<'a>(pub(super) &'a LayoutEngine);

impl<'a> std::ops::Deref for PageLayouter<'a> {
    type Target = LayoutEngine;
    fn deref(&self) -> &LayoutEngine {
        self.0
    }
}

impl<'a> PageLayouter<'a> {
    #[allow(unused_assignments)]
    /// `logical` carries the LOGICAL page number across IR pages: in = the
    /// number of the last page emitted before this one (0 if none), out = the
    /// number of the last page this call emits. S1291/S1294 state their rules
    /// in that number, and there is exactly one caller, so a scalar in/out is
    /// enough (S912's `page_numbers` recomputes the same walk as a post-pass).
    pub(super) fn layout_page(&self, page: &Page, logical: &mut u32, ir_index: usize, column_search: &mut GridColumnSearch) -> Vec<LayoutPage> {
        // S1294: the logical number of this IR page's FIRST layout page, and
        // the base such that logical(i) = logical_base + i. A restart moves the
        // base rather than every page's number.
        // S1294 SHIPPED default-ON 2026-09-06 (opt-out OXI_S1294_DISABLE) together
        // with S1335, the page-15 defect named below: the two were compensating
        // (0ea3ec86 0.7323 alone / 0.2813 with only this / 0.1040 with only S1335
        // / 0.9940 with both, W43/O43).
        // S1294 was HELD OPT-IN (`OXI_S1294=1`), not because the law is in doubt
        // but because a SECOND defect was cancelling it. With the blank page in
        // its right place `reference__0ea3ec86` reads Word pages 1-14 exactly
        // (they were all one page early before), and then a page-15 table that
        // Oxi breaks early sends the rest +1: the doc's score goes 0.7315 ->
        // 0.2822 and the gate reads WORSE. The missing blank and the early
        // table break were compensating, and only one of them is fixed. Turn
        // this on together with the page-15 fix.
        let s1294_on = std::env::var("OXI_S1294_DISABLE").is_err();
        let first_logical = section_page_number_start(page).unwrap_or(*logical + 1);
        let mut logical_base: i64 = first_logical as i64;
        // Vertical writing (tategaki) section: route to the dedicated path.
        if page.vertical_section && std::env::var("OXI_VERTICAL_DISABLE").is_err() {
            let out = self.layout_page_vertical(page);
            *logical = (logical_base + out.len() as i64 - 1).max(0) as u32;
            return out;
        }
        // R-05b: reduce body content width when the document has comments —
        // makes room for the right-margin balloon column. Header / footer /
        // floating-image / footnote widths intentionally use the full
        // un-reduced width (matches Word's behavior: only the body reflows).
        // S-01: only reduce when the engine's `show_comments` is true.
        let balloon_reservation = if self.show_comments {
            self.balloon_column_width
        } else {
            0.0
        };
        let total_content_width =
            page.size.width - page.margin.left - page.margin.right - balloon_reservation;
        // COM-confirmed (2026-04-03, order_08): when header extends below margin.top,
        // body content starts below the header (header pushes body down).
        // header_distance + header_content_height = header_bottom.
        // start_y = max(margin.top, header_bottom)
        // S755: page 1 of a titlePg section uses the FIRST-type header
        // (absent first reference = blank per ECMA-376 §17.10.2); the
        // default header drives pages 2+ via s755_geom below.
        let s755_on = std::env::var("OXI_S755_DISABLE").is_err();
        let s755_first_hdr: &[Block] = if s755_on && page.title_pg {
            &page.header_first
        } else if s755_on && page.even_odd_hf && first_logical % 2 == 0 {
            &page.header_even
        } else {
            &page.header
        };
        // S1174: this section's headers/footers re-resolve STYLEREF per page.
        // The search is SECTION-SCOPED both ways (reference__0061531a's
        // Schedule pages: Word leaves the Part/Division STYLEREF lines BLANK
        // although Part 4 precedes them and 'Part 1—Costs' follows — neither
        // the backward nor the forward search crosses the section boundary),
        // so the registries reset here and the forward-fallback prescan walks
        // THIS section's blocks only.
        let s1174_have_ref = S1174_ACTIVE.with(|c| c.get())
            && (LayoutEngine::s1174_has_ref(&page.header)
                || LayoutEngine::s1174_has_ref(&page.footer)
                || LayoutEngine::s1174_has_ref(&page.header_first)
                || LayoutEngine::s1174_has_ref(&page.footer_first)
                || LayoutEngine::s1174_has_ref(&page.header_even)
                || LayoutEngine::s1174_has_ref(&page.footer_even));
        if S1174_ACTIVE.with(|c| c.get()) {
            S1174_LAST.with(|m| m.borrow_mut().clear());
            S1174_FIRST.with(|m| m.borrow_mut().clear());
            for b in &page.blocks {
                LayoutEngine::s1174_ingest_block(b, true);
            }
        }
        let s1174_map0 = if s1174_have_ref { Some(LayoutEngine::s1174_map()) } else { None };
        let s1174_first_hdr_sub: Vec<Block>;
        let s755_first_hdr: &[Block] = if let Some(m) = s1174_map0.as_ref() {
            s1174_first_hdr_sub = LayoutEngine::s1174_substitute(s755_first_hdr, m);
            &s1174_first_hdr_sub
        } else {
            s755_first_hdr
        };
        let header_bottom = self.s755_header_bottom(s755_first_hdr, page);
        let mut start_y = page.body_start_y(header_bottom, self.s1381_header_band(s755_first_hdr, page));

        // §11.2.2 LM2 unified P0 formula (Round 23, COM-confirmed 2026-04-08).
        // In linesAndChars (LM2) mode, the FIRST body paragraph is allocated a
        // grid-snapped cell whose height = strict-greater snap of LM0_lh, and
        // the line box is vertically centered within that cell:
        //   P0_h = (floor(LM0_lh / pitch) + 1) * pitch
        //   P0_y = topMargin + (P0_h - LM0_lh) / 2
        // Subsequent paragraphs use the regular per-line grid snap.
        // Only applies when header_bottom <= topMargin (no header pushdown).
        if header_bottom <= page.margin.top {
            if let Some(pitch) = page.grid_line_pitch {
                if pitch > 0.0 {
                    if let Some(first_para) = page.blocks.iter().find_map(|b| match b {
                        Block::Paragraph(p) => Some(p),
                        _ => None,
                    }) {
                        // Round 28 (2026-04-08, COM-confirmed): lineSpacingRule="exact"
                        // completely DISABLES the LM2 first-paragraph centering. Word
                        // places P0_y = topMargin exactly regardless of font/size/value.
                        // Verified across TNR/MS Mincho × 10.5/12/14pt × exact_12/18/24/36
                        // — all 24 combinations measured P0_y = 72.00.
                        let rule = first_para.style.line_spacing_rule.as_deref();
                        if rule != Some("exact") {
                            // Use full inheritance chain (resolve_font_size) so the
                            // Normal style sz= value (e.g. b837: sz=24=12pt) is picked up
                            // when the run/pPr.rPr have no explicit size. Earlier manual
                            // chain bypassed default_run_style and fell back to 11pt.
                            let default_run_style = RunStyle::default();
                            let first_run_style = first_para
                                .runs
                                .first()
                                .map(|r| &r.style)
                                .unwrap_or(&default_run_style);
                            let fs = first_para
                                .style
                                .ppr_rpr
                                .as_ref()
                                .and_then(|r| r.font_size)
                                .unwrap_or_else(|| {
                                    self.resolve_font_size(first_run_style, &first_para.style)
                                });
                            let metrics = first_para
                                .runs
                                .first()
                                .map(|r| self.metrics_for(&r.style, &first_para.style))
                                .unwrap_or_else(|| {
                                    let rpr_ref = first_para
                                        .style
                                        .ppr_rpr
                                        .as_ref()
                                        .cloned()
                                        .unwrap_or_default();
                                    self.metrics_for_para_mark(&rpr_ref, &first_para.style)
                                });
                            // LM0 base line height (Round 9 lookup if available).
                            let lm0_lh = self
                                .registry
                                .lm0_lineauto_base(&metrics.family, fs)
                                .unwrap_or_else(|| metrics.word_line_height_no_grid(fs));
                            // Strict-greater snap to next pitch multiple.
                            let cells = (lm0_lh / pitch).floor() + 1.0;
                            let p0_h = cells * pitch;
                            // COM-confirmed (2026-04-13, db9c): Word does NOT add
                            // the centering offset to cursor_y. The cursor starts
                            // at topMargin; centering is achieved via text_y_offset
                            // (= (pitch - natural) / 2) in text_y_offset_for_line().
                            // Adding p0_offset to start_y caused 2+pt cursor drift
                            // that accumulated over the entire page (38 lines × 2pt
                            // drift in db9c = different page count).
                            // NOTE (Session 107, 2026-05-18): the half-leading IS
                            // applied at cursor.set(page_top) sites inside
                            // layout_paragraph for subsequent pages. See d77a p.2
                            // fix below — page 1's first paragraph is intentionally
                            // left at topMargin to avoid the cascade documented above.
                            let _p0_offset = (p0_h - lm0_lh) / 2.0;
                            // Previously: start_y += p0_offset;
                        }
                    }
                }
            }
        }
        // Body content area: reserves footer space at the bottom.
        // Word reserves footer height from the body content area. If body extends
        // past the footer-top position, content overlaps footer. COM-confirmed
        // on 04b88e (2026-04-17): Word body stops above footer, Oxi body extends
        // past it — causing 1 fewer page than Word.
        // Footer reservation = footer_distance + footer_height. Compare to
        // page.margin.bottom; use whichever is larger.
        let s755_first_ftr: &[Block] = if s755_on && page.title_pg {
            &page.footer_first
        } else if s755_on && page.even_odd_hf && first_logical % 2 == 0 {
            &page.footer_even
        } else {
            &page.footer
        };
        // S1174: footer STYLEREF fields resolve the same way.
        let s1174_first_ftr_sub: Vec<Block>;
        let s755_first_ftr: &[Block] = if let Some(m) = s1174_map0.as_ref() {
            s1174_first_ftr_sub = LayoutEngine::s1174_substitute(s755_first_ftr, m);
            &s1174_first_ftr_sub
        } else {
            s755_first_ftr
        };
        let (footer_reserved, footer_has_text) = self.s755_footer_geom(s755_first_ftr, page);

        let footer_tight = footer_reserved > page.margin.bottom + 0.05
            && footer_has_text
            && std::env::var("OXI_S726_DISABLE").is_err();
        let mut content_height = page.size.height - start_y - footer_reserved;
        // S755: per-page geometry variants. Some ONLY when first/even/odd
        // actually differ (>0.05pt) — the whole corpus is None (titlePg docs
        // have no header refs; albaluna variants are same-height 1-line;
        // bd90b00's first/default are both 1 line) → byte-identical by
        // construction. Page 1 = the (start_y, content_height) just computed.
        let s755_geom: Option<S755Geom> = if s755_on && (page.title_pg || page.even_odd_hf) {
            let hb_odd = self.s755_header_bottom(&page.header, page);
            let sy_odd = page.body_start_y(hb_odd, self.s1381_header_band(&page.header, page));
            let (fr_odd, _) = self.s755_footer_geom(&page.footer, page);
            let ch_odd = page.size.height - sy_odd - fr_odd;
            let (sy_even, ch_even) = if page.even_odd_hf {
                // Absent even reference with the flag set = BLANK even header
                // (ECMA-376), like the titlePg first-page rule.
                let hb = self.s755_header_bottom(&page.header_even, page);
                let sy = page.body_start_y(hb, self.s1381_header_band(&page.header_even, page));
                let (fr, _) = self.s755_footer_geom(&page.footer_even, page);
                (sy, page.size.height - sy - fr)
            } else {
                (sy_odd, ch_odd)
            };
            let differs = (sy_odd - start_y).abs() > 0.05
                || (ch_odd - content_height).abs() > 0.05
                || (sy_even - sy_odd).abs() > 0.05
                || (ch_even - ch_odd).abs() > 0.05;
            // Preserve the numbering origin even when both variants have equal height.
            if differs || page.even_odd_hf {
                Some(S755Geom {
                    first_even: first_logical % 2 == 0,
                    first: (start_y, content_height),
                    odd: (sy_odd, ch_odd),
                    even: (sy_even, ch_even),
                    page_override: None,
                })
            } else {
                None
            }
        } else {
            None
        };
        // S863: a continuous section's top/bottom and header/footer
        // distances govern physical pages that begin after its boundary. Build
        // the same first/odd/even geometry tuple as S755, but for every section
        // run; the block loop switches the active tuple without moving the
        // boundary page's current cursor.
        let s863_vertical_geoms: Vec<S755Geom> =
            if std::env::var("OXI_S863_DISABLE").is_err() && page.vertical_runs.len() > 1 {
                page.vertical_runs
                    .iter()
                    .enumerate()
                    .map(|(ri, (_, top, bottom, hd, fd))| {
                        let mut rp = page.clone();
                        rp.margin.top = *top;
                        rp.margin.bottom = *bottom;
                        rp.header_distance = *hd;
                        rp.footer_distance = *fd;
                        // S1553 (2026-09-25, default ON, opt-out OXI_S1553_DISABLE):
                        // a merged continuous section's pages take the header/
                        // footer set that section carries, not the first
                        // section's (see Page::header_runs).
                        if std::env::var_os("OXI_S1553_DISABLE").is_none() {
                            if let Some(hr) = page.header_runs.get(ri) {
                                rp.header = hr.header.clone();
                                rp.footer = hr.footer.clone();
                                rp.header_first = hr.header_first.clone();
                                rp.footer_first = hr.footer_first.clone();
                                rp.header_even = hr.header_even.clone();
                                rp.footer_even = hr.footer_even.clone();
                                rp.title_pg = hr.title_pg;
                                rp.even_odd_hf = hr.even_odd_hf;
                            }
                        }
                        let geom = |hdr: &[Block], ftr: &[Block]| {
                            let sy = rp.body_start_y(self.s755_header_bottom(hdr, &rp), self.s1381_header_band(hdr, &rp));
                            let (mut fr, _) = self.s755_footer_geom(ftr, &rp);
                            // A zero footer distance pins the footer to the physical
                            // page edge. Word then lets body flow use the bottom margin
                            // boundary (the footer stack does not enlarge that margin)
                            // and permits the usual ~device-pixel ink overhang there.
                            let edge_slack = if rp.footer_distance == Some(0.0) {
                                fr = rp.margin.bottom;
                                1.25
                            } else {
                                0.0
                            };
                            (sy, rp.size.height - sy - fr + edge_slack)
                        };
                        let first = if s755_on && rp.title_pg {
                            geom(&rp.header_first, &rp.footer_first)
                        } else if s755_on && rp.even_odd_hf && first_logical % 2 == 0 {
                            geom(&rp.header_even, &rp.footer_even)
                        } else {
                            geom(&rp.header, &rp.footer)
                        };
                        let odd = geom(&rp.header, &rp.footer);
                        let even = if s755_on && rp.even_odd_hf {
                            geom(&rp.header_even, &rp.footer_even)
                        } else {
                            odd
                        };
                        S755Geom { first_even: first_logical % 2 == 0, first, odd, even, page_override: None }
                    })
                    .collect()
            } else {
                Vec::new()
            };
        let mut s755_geom = s755_geom;
        let mut s863_vertical_run_idx: usize = 0;
        // S1227 (2026-08-26): the vertical-run geom that was ACTIVE WHEN THE
        // CURRENT PAGE BEGAN. Word's continuous-section top/bottom margins
        // govern whole physical pages — a mid-page section switch must not
        // shrink/grow the CURRENT page's content box (kyotei36spec p4: the
        // 裏面 page begins in the title section, bottom=284tw → content
        // bottom 581.1; the 2-col body section's bottom=510tw applies only
        // to pages BEGINNING in it. The per-block refresh below adopted the
        // switched geom immediately → Oxi's columns stopped one grid row
        // early, Word packs the row-47 line box to 580.35 ≤ 581.1). The
        // page-begin index updates whenever pages.len() changes, capturing
        // the run of the block that pushed the page. Opt-out OXI_S1227_DISABLE.
        let mut s863_page_begin_idx: usize = 0;
        let mut s863_last_pages_len: usize = 0;
        let s1227_on = std::env::var("OXI_S1227_DISABLE").is_err();
        // Round 29 (2026-04-08): per-page dynamic footnote reservation.
        // Footnotes are reserved at the bottom of the page where their reference
        // appears. The amount reserved varies per page based on which footnotes
        // are referenced. Tracked dynamically as the body layout progresses:
        // when a paragraph contains a footnoteReference, the corresponding note's
        // estimated body height is added to the running reservation; on page
        // break the reservation resets. The body's effective overflow check uses
        // (content_height - footnote_reserved_current_page).
        // Helper to estimate one footnote body height by id.
        // S727 (2026-07-03): footnote-height estimates pass the TYPED grid pitch
        // so grid-snapping footnote paragraphs reserve their real (snapped) line
        // height. Word render-truth (probefn, type=lines 360): footnote lines
        // snap to the 18pt grid (gaps 18.0 at 9pt font) and Oxi's RENDER already
        // snaps (FN_PLACE heights=18.0) — but the estimate passed gp=None →
        // natural ~11.3/line → the body under-reserved ~6pt/note → packed +2
        // body lines/page → probefn {-1:8}. The DISCRIMINATOR rides
        // estimate_para_height's own `para.style.snap_to_grid` gate: real docs'
        // footnote-text style sets snapToGrid=0 (b837 style a8) → natural,
        // byte-identical; style-less footnote paras (snap default true) snap.
        // No-type grids don't snap (S609/S571 family) → keep None.
        let fn_est_gp = if std::env::var("OXI_S727_DISABLE").is_err() && !page.doc_grid_no_type {
            page.grid_line_pitch
        } else {
            None
        };
        if std::env::var("OXI_FN_PROBE").is_ok() {
            eprintln!(
                "[FN_EST] fn_est_gp={:?} pitch={:?} no_type={}",
                fn_est_gp, page.grid_line_pitch, page.doc_grid_no_type
            );
        }
        let estimate_footnote_h = move |id: u32| -> f32 {
            let _fng = FnLayoutGuard::new();
            if let Some(note) = page.footnotes.iter().find(|n| n.number == id) {
                let cw = page.size.width - page.margin.left - page.margin.right;
                let mut h: f32 = 0.0;
                let mut first_para = true;
                // S804 (2026-07-12, opt-out OXI_S804_DISABLE): footnote paragraphs
                // inherit the style chain's before/after spacing (uklocalspending:
                // FootnoteText basedOn Normal before/after=240 -> Word inter-note
                // gap = line 11.5 + collapse 12 = 23.6, and the LAST note's after
                // sits inside the bottom-anchored stack). estimate_para_height
                // drops style-level spacing (the S709/S803 class), so the
                // reservation under-counted ~24pt/note -> fn pages over-packed
                // ~60-100pt (probe fn_probe.py: 2 notes cost 98.5pt of body in
                // Word vs ~61 reserved). Add the internal collapse gaps + the
                // trailing after; the first paragraph's before belongs to the
                // separator gap (footnote_sep_alloc). Gated per-para to
                // !has_direct_spacing = exactly when the estimate dropped it; JP
                // footnote styles carry no spacing -> +0, byte-identical.
                let s804 = std::env::var("OXI_S804_DISABLE").is_err();
                let mut s804_prev_sa: Option<f32> = None;
                for nb in &note.blocks {
                    if let Block::Paragraph(p) = nb {
                        if s804 && self.footnote_twip_spacing_supported(&p.style) {
                            h += self.footnote_twip_spacing_correction(
                                &p.style, &mut s804_prev_sa);
                        } else if s804 && !p.style.has_direct_spacing {
                            let sb = p.style.space_before.unwrap_or(0.0);
                            let sa = p.style.space_after.unwrap_or(0.0);
                            if let Some(prev) = s804_prev_sa {
                                h += prev.max(sb);
                            }
                            s804_prev_sa = Some(sa);
                            // S810 (2026-07-13): a non-auto (exact/atLeast) fn para
                            // KEEPS its style sb/sa inside estimate_para_height
                            // (should_reset is auto-only) — the S806(d) footer
                            // discovery applied to the fn stack. Strip so the gap
                            // accounting is single-source (ukframework
                            // FootnoteText: line=240 exact + after=60; Word fn
                            // pitch = 15.0 = exact 12 + after 3, NOT 18).
                            if !matches!(p.style.line_spacing_rule.as_deref(), None | Some("auto"))
                            {
                                h -= sb + sa;
                            }
                        } else {
                            s804_prev_sa = Some(0.0);
                        }
                        if first_para {
                            // Footnote rendering prepends a seq number to the first
                            // paragraph, which increases its width and may add a line.
                            // Clone and add prefix to match actual rendering.
                            let mut p2 = p.clone();
                            let seq = page
                                .footnotes
                                .iter()
                                .position(|n| n.number == id)
                                .map(|pos| (pos as u32) + 1)
                                .unwrap_or(id);
                            let prefix = format!("{}", seq);
                            if let Some(first_run) = p2.runs.first_mut() {
                                if first_run.text.is_empty() {
                                    first_run.text = prefix;
                                } else {
                                    first_run.text = format!("{}{}", prefix, first_run.text);
                                }
                            }
                            // S727: render-lh dispatch — a SNAPPING footnote para
                            // estimates at the render line height (grid-snapped,
                            // line_height_inner) so the reservation matches the
                            // emitted 18pt/line; non-snapping (footnote-text style
                            // snapToGrid=0, e.g. b837 a8) keeps the calibrated
                            // word_line_height_table_cell estimate byte-identically.
                            // S727: a SNAPPING footnote paragraph in a typed grid
                            // occupies whole grid cells per line (the footnote body
                            // renders through the BODY line-height path, which
                            // grid-snaps: probefn render = 18.0/line at 9pt; Word
                            // render-truth gaps = 18.0). The natural estimate
                            // under-reserved ~6pt/line → the body over-packed.
                            // lines = natural_h / per-line natural (uniform-font
                            // notes); cells/line = ceil(natural line / pitch).
                            // Non-snapping footnote paras (footnote-text style
                            // snapToGrid=0, e.g. b837 a8) keep the calibrated
                            // natural estimate byte-identically.
                            let ph_nat =
                                self.estimate_para_height(&p2, cw, None, None, false, None, None);
                            // Paragraph spacing is not a count of text lines.
                            let mut line_para = p2.clone();
                            line_para.style.space_before = Some(0.0);
                            line_para.style.space_after = Some(0.0);
                            line_para.style.before_lines = None;
                            line_para.style.after_lines = None;
                            let line_height = self.estimate_para_height(
                                &line_para, cw, None, None, false, None, None);
                            let paragraph_spacing = ph_nat - line_height;
                            let ph = if let (Some(pitch), true) = (fn_est_gp, p2.style.snap_to_grid)
                            {
                                let fs = self.resolve_font_size(
                                    p2.runs
                                        .iter()
                                        .find(|r| !r.text.trim().is_empty())
                                        .map(|r| &r.style)
                                        .unwrap_or(&RunStyle::default()),
                                    &p2.style,
                                );
                                let m = self.metrics_for(
                                    p2.runs
                                        .iter()
                                        .find(|r| !r.text.trim().is_empty())
                                        .map(|r| &r.style)
                                        .unwrap_or(&RunStyle::default()),
                                    &p2.style,
                                );
                                let per_line = m.word_line_height_table_cell(fs).max(1.0);
                                let lines = (line_height / per_line).round().max(1.0);
                                let cells =
                                    (m.word_line_height_no_grid(fs) / pitch).ceil().max(1.0);
                                lines * cells * pitch + paragraph_spacing
                            } else if !self.doc_body_has_real_cjk
                                && std::env::var("OXI_S808_DISABLE").is_err()
                                && matches!(
                                    p2.style.line_spacing_rule.as_deref(),
                                    None | Some("auto")
                                )
                            {
                                // S808 (2026-07-12): Latin footnote lines are the
                                // hhea natural (TNR10 11.499 - the S779/S805 line),
                                // not the estimate's word_line_height_table_cell
                                // (10.5) - uklocal fn13 est 21.0 vs Word 23.0.
                                // S810: auto-rule lines only — an exact-rule fn
                                // (ukframework line=240) uses its declared box.
                                // S828(b): skip the SUPERSCRIPT ref-mark run (the
                                // prefix "1" lands in it, making it the first
                                // non-empty run; its auto-shrunk 2/3 fs gave
                                // nyserda ph=7.93 for an 11.5 line).
                                let s828b = std::env::var("OXI_S828_DISABLE").is_err();
                                let rs = p2
                                    .runs
                                    .iter()
                                    .find(|r| {
                                        !r.text.trim().is_empty()
                                            && !(s828b
                                                && matches!(
                                                    r.style.vertical_align,
                                                    Some(VerticalAlign::Superscript)
                                                        | Some(VerticalAlign::Subscript)
                                                ))
                                    })
                                    .or_else(|| p2.runs.iter().find(|r| !r.text.trim().is_empty()))
                                    .map(|r| &r.style)
                                    .cloned()
                                    .unwrap_or_default();
                                let fs = self.resolve_font_size(&rs, &p2.style);
                                let m = self.metrics_for(&rs, &p2.style);
                                let per_line = m.word_line_height_table_cell(fs).max(1.0);
                                let lines = (line_height / per_line).round().max(1.0);
                                lines * m.natural_line_height_hhea(fs) + paragraph_spacing
                            } else {
                                ph_nat
                            };
                            if std::env::var("OXI_FN_PROBE").is_ok() {
                                eprintln!(
                                    "[FN_EST] id={} snap={} ph_nat={:.2} ph={:.2}",
                                    id, p2.style.snap_to_grid, ph_nat, ph
                                );
                            }
                            // 2026-05-05 Track A continuation: removed +2.0pt
                            // per-fn marker overhead. Empirically (b837 spill data
                            // 25 fns) Oxi's est = Word actual + exactly 2.0pt for
                            // every fn — the marker renders inline, no extra
                            // line-height. Over-reservation by 10pt per page (5
                            // fns × 2pt) prevented para 70 from fitting on p5.
                            h += ph;
                            // S807 (2026-07-12, opt-out OXI_S807_DISABLE): the
                            // note's FIRST line (the superscript ref-mark line)
                            // renders TALLER — the vertAlign run keeps its
                            // declared-fs line box RAISED by the superscript
                            // offset, so ref_line_h = plain + raise. DERIVED
                            // (_fn_refline_gen.py, 3 fonts x 4 sizes, Word PDF
                            // mark-span raise vs baseline pitch): growth ==
                            // the measured raise exactly (TNR10 +3.5 = the
                            // uklocalspending footnote; values are half-point
                            // pre-quantized: TNR {3,3.5,4,4.5} Arial {3,3,4,4}
                            // Calibri {2.5,3.5,4,4} @9-12pt). v1 raise =
                            // halfround(0.35*fs) — exact for TNR 9-11, Arial
                            // 9/11/12, Calibri 10-12; the +-0.5 residuals
                            // (TNR12, Arial10, Calibri9) await the font-metric
                            // rule. Latin scope; JP fn stack stays calibrated.
                            // S810: an EXACT-rule fn para's box clamps the
                            // superscript raise (Word ukframework pitch 15.0
                            // exact, no growth) — the raise applies to auto/
                            // atLeast lines only.
                            // S807 RETIRED to opt-in 2026-07-14 (OXI_S807=1):
                            // under the S833 declared-separator model the
                            // ref-line raise is a DOUBLE-COUNT (the fnr probe
                            // boxes show no raise term); uklocal natural
                            // 1.0000 requires it off. Kept as a knob for the
                            // legacy (S833-off) comparison state.
                            if !self.doc_body_has_real_cjk
                                && std::env::var("OXI_S807").is_ok()
                                && p2.style.line_spacing_rule.as_deref() != Some("exact")
                            {
                                let rs = p2
                                    .runs
                                    .iter()
                                    .find(|r| !r.text.trim().is_empty())
                                    .map(|r| &r.style)
                                    .cloned()
                                    .unwrap_or_default();
                                let fs = self.resolve_font_size(&rs, &p2.style);
                                h += (0.35 * fs * 2.0).round() / 2.0;
                            }
                            first_para = false;
                        } else {
                            let ph_nat =
                                self.estimate_para_height(p, cw, None, None, false, None, None);
                            // Paragraph spacing is not a count of text lines.
                            let mut line_para = p.clone();
                            line_para.style.space_before = Some(0.0);
                            line_para.style.space_after = Some(0.0);
                            line_para.style.before_lines = None;
                            line_para.style.after_lines = None;
                            let line_height = self.estimate_para_height(
                                &line_para, cw, None, None, false, None, None);
                            let paragraph_spacing = ph_nat - line_height;
                            h += if let (Some(pitch), true) = (fn_est_gp, p.style.snap_to_grid) {
                                let rs = p
                                    .runs
                                    .iter()
                                    .find(|r| !r.text.trim().is_empty())
                                    .map(|r| &r.style)
                                    .cloned()
                                    .unwrap_or_default();
                                let fs = self.resolve_font_size(&rs, &p.style);
                                let m = self.metrics_for(&rs, &p.style);
                                let per_line = m.word_line_height_table_cell(fs).max(1.0);
                                let lines = (line_height / per_line).round().max(1.0);
                                let cells =
                                    (m.word_line_height_no_grid(fs) / pitch).ceil().max(1.0);
                                lines * cells * pitch + paragraph_spacing
                            } else if !self.doc_body_has_real_cjk
                                && std::env::var("OXI_S808_DISABLE").is_err()
                                && matches!(
                                    p.style.line_spacing_rule.as_deref(),
                                    None | Some("auto")
                                )
                            {
                                // S808: Latin fn lines = hhea natural (see above).
                                // S810: auto-rule lines only.
                                let rs = p
                                    .runs
                                    .iter()
                                    .find(|r| !r.text.trim().is_empty())
                                    .map(|r| &r.style)
                                    .cloned()
                                    .unwrap_or_default();
                                let fs = self.resolve_font_size(&rs, &p.style);
                                let m = self.metrics_for(&rs, &p.style);
                                let per_line = m.word_line_height_table_cell(fs).max(1.0);
                                let lines = (line_height / per_line).round().max(1.0);
                                lines * m.natural_line_height_hhea(fs) + paragraph_spacing
                            } else {
                                ph_nat
                            };
                        }
                    }
                }
                if s804 {
                    if let Some(last) = s804_prev_sa {
                        h += last;
                    }
                }
                h
            } else {
                0.0
            }
        };

        // S596b (2026-06-21): footnote separator reservation.
        // Word renders the footnote separator as a paragraph that occupies a
        // full footnote-text line (~18pt for 10pt Yu Gothic), but Oxi reserved
        // only 6pt (separator line 2pt + padding 4pt) + OXI_FN_SEP_GAP_EXTRA
        // (6pt). For NO-DOCGRID docs this ~12pt under-reservation lets one extra
        // body line over-pack the page bottom on every footnote page
        // (bunkacontract: the first para of pages 3/5/7 was packed onto the
        // previous page = -1 x3). Reserving one footnote line for the separator
        // pushes those breaks to match Word. GRID footnote docs (b837 grid300,
        // kojin linesAndChars, etc.) snap body+footnote to the grid and the
        // small 12pt base is correct there (S160 calibration: sep_extra>=10
        // regressed b837); bunkacontract is the corpus's ONLY no-docGrid
        // footnote doc, so gating on grid_pitch.is_none() scopes this to it.
        // Default ON, opt-out OXI_S596B_DISABLE.
        let special_footnote_height = |num: u32| -> f32 {
                    page.footnotes
                        .iter()
                        .find(|n| n.number == num)
                        .and_then(|note| {
                            note.blocks.iter().find_map(|b| match b {
                                Block::Paragraph(p) => Some(p),
                                _ => None,
                            })
                        })
                        .map(|p| {
                            let _fng = FnLayoutGuard::new();
                            let mut h =
                                self.estimate_para_height(p, page.size.width - page.margin.left - page.margin.right, None, None, false, None, None);
                            // S900b (2026-07-17): a TEXT-EMPTY special para's line
                            // resolves through the DEFAULT PARAGRAPH STYLE's run
                            // props — Word sizes the separator by Normal (81e80:
                            // Arial 12 → 13.8; the est's ¶-mark fallback gave the
                            // docDefaults theme Calibri 11 = 12.649, one note-slot
                            // short at the area cutoff). uklocal/framework:
                            // Normal == docDefaults → value-identical.
                            if p.runs.iter().all(|r| r.text.trim().is_empty())
                                && !(p.style.line_spacing_rule.as_deref() == Some("exact")
                                    && std::env::var("OXI_EXACT_FN_SEPARATOR_DISABLE").is_err())
                            {
                                if let Some(drs) = p.style.default_run_style.as_ref() {
                                    if let Some(fs) = drs.font_size {
                                        let m = self.metrics_for(drs, &p.style);
                                        let line = m.natural_line_height_hhea(fs);
                                        if line > 0.0 {
                                            h = line;
                                        }
                                    }
                                }
                            }
                            // S804 convention: the estimate drops STYLE-level
                            // spacing; add it back unless the para carries
                            // direct spacing (the real notice's explicit 0/0).
                            if !p.style.has_direct_spacing {
                                h += p.style.space_before.unwrap_or(0.0)
                                    + p.style.space_after.unwrap_or(0.0);
                            }
                            h
                        })
                        .unwrap_or(0.0)
                };
        let legacy_notice_height = || -> f32 {
            if self.compat_mode < 15 && self.fn_special_declared
                && !self.doc_body_has_real_cjk
                && std::env::var("OXI_S833_DISABLE").is_err()
                && special_footnote_height(u32::MAX) > 0.0 {
                special_footnote_height(u32::MAX - 1)
            } else { 0.0 }
        };
        let footnote_sep_alloc = |first_id: u32| -> f32 {
            let sep_extra: f32 = std::env::var("OXI_FN_SEP_GAP_EXTRA")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(6.0);
            let base = 6.0 + sep_extra;
            // S833 (2026-07-13, opt-out OXI_S833_DISABLE): DECLARED custom
            // separator model. When settings.xml footnotePr declares the
            // special footnotes, Word reserves the separator region as the
            // CUSTOM separator paragraph at its FULL styled height (line +
            // style-chain sb/sa — uklocal: Normal-inherited 12 + 12.649 + 12
            // = 36.65 vs the built-in compact ~13) PLUS the continuationNotice
            // paragraph's styled height even on non-continuing pages (real
            // uklocal notice: direct spacing 0/0 -> line only, 12.65).
            // DERIVED via the _pb_fnres probes (P1-P4 decomposition; the
            // arithmetic closes on the real doc: 59.6 + 22.9 + 12.4 = 94.9 vs
            // the measured ~94.6; continuationSeparator declaration = 0 on
            // non-continuing pages). Latin scope: the JP footnotePr docs
            // (b837/kojin/bunkacontract) keep their calibrated base (their
            // special paras inherit spacing-less JP Normals, where the models
            // nearly coincide; reconciliation is a separate gated session).
            // ★SHIPPED default-ON 2026-07-14 (opt-out OXI_S833_DISABLE) as
            // part of the EN natural-flow endgame: the S559 exposures that
            // held it opt-in (ukframework {+1:1} / uklocalspending {+1:7} in
            // LRPB mode) are resolved by S835 (fn-boundary fs/16 relief) +
            // S836 (Latin saved-LRPB retirement) — the bundle state measures
            // 6/6 = 1.0000 on the EN gate.
            if self.fn_special_declared
                && !self.doc_body_has_real_cjk
                && std::env::var("OXI_S833_DISABLE").is_err()
            {
                let sep_h = special_footnote_height(u32::MAX);
                if std::env::var("OXI_FN_PROBE").is_ok() {
                    eprintln!("[FN_SEP] branch=s833 declared={} sep_h={:.3} notice={:.3}",
                        self.fn_special_declared, sep_h, special_footnote_height(u32::MAX - 1));
                }
                if sep_h > 0.0 {
                    // Experiment knob (default 0 = unchanged). The `_pb_fnkeep`
                    // sweep puts Oxi's keep boundary a flat 16tw = 0.8pt below
                    // Word's on this path, at both 3 and 4 notes — a constant,
                    // not a per-note error. OXI_SEPX sweeps that constant.
                    let sepx: f32 = std::env::var("OXI_SEPX")
                        .ok()
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0.0);
                    // This allocation precedes ordinary footnote placement.
                    // A continuation notice belongs to a split footnote, not
                    // every page containing a footnote. Its paragraph height
                    // must not reduce the ordinary body's available space.
                    return sep_h + legacy_notice_height() + sepx;
                }
            }
            // S1247 (opt-in OXI_S1247=1): S596b's scope was "no docGrid at all",
            // written when bunkacontract was the corpus's only such footnote doc.
            // The EN corpus writes <w:docGrid w:linePitch="N"/> with NO w:type,
            // and `_pb_fnkeep` re-run WITH that grid flips at the SAME spacer
            // (600 keeps / 604 breaks, both with and without) -- the separator
            // reservation is grid-independent, so an untyped grid must not fall
            // back to the legacy 12pt base. Latin scope, as derived.
            let s1247_no_type = std::env::var("OXI_S1247").ok().as_deref() == Some("1")
                && page.doc_grid_no_type
                && !self.doc_body_has_real_cjk;
            if (page.grid_line_pitch.is_none() || s1247_no_type)
                && std::env::var("OXI_S596B_DISABLE").is_err()
            {
                if let Some(note) = page.footnotes.iter().find(|n| n.number == first_id) {
                    if let Some(Block::Paragraph(p)) = note.blocks.first() {
                        let fs = self.resolve_font_size(
                            p.runs
                                .first()
                                .map(|r| &r.style)
                                .unwrap_or(&RunStyle::default()),
                            &p.style,
                        );
                        let metrics = p
                            .runs
                            .first()
                            .map(|r| self.metrics_for_text(&r.text, &r.style, &p.style))
                            .unwrap_or_else(|| {
                                let rpr = p.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                self.metrics_for_para_mark(&rpr, &p.style)
                            });
                        // Separator paragraph = one footnote text line. Never
                        // reserve LESS than the legacy base.
                        // S804: the separator region also carries the collapsed
                        // style spacing on both sides (body para after -> sep para
                        // -> first note before; probe: 12 + sep line + 12). Use
                        // 2 x the first note's style before as the two collapsed
                        // gaps (uniform-Normal docs; 0 when the style has none).
                        let s804_sep = if std::env::var("OXI_S804_DISABLE").is_err()
                            && !p.style.has_direct_spacing
                        {
                            2.0 * p.style.space_before.unwrap_or(0.0)
                        } else {
                            0.0
                        };
                        let v = metrics.word_line_height_no_grid(fs).max(base) + s804_sep;
                        if std::env::var("OXI_FN_PROBE").is_ok() {
                            eprintln!("[FN_SEP] branch=s596b first_note_fs={:.2} line={:.3} base={:.3} s804={:.3} -> {:.3}",
                                fs, metrics.word_line_height_no_grid(fs), base, s804_sep, v);
                        }
                        return v;
                    }
                }
            }
            // S804: same spacing addition for the grid-doc path (JP footnote
            // styles carry no spacing -> +0, byte-identical).
            let s804_sep = if std::env::var("OXI_S804_DISABLE").is_err() {
                page.footnotes
                    .iter()
                    .find(|n| n.number == first_id)
                    .and_then(|note| note.blocks.first())
                    .and_then(|b| match b {
                        Block::Paragraph(p) if !p.style.has_direct_spacing => {
                            Some(2.0 * p.style.space_before.unwrap_or(0.0))
                        }
                        _ => None,
                    })
                    .unwrap_or(0.0)
            } else {
                0.0
            };
            if std::env::var("OXI_FN_PROBE").is_ok() {
                eprintln!("[FN_SEP] branch=base base={:.3} s804={:.3} -> {:.3}",
                    base, s804_sep, base + s804_sep);
            }
            base + s804_sep
        };

        // Multi-column layout: compute column X positions and widths
        // COM-confirmed: col_x = margin + Σ(prev_width + prev_spacing)
        // S560 (2026-06-13): factored into a closure so per-section column
        // layouts (page.column_runs, populated when `continuous` section breaks
        // merge sections with DIFFERENT column counts) can be recomputed at
        // each section boundary inside the block loop below.
        let margin_left = page.margin.left;
        // S729: compute_cols takes the run's HORIZONTAL margins (ml + text
        // width) so per-section margin runs (merged continuous sections with
        // different left/right margins) produce the correct geometry. The
        // page-level call passes the page margins — byte-identical.
        let compute_cols = |cols: &Option<crate::ir::ColumnLayout>,
                            ml: f32,
                            tcw: f32|
         -> (usize, Vec<f32>, Vec<f32>) {
            let num = cols.as_ref().map(|c| c.num.max(1) as usize).unwrap_or(1);
            let mut xs: Vec<f32> = Vec::with_capacity(num);
            let mut ws: Vec<f32> = Vec::with_capacity(num);
            if num > 1 {
                if let Some(ref c) = cols {
                    if !c.columns.is_empty() {
                        // Unequal width columns: use explicit definitions
                        let mut x = ml;
                        for col_def in &c.columns {
                            xs.push(x);
                            ws.push(col_def.width);
                            x += col_def.width + col_def.space.unwrap_or(0.0);
                        }
                    } else {
                        // Equal width columns
                        let spacing = c.space.unwrap_or(36.0); // default 36pt
                        let col_w = (tcw - spacing * (num - 1) as f32) / num as f32;
                        let mut x = ml;
                        for _ in 0..num {
                            xs.push(x);
                            ws.push(col_w);
                            x += col_w + spacing;
                        }
                    }
                }
            }
            if xs.is_empty() {
                xs.push(ml);
                ws.push(tcw);
            }
            // Bidi (RTL) section: columns flow RIGHT-to-LEFT, so the first
            // reading column (fill order index 0) is the RIGHTMOST one.
            // Reversing both xs and ws keeps each column's (x, width) pairing
            // while mapping fill-order 0 -> rightmost physical column.
            // Word-confirmed (minimal repro + albalunaTaidan). Opt-out via
            // OXI_BIDICOL_DISABLE.
            if page.bidi_columns && xs.len() > 1 && std::env::var("OXI_BIDICOL_DISABLE").is_err() {
                xs.reverse();
                ws.reverse();
            }
            (xs.len(), xs, ws)
        };

        let (mut num_columns, mut col_x_positions, mut col_widths) =
            compute_cols(&page.columns, margin_left, total_content_width);

        // S560: per-section column runs. Switch per-section ONLY when the
        // merged page has HETEROGENEOUS column counts (e.g. kyotei36spec: a
        // 1-col form table + a continuous 2-col 記載心得 instruction block).
        // When all runs share one column count (the entire 269-doc baseline is
        // num=1), `heterogeneous` is false and the loop never switches → the
        // pre-S560 single-layout path runs byte-identically.
        let col_runs: Vec<(usize, usize, Vec<f32>, Vec<f32>)> = page
            .column_runs
            .iter()
            .enumerate()
            .map(|(i, (start, cols))| {
                // S729: use the run's own margins when a parallel margin run
                // exists (index-aligned with column_runs — both are seeded and
                // pushed together in the parser); fall back to page margins.
                let (ml, mr) = page
                    .margin_runs
                    .get(i)
                    .filter(|(ms, _, _)| ms == start)
                    .map(|(_, l, r)| (*l, *r))
                    .unwrap_or((page.margin.left, page.margin.right));
                let tcw = page.size.width - ml - mr;
                let (n, xs, ws) = compute_cols(cols, ml, tcw);
                (*start, n, xs, ws)
            })
            .collect();
        let heterogeneous = {
            let col_het = {
                let mut it = col_runs.iter().map(|r| r.1);
                match it.next() {
                    Some(first) => it.any(|n| n != first),
                    None => false,
                }
            };
            // S729: margin heterogeneity also demands per-run switching (a
            // continuous section with different left/right margins must lay
            // out at its own text width — probexmargins). Uniform-margin
            // corpora keep col_het semantics byte-identically.
            let mar_het = std::env::var("OXI_S729_DISABLE").is_err() && {
                let mut it = page.margin_runs.iter().map(|(_, l, r)| (*l, *r));
                match it.next() {
                    Some((l0, r0)) => {
                        it.any(|(l, r)| (l - l0).abs() > 0.01 || (r - r0).abs() > 0.01)
                    }
                    None => false,
                }
            };
            col_het || mar_het
        };
        if heterogeneous {
            // Base the page on the FIRST run's column layout; subsequent runs
            // switch in at their block boundaries.
            if let Some((_, n, xs, ws)) = col_runs.first() {
                num_columns = *n;
                col_x_positions = xs.clone();
                col_widths = ws.clone();
            }
        }
        let mut active_run_idx: usize = 0;
        let mut allocation_start = col_runs.first().map_or(0, |r| r.0);
        if std::env::var("OXI_DBG_COL").is_ok() {
            let runs: Vec<(usize, usize)> = col_runs.iter().map(|r| (r.0, r.1)).collect();
            eprintln!(
                "[COL] heterogeneous={} col_runs(start,ncol)={:?} blocks={}",
                heterogeneous,
                runs,
                page.blocks.len()
            );
        }

        let mut current_column: usize = 0;
        let mut start_x = col_x_positions[0];
        let mut content_width = col_widths[0];
        // S560: lowest column-bottom reached on the current page, so a
        // following column-section flows below ALL columns of the one it
        // succeeds. Only read on the heterogeneous (per-section column) path.
        let mut section_max_y = start_y;
        let mut section_prev_page = 0usize;
        // S638 (kyotei): vertAnchor="text" full-page float — the body flows in
        // the GAP above the float, then SKIPS the float's region. Records
        // (top, bottom, page) of the active float; a following block whose
        // cursor lands inside [top, bottom) on that page is bumped to bottom.
        // S1241 (2026-08-27): the region also records the float's X range
        // (fx0, fx1) so a flow in a DIFFERENT column lane (forms__000cf39c:
        // float in col1, inline table flowing in col2) is not bumped below it.
        // S1509: the sixth element is the free side lane (pt, net of the
        // tblpPr wrap distances) beside the float.
        let mut text_float_region: Option<(f32, f32, usize, f32, f32, f32)> = None;
        // S1195 (2026-08-22, default ON, opt-out `OXI_S1195_DISABLE`): a
        // wrap-below float that still leaves a usable side LANE. Word flows the
        // EMPTY paragraphs that follow such a float in that lane, beside the
        // table, and only drops real content below it — so those empties must
        // not spend a line under the float.
        // Records (below_y, page) of the float whose lane is still open.
        let mut float_lane_below: Option<(f32, usize)> = None;

        let mut grid_pitch = page.grid_line_pitch;
        // S735 (2026-07-03): per-section grid-pitch runs on a merged continuous
        // page. Only active when the runs actually DIFFER (grid_het) — the
        // uniform-pitch corpus is byte-identical. parttime (403/415) is the one
        // corpus doc with real pitch diversity. Opt-out OXI_S735_DISABLE.
        let s735_grid_het = std::env::var("OXI_S735_DISABLE").is_err() && {
            let mut it = page.grid_runs.iter().map(|(_, p)| *p);
            match it.next() {
                Some(first) => it.any(|p| match (p, first) {
                    (Some(a), Some(b)) => (a - b).abs() > 0.01,
                    (None, None) => false,
                    _ => true,
                }),
                None => false,
            }
        };
        let mut s735_run_idx: usize = 0;
        // S1336 (2026-09-06, HELD OPT-IN `OXI_S1336=1`; see the archive -- the
        // per-section pitch is Word's, but alone it reads -0.0008 on 0ea3ec86 and
        // 167853 because the mid-line 、） compression Oxi still applies at
        // compat 14 (Word: none) was compensating): a merged
        // continuous section whose w:charSpace differs from the page's first
        // section lays its blocks out against a Page variant carrying ITS
        // character pitch (raw pitch = default size + charSpace/4096, the S466
        // convention; the default size is recovered from the page's own pitch).
        // reference__0ea3ec86: sections alternate charSpace 3194 / 2048; Word
        // walks the 3194 sections at 11.76-11.78 per character (PDF p18: 20
        // characters per 235.6pt column line), Oxi at the first section's 11.5
        // and packed 21-22 -- the five residual -1 paragraphs of that document.
        let s1336_variants: Vec<(usize, Option<Page>)> = if std::env::var("OXI_S1336").ok().as_deref() == Some("1")
            && std::env::var("OXI_S466_DISABLE").is_err()
            && page.grid_char_runs.len() > 1
        {
            match page.grid_char_pitch {
                Some(pitch) => {
                    let cs_pt = |raw: Option<i32>| raw.map(|c| c as f32 / 4096.0).unwrap_or(0.0);
                    let base_fs = pitch - cs_pt(page.grid_char_space_raw);
                    page.grid_char_runs
                        .iter()
                        .map(|&(start, raw)| {
                            if raw == page.grid_char_space_raw {
                                (start, None)
                            } else {
                                let mut v = page.clone();
                                let p = base_fs + cs_pt(raw);
                                v.grid_char_pitch = Some(p);
                                if base_fs > 0.0 {
                                    v.grid_char_cw_ratio = Some(p / base_fs);
                                }
                                v.grid_char_space_raw = raw;
                                (start, Some(v))
                            }
                        })
                        .collect()
                }
                None => Vec::new(),
            }
        } else {
            Vec::new()
        };
        let page_orig: &Page = page;
        let mut mult_cumul_raw: f32 = 0.0;
        let mut pages: Vec<LayoutPage> = Vec::new();
        let mut elements: Vec<LayoutElement> = Vec::new();
        let mut cursor = LayoutCursor::new(start_y);
        let mut lm2_cells: usize = 0;
        let mut prev_para_style_id: Option<String> = None;
        let mut prev_contextual_spacing: bool = false;
        // S931: previous paragraph's numId when it carried afterAutospacing.
        let mut prev_autospacing_numid: Option<String> = None;
        // S658: the previous body paragraph's pBdr, for the border-merge gate.
        let mut prev_borders: Option<ParagraphBorders> = None;
        // S739: the previous body paragraph has keepNext (the follower of a
        // keepNext para gets the LENIENT natural page-bottom test, see the
        // centered-box rule in layout_paragraph).
        let mut prev_keep_next: bool = false;
        // S749 (2026-07-05): the vertical TOP of the current multi-column BAND.
        // A continuous multi-col section starting MID-PAGE flows its columns
        // from the section boundary (the SWITCH cursor), not the page top —
        // column 2 previously reset to page_top and gained the whole upper
        // page (probexcont2col: col2 filled from y=71 while col1's band began
        // at 503 → +432pt phantom capacity → -1x12). Reset to start_y on page
        // transitions (the band continues at the page top on later pages).
        let mut col_band_top: f32 = start_y;
        let mut prev_space_after: f32 = 0.0;
        // A continuous section marker has no line box. Its predecessor's
        // after-spacing still belongs to the ending section and must survive
        // column balancing before the next section establishes its origin.
        let mut pending_section_gap: f32 = 0.0;
        // Track Y position and layout page index for each block (for paragraph-relative TextBox positioning)
        let mut block_y_positions: Vec<f32> = Vec::with_capacity(page.blocks.len());
        // S1222 (2026-08-26): the COLUMN each block flows in, for anchored objects
        // whose positionH says relativeFrom="column" -- that means the column of
        // the ANCHOR PARAGRAPH, not the margin. forms__000cf39c346c1e59's bottom
        // box (7.35in wide, H column -3.93in, anchored in the RIGHT column of a
        // two-column section) rendered 3.5in off the left page edge without this.
        let mut block_col_x: Vec<f32> = Vec::with_capacity(page.blocks.len());
        // S1123 (2026-08-14): the page a block STARTED on, never overwritten.
        // `block_page_indices` serves two different consumers with two
        // different meanings: footnote bookkeeping wants the block's FINAL
        // page (the ow@7969 overwrite exists for it — its comment says "if
        // the paragraph moved entirely" but the code is unconditional), while
        // anchored-float resolution wants the ANCHOR page = where the
        // paragraph BEGINS. administrative__000727a4: paragraph 0 opens with
        // a page-break char (), its mark stays on p1, the overwrite set
        // its entry to p2, and all 18 anchored shapes rendered one page late
        // (Word draws them on p1 at start+offset). Keep both meanings in
        // separate arrays instead of re-conditioning the overwrite (which
        // the footnote path relies on).
        let mut block_start_page_indices: Vec<usize> = Vec::with_capacity(page.blocks.len());
        let mut block_page_indices: Vec<usize> = Vec::with_capacity(page.blocks.len());
        let mut current_page_idx: usize = 0;
        // S469 (2026-06-01): wrap-below floating tables (vertAnchor="text",
        // tblpX=0, full-width — see R7.75/R7.76) advance the FLOW cursor below
        // the table so body TEXT wraps under it (Word-confirmed, session 60).
        // BUT floating objects (textboxes/images) anchored to a paragraph that
        // follows such a table use that paragraph's NATURAL (pre-wrap) flow
        // position, NOT the wrapped cursor. 1ec1 root cause: its bottom note
        // box + 国税庁 logo + badge are all anchored to the (text-less) para
        // after a wrap-below floating table; Oxi recorded their anchor Y as the
        // wrapped cursor (~749pt) → anchor_y + posOffset overflowed the page →
        // clamp → objects ~46pt too low (note/logo overlap). Word anchors them
        // at the natural Y (~571pt). Fix: accumulate the wrap-below advance and
        // subtract it when recording block_y_positions (used ONLY for anchor
        // resolution); the flow cursor / pagination are untouched (Phase-1
        // safe). Reset per page. Default ON, opt-out OXI_S469_DISABLE.
        let s469_enabled = std::env::var("OXI_S469_DISABLE").is_err();
        let mut anchor_flow_offset: f32 = 0.0;
        let mut anchor_offset_page: usize = 0;
        // Round 29: dynamic per-page footnote reservation. Tracks the sum of
        // estimated heights for footnotes whose references appear on the current
        // layout page. Resets to 0 each time a new page is pushed. Subtracts from
        // effective content_height in overflow checks below.
        let mut footnote_reserve_current: f32 = 0.0;
        let mut footnote_ids_current_page: Vec<u32> = Vec::new();
        // Step 1 partial (Option B, 2026-04-22): per-page accumulation of
        // fn_ref ids whose markers actually render on each page, built from
        // layout_paragraph's per-line attribution (Step 0, e347cdf). Replaces
        // the block_page_indices-based collect_footnote_refs call in fn_area
        // render — block-level attribution mis-assigns mid-break refs (e.g.
        // b837 block 48 ref 15 marker on p3 but block_page_indices=p4).
        // NOTE: reserve seeding on NEW page (Step 1 full) was FALSIFIED on
        // b837 (-0.0828 net) due to body cascade; see
        // project_fn_reserve_option_b_step1_FALSIFIED.md.
        let mut page_fn_refs: Vec<Vec<u32>> = Vec::new();
        // S900: notes DEFERRED to a later page's area: (target_page, id, est
        // height incl the separator for the page's first entry). Folded into
        // footnote_reserve_current at the page-transition resets.
        let mut s900_pending_deferred: Vec<(usize, u32, f32)> = Vec::new();
        let s900_fold = |reserve: &mut f32,
                         ids: &mut Vec<u32>,
                         pending: &mut Vec<(usize, u32, f32)>,
                         pg: usize| {
            pending.retain(|&(p, id, h)| {
                if p <= pg {
                    *reserve += h;
                    ids.push(id);
                    false
                } else {
                    true
                }
            });
        };
        // S740: block indices of tables whose cell-footnote ids were attributed
        // per-page via layout_table (skip the coarse block-level attribution).
        let mut s740_attributed_tables: std::collections::HashSet<usize> =
            std::collections::HashSet::new();

        // R7.60 (Day 36 part 6, 2026-05-14): track floating-table Y ranges per page
        // for body-position (vertAnchor="page", tblpY > top_margin) full-width tables.
        // When a new floating table would Y-overlap with an already-placed one on the
        // current page, push it to the next page. Word's behavior for 459f05: both
        // floating tables span content width, so they cannot share a page; second
        // table's anchor effectively rolls to next page.
        // Header-position floats (tblpY <= top_margin, e.g. 1ec1/2ea81a) and
        // vertAnchor="text" floats (e.g. 3a4f9f/ed025c) are NOT tracked — they
        // co-locate with body content normally.
        let mut previous_table_probe: Option<(usize, f32, usize, f32, f32, Option<S755Geom>)> = None;
        let mut previous_table_probe_elements: Vec<LayoutElement> = Vec::new();
        let mut floating_tables_per_page: Vec<Vec<(f32, f32)>> = vec![Vec::new()];

        if std::env::var("OXI_DBG_BLOCKS").is_ok() {
            for (bi, block) in page.blocks.iter().enumerate() {
                let (kind, txt) = match block {
                    Block::Paragraph(p) => (
                        "P",
                        p.runs
                            .iter()
                            .flat_map(|r| r.text.chars())
                            .take(16)
                            .collect::<String>(),
                    ),
                    Block::Table(t) => (
                        "T",
                        t.rows
                            .first()
                            .and_then(|r| r.cells.first())
                            .and_then(|c| c.blocks.first())
                            .and_then(|b| {
                                if let Block::Paragraph(p) = b {
                                    Some(
                                        p.runs
                                            .iter()
                                            .flat_map(|r| r.text.chars())
                                            .take(16)
                                            .collect::<String>(),
                                    )
                                } else {
                                    None
                                }
                            })
                            .unwrap_or_default(),
                    ),
                    _ => ("?", String::new()),
                };
                eprintln!("[BLOCKS] {} {} {:?}", bi, kind, txt);
            }
        }

        // S676 (2026-06-27): pending drop-cap float. When a paragraph with
        // framePr dropCap="drop" is seen, its glyph is rendered at the left
        // (anchored to the NEXT paragraph's top) and the next paragraph's body
        // is indented by `Some(indent_pt)` so it wraps to the right of the cap.
        let mut pending_dropcap: Option<f32> = None;
        // S734 (2026-07-03): wrapTopAndBottom floating images RESERVE their
        // vertical band in the body flow (Word: text may not sit beside them;
        // the anchor paragraph and everything after flow BELOW the image).
        // wrap_type was parsed but never consumed — floats painted at absolute
        // positions with ZERO flow effect (probezwraptb {-1:8}: text flowed
        // straight through two 126pt images). Paragraph-relative anchors only
        // (the measured case); corpus has ZERO wrapTopAndBottom docs → the
        // band map is empty → byte-identical. Opt-out OXI_S734_DISABLE.
        let s734_bands: std::collections::HashMap<usize, f32> =
            if std::env::var("OXI_S734_DISABLE").is_err() {
                page.floating_images
                    .iter()
                    .filter(|img| {
                        img.wrap_type == Some(crate::ir::WrapType::TopAndBottom)
                            && img
                                .position
                                .as_ref()
                                .map_or(false, |p| p.v_relative.as_deref() == Some("paragraph"))
                    })
                    .map(|img| {
                        // S1513 (2026-09-21, default ON, opt-out OXI_S1513_DISABLE): the
                        // band reaches down to the picture's effectExtent b (its
                        // shadow/outline margin), not only its extent. educational__
                        // 004c4a3d: Figure 1 b=18415 EMU (1.45pt) is exactly Oxi's -1.45
                        // on the caption below; Figure 2 (b=1.65) then ended 769.2 vs
                        // a 769.9 text bottom in Oxi and 770.9 in Word, which pushed
                        // the host to the next page (tb_after_probe2.py: the host
                        // moves the moment the band bottom exceeds the text bottom).
                        let s1513_b = if std::env::var_os("OXI_S1513_DISABLE").is_none() { img.effect_extent_b.max(0.0) } else { 0.0 };
                        (
                            img.anchor_block_index,
                            img.height + s1513_b + img.position.as_ref().map_or(0.0, |p| p.y.max(0.0)),
                        )
                    })
                    .collect()
            } else {
                Default::default()
            };
        // anchor_block_index → (layout page, band top y): where the band was
        // actually reserved; the paint pass places the image THERE (the anchor
        // paragraph now sits BELOW the image, so resolving from the paragraph's
        // y would double-shift).
        let mut s734_flow_pos: std::collections::HashMap<usize, (usize, f32)> = Default::default();
        // S1500 (2026-09-20, default ON, opt-out OXI_S1500_DISABLE): the block
        // right after a wrapTopAndBottom band host anchors its own
        // paragraph-relative shapes to its UNPUSHED top -- the y it had before
        // the host's band dropped the cursor -- while its lines sit below the
        // band. MEASURED (box7_probe v2/v3, c5bb00 p8 slice: para 91 hosts
        // four bands, para 92 hosts a 196pt box at posOffset 134.6): Word
        // draws the box at 65 + 134.6 (para 91's line bottom + offset) though
        // para 92's text renders at 175; with one plain paragraph between,
        // the box sits at that paragraph's pushed bottom + offset (normal).
        // The document: box at 405 = 253.5 + 17 + 134.6, Oxi had 512.
        // (next block idx, page, unpushed y)
        let mut s1500_unpushed: Option<(usize, usize, f32)> = None;
        // S1497 (2026-09-20, default ON, opt-out OXI_S1497_DISABLE): a float
        // whose top sits BELOW its anchor paragraph's top (posOffset > 0) is a
        // band inside that paragraph: the lines that fit above it stay, the
        // first line whose box would cross it (and everything after) resumes
        // at the band bottom. Word (`tb_host_probe.py`, 200pt picture, 18pt
        // lines): off 18 keeps 1 line above, 45 keeps 2, 60 keeps 3, and off
        // 0 / 10 push the whole paragraph -- the S734 case, which is the same
        // rule with the first line cut. technical__c5bb0090235dfedb p1: the
        // 「下図は…」 line stays at 287.25 above its picture (off 20.35).
        let s1497_on = std::env::var_os("OXI_S1497_DISABLE").is_none();
        // S1497b: a wrapTopAndBottom float pushes only the lines of the
        // COLUMN(S) its rectangle crosses horizontally. educational__13ef8d6ec218af24
        // (two columns): a 175pt picture anchored in a column-1 paragraph but
        // placed at posH 277 (x 334-509, column 1 ends at 286) leaves column 1
        // untouched in Word -- `bisect13ef.py`: posH <= 220 pushes, >= 240 does
        // not; VML boxes, compat, host style, offset, WMF/PNG and layoutInCell
        // change nothing. A single-column page always overlaps.
        let s1497b_img_overlaps = |block_idx: usize, byp: &[f32], cy: f32, sx: f32, cw: f32| -> bool {
            if !s1497_on { return true; }
            page.floating_images.iter().any(|img| {
                img.anchor_block_index == block_idx
                    && img.wrap_type == Some(crate::ir::WrapType::TopAndBottom)
                    && img.position.as_ref().map_or(false, |p| p.v_relative.as_deref() == Some("paragraph"))
                    && {
                        let (fx, _) = self.resolve_floating_image_position(img, page, byp, cy);
                        fx < sx + cw - 0.5 && fx + img.width > sx + 0.5
                    }
            })
        };
        // S1552 (2026-09-25, default ON, opt-out OXI_S1552_DISABLE): a wrapSquare
        // text box that leaves no usable side lane (both lanes, net of the wrap
        // distance, below the S1031 floor: 41.5pt Latin / 100pt CJK) behaves as
        // a top-and-bottom band — the host paragraph moves to the next page when
        // the band does not fit (S734/S1513) and the flow resumes below it.
        // educational__0050e825: a 540pt-wide rubric box (positionV paragraph
        // +15.5, 362.7 tall) on a 468pt column; Word puts the host at the next
        // page's top (72.0), the box at 87.5..450.2 and the next paragraph at
        // 451.5; Oxi kept the host on the previous page and drew the box off
        // the bottom.
        let s1552_lane_min = if std::env::var("OXI_S1031_DISABLE").is_err()
            && !self.doc_body_has_real_cjk
        {
            41.5
        } else {
            100.0
        };
        let s1552_no_lane = |tb: &crate::ir::TextBox| -> bool {
            std::env::var_os("OXI_S1552_DISABLE").is_none()
                && tb.wrap_type == Some(crate::ir::WrapType::Square)
                && tb.position.as_ref().map_or(false, |tp| {
                    let col_l = page.margin.left;
                    let col_r = page.size.width - page.margin.right;
                    let (rl, rw) = match tp.h_relative.as_deref() {
                        Some("page") => (0.0, page.size.width),
                        _ => (col_l, col_r - col_l),
                    };
                    let x0 = match tp.h_align.as_deref() {
                        Some("center") => rl + (rw - tb.width) * 0.5,
                        Some("right") => rl + rw - tb.width,
                        Some("left") => rl,
                        _ => rl + tp.x,
                    };
                    let dl = tp.dist_l.unwrap_or(9.0);
                    let dr = tp.dist_r.unwrap_or(9.0);
                    let left_lane = x0 - dl - col_l;
                    let right_lane = col_r - (x0 + tb.width + dr);
                    left_lane.max(right_lane) < s1552_lane_min
                })
        };
        let s1497b_tb_overlaps = |block_idx: usize, byp: &[f32], bcx: &[f32], sx: f32, cw: f32| -> bool {
            if !s1497_on { return true; }
            page.text_boxes.iter().any(|tb| {
                tb.anchor_block_index == block_idx
                    && (tb.wrap_type == Some(crate::ir::WrapType::TopAndBottom) || s1552_no_lane(tb))
                    && tb.position.as_ref().map_or(false, |p| p.v_relative.as_deref() == Some("paragraph"))
                    && {
                        let (fx, _) = self.resolve_textbox_position(tb, page, byp, bcx);
                        fx < sx + cw - 0.5 && fx + tb.width > sx + 0.5
                    }
            })
        };
        let s1497_mid: std::collections::HashMap<usize, (f32, f32)> = if s1497_on {
            let mut m: std::collections::HashMap<usize, (f32, f32)> = Default::default();
            for (idx, off, h) in page
                .floating_images
                .iter()
                .filter(|img| img.wrap_type == Some(crate::ir::WrapType::TopAndBottom))
                .filter_map(|img| img.position.as_ref().filter(|p| p.v_relative.as_deref() == Some("paragraph")).map(|p| (img.anchor_block_index, p.y,
                    img.height + if std::env::var_os("OXI_S1513_DISABLE").is_none() { img.effect_extent_b.max(0.0) } else { 0.0 })))
                .chain(
                    page.text_boxes
                        .iter()
                        .filter(|tb| tb.wrap_type == Some(crate::ir::WrapType::TopAndBottom))
                        .filter_map(|tb| tb.position.as_ref().filter(|p| p.v_relative.as_deref() == Some("paragraph")).map(|p| (tb.anchor_block_index, p.y, tb.height))),
                )
            {
                if off > 0.0 {
                    let e = m.entry(idx).or_insert((off, h));
                    if off + h > e.0 + e.1 { *e = (off, h); }
                }
            }
            m
        } else {
            Default::default()
        };
        // S1089 (2026-08-07, opt-out OXI_S1089_DISABLE): S734 covers floating
        // IMAGES only — a wrapTopAndBottom float that is a wps SHAPE lands in
        // page.text_boxes (S839) and reserved NOTHING.  technical__002c6778's
        // title block anchors a 0.1pt Freeform rule (posOffset 1.25pt) in the
        // "Models: DC-6174" paragraph: Word draws the rule at 138.37 and starts
        // that paragraph's text at 145.83 = band bottom + the 7.3pt before, i.e.
        // exactly S734's model (band = pos.y + height, reserved BEFORE the
        // paragraph's own spacing).  Oxi drew the rule in the right place but
        // never advanced, so page 1 ran 1.57pt high and the keepNext pair at its
        // foot fit by 0.13pt where Word overflows by 1.43.
        // S1459 (2026-09-17, default ON, opt-out OXI_S1459_DISABLE): a
        // wrapTopAndBottom textbox anchored to the PAGE interrupts the flow at
        // its own absolute position -- the anchor paragraph stays ABOVE it and
        // the next block starts below its bottom. (The S734/S1089 bands model
        // the PARAGRAPH-anchored shape, which pushes the anchor itself down.)
        // golden parttime p1: the title sits at 71.25, the box spans the page's
        // 107.05..179.25, and Word starts 第１章 at 180.00. Oxi drew the box but
        // never advanced, so page 1 swallowed eleven extra paragraphs and the
        // document came out 6 pages against Word's 7.
        let s1459_page_boxes: std::collections::HashMap<usize, f32> =
            if std::env::var("OXI_S1459_DISABLE").is_err() {
                let mut m: std::collections::HashMap<usize, f32> = Default::default();
                for tb in &page.text_boxes {
                    if tb.wrap_type != Some(crate::ir::WrapType::TopAndBottom) {
                        continue;
                    }
                    if let Some(p) = tb.position.as_ref() {
                        if p.v_relative.as_deref() == Some("page") {
                            let e = m.entry(tb.anchor_block_index).or_insert(0.0_f32);
                            *e = e.max(s1486_wrap_bottom(p.y + tb.height, tb.stroke_width));
                        }
                    }
                }
                m
            } else {
                Default::default()
            };
        // S1461 (2026-09-17, default ON, opt-out OXI_S1461_DISABLE): a
        // wrapTopAndBottom shape attached to a body paragraph interrupts the
        // flow -- the anchor paragraph stays ABOVE it and the next block starts
        // below its bottom. (S734/S1089 model the band that pushes the anchor
        // itself down; this is the other shape.) golden parttime p1: the title
        // sits at 71.25, the box spans the page's 107.05..179.25 and Word starts
        // 第１章 at 180.00. Oxi drew the box and never advanced, so page 1
        // swallowed eleven extra paragraphs and the document came out 6 pages
        // against Word's 7. The written wrap kind lives in `anchor_wrap`.
        let s1461_tb_shapes: std::collections::HashMap<usize, f32> =
            if std::env::var("OXI_S1461_DISABLE").is_err() {
                let mut m: std::collections::HashMap<usize, f32> = Default::default();
                for (bi, b) in page.blocks.iter().enumerate() {
                    let Block::Paragraph(p) = b else { continue };
                    for sh in &p.shapes {
                        if sh.wrap_type.or(sh.anchor_wrap)
                            != Some(crate::ir::WrapType::TopAndBottom)
                        {
                            continue;
                        }
                        let Some(pos) = sh.position.as_ref() else { continue };
                        if pos.v_relative.as_deref() != Some("page") {
                            continue;
                        }
                        let e = m.entry(bi).or_insert(0.0_f32);
                        *e = e.max(s1486_wrap_bottom(pos.y + sh.height, sh.stroke_width));
                    }
                }
                m
            } else {
                Default::default()
            };
        let s1089_tb_bands: std::collections::HashMap<usize, f32> =
            if std::env::var("OXI_S1089_DISABLE").is_err() {
                let mut m: std::collections::HashMap<usize, f32> = Default::default();
                for tb in &page.text_boxes {
                    if tb.wrap_type != Some(crate::ir::WrapType::TopAndBottom) {
                        continue;
                    }
                    let off = match tb.position.as_ref() {
                        Some(p) if p.v_relative.as_deref() == Some("paragraph") => p.y.max(0.0),
                        _ => continue,
                    };
                    let e = m.entry(tb.anchor_block_index).or_insert(0.0_f32);
                    *e = e.max(off + tb.height);
                }
                m
            } else {
                Default::default()
            };
        // anchor_block_index → (page, band top): the paint keeps the band where
        // it was reserved (the anchor paragraph now sits BELOW it, so resolving
        // from the paragraph's y would double-shift) — the S734 contract.
        let mut s1089_flow_pos: std::collections::HashMap<usize, (usize, f32)> = Default::default();
        // S1552: push-only band (anchor block -> off + height) for wrapSquare
        // boxes without a usable lane. Only the host-push decision consults it;
        // the flow between the host and the box top is untouched
        // (policies__0beb595a: a 23pt box 84pt below its host, Word flows six
        // lines above it) and it never records a flow position.
        let s1552_bands: std::collections::HashMap<usize, f32> = {
            let mut m: std::collections::HashMap<usize, f32> = Default::default();
            for tb in &page.text_boxes {
                if !s1552_no_lane(tb) {
                    continue;
                }
                let off = match tb.position.as_ref() {
                    Some(p) if p.v_relative.as_deref() == Some("paragraph") => p.y.max(0.0),
                    _ => continue,
                };
                let e = m.entry(tb.anchor_block_index).or_insert(0.0_f32);
                *e = e.max(off + tb.height);
            }
            m
        };
        let mut shared_float_anchors: std::collections::HashMap<usize, (usize, f32)> = Default::default();
        // S758 (2026-07-06, default ON, opt-out OXI_S758_DISABLE): wrapSquare
        // floating-IMAGE side-wrap. Word narrows every LINE whose y-range
        // intersects the float's band to (float_left − distL) [right-side
        // floats] — imgfloat truth: image [382.7..524.4]×[286.9..400.2],
        // 7 lines narrowed to x1≈374, the band cuts MID-paragraph. v1 scope =
        // paragraph-anchored wrapSquare IMAGES only (every corpus wrapSquare
        // anchor is a TEXTBOX → the registry is empty for the whole corpus =
        // byte-identical by construction; only probeximgfloat activates).
        // distL/distR are not parsed yet — Word's default 114300EMU = 9.0pt
        // (matches the measured 8.7pt gap).
        let s758_on = std::env::var("OXI_S758_DISABLE").is_err();
        let s758_squares: std::collections::HashMap<usize, Vec<usize>> = if s758_on {
            let mut m: std::collections::HashMap<usize, Vec<usize>> = Default::default();
            for (ii, img) in page.floating_images.iter().enumerate() {
                if img.wrap_type == Some(crate::ir::WrapType::Square)
                    && img
                        .position
                        .as_ref()
                        .map_or(false, |ip| ip.v_relative.as_deref() == Some("paragraph"))
                {
                    m.entry(img.anchor_block_index).or_default().push(ii);
                }
            }
            m
        } else {
            Default::default()
        };
        // S758-TB (2026-07-07, default ON, opt-out OXI_S758_TB_DISABLE):
        // wrapSquare TEXTBOX side-wrap — the corpus's actual wrapSquare
        // anchors are all textboxes (29dc6e x1, 2ea81a x4, both word_png)
        // which the image registry ignores. Word truth (probeqtxbxwrap):
        // text flows BESIDE the box exactly like the image variant (Word 5
        // pages vs Oxi 4 without the band; FAIL 0.86 -> PASS 1.0).
        // ★DERIVED clamp-vs-push rule (2026-07-07 _sidewrap_clamp_sweep.py,
        // 20/20 configs): a wrapSquare float that does not fit at its
        // natural position (anchor + posV + h > content_bottom) is
        // (a) CLAMPED up against the CONTENT bottom with the anchor KEPT
        // when the clamped top stays at-or-below the anchor line
        // (content_bottom - h >= anchor_y), else (b) PUSHED to the next
        // page with its anchor. This unifies the probe push (tall box:
        // clamp would rise above the anchor) with 2ea81a's kept stamp
        // boxes (short box far below the anchor: clamp keeps it under).
        // The earlier posV<=30 registration gate is replaced by this rule.
        // S1455 (2026-09-17, default ON, opt-out OXI_S1455_DISABLE): a bare
        // DRAWING SHAPE with wrapSquare is a wrap obstacle exactly like the
        // image and textbox sources. reference__0ea3ec86 p16 right column: the
        // document's ONLY wrapSquare is a 130.4 x 172.9pt "正方形/長方形" with
        // noFill and no line -- an invisible spacer whose whole job is to hold
        // a lane open -- positioned relativeFrom="margin" <wp:align>right.
        // Word flows nine 8-character lines beside it (PDF x0 308.7 -> x1
        // 402.1, then 543.8 once past its bottom, and 172.9pt is exactly those
        // nine lines). The band registry read only `floating_images` and
        // `text_boxes`, so Oxi used the full column, fitted nine lines too many,
        // and every page from 17 to 24 sat one early.
        //
        // The shape is NOT in `page.shapes`: an anchor inside a run lands in
        // `Block::Paragraph(p).shapes`, which is why the painter draws it (the
        // dump shows it) while every wrap consumer missed it. Collect BOTH the
        // page-level and the paragraph-level shapes, the same pairing the
        // painter itself builds. The written wrap kind lives in `anchor_wrap`;
        // `wrap_type` is only filled behind the OXI_DRAWING_SHAPE_WRAP opt-in.
        type S1455Band = (f32, f32, f32, Option<String>, f32, f32, f32, bool);
        let s1455_shps: std::collections::HashMap<usize, Vec<S1455Band>> =
            if s758_on && std::env::var("OXI_S1455_DISABLE").is_err() {
                let mut m: std::collections::HashMap<usize, Vec<S1455Band>> = Default::default();
                let page_level = page.shapes.iter().map(|s| (s.anchor_block_index, s));
                let para_level = page.blocks.iter().enumerate().flat_map(|(i, b)| match b {
                    Block::Paragraph(p) => p.shapes.iter().map(move |s| (i, s)).collect::<Vec<_>>(),
                    _ => Vec::new(),
                });
                for (blk, sh) in page_level.chain(para_level) {
                    if sh.wrap_type.or(sh.anchor_wrap) != Some(crate::ir::WrapType::Square) {
                        continue;
                    }
                    let Some(sp) = sh.position.as_ref() else { continue };
                    if sp.v_relative.as_deref() != Some("paragraph") || sp.y < 0.0 {
                        continue;
                    }
                    // A text box's shape and text share one wrapping rectangle.
                    if page.text_boxes.iter().any(|tb| tb.anchor_block_index == blk
                        && (tb.width - sh.width).abs() < 0.01 && (tb.height - sh.height).abs() < 0.01
                        && tb.position.as_ref().is_some_and(|tp| tp.h_relative == sp.h_relative
                            && tp.v_relative == sp.v_relative && tp.h_align == sp.h_align
                            && (tp.x - sp.x).abs() < 0.01 && (tp.y - sp.y).abs() < 0.01)) {
                        continue;
                    }
                    if std::env::var("OXI_DBG_S1455").is_ok() {
                        eprintln!(
                            "[S1455] blk={} w={:.1} h={:.1} x={:.1} y={:.1} hrel={:?} halign={:?}",
                            blk, sh.width, sh.height, sp.x, sp.y, sp.h_relative, sp.h_align
                        );
                    }
                    m.entry(blk).or_default().push((
                        sp.y,
                        sh.width,
                        sh.height,
                        sp.h_align.clone(),
                        sp.x,
                        sp.dist_l.unwrap_or(9.0),
                        sp.dist_r.unwrap_or(9.0),
                        sp.h_relative.as_deref() == Some("column"),
                    ));
                }
                m
            } else {
                Default::default()
            };
        let s758_tbs: std::collections::HashMap<usize, Vec<usize>> =
            if s758_on && std::env::var("OXI_S758_TB_DISABLE").is_err() {
                let mut m: std::collections::HashMap<usize, Vec<usize>> = Default::default();
                for (ti, tb) in page.text_boxes.iter().enumerate() {
                    // S981 (2026-07-22, default ON, opt-out OXI_S981_DISABLE):
                    // wrapTIGHT side-wrap. Word render-truth (reports__0013bcb8,
                    // whose gray title box is Tight): the host run 'R' renders at
                    // x=476.62 BESIDE the box, so Tight wraps like Square and the
                    // band contract needs no change; that document goes 0.7849 ->
                    // 0.9140. Formerly HELD (S975 opt-in) because policies__0026b7f7
                    // flipped 1.0000 -> 0.7568 (2 -> 3 pages) — its behindDoc Tight
                    // sidebar was pushed to p3 by the content-bottom clamp. REPORT_N
                    // derived the pair fix (physical-bottom fit + identical-x band
                    // union, see the s758_srcs `physical` field): policies is now
                    // back to 2 pages (PASS). BLAST RADIUS (byte-compare, S981
                    // off/on): 2/200 EN docs change (policies + reports, both
                    // Word-correct); 0 golden / 0 real_en (ukframework PASS 1.0 both
                    // ways) / 0 JP / ssim_ab 0 changed.
                    let s981 = std::env::var("OXI_S981_DISABLE").is_err()
                        && tb.wrap_type == Some(crate::ir::WrapType::Tight);
                    if (tb.wrap_type == Some(crate::ir::WrapType::Square) || s981)
                        && tb.position.as_ref().map_or(false, |tp| {
                            tp.v_relative.as_deref() == Some("paragraph") && tp.y >= 0.0
                        })
                    {
                        m.entry(tb.anchor_block_index).or_default().push(ti);
                    }
                }
                m
            } else {
                Default::default()
            };
        // resolved bands: (page_idx, top, bottom, x0, x1)
        // S1387 (2026-09-13, default ON, opt-out OXI_S1387_DISABLE): a PAGE-
        // relative wrapSquare / wrapTight text box is a side band at its
        // absolute y, registered when its anchor block is laid out. The band
        // consumer already moves a line below a band whose lane is under
        // 30pt. reports__0045b085cb32ceba: a 472.5pt-wide box (page y 103.5,
        // h 333) in a 451pt column -- Word starts the next paragraph at
        // 441.75, below the box; Oxi laid it at 108 through the box.
        let s1387_tbs: std::collections::HashMap<usize, Vec<usize>> =
            if s758_on && std::env::var("OXI_S1387_DISABLE").is_err() {
                let mut m: std::collections::HashMap<usize, Vec<usize>> = Default::default();
                for (ti, tb) in page.text_boxes.iter().enumerate() {
                    if std::env::var("OXI_DBG1387").is_ok() {
                        eprintln!("[S1387] tb#{} wrap={:?} pos={:?} anchor={} w={:.1} h={:.1}", ti, tb.wrap_type,
                            tb.position.as_ref().map(|p| (p.v_relative.clone(), p.y, p.h_relative.clone(), p.x)), tb.anchor_block_index, tb.width, tb.height);
                    }
                    // Only a box that CLOSES the column (width within 30pt of it,
                    // the S1389 test): ukframework's 393pt tight cover box in a
                    // 451pt column left a lane and Word kept its empty anchor
                    // paragraph at the page top; a band there added a page.
                    let s1387_col_w = page.size.width - page.margin.left - page.margin.right;
                    if matches!(tb.wrap_type, Some(crate::ir::WrapType::Square | crate::ir::WrapType::Tight))
                        && tb.width >= s1387_col_w - 30.0
                        && tb.position.as_ref().map_or(false, |tp| tp.v_relative.as_deref() == Some("page"))
                    {
                        m.entry(tb.anchor_block_index).or_default().push(ti);
                    }
                }
                m
            } else {
                Default::default()
            };
        let mut s758_bands: Vec<(usize, f32, f32, f32, f32, bool, BodyWrapPolicy)> = Vec::new();
        // S847 (2026-07-14): CONSECUTIVE paragraphs sharing an IDENTICAL
        // page-anchored framePr form ONE text frame — Word stacks them
        // vertically from (x, y) and the continuation paragraphs INHERIT the
        // first paragraph's anchor (they typically omit hAnchor). Keyed on the
        // DECLARED (x, y): (decl_x, decl_y, first_fx, running_bottom_y). Reset
        // when a non-frame block intervenes or the declared key changes.
        let mut s847_frame: Option<(f32, f32, f32, f32)> = None;
        // S1379: the running yAlign frame group -- (declared x, yAlign, group top).
        let mut s1379_group: Option<(f32, String, f32)> = None;
        // S898b: a wrap="notBeside" page-anchored frame group excludes the
        // body from its whole vertical extent — the flow resumes below the
        // STACKED frame bottom (00054c43: Commonwealth letterhead frame
        // 35.25 + 4 lines = 104.25 = Word's measured flow start EXACT).
        let mut s898_notbeside_bottom: Option<f32> = None;
        // S863: consecutive identical vAnchor="text" framePr
        // paragraphs with a negative Y and an exact height form ONE frame.
        // State: (decl_x, decl_y, first_x, running_content_bottom,
        //         exact_frame_bottom, original_anchor_y).
        let mut s863_frame: Option<(f32, f32, f32, f32, f32, f32)> = None;
        // ANCHORPUSH experiment (opt-in OXI_ANCHORPUSH=1, ROWBOX2-family):
        // Word keeps an anchored drawing on the SAME PAGE as its anchor
        // paragraph — when the float's extent (anchor_y + posOffsetV + cy)
        // would cross the PHYSICAL page bottom, Word pushes the ANCHOR
        // paragraph to the next page (2ea81a pi27/pi28: the 34.4×30.4 stamp
        // circle at posV 47.5 and the ＜＜記載例＞＞ 42×57.85 box at posV 8
        // both land past p1's page edge at their anchors' natural position
        // → Word starts p2 with them; S758 shipped the wrapSquare variant).
        // PHYSICAL page bottom (not the content bottom): corpus stamps/seals
        // legitimately hang into the bottom MARGIN and Word keeps them.
        let anchorpush: std::collections::HashMap<usize, f32> =
            if std::env::var("OXI_ANCHORPUSH").is_ok() {
                let mut m: std::collections::HashMap<usize, f32> = Default::default();
                for tb in &page.text_boxes {
                    if let Some(p) = &tb.position {
                        if p.v_relative.as_deref() == Some("paragraph") && p.y >= 0.0 {
                            let need = p.y + tb.height;
                            let e = m.entry(tb.anchor_block_index).or_insert(0.0_f32);
                            if need > *e {
                                *e = need;
                            }
                        }
                    }
                }
                m
            } else {
                Default::default()
            };
        // S970 v2 (2026-07-21, default ON, opt-out OXI_S970_DISABLE): a table whose
        // document-order TERMINAL paragraph declares keepNext keeps the FOLLOWING
        // body paragraph with it — Word puts terminal and follower on the same page
        // at all 15 firing points in the corpus. The v1 decided this from
        // estimate_table_row_natural_h and over-fired badly (technical__00501ca: it
        // summed 420.8pt for a table that lays out over ~1700pt across four page
        // fragments, so the "fits a fresh page" guard passed for a table that fits
        // no page at all, and 10 pages = Word became 11). This version decides from
        // ACTUAL geometry like S960: the table must have laid out without pushing a
        // single page (that IS "the table fits one page"), and the follower must
        // then have whole-moved — only then are the already-emitted table elements
        // pulled onto the follower's page. Nothing here pushes a page, so the whole
        // page-transition state comes from the follower's own normal layout.
        // (table_block_idx, page_idx, elem_start, elem_end, table_top)
        let mut s970_pending: Option<(usize, usize, usize, usize, f32)> = None;
        // S1174: per-page STYLEREF resolution state. `ingested` is a cursor
        // over this section's blocks (paragraphs strictly BEFORE the current
        // block have completed layout = they sit on earlier pages when a new
        // page starts, which is exactly the causal last-before-page set).
        // `snapshots` records each page's resolved map for the emit loop.
        let mut s1174_ingested = 0usize;
        let mut s1174_snapshots: std::collections::HashMap<
            usize,
            std::collections::HashMap<String, String>,
        > = std::collections::HashMap::new();
        for (block_idx, block) in page.blocks.iter().enumerate() {
            // A paragraph-owned band must end even when an empty anchor takes
            // an early continuation path. Nested layouts restore their caller.
            let _paragraph_float_band = ParagraphFloatBandGuard::new();
            // In modern layout an over-height square text box owns the rest
            // of its anchor page; following body content starts on a new page.
            if block_idx > 0 && block_page_indices.get(block_idx - 1) == Some(&pages.len())
                && page.text_boxes.iter().any(|tb| tb.anchor_block_index == block_idx - 1
                    && tb.height > content_height && !self.legacy_square_textbox_clamp(tb)
                    && tb.wrap_type == Some(crate::ir::WrapType::Square)
                    && tb.position.as_ref().is_some_and(|pos|
                        pos.v_relative.as_deref() == Some("paragraph") && pos.y >= 0.0)) {
                cursor.set(cursor.cursor_y.max(start_y + content_height));
            }
            if let Ok(rng) = std::env::var("OXI_DBG_BLKTRACE") {
                let mut it = rng.split('-');
                let lo: usize = it.next().and_then(|v| v.parse().ok()).unwrap_or(0);
                let hi: usize = it.next().and_then(|v| v.parse().ok()).unwrap_or(usize::MAX);
                if block_idx >= lo && block_idx <= hi {
                    let head: String = match block {
                        Block::Paragraph(p) => p
                            .runs
                            .iter()
                            .flat_map(|r| r.text.chars())
                            .take(20)
                            .collect(),
                        Block::Table(_) => "<TABLE>".into(),
                        _ => "<?>".into(),
                    };
                    eprintln!(
                        "[BLKTRACE] blk={} page={} cur={:.1} {:?}",
                        block_idx,
                        pages.len() + 1,
                        cursor.cursor_y,
                        head
                    );
                }
            }
            if !s863_vertical_geoms.is_empty() {
                if std::env::var_os("OXI_MARGIN_TRACE").is_some() {
                    eprintln!("[MARGIN] blk={} page={} elements={} cursor={} start={} run={} begin={} runs={:?}", block_idx, pages.len()+1, elements.len(), cursor.cursor_y, start_y, s863_vertical_run_idx, s863_page_begin_idx, page.vertical_runs);
                }
                // S1227: a page begun by an earlier block is governed by the
                // run active at that push — capture it BEFORE this block's
                // run switch.
                let section_starts_fresh_page = pages.len() != s863_last_pages_len
                    && elements.is_empty()
                    && current_column == 0
                    && (cursor.cursor_y - start_y).abs() < 0.01;
                let previous_vertical_run = s863_vertical_run_idx;
                if pages.len() != s863_last_pages_len {
                    s863_last_pages_len = pages.len();
                    s863_page_begin_idx = s863_vertical_run_idx;
                }
                while s863_vertical_run_idx + 1 < page.vertical_runs.len()
                    && block_idx >= page.vertical_runs[s863_vertical_run_idx + 1].0
                {
                    s863_vertical_run_idx += 1;
                    s755_geom = Some(s863_vertical_geoms[s863_vertical_run_idx]);
                }
                // An empty physical page opened by the preceding block has
                // no body content belonging to the previous section. The
                // incoming section governs both its box and its cursor.
                if section_starts_fresh_page
                    && previous_vertical_run != s863_vertical_run_idx
                {
                    s863_page_begin_idx = s863_vertical_run_idx;
                    let g = &s863_vertical_geoms[s863_vertical_run_idx];
                    start_y = g.top(pages.len() + 1);
                    content_height = g.ch(pages.len() + 1);
                    cursor.set(start_y);
                    col_band_top = start_y;
                }
            }
            // S1174: ingest the paragraphs completed so far and, when a new
            // page has begun, re-resolve the header/footer STYLEREF text and
            // recompute this page's geometry (the resolved Part/Chapter name
            // can wrap to a different line count, moving header_bottom across
            // the top margin — reference__0061531a p52's +7pt pushdown).
            // Vertical multi-run sections keep their S863 geometry untouched.
            if S1174_ACTIVE.with(|c| c.get()) {
                // Ingest through the CURRENT block (v1.5): Word's rule is
                // first-on-page, and the page-starting block is most often the
                // very heading the header should show (ActHead pbb pages) — a
                // page break triggered INSIDE this block must already see its
                // style text, or a Schedule page's header renders the PREVIOUS
                // Part's 2-line name where Word shows the current 1-line one
                // (the +1x2 overshoot this replaced).
                while s1174_ingested <= block_idx {
                    LayoutEngine::s1174_ingest_block(&page.blocks[s1174_ingested], false);
                    s1174_ingested += 1;
                }
                if s1174_have_ref && s863_vertical_geoms.is_empty() {
                    let pno = pages.len() + 1;
                    // Geometry is recomputed EVERY block (a break inside THIS
                    // block must see its just-ingested style text); the emit
                    // snapshot keeps the page-START state (first write wins).
                    {
                        let map = LayoutEngine::s1174_map();
                        let hdr_sub = LayoutEngine::s1174_substitute(&page.header, &map);
                        let hb_odd = self.s755_header_bottom(&hdr_sub, page);
                        let sy_odd = page.body_start_y(hb_odd, self.s1381_header_band(&hdr_sub, page));
                        let (fr_odd, _) = self
                            .s755_footer_geom(&LayoutEngine::s1174_substitute(&page.footer, &map), page);
                        let ch_odd = page.size.height - sy_odd - fr_odd;
                        let (sy_even, ch_even) = if page.even_odd_hf {
                            let hdr_even_sub = LayoutEngine::s1174_substitute(&page.header_even, &map);
                            let hb = self.s755_header_bottom(&hdr_even_sub, page);
                            let sy = page.body_start_y(hb, self.s1381_header_band(&hdr_even_sub, page));
                            let (fr, _) = self.s755_footer_geom(
                                &LayoutEngine::s1174_substitute(&page.footer_even, &map),
                                page,
                            );
                            (sy, page.size.height - sy - fr)
                        } else {
                            (sy_odd, ch_odd)
                        };
                        let first = s755_geom
                            .as_ref()
                            .map(|g| g.first)
                            .unwrap_or((start_y, content_height));
                        if std::env::var("OXI_DBG1174").is_ok() {
                            eprintln!(
                                "[S1174] hook pno={} blk={} sy_odd={:.2} sy_even={:.2} map={} part={}",
                                pno,
                                block_idx,
                                sy_odd,
                                sy_even,
                                map.len(),
                                map.get("CharPartText").map(|s| s.len()).unwrap_or(0)
                            );
                        }
                        s755_geom = Some(S755Geom {
                    first_even: first_logical % 2 == 0,
                            first,
                            odd: (sy_odd, ch_odd),
                            even: (sy_even, ch_even),
                    page_override: None,
                        });
                        s1174_snapshots.entry(pno).or_insert(map);
                    }
                }
            }
            // S755: refresh the current page's header/footer geometry (a
            // previous block may have pushed pages internally).
            // S1227: on a vertical-multi-run page, the CURRENT page's box is
            // the geom of the run in which the page BEGAN, not the run this
            // block belongs to (Word: continuous-section top/bottom margins
            // govern whole physical pages).
            if s1227_on && !s863_vertical_geoms.is_empty() {
                let g = &s863_vertical_geoms[s863_page_begin_idx.min(s863_vertical_geoms.len() - 1)];
                start_y = g.top(pages.len() + 1);
                content_height = g.ch(pages.len() + 1);
            } else if let Some(g) = s755_geom.as_ref() {
                start_y = g.top(pages.len() + 1);
                content_height = g.ch(pages.len() + 1);
            }
            // S735: switch the active grid pitch at section-run boundaries.
            if s735_grid_het
                && s735_run_idx + 1 < page.grid_runs.len()
                && block_idx >= page.grid_runs[s735_run_idx + 1].0
            {
                while s735_run_idx + 1 < page.grid_runs.len()
                    && block_idx >= page.grid_runs[s735_run_idx + 1].0
                {
                    s735_run_idx += 1;
                }
                grid_pitch = page.grid_runs[s735_run_idx].1;
            }
            // S1336: this block's character grid (see the variants above).
            let page: &Page = s1336_variants
                .iter()
                .rev()
                .find(|(start, _)| block_idx >= *start)
                .and_then(|(_, v)| v.as_ref())
                .unwrap_or(page_orig);
            // S734: reserve the wrapTopAndBottom band ABOVE this anchor block.
            if let Some(&band_h) = s734_bands.get(&block_idx)
                .filter(|_| s1497b_img_overlaps(block_idx, &block_y_positions, cursor.cursor_y, start_x, content_width))
            {
                let band_h = if crate::layout::s1467_float_column_flow() {
                    band_h.max(s1089_tb_bands.get(&block_idx).copied().unwrap_or(0.0))
                } else { band_h };
                // S1513b: the paragraph-relative band hangs from the host's TOP,
                // i.e. after its (collapsed) space-before; measure the remaining
                // room from there (tb_after_probe2.py: the host moves to the next
                // page exactly when band bottom > text bottom; 004c4a3d's host
                // carries before=120 and its band ended 0.5pt short in Oxi).
                let s1513_before = if std::env::var_os("OXI_S1513_DISABLE").is_none() {
                    if let Block::Paragraph(para) = block {
                        self.paragraph_spacing_before(
                            para, page, grid_pitch, prev_para_style_id.as_deref(),
                            prev_contextual_spacing, prev_autospacing_numid.as_deref(),
                            prev_space_after, Some(block_idx), &pages, &elements,
                            cursor.cursor_y, start_y,
                        ).0.max(0.0)
                    } else { 0.0 }
                } else { 0.0 };
                let remaining = (start_y + content_height) - cursor.cursor_y - s1513_before;
                if band_h > remaining && band_h <= content_height && !elements.is_empty() {
                    if crate::layout::s1467_float_column_flow()
                        && current_column + 1 < num_columns
                    {
                        current_column += 1;
                        start_x = col_x_positions[current_column];
                        content_width = col_widths[current_column];
                        cursor.set(col_band_top);
                        lm2_cells = 0;
                    } else {
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(&mut elements),
                    });
                    if let Some(g) = s755_geom.as_ref() {
                        start_y = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(start_y);
                    lm2_cells = 0;
                    current_page_idx += 1;
                    footnote_reserve_current = 0.0;
                    footnote_ids_current_page.clear();
                    s900_fold(
                        &mut footnote_reserve_current,
                        &mut footnote_ids_current_page,
                        &mut s900_pending_deferred,
                        current_page_idx,
                    );
                        if crate::layout::s1467_float_column_flow() {
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            col_band_top = start_y;
                        }
                    }

                }
                // S1500: an in-paragraph band on the block after a band host
                // is measured from the unpushed top.
                let s1500_fy = s1500_unpushed
                    .filter(|&(b, pg, _)| b == block_idx && pg == current_page_idx && s1497_mid.contains_key(&block_idx))
                    .map(|(_, _, y)| y)
                    .unwrap_or(cursor.cursor_y);
                s734_flow_pos.insert(block_idx, (current_page_idx, s1500_fy));
                if let Some(&(off, h)) = s1497_mid.get(&block_idx) {
                    S1497_BAND.with(|c| c.set(Some((off - (cursor.cursor_y - s1500_fy), h))));
                } else {
                    cursor.advance(band_h);
                }
            }
            // S1089: the same reservation for a wrapTopAndBottom wps SHAPE.
            if let Some(&band_h) = s1089_tb_bands.get(&block_idx)
                .filter(|_| s1497b_tb_overlaps(block_idx, &block_y_positions, &block_col_x, start_x, content_width))
            {
                if crate::layout::s1467_float_column_flow()
                    && s734_flow_pos.contains_key(&block_idx)
                {
                    s1089_flow_pos.insert(block_idx, s734_flow_pos[&block_idx]);
                } else {
                let remaining = (start_y + content_height) - cursor.cursor_y;
                if band_h > remaining && band_h <= content_height && !elements.is_empty() {
                    if crate::layout::s1467_float_column_flow()
                        && current_column + 1 < num_columns
                    {
                        current_column += 1;
                        start_x = col_x_positions[current_column];
                        content_width = col_widths[current_column];
                        cursor.set(col_band_top);
                        lm2_cells = 0;
                    } else {
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(&mut elements),
                    });
                    if let Some(g) = s755_geom.as_ref() {
                        start_y = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(start_y);
                    lm2_cells = 0;
                    current_page_idx += 1;
                    footnote_reserve_current = 0.0;
                    footnote_ids_current_page.clear();
                    s900_fold(
                        &mut footnote_reserve_current,
                        &mut footnote_ids_current_page,
                        &mut s900_pending_deferred,
                        current_page_idx,
                    );
                        if crate::layout::s1467_float_column_flow() {
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            col_band_top = start_y;
                        }
                    }

                }
                let s1500_fy = s1500_unpushed
                    .filter(|&(b, pg, _)| b == block_idx && pg == current_page_idx && s1497_mid.contains_key(&block_idx))
                    .map(|(_, _, y)| y)
                    .unwrap_or(cursor.cursor_y);
                s1089_flow_pos.insert(block_idx, (current_page_idx, s1500_fy));
                if let Some(&(off, h)) = s1497_mid.get(&block_idx) {
                    S1497_BAND.with(|c| c.set(Some((off - (cursor.cursor_y - s1500_fy), h))));
                } else {
                    cursor.advance(band_h);
                }
                }
            }
            // S1552: a wrapSquare box without a usable lane moves its host to
            // the next page when the box would not fit below the host (same
            // test as S1089/S1513: band > remaining, band <= content height);
            // nothing else changes — the flow around the box is S758's.
            if let Some(&band_h) = s1552_bands.get(&block_idx)
                .filter(|_| !s1089_tb_bands.contains_key(&block_idx))
                .filter(|_| s1497b_tb_overlaps(block_idx, &block_y_positions, &block_col_x, start_x, content_width))
            {
                let remaining = (start_y + content_height) - cursor.cursor_y;
                if band_h > remaining && band_h <= content_height && !elements.is_empty()
                    && !(crate::layout::s1467_float_column_flow() && current_column + 1 < num_columns)
                {
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(&mut elements),
                    });
                    if let Some(g) = s755_geom.as_ref() {
                        start_y = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(start_y);
                    lm2_cells = 0;
                    current_page_idx += 1;
                    footnote_reserve_current = 0.0;
                    footnote_ids_current_page.clear();
                    s900_fold(
                        &mut footnote_reserve_current,
                        &mut footnote_ids_current_page,
                        &mut s900_pending_deferred,
                        current_page_idx,
                    );
                    if crate::layout::s1467_float_column_flow() {
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        col_band_top = start_y;
                    }
                }
            }
            // S842 (2026-07-14, opt-out OXI_S842_DISABLE): a PAGE-anchored
            // wrapTopAndBottom float pushes its anchor block below the band.
            // hmrc's top rule (Group 3426: positionV page 19.75, 2pt line,
            // anchored to p2's first empty para): Word starts that para at
            // band bottom 21.75, not the 15.75 top margin — without this the
            // whole p2 ran ~6pt high.
            if std::env::var("OXI_S842_DISABLE").is_err() {
                for tb in &page.text_boxes {
                    if tb.anchor_block_index == block_idx
                        && matches!(tb.wrap_type, Some(crate::ir::WrapType::TopAndBottom))
                    {
                        if let Some(tp) = tb.position.as_ref() {
                            if tp.v_relative.as_deref() == Some("page") {
                                let band_bottom = tp.y + tb.height;
                                if cursor.cursor_y < band_bottom && cursor.cursor_y + 10.0 > tp.y {
                                    cursor.set(band_bottom);
                                }
                            }
                        }
                    }
                }
            }
            // ANCHORPUSH: push the anchor block to the next page when its
            // paragraph-anchored float would cross the PHYSICAL page bottom.
            if let Some(&need) = anchorpush.get(&block_idx) {
                if cursor.cursor_y + need > page.size.height
                    && need <= content_height
                    && !elements.is_empty()
                {
                    if crate::layout::s1467_float_column_flow()
                        && current_column + 1 < num_columns
                    {
                        current_column += 1;
                        start_x = col_x_positions[current_column];
                        content_width = col_widths[current_column];
                        cursor.set(col_band_top);
                        lm2_cells = 0;
                    } else {
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(&mut elements),
                    });
                    if let Some(g) = s755_geom.as_ref() {
                        start_y = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(start_y);
                    lm2_cells = 0;
                    current_page_idx += 1;
                    footnote_reserve_current = 0.0;
                    footnote_ids_current_page.clear();
                    s900_fold(
                        &mut footnote_reserve_current,
                        &mut footnote_ids_current_page,
                        &mut s900_pending_deferred,
                        current_page_idx,
                    );
                        if crate::layout::s1467_float_column_flow() {
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            col_band_top = start_y;
                        }
                    }

                }
            }
            // Side-wrapping geometric shapes share the paragraph wrap bands.
            for shape in page.shapes.iter().filter(|s| s.anchor_block_index == block_idx)
                .chain(match &page.blocks[block_idx] {
                    Block::Paragraph(p) => p.shapes.iter(),
                    _ => [].iter(),
                }) {
                if matches!(shape.wrap_type, Some(crate::ir::WrapType::Square | crate::ir::WrapType::Tight)) {
                    if let Some(pos) = &shape.position {
                        let mut top = cursor.cursor_y + pos.y;
                        if (std::env::var("OXI_PARAGRAPH_FLOAT_SPACING").is_ok()
                || std::env::var("OXI_S1471_DISABLE").is_err())
                            && matches!(pos.v_relative.as_deref(), None | Some("paragraph"))
                        {
                            if let Block::Paragraph(para) = block {
                                top += self.paragraph_spacing_before(
                                    para, page, grid_pitch, prev_para_style_id.as_deref(),
                                    prev_contextual_spacing, prev_autospacing_numid.as_deref(),
                                    prev_space_after, Some(block_idx), &pages, &elements,
                                    cursor.cursor_y, start_y,
                                ).0;
                            }
                        }
                        let mut left = start_x + pos.x;
                        if std::env::var("OXI_DRAWING_SHAPE_WRAP").is_ok() {
                            let (rx, rw) = match pos.h_relative.as_deref() {
                                Some("page") => (0.0, page.size.width),
                                Some("margin") => (page.margin.left, total_content_width),
                                _ => (start_x, content_width),
                            };
                            left = match pos.h_align.as_deref() {
                                Some("right") => rx + rw - shape.width,
                                Some("center") => rx + (rw - shape.width) * 0.5,
                                Some("left") => rx,
                                _ => rx + pos.x,
                            };
                            let reference = match pos.v_relative.as_deref() {
                                Some("page") => Some((0.0, page.size.height)),
                                Some("margin") => Some((page.margin.top,
                                    page.size.height - page.margin.top - page.margin.bottom)),
                                _ => None,
                            };
                            if let Some((ry, rh)) = reference {
                                top = match pos.v_align.as_deref() {
                                    Some("top") => ry,
                                    Some("bottom") => ry + rh - shape.height,
                                    Some("center") => ry + (rh - shape.height) * 0.5,
                                    _ => ry + pos.y,
                                };
                            }
                        }
                        let bounds = if shape.wrap_type == Some(crate::ir::WrapType::Tight) {
                            shape.wrap_polygon.iter().fold((0.0_f32, 0.0_f32, 1.0_f32, 1.0_f32),
                                |(x0, y0, x1, y1), &(x, y)| (x0.min(x), y0.min(y), x1.max(x), y1.max(y)))
                        } else { (0.0, 0.0, 1.0, 1.0) };
                        s758_bands.push((current_page_idx, top + bounds.1 * shape.height, top + bounds.3 * shape.height,
                            left + bounds.0 * shape.width - pos.dist_l.unwrap_or(9.0),
                            left + bounds.2 * shape.width + pos.dist_r.unwrap_or(9.0),
                            shape.wrap_type == Some(crate::ir::WrapType::Tight), BodyWrapPolicy::OBJECT));
                    }
                }
            }
            // Page-relative shapes keep their physical Y even when the anchor
            // paragraph moves. Their wrapping still excludes body text; behindDoc
            // controls painting order, not the wrap rectangle.
            if std::env::var_os("OXI_PAGE_SHAPE_WRAP").is_some() {
                for tb in &page.text_boxes {
                    if tb.anchor_block_index != block_idx
                        || !matches!(tb.wrap_type, Some(crate::ir::WrapType::Square) | Some(crate::ir::WrapType::Tight))
                    { continue; }
                    let Some(pos) = tb.position.as_ref() else { continue };
                    if pos.v_relative.as_deref() != Some("page") { continue; }
                    let (ref_x, ref_w) = if pos.h_relative.as_deref() == Some("page") {
                        (0.0, page.size.width)
                    } else if pos.h_relative.as_deref() == Some("column") {
                        (start_x, content_width)
                    } else { (page.margin.left, total_content_width) };
                    let x = match pos.h_align.as_deref() {
                        Some("center") => ref_x + (ref_w - tb.width) * 0.5,
                        Some("right") => ref_x + ref_w - tb.width,
                        Some("left") => ref_x,
                        _ => ref_x + pos.x,
                    };
                    s758_bands.push((current_page_idx, pos.y, pos.y + tb.height,
                        x - pos.dist_l.unwrap_or(9.0), x + tb.width + pos.dist_r.unwrap_or(9.0),
                        tb.wrap_type == Some(crate::ir::WrapType::Tight), BodyWrapPolicy::OBJECT));
                }
            }
            // S758: resolve wrapSquare bands anchored to this block (band top =
            // the anchor paragraph's first-line top = the cursor here).
            // S981 (2026-07-22, gated under S975 opt-in): the trailing bool is
            // `physical` — a behindDoc Tight textbox occupies the BOTTOM MARGIN
            // (drawn to the physical page bottom), so it is fit against the
            // physical page height, NOT the content bottom, and it is NOT
            // whole-pushed to the next page. policies__0026b7f7: a behindDoc
            // wrapTight sidebar (y 528..840.75, past content bottom 770) that
            // Word keeps on p2; the content-bottom clamp/push sent it + its
            // anchor to p3 (S975=1: 2 -> 3 pages). REPORT_N.
            let s758_srcs: Vec<(f32, f32, f32, Option<String>, f32, f32, f32, bool, bool, bool)> = {
                // unified (pos_y, width, height, h_align, x, dist_l, dist_r, physical)
                // over image + textbox wrapSquare sources anchored to this block
                let mut v: Vec<(f32, f32, f32, Option<String>, f32, f32, f32, bool, bool, bool)> = Vec::new();
                if let Some(iis) = s758_squares.get(&block_idx) {
                    for &ii in iis {
                        let img = &page.floating_images[ii];
                        if let Some(ip) = img.position.as_ref() {
                            v.push((
                                ip.y,
                                img.width,
                                img.height,
                                ip.h_align.clone(),
                                ip.x,
                                ip.dist_l.unwrap_or(9.0),
                                ip.dist_r.unwrap_or(9.0),
                                false,
                                false,
                                ip.h_relative.as_deref() == Some("column"),
                            ));
                        }
                    }
                }
                if let Some(bands) = s1455_shps.get(&block_idx) {
                    for b in bands {
                        v.push((b.0, b.1, b.2, b.3.clone(), b.4, b.5, b.6, false, false, b.7));
                    }
                }
                if let Some(tis) = s758_tbs.get(&block_idx) {
                    for &ti in tis {
                        let tb = &page.text_boxes[ti];
                        if let Some(tp) = tb.position.as_ref() {
                            let s981_physical = std::env::var("OXI_S981_DISABLE").is_err()
                                && tb.behind_doc
                                && tb.wrap_type == Some(crate::ir::WrapType::Tight);
                            v.push((
                                tp.y,
                                tb.width,
                                tb.height,
                                tp.h_align.clone(),
                                tp.x,
                                tp.dist_l.unwrap_or(9.0),
                                tp.dist_r.unwrap_or(9.0),
                                s981_physical || self.legacy_square_textbox_clamp(tb),
                                tb.wrap_type == Some(crate::ir::WrapType::Tight),
                                tp.h_relative.as_deref() == Some("column"),
                            ));
                        }
                    }
                }
                v
            };
                // The anchor paragraph must carry visible text: ukframework's cover
                // box (lanes 4 / 8pt) hangs off an EMPTY paragraph that Word leaves
                // at the page top, unmoved (the S1195 empty-beside-float rule);
                // reports__0045b085's box hangs off its title. Hypothesis until a
                // probe separates "empty anchor" from "lane closed".
                let s1387_anchor_visible = matches!(&page.blocks[block_idx], Block::Paragraph(p)
                    if p.runs.iter().any(|r| r.text.chars().any(|c| !c.is_whitespace())));
                if let Some(tis) = s1387_tbs.get(&block_idx).filter(|_| s1387_anchor_visible) {
                    for &ti in tis {
                        let tb = &page.text_boxes[ti];
                        if let Some(tp) = tb.position.as_ref() {
                            let x0 = match (tp.h_relative.as_deref(), tp.h_align.as_deref()) {
                                (_, Some("right")) => page.margin.left + total_content_width - tb.width,
                                (_, Some("center")) => page.margin.left + (total_content_width - tb.width) * 0.5,
                                (_, Some("left")) => page.margin.left,
                                (Some("page"), _) => tp.x,
                                _ => page.margin.left + tp.x,
                            };
                            let dl = tp.dist_l.unwrap_or(9.0);
                            let dr = tp.dist_r.unwrap_or(9.0);
                            let nb = (current_page_idx, tp.y, tp.y + tb.height, x0 - dl, x0 + tb.width + dr,
                                      tb.wrap_type == Some(crate::ir::WrapType::Tight), BodyWrapPolicy::OBJECT);
                            if std::env::var("OXI_DBG1387").is_ok() {
                                eprintln!("[S1387] blk={} band={:?}", block_idx, nb);
                            }
                            s758_bands.push(nb);
                        }
                    }
                }
            if !s758_srcs.is_empty() {
                // Co-anchored floats share the origin from before a
                // top-and-bottom float reserved space in the text flow.
                let shared_origin = if std::env::var("OXI_SHARED_FLOAT_ANCHOR_DISABLE").is_err() {
                    [s734_flow_pos.get(&block_idx), s1089_flow_pos.get(&block_idx)]
                        .into_iter().flatten()
                        .filter(|(pg, _)| *pg == current_page_idx)
                        .map(|(_, y)| *y).reduce(f32::min)
                } else { None };
                let mut s758_anchor_y = shared_origin.unwrap_or(cursor.cursor_y);
                if (std::env::var("OXI_PARAGRAPH_FLOAT_SPACING").is_ok()
                || std::env::var("OXI_S1471_DISABLE").is_err()) {
                    if let Block::Paragraph(para) = block {
                        s758_anchor_y += self.paragraph_spacing_before(
                            para, page, grid_pitch, prev_para_style_id.as_deref(),
                            prev_contextual_spacing, prev_autospacing_numid.as_deref(),
                            prev_space_after, Some(block_idx), &pages, &elements,
                            s758_anchor_y, start_y,
                        ).0;
                    }
                }
                let shared_advance = cursor.cursor_y - shared_origin.unwrap_or(cursor.cursor_y);
                // Word pushes the float AND its anchor paragraph to the next
                // page when the band would cross the page bottom (a wrapSquare
                // float never splits): probeximgfloat float#2 anchored at
                // 第25条 — Word starts p3 with the paragraph + image at the
                // top. Same push shape as the S734 wrapTopAndBottom arm.
                let s758_content_bottom = start_y + content_height;
                // PUSH only when some source cannot be clamped below the
                // anchor line (derived rule branch (b)); a clampable
                // overflow keeps the anchor (branch (a)).
                let s758_needs_push =
                    s758_srcs
                        .iter()
                        .any(|(py, _w, h, _a, _x, _dl, _dr, physical, tight, _)| {
                            // Legacy square text boxes keep their anchor and may
                            // move above it to remain within the physical page.
                            if *physical && !*tight { return false; }
                            // S981: a physical (behindDoc Tight) source is fit against the
                            // physical page bottom — it may use the bottom margin.
                            let fit_bottom = if *physical {
                                page.size.height
                            } else {
                                s758_content_bottom
                            };
                            let nat_bottom = s758_anchor_y + py.max(0.0) + h;
                            nat_bottom > fit_bottom && (fit_bottom - h) < s758_anchor_y - 0.1
                        });
                let s758_max_bottom = s758_srcs
                    .iter()
                    .map(|(py, _w, h, _a, _x, _dl, _dr, _, _, _)| s758_anchor_y + py.max(0.0) + h)
                    .fold(f32::NEG_INFINITY, f32::max);
                if s758_needs_push
                    && (s758_max_bottom - s758_anchor_y <= content_height
                        || page.text_boxes.iter().any(|tb| tb.anchor_block_index == block_idx
                            && !self.legacy_square_textbox_clamp(tb)
                            && tb.wrap_type == Some(crate::ir::WrapType::Square)))
                    && !elements.is_empty()
                {
                    if crate::layout::s1467_float_column_flow()
                        && current_column + 1 < num_columns
                    {
                        current_column += 1;
                        start_x = col_x_positions[current_column];
                        content_width = col_widths[current_column];
                        cursor.set(col_band_top);
                        lm2_cells = 0;
                        s758_anchor_y = col_band_top;
                        if shared_origin.is_some() {
                            cursor.advance(shared_advance);
                            for positions in [&mut s734_flow_pos, &mut s1089_flow_pos] {
                                if let Some(value) = positions.get_mut(&block_idx) {
                                    *value = (current_page_idx, col_band_top);
                                }
                            }
                        }
                    } else {
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(&mut elements),
                    });
                    if let Some(g) = s755_geom.as_ref() {
                        start_y = g.top(pages.len() + 1);
                        content_height = g.ch(pages.len() + 1);
                    }
                    cursor.set(start_y);
                    lm2_cells = 0;
                    current_page_idx += 1;
                    s758_anchor_y = start_y;
                    if shared_origin.is_some() {
                        cursor.advance(shared_advance);
                        for positions in [&mut s734_flow_pos, &mut s1089_flow_pos] {
                            if let Some(value) = positions.get_mut(&block_idx) {
                                *value = (current_page_idx, start_y);
                            }
                        }
                    }
                    footnote_reserve_current = 0.0;
                    footnote_ids_current_page.clear();
                    s900_fold(
                        &mut footnote_reserve_current,
                        &mut footnote_ids_current_page,
                        &mut s900_pending_deferred,
                        current_page_idx,
                    );
                        if crate::layout::s1467_float_column_flow() {
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            col_band_top = start_y;
                        }
                    }

                }
                if shared_origin.is_some() {
                    shared_float_anchors.insert(block_idx, (current_page_idx, s758_anchor_y));
                }
                for (py, w, h, h_align, px, dl, dr, physical, tight, column_relative) in &s758_srcs {
                    let nat_top = s758_anchor_y + py.max(0.0);
                    // S981: a physical (behindDoc Tight) float may extend into the
                    // bottom margin, so clamp against the physical page bottom.
                    let cb = if *physical {
                        page.size.height
                    } else {
                        start_y + content_height
                    };
                    // derived rule branch (a): clamp an overflowing float up
                    // against the content bottom (never above the anchor).
                    let top = if nat_top + h > cb {
                        if *physical && !*tight { (cb - h).max(0.0) }
                        else { (cb - h).max(s758_anchor_y) }
                    } else {
                        nat_top
                    };
                    let bottom = top + h;
                    let column_scope = *column_relative
                        && crate::layout::s1467_float_column_flow();
                    let content_left = if column_scope { start_x } else { page.margin.left };
                    let band_width = if column_scope { content_width } else { total_content_width };
                    let x0 = match h_align.as_deref() {
                        Some("right") => content_left + band_width - w,
                        Some("center") => content_left + (band_width - w) * 0.5,
                        Some("left") => content_left,
                        _ => content_left + px,
                    };
                    // the keep-out rect includes the wp:anchor distL/distR
                    // margins (default 114300EMU = 9.0pt); the consumption
                    // side then narrows exactly to the rect edge.
                    let nb = (current_page_idx, top, bottom, x0 - dl, x0 + w + dr, *tight, BodyWrapPolicy::OBJECT);
                    // S981: a physical Tight band that vertically overlaps an
                    // already-registered band of the SAME x-extent (the leading
                    // wrapSquare sidebar the author chains to it) is UNIONed —
                    // the consumer picks one band by `.find()`, so without the
                    // union the following lines rebreak at the Square's shorter
                    // bottom and ignore the Tight extent.
                    if *physical && *tight {
                        if let Some(b) = s758_bands.iter_mut().rev().find(|b| {
                            b.0 == nb.0
                                && nb.1 <= b.2 + 0.5
                                && nb.2 >= b.1 - 0.5
                                && (b.3 - nb.3).abs() <= 0.5
                                && (b.4 - nb.4).abs() <= 0.5
                        }) {
                            b.1 = b.1.min(nb.1);
                            b.2 = b.2.max(nb.2);
                            continue;
                        }
                    }
                    s758_bands.push(nb);
                }
            }
            // S638 (kyotei): if a vertAnchor="text" full-page float is active and
            // this block's cursor has reached the float's region (the gap above it
            // is now consumed), skip the cursor past the float (body wraps below).
            if let Some((ft_top, ft_bot, ft_page, ft_x0, ft_x1, ft_lane)) = text_float_region {
                // S1509: an EMPTY paragraph flows in a side lane of at least 18.5pt
                // (the S1195 floor) instead of being bumped below the float.
                let s1509_keep = std::env::var_os("OXI_S1509_DISABLE").is_none()
                    && ft_lane >= 18.5
                    && matches!(block, Block::Paragraph(p) if p.runs.iter().all(|r| r.text.is_empty()));
                // S1241: the float bump applies only when the CURRENT lane
                // ([start_x, start_x + content_width]) overlaps the float's X
                // range — a col2 flow passes a col1 float untouched (Word:
                // forms__000cf39c's inline table lives beside the float).
                // Note: the region is NOT cleared for a non-overlapping lane —
                // a later return to an overlapping lane (next page reset) still
                // sees it.
                let s1241_lane_overlaps = std::env::var("OXI_S1241_DISABLE").is_ok()
                    || (ft_x0 < start_x + content_width - 1.0 && ft_x1 > start_x + 1.0);
                if current_page_idx == ft_page && s1241_lane_overlaps {
                    // S872 (2026-07-16, default ON, opt-out OXI_S872_DISABLE): an
                    // EMPTY paragraph whose line BOX would cross the float's top
                    // is bumped below the float — the cursor-start-only test let
                    // it overlap. policies__0009e9db Word COM truth: 7 empties
                    // follow the centered Pathway float; Word puts #1 above the
                    // float (box 367.5..382 fits above top 401.9) and bumps #2
                    // (box 390..404.5 CROSSES) below to 642; Oxi placed #2 at
                    // 390.4 in the gap. The downstream shift is what put the
                    // vertAnchor=page Cognition float's flow position on page 1
                    // (Word: page 2 — a float renders at tblpY on the page of
                    // its FLOW position; probe _pb_floatanchor_gen: the float
                    // stays with its flow position even at 1.5pt room, and the
                    // flow position advances a page when the empties push past
                    // the content bottom). Scoped to EMPTY paragraphs (the
                    // S562b/S736 "empty paras use the full box" principle) —
                    // text lines keep the cursor-only test (kyotei's 様式 label
                    // sits in a gap its box may graze; S638 canary).
                    // The line-box test must include the collapsed before-spacing
                    // (max(prev_sa, own sb)) — the snap runs BEFORE layout_paragraph
                    // applies it (policies pi10: cursor 382.36 + 8 + 14.49 = 404.85
                    // crosses top 399.2; without the spacing 396.85 misses).
                    // ★LATIN scope (!doc_body_has_real_cjk): the JP S638 gap-flow
                    // (2ea81a/kyotei) is a calibrated balance under the cursor-only
                    // rule — unscoped, S872 bumped 2ea81a's gap empty and cost
                    // −0.0718 SSIM (ssim_ab, the only changed doc). The empty-box
                    // rule is measured on policies (Word COM); a JP derivation
                    // would need its own session.
                    // S1489 (2026-09-19, default ON, opt-out OXI_S1489_DISABLE): the
                    // same empty-box rule for a CJK body. golden parttime p2: the
                    // empty exact-11 paragraph before the 年次有給休暇 float sits at
                    // 273.5..284.5 against the band top 281.2 and Word puts it below
                    // the table (「２ 年次…」 starts at 398.05 = table bottom 387.05 +
                    // 11); Oxi kept it above, -9.3pt for the rest of the page.
                    let s872_empty_cross = std::env::var("OXI_S872_DISABLE").is_err()
                        && (!self.doc_body_has_real_cjk || std::env::var_os("OXI_S1489_DISABLE").is_none())
                        && cursor.cursor_y < ft_top - 0.1
                        && matches!(block, Block::Paragraph(p)
                            if p.runs.iter().all(|r| r.text.is_empty())
                                && cursor.cursor_y
                                    + prev_space_after.max(p.style.space_before.unwrap_or(0.0))
                                    + self.estimate_para_height(p, self.s1211c_floor_body_width(p, content_width, page.grid_char_pitch, page.grid_char_cw_ratio), grid_pitch,
                                        None, false, None, None)
                                    > ft_top + 0.1);
                    // S1230 (2026-08-26, opt-out OXI_S1230_DISABLE): a TEXT line
                    // whose line BOX crosses the float band's top goes below the
                    // float — the JP text-line sibling of S872's Latin empty-box
                    // rule. Derived on kyotei36spec p3 (float form tblpY=268:
                    // Word keeps the title in the gap, box 48.8..60.3 <= band top
                    // ~60.55, and bumps 成立年月日, box 60.3..73.8, to 431.2 =
                    // band bottom) + _pb_floatband_gen COM arms (y268_noX: the
                    // following para's box 70.5..84.0 crosses the true band top
                    // 83.9 by 0.1pt -> Word puts it at 125.25 = below; y500/y1000:
                    // boxes that clear the band top stay in the gap). First-line
                    // height = the typed grid pitch when present (kyotei 11.5),
                    // else the paragraph estimate. EMPTY paragraphs keep the
                    // cursor-only rule (2ea81a's calibrated gap-empty balance —
                    // see the S872 scope note).
                    // 2026-08-26 (same day): first held opt-in — it exposed
                    // table3's +9pt height error (the bump landed at Oxi's
                    // inflated float bottom 440 vs Word 431, tail +12). S1231
                    // fixed that error (size-less empty cell ¶ = style-chain
                    // size); with it the bump lands at 433 vs Word 431.2 and
                    // the pair nets +0.052 on kyotei -> promoted default-ON.
                    let s1230_text_cross = std::env::var("OXI_S1230_DISABLE").is_err()
                        && self.doc_body_has_real_cjk
                        && cursor.cursor_y < ft_top - 0.1
                        && matches!(block, Block::Paragraph(p)
                            if p.runs.iter().any(|r| !r.text.is_empty())
                                && {
                                    let first_line_only = std::env::var_os("OXI_TEXT_FLOAT_FIRST_LINE_DISABLE").is_none(); // S1410
                                    let mut measure = cell_float::Measurement::default();
                                    let est = self.estimate_para_height_inner(
                                        p, self.s1211c_floor_body_width(p, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                        grid_pitch, None, false, None, None, false, first_line_only,
                                        if first_line_only { Some(&mut measure) } else { None },
                                    );
                                    // The collision box excludes paragraph spacing and later lines.
                                    // Spacing before is collapsed once by the caller below.
                                    let line_box = if first_line_only {
                                        measure.heights.first().copied().unwrap_or(est)
                                    } else { est };
                                    let l1 = match grid_pitch {
                                        Some(gp) if gp > 0.0 && p.style.snap_to_grid => line_box.min(gp),
                                        _ => line_box,
                                    };
                                    if std::env::var("OXI_DEBUG_TEXT_FLOAT_LINE").is_ok() {
                                        eprintln!("[TEXT_FLOAT_LINE] cursor={} before={} estimate={} first={} top={} bottom={} text={:?}",
                                            cursor.cursor_y, prev_space_after.max(p.style.space_before.unwrap_or(0.0)),
                                            est, l1, ft_top, ft_bot,
                                            p.runs.iter().flat_map(|r| r.text.chars()).take(30).collect::<String>());
                                    }
                                    cursor.cursor_y
                                        + prev_space_after
                                            .max(p.style.space_before.unwrap_or(0.0))
                                        + l1
                                        > ft_top + if first_line_only { 0.0 } else { 0.1 }
                                });
                    if s1509_keep && cursor.cursor_y < ft_bot {
                        // the empty line rides the lane; the region stays armed
                    } else if cursor.cursor_y >= ft_top - 0.1 && cursor.cursor_y < ft_bot {
                        cursor.set(ft_bot);
                        text_float_region = None;
                    } else if s1230_text_cross {
                        cursor.set(ft_bot);
                        text_float_region = None;
                    } else if s872_empty_cross {
                        cursor.set(ft_bot);
                        text_float_region = None;
                    } else if cursor.cursor_y >= ft_bot {
                        text_float_region = None;
                    }
                } else if current_page_idx != ft_page {
                    text_float_region = None;
                }
                // else: same page, non-overlapping lane (S1241) — keep the
                // region armed and do not bump.
            }
            // S1195: the lane beside a wrap-below float. An EMPTY paragraph is a
            // zero-width line, so Word fits it in the lane and it costs nothing
            // below the float; the first block with content resumes at the
            // float's bottom (or wherever the lane cursor has reached, if the
            // empties already carried it past). Measured on
            // `_pb_floatlane2_gen.py` (the lane floor) and on ed025cbecffb page 6
            // (one empty in a 28pt lane: Word puts the note directly under the
            // float, Oxi spent a line on the empty and ran 18pt low from there).
            if let Some((below_y, fpage)) = float_lane_below {
                if current_page_idx != fpage {
                    float_lane_below = None;
                } else {
                    let lane_ok = cursor.cursor_y < below_y
                        && matches!(block, Block::Paragraph(p)
                            if fits_blank_float_lane(p));
                    if !lane_ok {
                        if cursor.cursor_y < below_y {
                            cursor.set(below_y);
                        }
                        float_lane_below = None;
                    }
                }
            }
            // S560: on a fresh page the section-bottom tracker resets to the
            // top content origin (the deep value belongs to the prior page).
            if heterogeneous && current_page_idx != section_prev_page {
                section_max_y = start_y;
                section_prev_page = current_page_idx;
                col_band_top = start_y; // S749: band continues at the page top
            }
            // S560: switch column geometry at a section boundary within a
            // merged continuous-section page (only when heterogeneous, i.e.
            // the page mixes column counts — kyotei36spec's 1-col form table
            // followed by a continuous 2-col 記載心得 block). Word flows the
            // new section continuously below the previous section's content;
            // a 1-col section must NOT inherit the trailing 2-col geometry.
            let mut section_geometry_changed = false;
            if heterogeneous
                && active_run_idx + 1 < col_runs.len()
                && block_idx >= col_runs[active_run_idx + 1].0
            {
                while active_run_idx + 1 < col_runs.len()
                    && block_idx >= col_runs[active_run_idx + 1].0
                {
                    active_run_idx += 1;
                }
                let run_cols = col_runs[active_run_idx].1;
                // Only re-flow when the column COUNT changes; consecutive
                // same-count sections keep flowing in the current column.
                // S729: ALSO re-flow when the GEOMETRY changes (same column
                // count but different x/width — a continuous section with
                // different left/right margins, probexmargins).
                let run_geom_differs = col_runs[active_run_idx].2 != col_x_positions
                    || col_runs[active_run_idx].3 != col_widths;
                if run_cols != num_columns || run_geom_differs {
                    section_geometry_changed = true;
                    // S750 (2026-07-05, default ON, opt-out OXI_S750_DISABLE):
                    // Word BALANCES the columns of a continuous multi-column
                    // section on its FINAL page (equal column heights, splitting
                    // mid-paragraph) before the next continuous section flows
                    // below. Oxi's newspaper fill left col1 at full page depth
                    // (probexcont2col {+1:5} after S749). Post-hoc rebalance:
                    // the final band page's elements are still in `elements`,
                    // so move the tail rows of col1 to col2 until the columns
                    // are within one row of equal, then continue below the
                    // balanced band. v1 scope: 2 columns.
                    if num_columns == 2
                        && run_cols != num_columns
                        && std::env::var("OXI_S750_DISABLE").is_err()
                    {
                        let allocated_bottom = column_search.observe(
                            ir_index, allocation_start, col_runs[active_run_idx].0,
                            page, current_page_idx, col_band_top, start_y + content_height,
                            &col_widths, &elements, pending_section_gap);
                        if let Some(bottom) = if allocated_bottom.is_some() { allocated_bottom }
                        else if std::env::var("OXI_TEXT_BALANCE_DISABLE").is_err() {
                            LayoutEngine::rebalance_text_columns(&mut elements, col_band_top, &col_x_positions, &page.blocks, if std::env::var("OXI_COLUMN_COMPAT_BALANCE_DISABLE").is_err() { pending_section_gap } else { 0.0 }, self.compat_mode, if page.doc_grid_no_type { None } else { page.grid_line_pitch })
                        } else { None } {
                            if std::env::var("OXI_COLUMN_COMPAT_BALANCE_DISABLE").is_err() { pending_section_gap = 0.0; }
                            cursor.set(bottom);
                            section_max_y = bottom;
                        } else {
                        let col1_x0 = col_x_positions[0];
                        let col2_x0 = col_x_positions[1];
                        let split_x = col2_x0 - 1.0;
                        let mut rows1: Vec<f32> = Vec::new();
                        let mut rows2: Vec<f32> = Vec::new();
                        for e in elements.iter() {
                            if e.y < col_band_top - 0.1 {
                                continue;
                            }
                            let dst = if e.x < split_x {
                                &mut rows1
                            } else {
                                &mut rows2
                            };
                            let yk = (e.y * 10.0).round() / 10.0;
                            if !dst.iter().any(|v| (v - yk).abs() < 0.5) {
                                dst.push(yk);
                            }
                        }
                        rows1.sort_by(|a, b| a.partial_cmp(b).unwrap());
                        rows2.sort_by(|a, b| a.partial_cmp(b).unwrap());
                        let n1 = rows1.len();
                        let n2 = rows2.len();
                        // S1608 (2026-09-29): balancing cannot move content into a
                        // last column that already reaches the page bottom -- the
                        // balanced height is then the full column and Word leaves the
                        // band as filled. blind-G policies__1e87d3e6 p26: col 1 holds
                        // a table (55 distinct row y's) and col 2 runs to 784.5 of
                        // 785.2; S750 counted rows, moved 8 into col 2 over its own
                        // text and started the next 1-column section at 666.8 (Word:
                        // next page).
                        let s1608_col2_full = std::env::var_os("OXI_S1608_DISABLE").is_none() && {
                            let body_bottom = start_y + content_height;
                            let col2_bottom = elements.iter()
                                .filter(|e| e.y >= col_band_top - 0.1 && e.x >= split_x
                                    && matches!(e.content, LayoutContent::Text { .. }))
                                .map(|e| e.y + e.height)
                                .fold(f32::NEG_INFINITY, f32::max);
                            let pitch1 = {
                                let mut d: Vec<f32> = rows1.windows(2).map(|w| w[1] - w[0]).filter(|d| *d > 1.0).collect();
                                d.sort_by(|a, b| a.partial_cmp(b).unwrap());
                                d.get(d.len() / 2).copied().unwrap_or(0.0)
                            };
                            col2_bottom > body_bottom - pitch1
                        };
                        if std::env::var("OXI_DBG_COL").is_ok() && s1608_col2_full {
                            eprintln!("[COL] S1608 col2 full -> no S750 balance (n1={} n2={})", n1, n2);
                        }
                        if n1 >= 2 && n1 > n2 + 1 && !s1608_col2_full {
                            // uniform row pitch from col1's row diffs (median)
                            let mut diffs: Vec<f32> = rows1
                                .windows(2)
                                .map(|w| w[1] - w[0])
                                .filter(|d| *d > 1.0)
                                .collect();
                            diffs.sort_by(|a, b| a.partial_cmp(b).unwrap());
                            let pitch = diffs.get(diffs.len() / 2).copied().unwrap_or(0.0);
                            if pitch > 1.0 {
                                // S1338 (2026-09-06, default ON, opt-out OXI_S1338_DISABLE):
                                // Word balances GRID LINES, not rows -- a row twice the
                                // pitch tall (0ea3ec86 p3's 「❖ 障害者に関するマーク」
                                // heading, 42pt over a 20.55 pitch) counts as two. That
                                // section (heading + 6 lines) is 3 rows / 4 rows in
                                // Word's PDF (heading + 2 | 4 = 4 lines each); by rows
                                // it was 4 / 3. Every other balanced band of 0ea3ec86,
                                // 167853 and 0b6f3b32 (the `_colbalance_census.py`
                                // sweep, ~50 bands) has single-line rows, where lines
                                // and rows agree. A row's lines = floor(height / pitch),
                                // at least 1; column 1 keeps the first rows whose lines
                                // reach ceil(total / 2).
                                // HELD OPT-IN (`OXI_S1338=1`) 2026-09-06: kyotei36spec's
                                // band (46 rows / 15, rows separated by table gaps) reads
                                // so many "lines" in column 2 that the target is never
                                // reached inside column 1 and NOTHING moves (SSIM -0.0479
                                // against the row count's k=15). Row height is not line
                                // count where rows are table cells; weight only rows that
                                // are text lines before defaulting this.
                                let k = if std::env::var("OXI_S1338_DISABLE").is_err() {
                                    // a row's height runs from the previous row's y (the
                                    // band top for the first row of a column, so a
                                    // heading's space-before counts) to its own y
                                    let mut lines: Vec<usize> = Vec::new();
                                    for col in [&rows1, &rows2] {
                                        let mut prev = col_band_top;
                                        for (i, &y) in col.iter().enumerate() {
                                            let h = if i + 1 < col.len() {
                                                col[i + 1] - y
                                            } else {
                                                (y - prev).max(pitch)
                                            };
                                            let h = if i == 0 { (col[1.min(col.len() - 1)] - col_band_top).max(h) } else { h };
                                            // a row more than 2.5 pitches tall is a table gap or
                                            // an image, not lines (kyotei36spec): count it once
                                            let l = if h <= 2.5 * pitch {
                                                ((h / pitch + 0.1).floor() as usize).max(1)
                                            } else {
                                                1
                                            };
                                            lines.push(l);
                                            prev = y;
                                        }
                                    }
                                    let total: usize = lines.iter().sum();
                                    let target = (total + 1) / 2;
                                    let mut cum = 0usize;
                                    let mut keep = 0usize;
                                    for (i, &l) in lines.iter().enumerate() {
                                        cum += l;
                                        if cum >= target {
                                            keep = i + 1;
                                            break;
                                        }
                                    }
                                    n1.saturating_sub(keep.max(1))
                                } else {
                                    (n1 - n2) / 2
                                };
                                if k > 0 {
                                    let cut_y = rows1[n1 - k] - 0.5;
                                    let dx = col2_x0 - col1_x0;
                                    let col2_next = col_band_top + n2 as f32 * pitch;
                                    let dy = col2_next - rows1[n1 - k];
                                    // S750b (2026-07-15): Word ROW-ALIGNS the balanced
                                    // columns — col2 row i sits at the same Y as col1
                                    // row (n2+i), not a UNIFORM-pitch offset (which
                                    // misaligns when the source rows vary in height:
                                    // forms' drug/indication lists mix 14.5pt items
                                    // and 17pt headings, so the median-pitch dy drifted
                                    // ~3pt/row). Map each moved row to its target col1
                                    // row Y (preserving any intra-row element offset);
                                    // fall back to the uniform dy when the row index is
                                    // ambiguous or the target is missing.
                                    let moved_start = n1 - k;
                                    for e in elements.iter_mut() {
                                        if e.y >= cut_y
                                            && e.x < split_x
                                            && e.y >= col_band_top - 0.1
                                        {
                                            e.x += dx;
                                            let yk = (e.y * 10.0).round() / 10.0;
                                            let ri =
                                                rows1.iter().position(|&r| (r - yk).abs() < 0.5);
                                            match ri.and_then(|ri| {
                                                let j = ri.checked_sub(moved_start)?;
                                                rows1.get(n2 + j).map(|ty| (ri, *ty))
                                            }) {
                                                Some((ri, ty)) => {
                                                    e.y = ty + (e.y - rows1[ri]);
                                                }
                                                None => {
                                                    e.y += dy;
                                                }
                                            }
                                        }
                                    }
                                    let new_len = ((n1 - k).max(n2 + k)) as f32 * pitch;
                                    let balanced_bottom = col_band_top + new_len;
                                    if std::env::var("OXI_DBG_COL").is_ok() {
                                        eprintln!("[COL] S750 balance n1={} n2={} k={} pitch={:.1} bottom={:.1}", n1, n2, k, pitch, balanced_bottom);
                                    }
                                    cursor.set(balanced_bottom);
                                    section_max_y = balanced_bottom;
                                }
                            }
                        }
                        }
                    }
                    // New section continues below ALL columns of the section
                    // it succeeds (continuous flow, same page if room).
                    section_max_y = section_max_y.max(cursor.cursor_y);
                    cursor.set(section_max_y);
                    allocation_start = col_runs[active_run_idx].0;
                    let run = &col_runs[active_run_idx];
                    num_columns = run.1;
                    col_x_positions = run.2.clone();
                    col_widths = run.3.clone();
                    current_column = 0;
                    start_x = col_x_positions[0];
                    content_width = col_widths[0];
                    section_max_y = cursor.cursor_y;
                    // S749: the multi-col band starts HERE; column advances on
                    // this page return to this y, not the page top.
                    col_band_top = if num_columns > 1 && std::env::var("OXI_S749_DISABLE").is_err()
                    {
                        cursor.cursor_y
                    } else {
                        start_y
                    };
                    if std::env::var("OXI_DBG_COL").is_ok() {
                        eprintln!("[COL] SWITCH block_idx={} -> ncol={} at page={} cursor_y={:.1} band_top={:.1}", block_idx, num_columns, current_page_idx, cursor.cursor_y, col_band_top);
                    }
                }
            }
            if pending_section_gap != 0.0 {
                cursor.advance(pending_section_gap);
                section_max_y = section_max_y.max(cursor.cursor_y);
                if section_geometry_changed && num_columns > 1 {
                    col_band_top = cursor.cursor_y;
                }
                pending_section_gap = 0.0;
            }
            // S469: the wrap-below anchor offset is page-local. Reset it when the
            // flow has advanced to a new page since the previous block.
            if current_page_idx != anchor_offset_page {
                anchor_flow_offset = 0.0;
                anchor_offset_page = current_page_idx;
            }
            // wrapTopAndBottom: for inline TABLE blocks, push below overlapping TextBoxes
            // Skip for floating tables (tblpPr) as they have explicit positioning
            let is_floating_table = matches!(block, Block::Table(t) if t.style.position.is_some());
            if matches!(block, Block::Table(_)) && !is_floating_table {
                for tb in &page.text_boxes {
                    // Skip wrapNone text boxes (they don't affect text flow)
                    if tb.wrap_type == Some(crate::ir::WrapType::None) {
                        continue;
                    }
                    if tb.anchor_block_index < block_idx {
                        // S1437 (2026-09-16, default ON, opt-out OXI_S1437_DISABLE): the
                        // OXI_TABLE_WRAP_PAGE_SCOPE opt-in promoted. Without it a wrap
                        // text box anchored to an earlier block on ANOTHER page, whose
                        // page-local y-span happens to contain the cursor, pushed a
                        // table below it (reports__28abf02c p3: the p2 answer-bracket
                        // boxes at 467..586 shoved the (16) table 119pt down, one
                        // page over for 10 paragraphs).
                        if std::env::var("OXI_TABLE_WRAP_PAGE_SCOPE").is_ok()
                            || std::env::var_os("OXI_S1437_DISABLE").is_none() {
                            // A wrap obstacle belongs to the page on which it is drawn.
                            let anchor_pages = if std::env::var("OXI_S1123_DISABLE").is_err() {
                                &block_start_page_indices
                            } else {
                                &block_page_indices
                            };
                            let flow_page = if tb.wrap_type == Some(crate::ir::WrapType::TopAndBottom)
                                && tb.position.as_ref().and_then(|p| p.v_relative.as_deref()) == Some("paragraph")
                            {
                                s1089_flow_pos.get(&tb.anchor_block_index).map(|v| v.0)
                            } else { None };
                            let obstacle_page = flow_page.or_else(|| anchor_pages.get(tb.anchor_block_index).copied()).unwrap_or(0);
                            if obstacle_page != current_page_idx {
                                continue;
                            }
                        }
                        if let Some(ref pos) = tb.position {
                            let anchor_y = block_y_positions
                                .get(tb.anchor_block_index)
                                .copied()
                                .unwrap_or(0.0);
                            let tb_top = match pos.v_relative.as_deref() {
                                Some("paragraph") | Some("line") => anchor_y + pos.y,
                                Some("margin") => page.margin.top + pos.y,
                                Some("page") => pos.y,
                                _ => anchor_y + pos.y,
                            };
                            let tb_bottom = tb_top + tb.height;
                            if cursor.cursor_y >= tb_top && cursor.cursor_y < tb_bottom {
                                cursor.set(tb_bottom);
                            }
                        }
                    }
                }
            }
            // S1294 (2026-09-03, opt-out OXI_S1294_DISABLE): a merged CONTINUOUS
            // section that declares `<w:pgNumType w:start>` and BEGINS a page is
            // subject to the same alternation rule as any other section start
            // (S1291). Whether it is `continuous` is not the question -- whether
            // it starts a page is.
            //
            // Derived by taking `reference__0ea3ec86480140c2` apart one attribute
            // at a time (`_pb_0ea3_bisect.py`, Word PDF page census): its blank
            // page 2 needs evenAndOddHeaders AND section 1's start=88, and
            // section 2's own start only MOVES the blank. Then the h arms of
            // `_pb_blankpage_gen.py` swept the cover length:
            //
            //   cover 48 lines (fits page 1)   conflict -> no pad, sec2 mid-page
            //   cover 50 lines (FILLS page 1)  conflict -> blank p2, sec2 on p3
            //   cover 50 lines                 agree    -> no blank, sec2 on p2
            //   cover 52 lines (overflows)     either   -> no pad, sec2 mid-page 2
            //
            // Only the 50-line arm discriminates. A restart landing MID-page is
            // ignored outright -- Word's own logical numbers say so: 0ea3ec86's
            // section 3 restarts at 88 and its page still reports 90, the number
            // of the section that STARTED it.
            //
            // The decision is taken AFTER the block is laid out (see the end of
            // this loop): "does this section begin a page" is not knowable
            // before, because the section's own first block is what pushes the
            // page. 0ea3ec86's section 1 fills page 1 exactly and section 2's
            // first block overflows onto page 2 -- checking beforehand reads
            // "mid-page" and pads nothing.
            let s1294_restart: Option<u32> = if s1294_on && block_idx > 0 {
                page.page_number_runs
                    .iter()
                    .find(|(b, st)| *b == block_idx && st.is_some())
                    .and_then(|(_, st)| *st)
            } else {
                None
            };
            let s1294_pages_before = pages.len();
            let s1294_at_top_before =
                elements.is_empty() && cursor.cursor_y <= start_y + 0.01;
            // S469: record the NATURAL (pre-wrap) Y for anchor resolution by
            // subtracting any accumulated wrap-below advance on this page.
            let natural_anchor_y = shared_float_anchors.get(&block_idx)
                .filter(|(pg, _)| *pg == current_page_idx)
                .map_or(cursor.cursor_y, |(_, y)| *y);
            let anchor_spacing = if (std::env::var("OXI_PARAGRAPH_FLOAT_SPACING").is_ok()
                || std::env::var("OXI_S1471_DISABLE").is_err()) {
                if let Block::Paragraph(para) = block {
                    self.paragraph_spacing_before(
                        para, page, grid_pitch, prev_para_style_id.as_deref(),
                        prev_contextual_spacing, prev_autospacing_numid.as_deref(),
                        prev_space_after, Some(block_idx), &pages, &elements,
                        natural_anchor_y, start_y,
                    ).0
                } else { 0.0 }
            } else { 0.0 };
            block_y_positions.push(natural_anchor_y + anchor_spacing - anchor_flow_offset);
            block_col_x.push(start_x); // S1222: the current column's left edge
            block_page_indices.push(current_page_idx);
            block_start_page_indices.push(current_page_idx);
            // What a float ANCHORED to this block will resolve against. A
            // floating box lands where these two say, so when a box appears on
            // the wrong page this is the pair to read first.
            if std::env::var("OXI_DBG_ANCHOR").is_ok() {
                let (preview, nruns, pbb, pba) = match block {
                    Block::Paragraph(p) => (
                        p.runs
                            .iter()
                            .flat_map(|r| r.text.chars())
                            .take(14)
                            .collect::<String>(),
                        p.runs.len(),
                        p.style.page_break_before,
                        p.style.page_break_after,
                    ),
                    _ => ("<non-paragraph>".to_string(), 0, false, false),
                };
                eprintln!(
                    "[ANCHOR] blk={} page={} y={:.2} flow_off={:.2} runs={} pbb={} pba={} text={:?}",
                    block_start_page_indices.len() - 1,
                    current_page_idx,
                    cursor.cursor_y,
                    anchor_flow_offset,
                    nruns,
                    pbb,
                    pba,
                    preview
                );
            }
            match block {
                Block::Paragraph(para) => {
                    // S945 (2026-07-19): an EMPTY section-ending paragraph (in-body
                    // sectPr carrier) contributes NOTHING — including its style's
                    // pageBreakBefore/keepNext pre-pushes (NDIS 0043bfe0: an empty
                    // Heading1+sectPr para's style pageBreakBefore manufactured a
                    // phantom page before every chapter section). It is the last
                    // block before the next section. Merged continuous sections
                    // still need its style identity for neighbouring spacing.
                    // S1501 v2: the carrier of a nextPage section that holds only
                    // that carrier becomes block 0 of its page run (the following
                    // continuous section merges behind it and S730 marks it
                    // continuous) -- Word gives that mark a line at the page top.
                    let s1501_keep = std::env::var_os("OXI_S1501_DISABLE").is_none()
                        && para.style.page_section_break
                        && para.runs.iter().all(|r| r.text.is_empty())
                        && block_idx == 0
                        && cursor.cursor_y <= start_y + 0.01;
                    if std::env::var_os("OXI_DBG1501").is_some() && para.style.page_section_break {
                        eprintln!("[S1501] block={} keep={} cursor={:.2} start_y={:.2} cont={} empty={}", block_idx, s1501_keep, cursor.cursor_y, start_y, para.style.continuous_section_break, para.runs.iter().all(|r| r.text.is_empty()));
                    }
                    // S1576 (2026-09-26, default ON, opt-out OXI_S1576_DISABLE): the
                    // carrier of a CONTINUOUS section break keeps its mark line when a
                    // TABLE follows (policies__0084b6ad p5/6: an empty TNR-12 carrier
                    // between two tables; Word truth puts it at 242.2 and the next table
                    // at 256.5, Oxi dropped it and the page ran 16pt short). S945 was
                    // derived on a nextPage carrier.
                    let s1576_keep = std::env::var_os("OXI_S1576_DISABLE").is_none()
                        && para.style.continuous_section_break
                        && (matches!(page.blocks.get(block_idx + 1), Some(Block::Table(_)))
                            // A continuous carrier after a table is also a
                            // real paragraph line, even when text follows it.
                            || block_idx.checked_sub(1).is_some_and(|i|
                                matches!(page.blocks.get(i), Some(Block::Table(_)))));
                    S1501_KEEP.with(|c| c.set(s1501_keep || s1576_keep));
                    // A terminal empty continuous mark may exhaust this page,
                    // but does not become a blank line on the following page.
                    // Leave the transition to the incoming section's content.
                    if s1576_keep && para.runs.iter().all(|r| r.text.is_empty())
                        && !para.style.page_break_after && !para.style.page_break_before
                        && block_idx.checked_sub(1).is_some_and(|i|
                            matches!(page.blocks.get(i), Some(Block::Table(_))))
                    {
                        let mark_height = self.estimate_para_height(
                            para, content_width, grid_pitch, None, false,
                            page.grid_char_pitch, page.grid_char_cw_ratio,
                        );
                        if cursor.cursor_y + mark_height > start_y + content_height {
                            cursor.advance(mark_height);
                            continue;
                        }
                    }
                    if std::env::var("OXI_S945_DISABLE").is_err()
                        && !s1501_keep
                        && !s1576_keep
                        && !(para.style.page_break_after
                            && (std::env::var("OXI_SECTION_EXPLICIT_BREAKS").is_ok()
                                || (para.style.continuous_section_break
                                    && std::env::var("OXI_S1454_DISABLE").is_err())))
                        && para.style.page_section_break
                        && para.runs.iter().all(|r| r.text.is_empty())
                    {
                        if para.style.continuous_section_break {
                            if !prev_contextual_spacing
                                || prev_para_style_id != para.style.style_id
                            {
                                pending_section_gap += prev_space_after;
                            }
                            prev_space_after = 0.0;
                            prev_para_style_id = para.style.style_id.clone();
                            prev_contextual_spacing = para.style.contextual_spacing;
                        }
                        continue;
                    }
                    // S676 (2026-06-27): drop-cap float. A paragraph whose framePr is
                    // dropCap="drop" is a floating drop cap — Word renders its glyph at
                    // the left, anchored to the FOLLOWING paragraph's top, WITHOUT
                    // reserving vertical block space; the next paragraph's body wraps to
                    // the right (indented by the cap width). Oxi previously laid it out as
                    // a normal full-height block (the +61pt over-reservation the
                    // perturbation harness flagged). Gate-safe: 0 corpus docs use dropCap,
                    // so non-dropCap paras never enter this branch (byte-identical).
                    if std::env::var("OXI_S676_DISABLE").is_err()
                        && pending_dropcap.is_none()
                        && para
                            .style
                            .frame_pr
                            .as_ref()
                            .map(|fp| fp.drop_cap.as_deref() == Some("drop"))
                            .unwrap_or(false)
                    {
                        let cap_text: String =
                            para.runs.iter().flat_map(|r| r.text.chars()).collect();
                        if !cap_text.trim().is_empty() {
                            if let Some(first) = para.runs.iter().find(|r| !r.text.is_empty()) {
                                let fs = self.resolve_font_size(&first.style, &para.style);
                                let metrics = self.metrics_for(&first.style, &para.style);
                                let cap_w: f32 = cap_text
                                    .chars()
                                    .map(|c| {
                                        self.registry.char_width_pt_with_fallback(c, fs, metrics)
                                    })
                                    .sum();
                                let family = self
                                    .resolve_font_family_for_text(
                                        &cap_text,
                                        &first.style,
                                        &para.style,
                                    )
                                    .map(|s| s.to_string());
                                let color = self
                                    .resolve_color(&first.style, &para.style)
                                    .map(|s| s.to_string());
                                // Anchor the cap top to the body's top (current cursor).
                                let cap_y = cursor.cursor_y;
                                elements.push(LayoutElement::new(
                                    start_x,
                                    cap_y,
                                    cap_w,
                                    fs * 1.2,
                                    LayoutContent::Text {
                                        text: cap_text.clone(),
                                        font_size: fs,
                                        font_family: family,
                                        bold: first.style.bold,
                                        italic: first.style.italic,
                                        underline: false,
                                        underline_style: None,
                                        strikethrough: false,
                                        double_strikethrough: false,
                                        color,
                                        highlight: None,
                                        field_type: None,
                                        character_spacing: 0.0,
                                        text_scale: 100.0,
                                        is_vertical: false,
                                        effects: TextEffects::default(),
                                    },
                                ));
                                // Body indent = cap width + the framePr hSpace gap.
                                let h_space = para
                                    .style
                                    .frame_pr
                                    .as_ref()
                                    .map(|fp| fp.h_space)
                                    .unwrap_or(0.0)
                                    .max(0.0);
                                pending_dropcap = Some(cap_w + h_space);
                            }
                        }
                        // Float: do NOT advance the cursor; skip normal block processing.
                        continue;
                    }
                    let same_text_frame = |a: &crate::ir::FrameProperties, b: &crate::ir::FrameProperties| {
                        a.v_anchor.as_deref() == Some("text") && b.v_anchor == a.v_anchor
                            && a.x == b.x && a.y == b.y && a.width == b.width
                            && a.height == b.height && a.height_rule == b.height_rule
                            && a.h_anchor == b.h_anchor && a.x_align == b.x_align
                            && a.y_align == b.y_align && a.wrap == b.wrap
                            && a.h_space == b.h_space && a.v_space == b.v_space
                            && a.drop_cap == b.drop_cap && a.lines == b.lines
                    };
                    let grouped_text_frame = para.style.frame_pr.as_ref().is_some_and(|fp| {
                        fp.drop_cap.is_none() && fp.width.is_some() && fp.wrap.as_deref() != Some("none")
                            && [block_idx.checked_sub(1), Some(block_idx + 1)].into_iter().flatten()
                                .any(|i| matches!(page.blocks.get(i), Some(Block::Paragraph(p))
                                    if p.style.frame_pr.as_ref().is_some_and(|n| same_text_frame(fp, n))))
                    });
                    // S863: Word groups consecutive paragraphs that repeat the
                    // same exact-height text-anchored framePr into one frame.
                    // The frame top is relative to the first paragraph's flow
                    // anchor; its declared height is consumed only once.
                    let s863_fp = para.style.frame_pr.as_ref().filter(|fp| {
                        fp.drop_cap.is_none()
                            && fp.v_anchor.as_deref() == Some("text")
                            && fp.h_anchor.as_deref() == Some("page")
                            && fp.height_rule.as_deref() == Some("exact")
                            && fp.height.unwrap_or(0.0) > 0.0
                            && fp.y < 0.0
                            && fp.wrap.as_deref() == Some("around")
                    });
                    let s863_matches = |other: &Block, fp: &crate::ir::FrameProperties| {
                        matches!(other, Block::Paragraph(p) if p.style.frame_pr.as_ref().map_or(false, |n|
                            n.drop_cap.is_none()
                                && n.v_anchor.as_deref() == Some("text")
                                && n.h_anchor.as_deref() == Some("page")
                                && n.height_rule.as_deref() == Some("exact")
                                && n.wrap.as_deref() == Some("around")
                                && (n.x - fp.x).abs() < 0.1
                                && (n.y - fp.y).abs() < 0.1
                                && (n.height.unwrap_or(0.0) - fp.height.unwrap_or(0.0)).abs() < 0.1
                        ))
                    };
                    let s863_in_run = s863_fp.map_or(false, |fp| {
                        s863_frame.map_or(false, |(x, y, _, _, _, _)| {
                            (x - fp.x).abs() < 0.1 && (y - fp.y).abs() < 0.1
                        }) || page
                            .blocks
                            .get(block_idx + 1)
                            .map_or(false, |b| s863_matches(b, fp))
                    });
                    if std::env::var("OXI_S863_DISABLE").is_err() && s863_in_run {
                        let fp = s863_fp.unwrap();
                        let (fx, fy, frame_bottom, anchor_y) = match s863_frame {
                            Some((_, _, first_x, running_bottom, exact_bottom, anchor)) => {
                                (first_x, running_bottom, exact_bottom, anchor)
                            }
                            None => {
                                let anchor = cursor.cursor_y;
                                let top = anchor + fp.y;
                                (fp.x, top, top + fp.height.unwrap_or(0.0), anchor)
                            }
                        };
                        let fw = fp
                            .width
                            .unwrap_or(page.size.width - fx - page.margin.right)
                            .max(20.0);
                        let mut fcy = LayoutCursor::new(fy);
                        let empty_fn_frame = std::collections::HashMap::new();
                        let (frame_els, _, _) = self.layout_paragraph(
                            para,
                            fx,
                            &mut fcy,
                            fw,
                            page.size.height,
                            fy,
                            page,
                            &mut Vec::new(),
                            &mut Vec::new(),
                            grid_pitch,
                            None,
                            false,
                            None,
                            None,
                            false,
                            false,
                            0.0,
                            Some(block_idx),
                            None,
                            None,
                            false,
                            false,
                            None,
                            0.0,
                            &empty_fn_frame,
                            1,
                            0,
                            &[],
                            0.0,
                            false,
                            false,
                            None,
                            None,          // S758
                            None,          // S-TWOSEG
                            false,
                            0.0,
                            None,  // S900
                            None,  // S903
                            false, // S916
                            None,
                        );
                        elements.extend(frame_els);
                        let has_next = page
                            .blocks
                            .get(block_idx + 1)
                            .map_or(false, |b| s863_matches(b, fp));
                        if std::env::var("OXI_DBG863").is_ok() {
                            eprintln!(
                                "[S863] blk={} anchor={:.2} y={:.2} bottom={:.2} next={}",
                                block_idx, anchor_y, fy, fcy.cursor_y, has_next
                            );
                        }
                        if has_next {
                            s863_frame =
                                Some((fp.x, fp.y, fx, fcy.cursor_y, frame_bottom, anchor_y));
                        } else {
                            cursor.set(anchor_y.max(frame_bottom));
                            s863_frame = None;
                        }
                        prev_space_after = 0.0;
                        continue;
                    }
                    s863_frame = None;
                    // S898a (2026-07-17): a vAnchor="text" framePr with NEGATIVE y
                    // reaches ABOVE its anchor — that region is already laid out,
                    // so Word FLOATS the frame there (no flow consumption). The
                    // 00054c43 state-seal frame (y=-951tw, an inline 90.75pt image)
                    // renders at anchor-47.55 = Word's 24.75 while the body flows
                    // on; Oxi's in-flow path reserved ~98pt (the whole +65 p1
                    // drift the missing style-autospacing was compensating).
                    // Exact-height "around" runs are consumed by S863 above.
                    if std::env::var("OXI_S898_DISABLE").is_err()
                        && para.style.frame_pr.as_ref().map_or(false, |fp| {
                            fp.drop_cap.is_none()
                                && fp.v_anchor.as_deref() == Some("text")
                                && fp.y < -0.01 && !grouped_text_frame
                        })
                    {
                        let fp = para.style.frame_pr.as_ref().unwrap();
                        let fx = if fp.h_anchor.as_deref() == Some("page") {
                            fp.x
                        } else {
                            page.margin.left + fp.x
                        };
                        let fy = (cursor.cursor_y + fp.y).max(1.0);
                        let fw = fp.width.unwrap_or(content_width * 0.3).max(20.0);
                        let mut fcy = LayoutCursor::new(fy);
                        let empty_fn_frame = std::collections::HashMap::new();
                        let (frame_els, _, _) = self.layout_paragraph(
                            para,
                            fx,
                            &mut fcy,
                            fw,
                            page.size.height,
                            fy,
                            page,
                            &mut Vec::new(),
                            &mut Vec::new(),
                            grid_pitch,
                            None,
                            false,
                            None,
                            None,
                            false,
                            false,
                            0.0,
                            Some(block_idx),
                            None,
                            None,
                            false,
                            false,
                            None,
                            0.0,
                            &empty_fn_frame,
                            1,
                            0,
                            &[],
                            0.0,
                            false,
                            false,
                            None,
                            None,          // S758
                            None,          // S-TWOSEG
                            false,
                            0.0,
                            None,  // S900
                            None,  // S903
                            false, // S916
                            None,
                        );
                        elements.extend(frame_els);
                        if std::env::var("OXI_DBG898").is_ok() {
                            eprintln!(
                                "[S898a] blk={} anchor={:.2} fy={:.2} bottom={:.2}",
                                block_idx, cursor.cursor_y, fy, fcy.cursor_y
                            );
                        }
                        prev_space_after = 0.0;
                        continue;
                    }
                    // S758c (2026-07-06): a PAGE-anchored framePr paragraph is a
                    // fixed-position frame — the body flows AROUND it (side-wrap
                    // band) and the frame consumes NO flow height (probeqframepg:
                    // Word frame box [340..476.7]×[225..338] with body lines
                    // narrowed to x1≈333 beside it; Oxi laid it in-flow full-width
                    // → −1×3). The vAnchor="text" frames keep the existing
                    // in-flow path (probexframes PASSES with it).
                    // S1379 (2026-09-13, default ON, opt-out OXI_S1379_DISABLE): a
                    // framePr paragraph with yAlign=top|bottom is a fixed frame on
                    // the margin box (vAnchor absent or "margin") or on the page
                    // (vAnchor="page"); consecutive same-key frame paragraphs form
                    // ONE group, stacked at the line pitch with no inset, whose top
                    // sits on the reference top (yAlign=top) or whose bottom sits on
                    // the reference bottom (yAlign=bottom). The body flow never
                    // moves. MEASURED (`_pb_frame_ybottom_gen.py`, Word COM + PDF,
                    // three 9pt Calibri frame paragraphs, A4 margins 72/72):
                    //   yAlign=bottom hAnchor=page x=8971  737.25/748.5/759.0 @448.5
                    //                                      (bottom 770 = margin bottom)
                    //   + vAnchor=page                    809.25/820.5/831.0 (page bottom 842)
                    //   yAlign=top                        72.0/83.25/93.75 (margin top)
                    //   hAnchor=margin x=2000             same y, x = 72 + 100
                    //   one paragraph                     759.0
                    //   body lines 72/85.5/99/... in every arm (no flow consumption).
                    // Real witness: correspondence__0059143bed49147b, a 19-paragraph
                    // address column in the right margin (Word 612..792 @448.5);
                    // Oxi laid it in the body flow and every body line sat 189pt low.
                    let s1379_fp = |fp: &crate::ir::FrameProperties| -> bool {
                        std::env::var("OXI_S1379_DISABLE").is_err()
                            && fp.drop_cap.is_none()
                            && fp.wrap.as_deref() != Some("none")
                            && matches!(fp.y_align.as_deref(), Some("top") | Some("bottom"))
                            && matches!(fp.v_anchor.as_deref(), None | Some("margin") | Some("page"))
                    };
                    let s1379 = para.style.frame_pr.as_ref().map_or(false, |fp| s1379_fp(fp));
                    let positioned_frame = para.style.frame_pr.as_ref().is_some_and(|fp| {
                        fp.drop_cap.is_none() && fp.width.is_some()
                            && fp.wrap.as_deref() != Some("none")
                            && fp.y_align.is_none()
                            && (fp.v_anchor.as_deref() == Some("margin")
                                || (fp.v_anchor.as_deref() == Some("text") && (fp.y >= 0.0 || grouped_text_frame)))
                    });
                    if std::env::var("OXI_S758_DISABLE").is_err()
                        && (positioned_frame || s1379 || para.style.frame_pr.as_ref().map_or(false, |fp| {
                            fp.drop_cap.is_none()
                                && (fp.v_anchor.as_deref() == Some("page")
                                    || (std::env::var_os("OXI_FRAME_BOTTOM_MARGIN").is_some()
                                        && matches!(fp.v_anchor.as_deref(), None | Some("margin"))
                                        && fp.y_align.as_deref() == Some("bottom")))
                                && fp.wrap.as_deref() != Some("none")
                        }))
                    {
                        let fp = para.style.frame_pr.as_ref().unwrap();
                        // S847: match the group on the DECLARED (x, y). A
                        // continuation paragraph inherits the FIRST para's fx
                        // (its own hAnchor is often omitted → mis-computed) and
                        // starts at the previous frame para's bottom.
                        let fx0_decl = if fp.h_anchor.as_deref() == Some("page") {
                            fp.x
                        } else {
                            page.margin.left + fp.x
                        };
                        let fy_decl = if s1379 {
                            let ya = fp.y_align.as_deref().unwrap_or("").to_string();
                            match s1379_group {
                                Some((gx, ref ga, t))
                                    if s847_frame.is_some()
                                        && (gx - fp.x).abs() < 0.1
                                        && *ga == ya =>
                                {
                                    t
                                }
                                _ => {
                                    // Dry-run every consecutive same-key member for
                                    // the group height, then anchor the group.
                                    let (ref_top, ref_bottom) =
                                        if fp.v_anchor.as_deref() == Some("page") {
                                            (0.0, page.size.height)
                                        } else {
                                            (page.margin.top, page.size.height - page.margin.bottom)
                                        };
                                    let mut h_total = 0.0f32;
                                    let mut j = block_idx;
                                    while let Some(Block::Paragraph(q)) = page.blocks.get(j) {
                                        let n = match q.style.frame_pr.as_ref() {
                                            Some(n)
                                                if s1379_fp(n)
                                                    && (n.x - fp.x).abs() < 0.1
                                                    && n.y_align == fp.y_align
                                                    && n.h_anchor == fp.h_anchor
                                                    && n.v_anchor == fp.v_anchor =>
                                            {
                                                n
                                            }
                                            _ => break,
                                        };
                                        let qw = n.width.unwrap_or(content_width * 0.3).max(20.0);
                                        let mut dcy = LayoutCursor::new(0.0);
                                        let empty_fn_dry = std::collections::HashMap::new();
                                        let _ = self.layout_paragraph(
                                            q,
                                            fx0_decl,
                                            &mut dcy,
                                            qw,
                                            page.size.height,
                                            0.0,
                                            page,
                                            &mut Vec::new(),
                                            &mut Vec::new(),
                                            grid_pitch,
                                            None,
                                            false,
                                            None,
                                            None,
                                            false,
                                            false,
                                            0.0,
                                            Some(j),
                                            None,
                                            None,
                                            false,
                                            false,
                                            None,
                                            0.0,
                                            &empty_fn_dry,
                                            1,
                                            0,
                                            &[],
                                            0.0,
                                            false,
                                            false,
                                            None,  // S755
                                            None,  // S758
                                            None,  // S-TWOSEG
                                            false, // S835
                                            0.0,
                                            None,  // S900
                                            None,  // S903
                                            false, // S916
                                            None,
                                        );
                                        h_total += dcy.cursor_y.max(0.0);
                                        j += 1;
                                    }
                                    let t = if ya == "top" {
                                        ref_top
                                    } else {
                                        (ref_bottom - h_total).max(ref_top)
                                    };
                                    if std::env::var("OXI_DBG1379").is_ok() {
                                        eprintln!("[S1379] blk={}..{} yalign={} ref=({:.1},{:.1}) h={:.2} top={:.2}",
                                            block_idx, j, ya, ref_top, ref_bottom, h_total, t);
                                    }
                                    s1379_group = Some((fp.x, ya, t));
                                    t
                                }
                            }
                        } else if positioned_frame {
                            fp.y + if fp.v_anchor.as_deref() == Some("text") { cursor.cursor_y } else { page.margin.top }
                        } else {
                            fp.y
                        };
                        let aligned_bottom = !s1379 && std::env::var_os("OXI_FRAME_BOTTOM_MARGIN").is_some()
                            && fp.y_align.as_deref() == Some("bottom")
                            && matches!(fp.v_anchor.as_deref(), None | Some("margin") | Some("page"));
                        // S847 (opt-in OXI_S847=1, default OFF = byte-identical):
                        // consecutive same-(x,y) page frames stack vertically
                        // (Word groups them into one frame; continuations
                        // inherit the first para's fx). Structurally correct
                        // (title block matches Word) but HELD opt-in — on the
                        // sole affected corpus doc (correspondence Massachusetts
                        // letterhead) it nets −0.043 SSIM: the stacked title's
                        // lower lines drift sub-pixel vs Word's line heights AND
                        // the officials block (a separate vAnchor="text"
                        // negative-y frame + flow paras) is still unpositioned.
                        // Ships default-ON once frame line-height precision + the
                        // vAnchor=text officials frame are also handled.
                        // S898b: notBeside frames MUST stack (the flow-push needs
                        // the true group bottom), so grouping is default-ON for
                        // them; other page frames keep the S847 opt-in hold.
                        let s847_on = grouped_text_frame || aligned_bottom || std::env::var("OXI_S847").is_ok()
                            || s1379
                            || (std::env::var("OXI_S898_DISABLE").is_err()
                                && fp.wrap.as_deref() == Some("notBeside"));
                        let (fx, fy) = match s847_frame {
                            Some((px, py, first_fx, run_bottom))
                                if s847_on
                                    && (px - fp.x).abs() < 0.1
                                    && (py - fy_decl).abs() < 0.1 =>
                            {
                                (first_fx, run_bottom)
                            }
                            _ => {
                                let fx0 = if fp.h_anchor.as_deref() == Some("page") {
                                    fp.x
                                } else {
                                    page.margin.left + fp.x
                                };
                                let group_top = if aligned_bottom {
                                    let mut height = 0.0;
                                    for block in page.blocks.iter().skip(block_idx) {
                                        let Block::Paragraph(next) = block else { break; };
                                        let Some(nf) = next.style.frame_pr.as_ref() else { break; };
                                        if nf.x != fp.x || nf.y != fp.y || nf.width != fp.width || nf.height != fp.height || nf.h_anchor != fp.h_anchor || nf.v_anchor != fp.v_anchor || nf.x_align != fp.x_align || nf.y_align != fp.y_align || nf.wrap != fp.wrap || nf.h_space != fp.h_space || nf.v_space != fp.v_space || nf.height_rule != fp.height_rule || nf.drop_cap != fp.drop_cap || nf.lines != fp.lines { break; }
                                        height += self.estimate_para_height(next,
                                            fp.width.unwrap_or(content_width * 0.3).max(20.0),
                                            grid_pitch, None, false, None, None);
                                    }
                                    let bottom = page.size.height - if fp.v_anchor.as_deref() == Some("page") {
                                        0.0
                                    } else { page.margin.bottom };
                                    bottom - height.max(fp.height.unwrap_or(0.0))
                                } else { fy_decl };
                                (fx0, group_top)
                            }
                        };
                        let fw = fp.width.unwrap_or(content_width * 0.3).max(20.0);
                        let fx = if positioned_frame {
                            let (left, width) = if fp.h_anchor.as_deref() == Some("page") {
                                (0.0, page.size.width)
                            } else { (page.margin.left, page.size.width - page.margin.left - page.margin.right) };
                            match fp.x_align.as_deref() {
                                Some("center") => left + (width - fw) * 0.5,
                                Some("right") => left + width - fw,
                                Some("left") => left,
                                _ => left + fp.x,
                            }
                        } else { fx };
                        let (inset_x, inset_y) = if positioned_frame || s1379 || aligned_bottom { (0.0, 0.0) } else { (1.5, 3.0) };
                        let mut fcy = LayoutCursor::new(fy + inset_y);
                        let empty_fn_frame = std::collections::HashMap::new();
                        let (frame_els, _, _) = self.layout_paragraph(
                            para,
                            fx + inset_x,
                            &mut fcy,
                            (fw - 2.0 * inset_x).max(10.0),
                            page.size.height,
                            fy,
                            page,
                            &mut Vec::new(),
                            &mut Vec::new(),
                            grid_pitch,
                            None,
                            false,
                            None,
                            None,
                            false,
                            false,
                            0.0,
                            Some(block_idx),
                            None,
                            None,
                            false,
                            false,
                            None,
                            0.0,
                            &empty_fn_frame,
                            1,
                            0,
                            &[],
                            0.0,
                            false,
                            false,
                            None,  // S755
                            None,  // S758
                            None,  // S-TWOSEG
                            false, // S835
                            0.0,
                            None,  // S900
                            None,  // S903
                            false, // S916
                            None,
                        );
                        elements.extend(frame_els);
                        // Band height: framePr h (hRule atLeast semantics — the
                        // content may exceed it).
                        let content_h_frame = (fcy.cursor_y - fy).max(14.0);
                        let fh = if positioned_frame && fp.height_rule.as_deref() == Some("exact") {
                            fp.height.unwrap_or(content_h_frame)
                        } else { fp.height.unwrap_or(0.0).max(content_h_frame) };
                        let (hs, vs) = if positioned_frame { (fp.h_space, fp.v_space) } else { (9.0, 0.0) };
                        let text_not_beside = positioned_frame && fp.v_anchor.as_deref() == Some("text")
                            && fp.wrap.as_deref() == Some("notBeside");
                        let (wrap_left, wrap_right) = if text_not_beside {
                            (0.0, page.size.width)
                        } else { (fx - hs, fx + fw + hs) };
                        let text_frame = positioned_frame && fp.v_anchor.as_deref() == Some("text");
                        let frame_policy = if text_frame { BodyWrapPolicy::TEXT_FRAME } else { BodyWrapPolicy::OBJECT };
                        let excludes_column = text_not_beside || (text_frame
                            && (wrap_left - start_x).max(0.0).max((start_x + content_width - wrap_right).max(0.0)) < frame_policy.minimum_lane_width);
                        s758_bands.push((current_page_idx, fy - vs, fy + fh + vs, wrap_left, wrap_right, false, frame_policy));
                        let has_next_frame = matches!(page.blocks.get(block_idx + 1), Some(Block::Paragraph(p))
                            if p.style.frame_pr.as_ref().is_some_and(|n| same_text_frame(fp, n)));
                        if excludes_column && !has_next_frame {
                            let mut first = block_idx;
                            while first > 0 && matches!(page.blocks.get(first - 1), Some(Block::Paragraph(p))
                                if p.style.frame_pr.as_ref().is_some_and(|n| same_text_frame(fp, n))) { first -= 1; }
                            let group_top = elements.iter().filter(|e| e.paragraph_index.is_some_and(|i| i >= first && i <= block_idx))
                                .map(|e| e.y).fold(fy, f32::min) - vs;
                            let group_bottom = fy + fh + vs;
                            let preceding = elements.iter().filter_map(|e| e.paragraph_index)
                                .filter(|&i| i < first && matches!(page.blocks.get(i), Some(Block::Paragraph(p)) if p.style.frame_pr.is_none())).max();
                            let mut shifted = 0.0_f32;
                            if let Some(i) = preceding {
                                let intersect_top = elements.iter().filter(|e| e.paragraph_index == Some(i)
                                    && e.y < group_bottom && e.y + e.height > group_top + 0.01)
                                    .map(|e| e.y).fold(f32::INFINITY, f32::min);
                                if intersect_top.is_finite() {
                                    shifted = (group_bottom - intersect_top).max(0.0);
                                    for e in elements.iter_mut().filter(|e| e.paragraph_index == Some(i) && e.y >= intersect_top - 0.01) { e.y += shifted; }
                                }
                            }
                            cursor.set((cursor.cursor_y + shifted).max(group_bottom));
                        }
                        if std::env::var("OXI_DBG847").is_ok() {
                            let txt: String = para
                                .runs
                                .iter()
                                .flat_map(|r| r.text.chars())
                                .take(16)
                                .collect();
                            eprintln!("[S847] blk={} fx={:.1} fy_decl={:.1} fy_used={:.1} bottom={:.1} txt={:?}",
                                block_idx, fx, fy_decl, fy, fcy.cursor_y, txt);
                        }
                        // S847: remember this frame's declared key + first fx +
                        // bottom so the NEXT consecutive same-key paragraph
                        // inherits fx and stacks below it.
                        s847_frame = Some((fp.x, fy_decl, fx, fcy.cursor_y));
                        // S898b: record the notBeside group's running bottom.
                        if !positioned_frame && std::env::var("OXI_S898_DISABLE").is_err()
                            && fp.wrap.as_deref() == Some("notBeside")
                        {
                            let b = fcy.cursor_y;
                            s898_notbeside_bottom =
                                Some(s898_notbeside_bottom.map_or(b, |p: f32| p.max(b)));
                        }
                        // Float: no flow consumption.
                        prev_space_after = 0.0;
                        continue;
                    }
                    // S847: a non-frame block breaks the consecutive-frame run.
                    s847_frame = None;
                    s1379_group = None;
                    // S898b: the first normal block after a notBeside frame
                    // group resumes BELOW the stacked frame bottom.
                    if let Some(b) = s898_notbeside_bottom.take() {
                        if b > cursor.cursor_y && b - cursor.cursor_y < content_height {
                            if std::env::var("OXI_DBG898").is_ok() {
                                eprintln!(
                                    "[S898b] blk={} push {:.2} -> {:.2}",
                                    block_idx, cursor.cursor_y, b
                                );
                            }
                            cursor.set(b);
                            prev_space_after = 0.0;
                        }
                    }
                    // Round 29: compute footnote contribution if this paragraph is
                    // laid out on the CURRENT page (delta added) vs a NEW page
                    // (full from-scratch). Used by overflow checks below.
                    // S168 (2026-05-22) Phase B-2 holistic: per-line fn heights map.
                    let mut para_fn_heights_map: std::collections::HashMap<u32, f32> =
                        std::collections::HashMap::new();
                    let (delta_if_current, full_if_new): (f32, f32) = if page.footnotes.is_empty() {
                        (0.0, 0.0)
                    } else {
                        let mut delta = 0.0_f32;
                        let mut full = 0.0_f32;
                        let mut seen_new: Vec<u32> = Vec::new();
                        for r in &para.runs {
                            if let Some(id) = r.footnote_ref {
                                if !seen_new.contains(&id) {
                                    seen_new.push(id);
                                    let h = estimate_footnote_h(id);
                                    para_fn_heights_map.insert(id, h);
                                    // First footnote on page includes separator overhead.
                                    // S160 (2026-05-21): Word measurement on b837 page 1
                                    // shows body→fn gap = ~27pt, but Oxi reserves only 6pt
                                    // (sep line 2pt + padding 4pt). Add OXI_FN_SEP_GAP_EXTRA
                                    // env gate for the missing ~21pt = body_line_height
                                    // worth of leading above the separator. Default off
                                    // pending verify across b837's pages (per memory:
                                    // page-to-page load-bearing risk).
                                    if footnote_ids_current_page.is_empty() && seen_new.len() == 1 {
                                        // S160 (2026-05-21): Word body→fn gap is wider
                                        // than Oxi's 6pt (sep_line 2pt + padding 4pt).
                                        // Empirically on b837 page 1: Word gap=27pt,
                                        // Oxi gap=3.5pt → Oxi under-reserves ~21pt
                                        // (1 body line + padding). Adding 6pt extra
                                        // shifts page 1 body break by 1 line, matching
                                        // Word, without page-shifting later pages
                                        // (sweep showed sep_extra=5-9 all give same
                                        // result, sep_extra>=10 regresses Phase 1).
                                        // b837 IoU 0.5855 → 0.6921 (+0.1066).
                                        // S240 (2026-05-23): removed OXI_LEGACY_FN_SEP_GAP
                                        // legacy env-var fallback during hardening pass.
                                        // OXI_FN_SEP_GAP_EXTRA tuning knob preserved.
                                        // S596b: no-docGrid docs reserve one footnote
                                        // line for the separator (see footnote_sep_alloc).
                                        let sep = footnote_sep_alloc(id);
                                        full += sep;
                                        delta += sep;
                                    }
                                    full += h;
                                    if !footnote_ids_current_page.contains(&id) {
                                        delta += h;
                                    }
                                }
                            }
                        }
                        (delta, full)
                    };
                    let effective_content_h =
                        (content_height - (footnote_reserve_current + delta_if_current)).max(0.0);
                    // Diagnostic allocation of a column band without changing
                    // physical page geometry or the references of floating shapes.
                    let effective_content_h = if num_columns > 1 {
                        std::env::var("OXI_COLUMN_FLOW_HEIGHT_PT").ok()
                            .and_then(|v| v.parse::<f32>().ok())
                            .filter(|v| v.is_finite() && *v > 0.0)
                            .map_or(effective_content_h, |height|
                                effective_content_h.min((col_band_top + height - start_y).max(0.0)))
                    } else { effective_content_h };
                    let effective_content_h = column_search.capacity(ir_index, allocation_start, current_page_idx)
                        .map_or(effective_content_h, |height|
                            effective_content_h.min((col_band_top + height - start_y).max(0.0)));
                    let effective_content_h_new_page = (content_height - full_if_new).max(0.0);
                    let _ = effective_content_h_new_page;

                    // Helper closure: commit this paragraph's footnotes to the
                    // current page's running reservation (called once we've
                    // decided which page the paragraph lands on).
                    let commit_para_footnotes =
                        |reserve: &mut f32, ids: &mut Vec<u32>, page_i: usize, blk_i: usize| {
                            if page.footnotes.is_empty() {
                                return;
                            }
                            for r in &para.runs {
                                if let Some(id) = r.footnote_ref {
                                    if !ids.contains(&id) {
                                        // First footnote: separator line (2pt + 4pt padding).
                                        // S160 env gate OXI_FN_SEP_GAP_EXTRA adds leading
                                        // above separator (Word measurement: ~21pt missing).
                                        if ids.is_empty() {
                                            // S160: see estimate-path comment near line 1934.
                                            // S596b: no-docGrid separator = one footnote line.
                                            *reserve += footnote_sep_alloc(id);
                                        }
                                        ids.push(id);
                                        // estimate + per-note rendering overhead
                                        // (superscript marker consumes ~10pt extra Y space
                                        // not captured by estimate_para_height)
                                        // Per-note overhead accounts for superscript marker
                                        // vertical space in actual rendering vs estimate
                                        let h = estimate_footnote_h(id);
                                        *reserve += h;
                                        if std::env::var("OXI_FN_PROBE").is_ok() {
                                            eprintln!("[FN_COMMIT] page_idx={} block_idx={} id={} h={:.1} reserve_now={:.1}",
                                            page_i, blk_i, id, h, *reserve);
                                        }
                                    }
                                }
                            }
                        };

                    // SOFT lastRenderedPageBreak (ECMA-376 §17.3.1.18, Session 56 Day 4):
                    // ANY run carrying <w:lastRenderedPageBreak/> indicates Word's
                    // saved render had a page break before this point. The naive
                    // "always force" implementation cascaded badly in over-packed
                    // docs (Day 3: 0e7af 1.0→0.26, d77a 0.96→0.27 from extra breaks).
                    // SOFT rule: force break only when BOTH conditions hold:
                    //   1. The paragraph would naturally fit on the current page
                    //      (i.e., we have not already overflowed past Word's break)
                    //   2. The current page is already substantially filled
                    //      (cursor more than halfway down the body area). Without
                    //      this, LRPB fires near the top of an Oxi page that already
                    //      aligns with Word's break — wrongly pushing content to
                    //      next page (bd90b00 cascade: 0.96→0.74 with rule-1-only).
                    // R7.45 (Day 34 part 14, 2026-05-13): only fire SOFT LRPB
                    // when the marker is on the FIRST run (paragraph-start
                    // break). When LRPB is on a later run, Word broke
                    // mid-paragraph — force-breaking the whole paragraph
                    // here moves both lines to the next page, but Word
                    // actually leaves line 0 on the current page. Let the
                    // natural per-line break handle the mid-paragraph case
                    // (34140 w_i=535 example).
                    // OXI_LRPB_DISABLE=1 (2026-07-11, opt-IN measurement knob,
                    // default off = byte-identical): suppress the block-level
                    // SOFT LRPB respect — used with OXI_S391_PER_LINE_LRPB=0 to
                    // measure the engine's NATURAL flow against fresh Word (the
                    // LRPB-off divergence catalog; nyserda's 28/56 page-start
                    // alignment collapses to 2/56 without LRPBs = the honest
                    // baseline of the Latin flow).
                    // S811: metric-incompatible-substitution docs distrust
                    // their saved LRPBs entirely (see doc_lrpb_distrust).
                    // S836 (2026-07-14, default ON, opt-out OXI_S836_DISABLE):
                    // Latin docs RETIRE the block-level SOFT LRPB respect —
                    // the EN natural flow (with S833/S834/S835 + S807 off)
                    // measures 6/6 = 1.0000 without it, and the saved breaks
                    // FIGHT the corrected fn/footer reservations (stale-LRPB
                    // class: nyserda p18 kept a saved break with 6.6 lines of
                    // fresh room). S391 per-line respect is untouched (it did
                    // not fire against the EN gate; JP keeps everything).
                    // S1491 (2026-09-19, default ON, opt-in OXI_S1491_LEGACY_LRPB
                    // restores the saved-break model): CJK bodies retire the
                    // saved page breaks too. The natural flow measures golden
                    // 187/187, EN 298/298 and JA blind 200/200 with both
                    // respects off, while the saved marks alone cost JA 3 docs
                    // (legal__0493f12e / policies__0820fb07 / correspondence__0b652248,
                    // each a mark one line stale against fresh Word).
                    let lrpb_knob_off = std::env::var("OXI_LRPB_DISABLE").is_ok()
                        || std::env::var_os("OXI_S1491_LEGACY_LRPB").is_none()
                        || self.doc_lrpb_distrust
                        || self.lrpb_count_distrust.get()
                        || (!self.doc_body_has_real_cjk
                            && std::env::var("OXI_S836_DISABLE").is_err());
                    let has_lrpb_at_start = para
                        .runs
                        .first()
                        .map(|r| r.has_last_rendered_page_break)
                        .unwrap_or(false);
                    let lrpb_should_break =
                        if has_lrpb_at_start && !elements.is_empty() && !lrpb_knob_off {
                            let est_h = self.estimate_para_height(
                                para,
                                self.s1211c_floor_body_width(para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                grid_pitch,
                                None,
                                false,
                                None,
                                None,
                            );
                            let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                            let consumed = cursor.cursor_y - start_y;
                            let half_page = effective_content_h * 0.5;
                            // S822 ATTEMPTED + REVERTED (2026-07-13): gating the
                            // SOFT LRPB on `remaining − est < K` (the "a real
                            // page-bottom break leaves little room" physical test,
                            // S814-v2's row analog) is Oxi-geometry-relative and
                            // FAILS the S581 wall: uklocalspending's PASS is still
                            // LRPB-supported at many boundaries where Oxi's natural
                            // flow under-fills (K=28 → uklocal 1.0→0.31 {−1:634},
                            // nyserda −1×22 new). usnyserda's +1×20 stale-LRPB
                            // subset (p18: a saved break with 6.6 lines of fresh
                            // room, fresh Word keeps) has NO local discriminator —
                            // the fix is the nyserda natural-flow catalog (make the
                            // LRPBs redundant, the b837-pi=89 resolution pattern).
                            let fires = est_h <= remaining && consumed > half_page;
                            // OXI_DBG_LRPB: one line per LRPB site with the
                            // geometry S822 tried to gate on. A saved break that
                            // Word no longer takes (policies__0353d0b2a7f98e13
                            // p34) leaves a LOT of fresh room; a live one leaves
                            // little. Dump both populations before proposing a
                            // discriminator -- S822 picked K=28 against a corpus
                            // that has since changed.
                            if std::env::var("OXI_DBG_LRPB").is_ok() {
                                let preview: String = para
                                    .runs
                                    .iter()
                                    .flat_map(|r| r.text.chars())
                                    .take(18)
                                    .collect();
                                eprintln!(
                                    "[LRPB] pg={} est={:.2} remaining={:.2} consumed={:.2} half={:.2} slack={:.2} fires={} text={:?}",
                                    pages.len() + 1,
                                    est_h,
                                    remaining,
                                    consumed,
                                    half_page,
                                    remaining - est_h,
                                    fires,
                                    preview
                                );
                            }
                            fires
                        } else {
                            false
                        };

                    // pageBreakBefore: force a new page (not just next column)
                    if (para.style.page_break_before || lrpb_should_break) && !elements.is_empty() {
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        if let Some(g) = s755_geom.as_ref() {
                            start_y = g.top(pages.len() + 1);
                            content_height = g.ch(pages.len() + 1);
                        }
                        cursor.set(start_y);
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        lm2_cells = 0;
                        current_page_idx += 1;
                        lm2_cells = 0; // Reset cumul line index for new page
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        s900_fold(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            &mut s900_pending_deferred,
                            current_page_idx,
                        );
                        commit_para_footnotes(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            current_page_idx,
                            block_idx,
                        );
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                    } else {
                        // R7.53 (2026-05-13): pre-commit DEFERRED to after
                        // layout_paragraph. Previously pre-committed here,
                        // which inflated pg_bot and rejected b837808 i=49/
                        // 60/72/90 paragraphs at their first line. Now the
                        // first-line break check uses lenient effective_h
                        // via `first_line_extra_content_h = delta_if_current`
                        // passed to layout_paragraph. Subsequent lines use
                        // strict (this para's fns are NOT yet committed but
                        // delta_if_current is accounted via the param).
                        // Post-layout commit below handles both non-spanning
                        // and spanning cases uniformly.
                    }

                    // keepLines: if doesn't fit, advance column or page
                    if para.style.keep_lines && !elements.is_empty() {
                        let est_h = self.estimate_para_height(
                            para,
                            self.s1211c_floor_body_width(para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                            grid_pitch,
                            None,
                            false,
                            None,
                            None,
                        );
                        let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                        let estimated_overflow = est_h > remaining && est_h <= effective_content_h;
                        // An estimate includes paragraph spacing and complete line
                        // advances. Confirm overflow with the natural line breaker:
                        // trailing spacing and the last line's leading may extend
                        // below the page without moving the paragraph's text.
                        let actual_overflow = if estimated_overflow {
                            let mut trial_cursor = LayoutCursor {
                                cursor_y: cursor.cursor_y,
                                visual_y: cursor.visual_y,
                                lm2_ideal_y: cursor.lm2_ideal_y,
                            };
                            let mut trial_pages: Vec<LayoutPage> = (0..pages.len())
                                .map(|_| LayoutPage { width: page.size.width, height: page.size.height, elements: Vec::new() })
                                .collect();
                            let mut trial_elements = elements.clone();
                            let mut trial_cells = lm2_cells;
                            let mut trial_mult = mult_cumul_raw;
                            let dc = pending_dropcap.unwrap_or(0.0);
                            let (trial_band, trial_two_seg, trial_wrap_advance) = self.body_paragraph_wrap_bands(
                                para, page, &s758_bands, current_page_idx,
                                cursor.cursor_y, start_x, content_width,
                            );
                            trial_cursor.advance(trial_wrap_advance);
                            let (_, _, final_column) = self.layout_paragraph(
                                para, start_x + dc, &mut trial_cursor, content_width - dc,
                                effective_content_h, start_y, page,
                                &mut trial_pages, &mut trial_elements, grid_pitch,
                                prev_para_style_id.as_deref(), prev_contextual_spacing,
                                prev_autospacing_numid.as_deref(), prev_borders.as_ref(),
                                prev_keep_next, false, prev_space_after, Some(block_idx),
                                Some(&mut trial_cells), Some(&mut trial_mult),
                                LayoutEngine::body_adjacent_to_empty_run(para, page, block_idx),
                                matches!(page.blocks.get(block_idx + 1), Some(Block::Table(_))),
                                None, delta_if_current, &para_fn_heights_map,
                                num_columns, current_column, &col_x_positions, col_band_top,
                                false, footer_tight, s755_geom.as_ref(), trial_band, trial_two_seg,
                                (footnote_reserve_current + delta_if_current) > 0.0,
                                footnote_reserve_current, None,
                                page.blocks.get(block_idx + 1).and_then(|b| match b {
                                    Block::Paragraph(p) => p.style.borders.as_ref(), _ => None,
                                }), false,
                                Some((&s758_bands, current_page_idx)),
                            );
                            trial_pages.len() > pages.len() || final_column != current_column
                        } else { false };
                        if actual_overflow {
                            if num_columns > 1 && current_column + 1 < num_columns {
                                current_column += 1;
                                start_x = col_x_positions[current_column];
                                content_width = col_widths[current_column];
                                cursor.set(col_band_top);
                            } else {
                                dbg_page_push(pages.len(), 0);
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: std::mem::take(&mut elements),
                                });
                                if let Some(g) = s755_geom.as_ref() {
                                    start_y = g.top(pages.len() + 1);
                                    content_height = g.ch(pages.len() + 1);
                                }
                                cursor.set(start_y);
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                                lm2_cells = 0;
                                current_page_idx += 1;
                                // Round 29: page push moves this paragraph (and
                                // its footnote refs) to the new page.
                                footnote_reserve_current = 0.0;
                                footnote_ids_current_page.clear();
                                s900_fold(
                                    &mut footnote_reserve_current,
                                    &mut footnote_ids_current_page,
                                    &mut s900_pending_deferred,
                                    current_page_idx,
                                );
                                commit_para_footnotes(
                                    &mut footnote_reserve_current,
                                    &mut footnote_ids_current_page,
                                    current_page_idx,
                                    block_idx,
                                );
                            }
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        }
                    }

                    // keepNext: advance column or page if pair doesn't fit.
                    // Word behavior: keepNext is best-effort. If the heading itself fits
                    // on the current page but heading+next doesn't, Word keeps the heading
                    // and sends the next paragraph to the next page. Only advance page when
                    // the heading itself doesn't fit.
                    // S916 (2026-07-18): set by the keepNext lookahead when a
                    // MULTI-LINE keepNext paragraph should SPLIT (keep n-2 head
                    // lines, move a 2-line tail + follower) instead of whole-moving.
                    let mut s916_split = false;
                    if para.style.keep_next && !elements.is_empty() {
                        // S1420 (2026-09-16, default ON, opt-out OXI_S1420_DISABLE): a
                        // keepNext paragraph whose follower is a TABLE keeps with the
                        // table's FIRST ROW. MEASURED (`_pb_keepnext_tbl_gen.py`,
                        // tests/fixtures/keepnext_tbl, Word COM, 18pt grid, 2-line
                        // cantSplit rows, content bottom 785.2): heading at 725.25 with
                        // row 1 at 741.75 (+36 fits) stays; heading at 743.25 with row 1
                        // at 759.75 (+36 = 795.75 > 785.2) moves to p2 under keepNext
                        // and stays on p1 without it. policies__094c44cd 「例示と好ましい
                        // 選択肢」 and policies__07543a6b are this class. The look-ahead
                        // below only knew a paragraph follower, so the heading stayed
                        // at the page bottom; the table's first row is folded into a
                        // one-line EXACT-spaced stand-in and the paragraph arm decides.
                        let s1420_synth: Option<Paragraph> = match page.blocks.get(block_idx + 1) {
                            Some(Block::Table(tbl))
                                if std::env::var_os("OXI_S1420_DISABLE").is_none()
                                    // CJK-body scope: Latin documents keep their
                                    // calibrated table-follower path (the
                                    // `next_block_is_table` arm of layout_paragraph
                                    // and the S960/S970 back-pulls); with both active
                                    // legal__0010437a / 001410a8 / 001beddec kept a
                                    // «Table» / «Form 22» heading Word pushes.
                                    && self.doc_body_has_real_cjk
                                    && tbl.style.position.is_none()
                                    && !tbl.rows.is_empty() =>
                            {
                                let cw = self.resolve_table_col_widths_n(tbl, content_width, false);
                                let dp = tbl.style.default_cell_margins.as_ref();
                                let (pl, pr, pt, pb) = (
                                    dp.and_then(|m| m.left).unwrap_or(5.4),
                                    dp.and_then(|m| m.right).unwrap_or(5.4),
                                    dp.and_then(|m| m.top).unwrap_or(0.0),
                                    dp.and_then(|m| m.bottom).unwrap_or(0.0),
                                );
                                let row_h = self.estimate_table_row_natural_h(
                                    &tbl.rows[0], &cw, pl, pr, pt, pb, tbl,
                                    page.grid_line_pitch, page.grid_char_pitch, None,
                                );
                                // S1428: with leading tblHeader rows the heading keeps
                                // with the header rows AND the first data row's
                                // page-bottom requirement (its minimum height when
                                // declared, else its natural height when cantSplit,
                                // else one line of it).
                                let n_hdr = tbl.rows.iter().take_while(|r| r.header).count();
                                let row_h = if n_hdr > 0 && n_hdr < tbl.rows.len()
                                    && std::env::var_os("OXI_S1428_DISABLE").is_none()
                                {
                                    let hdr_h: f32 = tbl.rows[..n_hdr].iter().map(|r| {
                                        let nat = self.estimate_table_row_natural_h(
                                            r, &cw, pl, pr, pt, pb, tbl,
                                            page.grid_line_pitch, page.grid_char_pitch, None,
                                        );
                                        r.height.map_or(nat, |h| h.max(nat))
                                    }).sum();
                                    let data = &tbl.rows[n_hdr];
                                    let data_nat = self.estimate_table_row_natural_h(
                                        data, &cw, pl, pr, pt, pb, tbl,
                                        page.grid_line_pitch, page.grid_char_pitch, None,
                                    );
                                    let data_req = match data.height {
                                        Some(h) => h,
                                        None if data.cant_split => data_nat,
                                        None => data_nat.min(page.grid_line_pitch.unwrap_or(14.0)),
                                    };
                                    hdr_h + data_req
                                } else {
                                    row_h
                                };
                                let mut synth = para.clone();
                                synth.runs.truncate(1);
                                if let Some(r) = synth.runs.first_mut() {
                                    r.text = "\u{3000}".to_string();
                                }
                                synth.shapes.clear();
                                synth.style.keep_next = false;
                                synth.style.widow_control = true;
                                synth.style.page_break_before = false;
                                synth.style.line_spacing_rule = Some("exact".to_string());
                                synth.style.line_spacing = Some(row_h.max(1.0));
                                synth.style.space_before = None;
                                synth.style.space_after = None;
                                synth.style.before_lines = None;
                                synth.style.after_lines = None;
                                synth.style.has_direct_spacing = true;
                                synth.style.has_direct_before = true;
                                Some(synth)
                            }
                            _ => None,
                        };
                        let s1420_next: Option<&Paragraph> = match page.blocks.get(block_idx + 1) {
                            // The body discards an empty nextPage section carrier
                            // under S945. Its predecessor must not reserve a line
                            // for that discarded box when applying keepNext.
                            // A follower cannot be the first-block mark kept by
                            // S1501. Continuous marks retain their table handling.
                            Some(Block::Paragraph(p))
                                if std::env::var_os("OXI_KEEP_EMPTY_SECTION_MARK_DISABLE").is_none()
                                    && std::env::var("OXI_S945_DISABLE").is_err()
                                    && p.style.page_section_break
                                    && !p.style.continuous_section_break
                                    && p.runs.iter().all(|r| r.text.is_empty())
                                    && !(p.style.page_break_after
                                        && std::env::var("OXI_SECTION_EXPLICIT_BREAKS").is_ok()) => None,
                            Some(Block::Paragraph(p)) => Some(p),
                            Some(Block::Table(_)) => s1420_synth.as_ref(),
                            _ => None,
                        };
                        if let Some(next_para) = s1420_next {
                            let this_h0 = self.estimate_para_height(
                                para,
                                self.s1211c_floor_body_width(para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                grid_pitch,
                                None,
                                false,
                                None,
                                None,
                            );
                            let next_h = self.estimate_para_height(
                                next_para,
                                self.s1211c_floor_body_width(next_para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                grid_pitch,
                                None,
                                false,
                                None,
                                None,
                            );
                            // S925: estimate_para_height drops a body follower's
                            // style-defined space_before under the same reset rule
                            // handled for `this_h` below.  The keepNext pair still
                            // consumes the collapsed inter-paragraph gap, so omitting
                            // it can claim the pair fits and then let the real layout
                            // push the follower, violating keepNext (legal Schedule 4:
                            // ySubsection + yIndenta at the wp196 bottom).  Add only
                            // the gap the estimate dropped; direct/atLeast/exact
                            // spacing remains accounted for by the estimate itself.
                            let next_h_with_gap = if std::env::var("OXI_S925_DISABLE").is_err()
                                && !next_para.style.has_direct_spacing
                                && next_para.style.line_spacing_rule.as_deref() != Some("exact")
                                && next_para.style.line_spacing_rule.as_deref() != Some("atLeast")
                            {
                                let next_sb = if let (Some(bl), Some(pitch)) =
                                    (next_para.style.before_lines, grid_pitch)
                                {
                                    bl / 100.0 * s1480_lines_unit(page.doc_grid_no_type, pitch)
                                } else {
                                    next_para.style.space_before.unwrap_or(0.0)
                                };
                                next_h + para.style.space_after.unwrap_or(0.0).max(next_sb)
                            } else {
                                next_h
                            };
                            // S709 (2026-06-30): estimate_para_height drops a paragraph's
                            // STYLE-defined space_before for body paras (its `should_reset`
                            // is a table-cell rule that mis-fires when has_direct_spacing=false
                            // and the line rule is auto). A keepNext HEADING (Heading1/2 etc.)
                            // carries its space_before in the STYLE → this_h0 = line only,
                            // under-counting the heading's real consumed height by ~space_before.
                            // The heading then "fits" at the page bottom where Word (honouring
                            // the heading's space_before) pushes the whole heading+follower to
                            // the next page (gen2 family "Scope"/"Procedure" headings orphaned).
                            // Add the effective space_before (= max(prev space_after, this
                            // space_before); cursor sits at the prev para's last-line bottom,
                            // i.e. it does NOT include prev space_after) when the estimate
                            // dropped it. Scoped to the keepNext lookahead only (the shared
                            // estimate is left untouched — high blast radius on Phase-1).
                            let s709 = std::env::var("OXI_S709_DISABLE").is_err();
                            let this_h = if s709 {
                                // S1087 (2026-08-07, opt-out OXI_S1087_DISABLE): the estimate's
                                // ACTUAL reset key is `has_direct_before` — S855 moved
                                // cell_spacing_reset_sides onto the per-SIDE flags ("the
                                // before/after RESET keys on whether the DIRECT pPr set
                                // before/after, NOT on has_direct_spacing") but S709's mirror
                                // was left on the pre-S855 predicate. A heading whose direct
                                // pPr sets ONLY `after` (reports__00196a Heading2
                                // `<w:spacing w:after="120"/>` over a style before=200) then
                                // kept head_reset=false while the estimate still dropped its
                                // 10pt space_before → this_h under-counted by exactly that
                                // (23.2 vs 33.2), the follower's 2 lines "fit" the 52.0pt
                                // remainder by 1.8pt and the heading stayed at the page bottom
                                // where Word pushes the pair. Mirror reset_before exactly.
                                let s1087 = std::env::var("OXI_S1087_DISABLE").is_err();
                                let reset_key = if s1087 {
                                    !para.style.has_direct_before
                                } else {
                                    !para.style.has_direct_spacing
                                };
                                let head_reset = reset_key
                                    && para.style.line_spacing_rule.as_deref() != Some("exact")
                                    && para.style.line_spacing_rule.as_deref() != Some("atLeast");
                                if head_reset {
                                    let this_sb = if let (Some(bl), Some(pitch)) =
                                        (para.style.before_lines, grid_pitch)
                                    {
                                        bl / 100.0 * s1480_lines_unit(page.doc_grid_no_type, pitch)
                                    } else {
                                        para.style.space_before.unwrap_or(0.0)
                                    };
                                    let prev_sa = if block_idx > 0 {
                                        match page.blocks.get(block_idx - 1) {
                                            Some(Block::Paragraph(pp)) => {
                                                pp.style.space_after.unwrap_or(0.0)
                                            }
                                            _ => 0.0,
                                        }
                                    } else {
                                        0.0
                                    };
                                    // S1087: the correction restores EXACTLY what the
                                    // estimate dropped — the paragraph's OWN space_before.
                                    // The `max(prev_sa, ...)` form is only accidentally
                                    // right when prev_sa <= this_sb: it also fires on a
                                    // paragraph with NO space_before at all, adding the
                                    // PREVIOUS paragraph's after (which `remaining` already
                                    // accounts for). technical__0056b52f's S.2.1/UR.2
                                    // headings (sb=None, prev after=12) were pushed on that
                                    // phantom 12pt and the doc collapsed 0.9849 -> 0.4121.
                                    this_h0
                                        + if s1087 { this_sb } else { prev_sa.max(this_sb) }
                                } else {
                                    this_h0
                                }
                            } else {
                                this_h0
                            };
                            let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                            // S1545 (2026-09-25, opt-out OXI_S1545_DISABLE): the pair test
                            // measured `remaining` from the raw cursor, but the heading is
                            // placed AFTER the wrap band of an anchored figure (the keepLines
                            // trial above already advances by `trial_wrap_advance`).
                            // policies__0016b30b0d5ab632 p3: decision at 652.1 (rem 81.5),
                            // heading placed at 688.7 (rem 44.9) below a wrapSquare chart;
                            // heading + 2 follower lines no longer fit, Word moves the
                            // heading to p4 with its paragraph, Oxi left it alone at the
                            // page bottom.
                            let remaining = if std::env::var_os("OXI_S1545_DISABLE").is_none() {
                                let (_, _, kn_wrap_advance) = self.body_paragraph_wrap_bands(
                                    para, page, &s758_bands, current_page_idx,
                                    cursor.cursor_y, start_x, content_width,
                                );
                                remaining - kn_wrap_advance.max(0.0)
                            } else {
                                remaining
                            };
                            // S635: a keepNext heading is pushed WITH its follower when the
                            // follower would move WHOLLY to the next page. A ≤3-line para
                            // can't split (any break leaves <2 lines on one side =
                            // widow/orphan) → it moves wholly, dragging the keepNext heading
                            // (ailitguide "10."+3-line follower). A ≥4-line follower splits
                            // (≥2 each side) → the heading stays. Line count uses a 1-line
                            // estimate at huge width (consistent with estimate_para_height —
                            // word_line_height_no_grid gives a DIFFERENT, smaller value).
                            // Follower line count via a 1-line estimate at huge width
                            // (consistent with estimate_para_height; word_line_height_no_grid
                            // gives a different, smaller value → over-counts).
                            let mut single_line_probe;
                            let one_line_para = if std::env::var("OXI_KEEP_HARD_BREAK_METRICS_DISABLE").is_err()
                                && next_para.runs.iter().any(|r| r.text.contains('\n'))
                            {
                                single_line_probe = next_para.clone();
                                for run in &mut single_line_probe.runs {
                                    run.text = run.text.replace('\n', "");
                                }
                                &single_line_probe
                            } else {
                                next_para
                            };
                            let one_line_h = self.estimate_para_height(
                                one_line_para, 1.0e6, grid_pitch, None, false, None, None,
                            );
                            // S915 (2026-07-18, opt-out OXI_S915_DISABLE):
                            // estimate_para_height adds the follower's
                            // space_before to BOTH next_h and one_line_h, so
                            // round(next_h/one_line_h) UNDER-counts a follower
                            // whose style carries a non-reset space_before. The
                            // WA-regulation `Subsection` follower (before 8-10pt +
                            // lineRule="atLeast" → the reset is suppressed) is the
                            // case: legal pi=1363 «111G. Berth operator's duties»
                            // next_h 65.2 / one_line_h 23.8 → round(2.74)=3, so
                            // S635's n<=3 branch whole-moves a 4-line follower that
                            // Word SPLITS 2+2, stranding the Heading5 above it
                            // (wp107-116, +1×~19). Subtracting the sb that the
                            // estimate itself added gives the exact count:
                            // (65.2-10)/(23.8-10)=4 → n<=3 false → not pushed
                            // (Word-correct). The S914 ship-note flagged exactly
                            // this as "a safe future refinement … strictly LESS
                            // eager". MONOTONIC (subtracting sb only INCREASES
                            // next_lines → can only REMOVE pushes, never add).
                            // MEASURED blast radius (KN635 trace, 519 docs):
                            // golden-test 0, docx_corpus/ja 0, word_png 0
                            // (byte-identical by construction); 5 EN decisions in
                            // 2 legal-style regulation docs (legal 1,
                            // policies__00148f8d 4), all removing spurious pushes —
                            // unscoped, rule-10 clean like S914. S914's wp90 orphan
                            // (pi=1105, nlines 5→6) is PRESERVED (still orphans:
                            // one_line_h+line_h_next 31.6 > rem-this_h 30.1).
                            // `follower_sb` mirrors estimate_para_height's own
                            // in_cell=false / table_para_style=None sb (the S855
                            // cell_spacing_reset_sides body path).
                            let follower_sb = if std::env::var("OXI_S915_DISABLE").is_ok() {
                                0.0
                            } else {
                                let raw_lr = next_para.style.line_spacing_rule.as_deref();
                                let explicit_rule =
                                    raw_lr == Some("exact") || raw_lr == Some("atLeast");
                                if !next_para.style.has_direct_before && !explicit_rule {
                                    0.0
                                } else if let (Some(bl), Some(pitch)) =
                                    (next_para.style.before_lines, grid_pitch)
                                {
                                    bl / 100.0 * s1480_lines_unit(page.doc_grid_no_type, pitch)
                                } else {
                                    next_para.style.space_before.unwrap_or(0.0)
                                }
                            };
                            // S1124 (2026-08-15, opt-out OXI_S1124_DISABLE): S915
                            // subtracts the follower's space_BEFORE from both sides
                            // of this division but forgot space_AFTER, so the
                            // denominator is (line + after), not the per-line
                            // increment. NDIS 0043bfe0 "Additional Supports"
                            // follower (docDefaults before=5 after=5 line=15
                            // atLeast): next_h=160, one_line=25 -> old
                            // round(155/20)=8 lines where the REAL count is 10
                            // (Word PDF: 2 on p39 + 8 on p40; Oxi render: 10 —
                            // the paragraph estimate itself was right all along).
                            // nlines=8 inflated line_h_next to (160-25)/7=19.3
                            // (real 15.0), follower_orphans tripped by 1.6pt, the
                            // heading whole-moved, and 70 pages of the price guide
                            // cascaded +1 ({+1:2813}). Mirror the sb gate for sa.
                            let follower_sa = if std::env::var("OXI_S1124_DISABLE").is_ok() {
                                0.0
                            } else {
                                let raw_lr = next_para.style.line_spacing_rule.as_deref();
                                let explicit_rule =
                                    raw_lr == Some("exact") || raw_lr == Some("atLeast");
                                if !next_para.style.has_direct_after && !explicit_rule {
                                    0.0
                                } else {
                                    next_para.style.space_after.unwrap_or(0.0)
                                }
                            };
                            let next_lines = (((next_h - follower_sb - follower_sa)
                                / (one_line_h - follower_sb - follower_sa).max(0.01))
                            .round() as usize)
                                .max(1);
                            // ★The follower moves WHOLLY only when WIDOW/ORPHAN control is
                            // ON for it: a ≤3-line para can't split without leaving <2 lines
                            // on a side. With widowControl OFF (0e7af/digitalcontract docDefaults
                            // val=0) Word splits even a 3-line follower (1+2), so the heading
                            // STAYS — exactly the old gate. Discriminator = next_para widow_control.
                            let s635 = std::env::var("OXI_S635_DISABLE").is_err();
                            // S914 (2026-07-17, opt-out OXI_S914_DISABLE): S635's own
                            // derivation above ("a ≤3-line para can't split without
                            // leaving <2 lines on a side") has TWO halves; only the
                            // "can NEVER split" half (n≤3) was implemented. The missing
                            // half is "cannot fit 2 lines HERE": an n≥4 follower whose
                            // first TWO lines do not fit in the space left after the
                            // heading is an ORPHAN, so Oxi's own orphan arm (the
                            // line_idx==0 "would leave only 1 line → push entire
                            // paragraph" rule) whole-moves it — and the lookahead, not
                            // knowing that, leaves the keepNext heading STRANDED at the
                            // page bottom. The two contradicted each other.
                            // legal__0001482d pi=1105 (a `<w:keepNext/>` Subsection,
                            // 5-line follower): rem_after_heading 30.10 vs the first two
                            // lines' 31.60 → short by 1.50pt → 1 line fits → orphan.
                            // Word agrees within 0.1pt (its own arithmetic says the
                            // second line overflows by 1.41pt) and puts heading+follower
                            // on the next page; Oxi kept the heading. ailitguide's S635
                            // specimen was a 3-line follower, so this gap never surfaced.
                            // MEASURED blast radius (KN635 trace, 519 docs): golden-test
                            // 369 → 0 changed decisions, docx_corpus/ja 50 → 0,
                            // word_png 238 → 0 (byte-identical by construction); only 4
                            // EN docs change (10 decisions), of which exactly ONE is
                            // gated — legal__0001482d, the target. 4 of those 10 are
                            // k=0 (not a single follower line fits) = unambiguous
                            // keepNext violations today. Unscoped: no JP doc reacts, so
                            // no language gate is needed (rule 10 clean).
                            let s914 = std::env::var("OXI_S914_DISABLE").is_err();
                            let line_h_next = if next_lines > 1 {
                                (next_h - one_line_h) / (next_lines - 1) as f32
                            } else {
                                one_line_h
                            };
                            // An orphan check mirroring the orphan arm that will actually
                            // run: does the follower get ≥2 lines in what's left?
                            //
                            // S934 (2026-07-18, default ON, opt-out OXI_S934_DISABLE):
                            // the space left for follower lines must also pay the
                            // heading→follower COLLAPSED GAP the S925 term already
                            // derived (max(heading style after, follower style before)
                            // when the estimate dropped both). legal__000ad039's
                            // TitleTitre headings carry after=720tw (36pt!): the 'D.'
                            // decision has rem 97.7 − this_h 36.8 = 60.9 ≥ 2 lines
                            // (55.2) WITHOUT the gap → Oxi kept the heading and split
                            // the follower 2+5; WITH the gap 24.9 < 55.2 → orphan →
                            // whole push = Word (render truth: p26 ends at [33],
                            // heading+follower on p27; the in-doc 'C.' heading gap
                            // measures 49.6 = line 13.8 + after 36 exactly).
                            let s934_gap = if std::env::var("OXI_S934_DISABLE").is_err() {
                                (next_h_with_gap - next_h).max(0.0)
                            } else {
                                0.0
                            };
                            // S1170 (2026-08-19, default ON, opt-out OXI_S1170_DISABLE):
                            // the second of these two lines IS the group's LAST line at
                            // the page bottom, and Word measures a last line by its
                            // UNMULTIPLIED height -- the line-spacing multiplier's
                            // trailing leading may hang past the bottom margin. That is
                            // the same rule S779/S827 already apply to a plain line;
                            // this check was the one place still reserving the whole
                            // multiplied box, so a keepNext heading needed a full extra
                            // pitch of room to stay.
                            // DERIVED (`_pb_keepnext_gen.py`, 24 arms = 6 filler counts x
                            // 4 sub-pitch phases, double-spaced, heading = keepNext +
                            // keepLines + before=200 like Heading2/3/4):
                            //   Word  stays at free 81.43, MOVES at 74.43
                            //         -> needs 37.6 + 27.6 + 13.8 = 79.0
                            //   Oxi   stays at free 96.40, MOVES at 89.40
                            //         -> needed 37.6 + 27.6 + 27.6 = 92.8
                            // exactly one trailing leading (13.8pt) too much. The
                            // companion sweep `_pb_lastline_gen.py` (21 arms) shows Word
                            // keeping plain lines whose BOX overruns the text bottom by
                            // up to +11.36 while the ink stays inside, and Oxi already
                            // matching that 21/21 -- so only the group test was wrong.
                            // educational__00158a7d549f9f51 p42 is this case: 86.04pt
                            // free, Word puts the heading + 2 body lines (last ink 712.96
                            // inside 720), Oxi leaves the page blank and moves both,
                            // losing a line per page for the next 25 pages.
                            // Scoped to a multiplied rule (>1), so single-spaced
                            // documents are byte-identical by construction.
                            let s1170_mult = if matches!(
                                next_para.style.line_spacing_rule.as_deref(),
                                None | Some("auto")
                            ) {
                                next_para.style.line_spacing.unwrap_or(1.0).max(1.0)
                            } else {
                                1.0
                            };
                            let last_line_h = if s1170_mult > 1.0
                                && std::env::var("OXI_S1170_DISABLE").is_err()
                            {
                                line_h_next / s1170_mult
                            } else {
                                line_h_next
                            };
                            let orphan_spacing_credit = if std::env::var("OXI_KEEP_FRAGMENT_SPACING_DISABLE").is_err() {
                                let head_includes_spacing = para.style.has_direct_spacing
                                    || matches!(para.style.line_spacing_rule.as_deref(), Some("exact") | Some("atLeast"));
                                let head_after = para.style.space_after.unwrap_or(0.0);
                                let next_before = next_para.style.space_before.unwrap_or(0.0);
                                let included_after = if head_includes_spacing { head_after } else { 0.0 };
                                (included_after + follower_sb - head_after.max(next_before)).max(0.0)
                                    + follower_sa
                            } else { 0.0 };
                            let follower_orphans = s914
                                && (one_line_h + last_line_h - orphan_spacing_credit)
                                    > (remaining - this_h - s934_gap);
                            // S1039 (2026-07-29, opt-out OXI_S1039_DISABLE): a follower
                            // that declares <w:keepLines/> cannot be split at all, so it
                            // ALWAYS moves wholly - the keepNext heading above it must
                            // follow. S635 derived "moves wholly" from widowControl only
                            // (its ailitguide specimen had no keepLines), so a keepLines
                            // follower with widowControl=0 read as splittable and the
                            // heading was stranded at the page bottom. This is the mirror
                            // of S1023, which stopped S916/S978 from SPLITTING a keepLines
                            // paragraph; here the same cohesion is applied to the pair
                            // decision. policies__0021ede1's "7 Research data" /
                            // "8 Metadata and documentation" headings (keepNext+keepLines,
                            // widowControl=0) head keepLines FirstLevel clauses: Word
                            // pushes heading+clause together, Oxi left the heading behind.
                            // S1090 (2026-08-07, opt-out OXI_S1090_DISABLE): a
                            // ONE-LINE follower always moves wholly — there is
                            // nothing to split, so widowControl is irrelevant.
                            // S635 derived the <=3 arm on a widowControl=ON doc
                            // and gated the whole thing on widow_control;
                            // technical__002c6778's "1.03 Submittals" (H4,
                            // keepNext) is followed by a 1-line "A. Product
                            // Data" with widowControl OFF, and Word pushes both.
                            let follower_moves_wholly = (next_para.style.widow_control
                                && (next_lines <= 3 || follower_orphans))
                                || (next_lines <= 1
                                    && std::env::var("OXI_S1090_DISABLE").is_err())
                                || (next_para.style.keep_lines
                                    && std::env::var("OXI_S1039_DISABLE").is_err());
                            // S802 (2026-07-12, opt-out OXI_S802_DISABLE): keepNext CHAIN —
                            // when the follower is ITSELF keepNext (H1→H2), the keep unit
                            // extends transitively to H2's follower (Word keeps the whole
                            // chain together). Without this, H1's pair check (H1+H2 fits)
                            // passed, then H2's own check pushed H2+body, STRANDING H1 at
                            // the page bottom (ukframework «Reviews and winding up
                            // arrangements» H1 alone at p37 y690.7 while Word starts p38
                            // with the chain). The chain unit adds the inter-para collapse
                            // gaps (max(after, before) — estimate_para_height excludes
                            // space_after) and the FINAL follower's requirement per the
                            // S635 rule: whole when it moves wholly (widowControl &&
                            // ≤3 lines), else its 2-line orphan minimum (a 1-line start
                            // at the page bottom is orphan-pushed, which would strand the
                            // chain anyway).
                            // The chain arm fires ONLY when H2 itself WOULD push (the
                            // strand case) — evaluated one level deep with the same
                            // S635 semantics; a chain with room stays untouched (the
                            // unconditional-unit first cut mass-pushed mid-page chains,
                            // {+1:66}).
                            // ★HELD OPT-IN (OXI_S802=1, default OFF byte-identical):
                            // even strand-scoped, the ESTIMATE-based one-level
                            // simulation over-fires (framework {+1:27} — lead-in
                            // keepNext paras before short bullets read as pushing
                            // chains where Word keeps them; estimate heights carry
                            // the S709 spacing corrections and inflate). The robust
                            // fix is a POST-HOC BACK-PULL with actual geometry (when
                            // a keepNext follower pushes, pull the stranded heading's
                            // already-emitted elements to the new page — the
                            // S750-rebalance / S728-replay pattern), a dedicated
                            // session. The wp38 strand case itself IS real: Word
                            // starts p38 with the whole H1→H2→body chain.
                            let s802 = std::env::var("OXI_S802").is_ok();
                            let (unit_next_h, follower_moves_wholly) = if s802
                                && next_para.style.keep_next
                            {
                                if let Some(Block::Paragraph(nn)) = page.blocks.get(block_idx + 2) {
                                    let nn_h = self.estimate_para_height(
                                        nn,
                                        self.s1211c_floor_body_width(nn, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                        grid_pitch,
                                        None,
                                        false,
                                        None,
                                        None,
                                    );
                                    let nn_one = self.estimate_para_height(
                                        nn, 1.0e6, grid_pitch, None, false, None, None,
                                    );
                                    let nn_lines =
                                        ((nn_h / nn_one.max(0.01)).round() as usize).max(1);
                                    let nn_fmw = nn.style.widow_control && nn_lines <= 3;
                                    let gap1 = para
                                        .style
                                        .space_after
                                        .unwrap_or(0.0)
                                        .max(next_para.style.space_before.unwrap_or(0.0));
                                    let gap2 = next_para
                                        .style
                                        .space_after
                                        .unwrap_or(0.0)
                                        .max(nn.style.space_before.unwrap_or(0.0));
                                    let tail = if nn_fmw { nn_h } else { 2.0 * nn_one };
                                    let remaining_after = remaining - this_h - gap1;
                                    let h2_unit = next_h + gap2 + tail;
                                    let h2_would_push = h2_unit > remaining_after
                                        && (next_h > remaining_after || nn_fmw)
                                        && h2_unit <= effective_content_h;
                                    if h2_would_push {
                                        (gap1 + next_h + gap2 + tail, true)
                                    } else {
                                        (next_h, follower_moves_wholly)
                                    }
                                } else {
                                    (next_h, follower_moves_wholly)
                                }
                            } else {
                                (next_h_with_gap, follower_moves_wholly)
                            };
                            // S948 (2026-07-19, opt-out OXI_S948_DISABLE): the pair sum
                            // double-counts spacing when the ESTIMATES included it (the
                            // non-reset case: direct spacing or an exact/atLeast rule —
                            // the S925 arm correctly composes the RESET case and is
                            // skipped here by construction). Word composition = head sb +
                            // head lines + max(head after, follower sb) + follower lines,
                            // with the follower's TRAILING after hanging at the page
                            // bottom. NDIS wp15 "Claiming for Time of Day": pair 96.1 vs
                            // rem 87.9 pushed; Word arithmetic 87.1 <= 87.9 fits — the
                            // +9 = after_head 4 + sb_next 5 uncollapsed (gap is 5) +
                            // trailing after 5. Correction applies only to what the
                            // estimates actually included.
                            let s948_corr = if std::env::var("OXI_S948_DISABLE").is_err() {
                                let inc = |st: &ParagraphStyle| {
                                    st.has_direct_spacing
                                        || matches!(
                                            st.line_spacing_rule.as_deref(),
                                            Some("exact") | Some("atLeast")
                                        )
                                };
                                let head_inc = inc(&para.style);
                                let next_inc = inc(&next_para.style);
                                let ha_e = para.style.space_after.unwrap_or(0.0);
                                let nsb_e = next_para.style.space_before.unwrap_or(0.0);
                                let na_e = next_para.style.space_after.unwrap_or(0.0);
                                let ha_i = if head_inc { ha_e } else { 0.0 };
                                let nsb_i = if next_inc { nsb_e } else { 0.0 };
                                let na_i = if next_inc { na_e } else { 0.0 };
                                (ha_i + nsb_i - ha_e.max(nsb_e)).max(0.0) + na_i
                            } else {
                                0.0
                            };
                            // The final line of the kept follower uses the same fit
                            // height as the orphan look-ahead, not its full advance.
                            let tail_leading = if !s802 && follower_moves_wholly {
                                (line_h_next - last_line_h).max(0.0)
                            } else { 0.0 };
                            let pair_fit_height = this_h + unit_next_h - s948_corr - tail_leading;
                            let pair_overflows = pair_fit_height > remaining
                                && pair_fit_height <= effective_content_h;
                            let mut do_push = if s635 {
                                pair_overflows && (this_h > remaining || follower_moves_wholly)
                            } else {
                                pair_overflows && this_h > remaining
                            };
                            // S916 (2026-07-18, opt-out OXI_S916_DISABLE): when
                            // the push fires in CASE B — `this_h <= remaining` (the
                            // keepNext para itself FITS; the push is driven only by
                            // `follower_moves_wholly`, NOT this_h > remaining) — and
                            // the para is a MULTI-LINE body paragraph (>= 4 lines),
                            // do NOT whole-move it. SPLIT: keep n-2 head lines on
                            // this page, move a 2-line tail + the follower to the
                            // next page (Word only requires the para's LAST line to
                            // stay with the follower). legal pi=2128 (a 5-line
                            // `<w:keepNext/>` Indenti para): Word's own saved
                            // mid-paragraph LRPB splits it 3+2 where Oxi whole-moved
                            // -> the wp158-162 +1x5 cascade. VERIFIED "keep n-2 /
                            // move 2-line tail" matches Word on pi=2128 (5->3+2) and
                            // pi=386 (wp40, 4->2+2). The line estimate under-counts
                            // (<= real), so gating on estimate>=4 guarantees the
                            // layout-side real-count gate (>=4) also holds. Blast
                            // radius (Unit B3, KN635 trace, 519 docs): golden 0, ja
                            // 0, word_png 0; 2 EN decisions (legal only). Unscoped
                            // (rule-10 clean).
                            // S978 (2026-07-22, opt-out OXI_S978_DISABLE): the same
                            // split rule in CASE A — `this_h > remaining`, i.e. the
                            // keepNext paragraph does NOT fit the rest of the page. It
                            // will therefore SPLIT at the page bottom, and keepNext only
                            // requires its LAST line to sit with the follower — which a
                            // split guarantees (the tail lands on the next page, the
                            // follower right after it). Whole-moving it instead is what
                            // let S802B's chain back-pull sweep a long run-in-heading
                            // chain off the page (technical__0056b52f: 18 consecutive
                            // keepNext paragraphs, ~440pt, dragged to the next page and
                            // leaving p6 60% empty, where Word breaks INSIDE the chain).
                            // Short paragraphs (< 4 lines) cannot split without a
                            // widow/orphan, so they keep the whole-move — the S916 gate.
                            let s978 = do_push
                                && s635
                                && this_h > remaining
                                && std::env::var("OXI_S978_DISABLE").is_err();
                            if do_push
                                && s635
                                && (this_h <= remaining || s978)
                                && std::env::var("OXI_S916_DISABLE").is_err()
                            {
                                let para_one_line = self.estimate_para_height(
                                    para, 1.0e6, grid_pitch, None, false, None, None,
                                );
                                // Accurate line count (S915 technique): subtract the
                                // estimate-included space_before from BOTH this_h0 and
                                // para_one_line so round() is not dragged below the real
                                // count by the spacing (plain this_h0/para_one_line
                                // under-counts pi=386 4->3). Still cannot OVER-count (a
                                // non-zero space_after only lowers it), so estimate>=4 =>
                                // real lines>=4 — the layout-side gate stays sound.
                                let raw_lr = para.style.line_spacing_rule.as_deref();
                                let p_explicit_rule =
                                    raw_lr == Some("exact") || raw_lr == Some("atLeast");
                                let para_sb = if !para.style.has_direct_before && !p_explicit_rule {
                                    0.0
                                } else if let (Some(bl), Some(pitch)) =
                                    (para.style.before_lines, grid_pitch)
                                {
                                    bl / 100.0 * s1480_lines_unit(page.doc_grid_no_type, pitch)
                                } else {
                                    para.style.space_before.unwrap_or(0.0)
                                };
                                let para_lines =
                                    (((this_h0 - para_sb) / (para_one_line - para_sb).max(0.01))
                                        .round() as usize)
                                        .max(1);
                                // S978: the S915 form subtracts only space_BEFORE, so a
                                // paragraph carrying a direct space_after is still
                                // under-counted (4 real lines of 11.5 with after=12 read
                                // as round(58/23.5)=2). Case A needs the exact count, so
                                // subtract the space_after the estimate also added. Kept
                                // LOCAL to the new gate — S916's case-B count keeps its
                                // measured blast radius unchanged.
                                let para_sa = if !para.style.has_direct_after && !p_explicit_rule {
                                    0.0
                                } else {
                                    para.style.space_after.unwrap_or(0.0)
                                };
                                let para_lines_exact = (((this_h0 - para_sb - para_sa)
                                    / (para_one_line - para_sb - para_sa).max(0.01))
                                .round()
                                    as usize)
                                    .max(1);
                                let eff_lines = if s978 { para_lines_exact } else { para_lines };
                                // S1023 (2026-07-27, opt-out OXI_S1023_DISABLE): neither
                                // S916 (Case B) nor S978 (Case A) may SPLIT a keepLines
                                // paragraph — Word whole-moves a keepNext+keepLines para.
                                // legal__001410a8 wi=528 (direct keepNext+keepLines, 4-line
                                // body): Word puts all 4 lines on p39; S916 Case-B split it
                                // 3+1 (do_push=false + s916_split). keepLines cohesion
                                // outranks the S916 split. Uses the RESOLVED keep_lines
                                // (style-inherited keepLines is the same OOXML contract).
                                // The S916 canary (legal__0001482d Indenti paras) is
                                // keepLines=false → still splits (orthogonal).
                                let s1023_kl = para.style.keep_lines
                                    && std::env::var("OXI_S1023_DISABLE").is_err();
                                if eff_lines >= 4 && !s1023_kl {
                                    // S978: in case A the page-bottom break already
                                    // splits the paragraph — no forced split point.
                                    s916_split = !s978;
                                    do_push = false;
                                }
                            }
                            if std::env::var("OXI_DBG_KN635").is_ok() {
                                let t: String = para
                                    .runs
                                    .iter()
                                    .flat_map(|r| r.text.chars())
                                    .take(16)
                                    .collect();
                                eprintln!("[KN635] {:?} this_h0={:.1} this_h={:.1} next_h={:.1} rem={:.1} 1line={:.1} nlines={} wc={} pair_ov={} fmw={} do_push={} sb={:?} hds={} hdba={} rule={:?}",
                                    t, this_h0, this_h, next_h, remaining, one_line_h, next_lines, next_para.style.widow_control, pair_overflows, follower_moves_wholly, do_push,
                                    para.style.space_before, para.style.has_direct_spacing, para.style.has_direct_before_after, para.style.line_spacing_rule);
                            }
                            if do_push {
                                // S802B (2026-07-13, default ON, opt-out OXI_S802B_DISABLE):
                                // keepNext CHAIN back-pull with ACTUAL geometry. The pairwise
                                // lookahead is greedy: H1's check passes (H1+H2 fit the page
                                // bottom), then H2's own check pushes H2 — STRANDING H1
                                // (ukframework natural: «Reviews and winding up arrangements»
                                // H1 25pt alone at the p37 bottom; Word starts p38 with the
                                // whole H1→H2→body chain; the −1×7 wp38-41 cascade). Word
                                // treats keepNext transitively. The v1/v2 ESTIMATE-based chain
                                // simulations over-fired ({+1:66}/{+1:27}, held opt-in
                                // OXI_S802) — this back-pull uses the ALREADY-LAID geometry
                                // instead: when the push fires, walk back over immediately
                                // preceding keepNext paragraphs still on this page, drain
                                // their emitted elements, and re-emit them at the new page
                                // top (the S728-replay / S750-rebalance element-move
                                // pattern); the cursor continues below them so this
                                // paragraph lays after the moved chain with its normal
                                // spacing collapse. space_before drops at the page top
                                // (Word p38: the moved title box sits AT the content top).
                                // Latin-doc scope v1 (!doc_body_has_real_cjk) — the JP
                                // corpus keeps its calibrated keepNext behavior
                                // byte-identically. Caveat (v1): the element split is
                                // position-based (y >= the chain-start cursor), so a float
                                // anchored earlier but rendered below the chain start would
                                // be swept along — none in the EN corpus headings.
                                // S1570 (2026-09-26, default ON, opt-out OXI_S1570_DISABLE):
                                // the JP side too. policies__1a7a3fec p31/32: 「3.15.2 偶発的
                                // 曝露…」 (heading, keepNext) then 「3.15.2.1 偶発的曝露」
                                // (keepNext) then body -- Word starts p32 with both headings,
                                // Oxi pushed only the second and stranded the first.
                                let s802b = (!self.doc_body_has_real_cjk
                                    || std::env::var_os("OXI_S1570_DISABLE").is_none())
                                    && std::env::var("OXI_S802B_DISABLE").is_err();
                                let mut pull_from = block_idx;
                                if s802b && !(num_columns > 1 && current_column + 1 < num_columns) {
                                    while pull_from > 0 {
                                        if let Some(Block::Paragraph(pp)) =
                                            page.blocks.get(pull_from - 1)
                                        {
                                            if pp.style.keep_next
                                                && block_page_indices.get(pull_from - 1)
                                                    == Some(&current_page_idx)
                                            {
                                                pull_from -= 1;
                                                continue;
                                            }
                                        }
                                        break;
                                    }
                                }
                                let pulled: Vec<LayoutElement> = if pull_from < block_idx {
                                    let y0 = block_y_positions[pull_from] - 0.1;
                                    let (keep, moved): (Vec<LayoutElement>, Vec<LayoutElement>) =
                                        elements.drain(..).partition(|e| {
                                            if (std::env::var("OXI_KEEP_CHAIN_SOURCE").is_ok()
                || std::env::var("OXI_S1474_DISABLE").is_err()) {
                                                if let Some(index) = e.paragraph_index {
                                                    return index < pull_from || index >= block_idx;
                                                }
                                            }
                                            e.y < y0
                                        });
                                    elements = keep;
                                    moved
                                } else {
                                    Vec::new()
                                };
                                let chain_end_old = cursor.cursor_y;
                                if num_columns > 1 && current_column + 1 < num_columns {
                                    current_column += 1;
                                    start_x = col_x_positions[current_column];
                                    content_width = col_widths[current_column];
                                    cursor.set(col_band_top);
                                } else {
                                    dbg_page_push(pages.len(), 0);
                                    pages.push(LayoutPage {
                                        width: page.size.width,
                                        height: page.size.height,
                                        elements: std::mem::take(&mut elements),
                                    });
                                    if let Some(g) = s755_geom.as_ref() {
                                        start_y = g.top(pages.len() + 1);
                                        content_height = g.ch(pages.len() + 1);
                                    }
                                    cursor.set(start_y);
                                    current_column = 0;
                                    start_x = col_x_positions[0];
                                    content_width = col_widths[0];
                                    lm2_cells = 0;
                                    current_page_idx += 1;
                                    footnote_reserve_current = 0.0;
                                    footnote_ids_current_page.clear();
                                    s900_fold(
                                        &mut footnote_reserve_current,
                                        &mut footnote_ids_current_page,
                                        &mut s900_pending_deferred,
                                        current_page_idx,
                                    );
                                    commit_para_footnotes(
                                        &mut footnote_reserve_current,
                                        &mut footnote_ids_current_page,
                                        current_page_idx,
                                        block_idx,
                                    );
                                }
                                if !pulled.is_empty() {
                                    let min_y = pulled.iter().map(|e| e.y).fold(f32::MAX, f32::min);
                                    let dy = cursor.cursor_y - min_y;
                                    if std::env::var("OXI_DBG_KN635").is_ok() {
                                        eprintln!("[S802B] pull {} blocks ({} els) dy={:.1} chain_end_old={:.1}",
                                            block_idx - pull_from, pulled.len(), dy, chain_end_old);
                                    }
                                    for mut e in pulled {
                                        e.y += dy;
                                        // Mirror the S728 clone-shift: border content carries
                                        // its own y1/y2 coordinates.
                                        if let LayoutContent::TableBorder {
                                            ref mut y1,
                                            ref mut y2,
                                            ..
                                        } = e.content
                                        {
                                            *y1 += dy;
                                            *y2 += dy;
                                        }
                                        elements.push(e);
                                    }
                                    // Cursor continues below the moved chain, preserving the
                                    // chain-internal advance (end − first element top).
                                    cursor.set(chain_end_old + dy);
                                    for bi in pull_from..block_idx {
                                        if let Some(p) = block_page_indices.get_mut(bi) {
                                            *p = current_page_idx;
                                        }
                                        if let Some(y) = block_y_positions.get_mut(bi) {
                                            *y += dy;
                                        }
                                    }
                                }
                                *block_page_indices.last_mut().unwrap() = current_page_idx;
                                *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                            }
                        } else if let Some(Block::Image(next_img)) = page.blocks.get(block_idx + 1)
                        {
                            // S959 (2026-07-20, opt-out OXI_S959_DISABLE): the follower is
                            // an IMAGE. S537 replaces an image-only paragraph with a bare
                            // Block::Image, so this lookahead — which only ever matched a
                            // Paragraph follower — silently ignored keepNext across that
                            // boundary. policies__00148f8d p54: «(5) A parking light…»
                            // (keepNext) is followed by an image-only Graphics paragraph
                            // and its caption; Word puts all three on p55, Oxi left the
                            // text on p54 and pushed only the image. An image cannot
                            // split, so the follower always "moves wholly" — push this
                            // paragraph when the pair does not fit the remaining space.
                            if std::env::var("OXI_S959_DISABLE").is_err() {
                                let this_h = self.estimate_para_height(
                                    para,
                                    self.s1211c_floor_body_width(para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                    grid_pitch,
                                    None,
                                    false,
                                    None,
                                    None,
                                );
                                let next_h = next_img.height;
                                let remaining = start_y + content_height - cursor.cursor_y;
                                // Follow resolved keep links through inline image hosts as well
                                // as text paragraphs. A chain already at the page top, or too
                                // tall for a fresh page, must be allowed to split in place.
                                let mut image_chain_start = block_idx;
                                let mut has_image_link = false;
                                if !(num_columns > 1 && current_column + 1 < num_columns) {
                                    while image_chain_start > 0 {
                                        let previous = image_chain_start - 1;
                                        if block_page_indices.get(previous) != Some(&current_page_idx) { break; }
                                        let keep = match page.blocks.get(previous) {
                                            Some(Block::Paragraph(p)) => p.style.keep_next,
                                            Some(Block::Image(image)) if image.position.is_none() => {
                                                let keep = image.host_paragraph.as_ref()
                                                    .map_or(false, |p| p.style.keep_next);
                                                has_image_link |= keep;
                                                keep
                                            }
                                            _ => false,
                                        };
                                        if !keep { break; }
                                        image_chain_start = previous;
                                    }
                                }
                                let chain_top = block_y_positions.get(image_chain_start)
                                    .copied().unwrap_or(cursor.cursor_y);
                                let fresh_height = s755_geom.as_ref()
                                    .map_or(content_height, |g| g.ch(pages.len() + 2));
                                let image_chain_can_move = !has_image_link
                                    || (chain_top > start_y + 0.1
                                        && cursor.cursor_y - chain_top + this_h + next_h <= fresh_height);
                                if this_h + next_h > remaining && this_h <= remaining
                                    && image_chain_can_move {
                                    if std::env::var("OXI_DBG_KN635").is_ok() {
                                        let t: String = para
                                            .runs
                                            .iter()
                                            .flat_map(|r| r.text.chars())
                                            .take(16)
                                            .collect();
                                        eprintln!("[KN635-IMG] {:?} this_h={:.1} img_h={:.1} rem={:.1} push=true",
                                            t, this_h, next_h, remaining);
                                    }
                                    // S963b: the image arm gets the SAME transitive
                                    // back-pull as the paragraph arm (S802B). Word
                                    // treats a keepNext run as one unit whatever the
                                    // terminal block is — measured for paragraphs by
                                    // _pb_kchain_gen.py and equally true when the
                                    // terminal is a figure. policies__00148f8d p65:
                                    // «(6) If only one brake light» (keepNext) →
                                    // «(7) Subrule (6) applies» (keepNext) → image;
                                    // S959 pushed the (7)+image pair and stranded (6).
                                    let s963b = !self.doc_body_has_real_cjk
                                        && std::env::var("OXI_S802B_DISABLE").is_err()
                                        && std::env::var("OXI_S963_DISABLE").is_err();
                                    let mut pull_from = if has_image_link { image_chain_start } else { block_idx };
                                    if s963b
                                        && !(num_columns > 1 && current_column + 1 < num_columns)
                                    {
                                        while pull_from > 0 {
                                            if let Some(Block::Paragraph(pp)) =
                                                page.blocks.get(pull_from - 1)
                                            {
                                                if pp.style.keep_next
                                                    && block_page_indices.get(pull_from - 1)
                                                        == Some(&current_page_idx)
                                                {
                                                    pull_from -= 1;
                                                    continue;
                                                }
                                            }
                                            break;
                                        }
                                    }
                                    let pulled: Vec<LayoutElement> = if pull_from < block_idx {
                                        let y0 = block_y_positions[pull_from] - 0.1;
                                        let (keep, moved): (
                                            Vec<LayoutElement>,
                                            Vec<LayoutElement>,
                                        ) = elements.drain(..).partition(|e| {
                                            if (std::env::var("OXI_KEEP_CHAIN_SOURCE").is_ok()
                || std::env::var("OXI_S1474_DISABLE").is_err()) {
                                                if let Some(index) = e.paragraph_index {
                                                    return index < pull_from || index >= block_idx;
                                                }
                                            }
                                            e.y < y0
                                        });
                                        elements = keep;
                                        moved
                                    } else {
                                        Vec::new()
                                    };
                                    let chain_end_old = cursor.cursor_y;
                                    if num_columns > 1 && current_column + 1 < num_columns {
                                        current_column += 1;
                                        start_x = col_x_positions[current_column];
                                        content_width = col_widths[current_column];
                                        cursor.set(col_band_top);
                                    } else {
                                        dbg_page_push(pages.len(), 0);
                                        pages.push(LayoutPage {
                                            width: page.size.width,
                                            height: page.size.height,
                                            elements: std::mem::take(&mut elements),
                                        });
                                        if let Some(g) = s755_geom.as_ref() {
                                            start_y = g.top(pages.len() + 1);
                                            content_height = g.ch(pages.len() + 1);
                                        }
                                        cursor.set(start_y);
                                        // S963 (2026-07-21, opt-out OXI_S963_DISABLE):
                                        // bring this push up to parity with the
                                        // paragraph arm's (5445). S959 shipped with a
                                        // bare pages.push — it never advanced
                                        // current_page_idx, reset the column/lm2 state
                                        // or re-pointed this block's page index, and
                                        // because it fires BEFORE `pages_before` is
                                        // sampled the later `pages_added` bookkeeping
                                        // cannot compensate. Everything keyed off the
                                        // page index (footnote attribution, float
                                        // anchors, block_page_indices) was therefore
                                        // one page stale for the rest of the section.
                                        if std::env::var("OXI_S963_DISABLE").is_err() {
                                            current_column = 0;
                                            start_x = col_x_positions[0];
                                            content_width = col_widths[0];
                                            lm2_cells = 0;
                                            current_page_idx += 1;
                                            footnote_reserve_current = 0.0;
                                            footnote_ids_current_page.clear();
                                            s900_fold(
                                                &mut footnote_reserve_current,
                                                &mut footnote_ids_current_page,
                                                &mut s900_pending_deferred,
                                                current_page_idx,
                                            );
                                            commit_para_footnotes(
                                                &mut footnote_reserve_current,
                                                &mut footnote_ids_current_page,
                                                current_page_idx,
                                                block_idx,
                                            );
                                            if !pulled.is_empty() {
                                                let min_y = pulled
                                                    .iter()
                                                    .map(|e| e.y)
                                                    .fold(f32::MAX, f32::min);
                                                let dy = cursor.cursor_y - min_y;
                                                if std::env::var("OXI_DBG_KN635").is_ok() {
                                                    eprintln!(
                                                        "[S963B] pull {} blocks ({} els) dy={:.1}",
                                                        block_idx - pull_from,
                                                        pulled.len(),
                                                        dy
                                                    );
                                                }
                                                for mut e in pulled {
                                                    e.y += dy;
                                                    if let LayoutContent::TableBorder {
                                                        ref mut y1,
                                                        ref mut y2,
                                                        ..
                                                    } = e.content
                                                    {
                                                        *y1 += dy;
                                                        *y2 += dy;
                                                    }
                                                    elements.push(e);
                                                }
                                                cursor.set(chain_end_old + dy);
                                                for bi in pull_from..block_idx {
                                                    if let Some(p) = block_page_indices.get_mut(bi)
                                                    {
                                                        *p = current_page_idx;
                                                    }
                                                    if let Some(y) = block_y_positions.get_mut(bi) {
                                                        *y += dy;
                                                    }
                                                }
                                            }
                                            *block_page_indices.last_mut().unwrap() =
                                                current_page_idx;
                                            *block_y_positions.last_mut().unwrap() =
                                                cursor.cursor_y;
                                        }
                                    }
                                }
                            }
                        } else if let Some(Block::Table(next_table)) =
                            page.blocks.get(block_idx + 1)
                        {
                            // S1024 (2026-07-27, opt-out OXI_S1024_DISABLE): the follower
                            // is a TABLE. The keepNext lookahead only matched a Paragraph
                            // (and Image via S959) — a table follower was ignored, so a
                            // keepNext heading followed by a table that overflows the
                            // remaining space split the table across the page boundary.
                            // legal__001410a8 Form 22 (keepNext) → 14-row table: Word
                            // whole-moves heading+table to a fresh page; Oxi kept the
                            // heading on p64 and split the table p64/p65. Word applies the
                            // heading's keepNext to the table follower: when the table
                            // would SPLIT on the current page but fits ONE fresh page, move
                            // the heading + table to the fresh page. DECIDED BY ACTUAL
                            // GEOMETRY — a throwaway layout_table probe on a fresh page
                            // (estimate is FORBIDDEN, S970: estimate_table_row_natural_h is
                            // 4× off). v1 scope: plain, non-floating table with no nested
                            // table / cell footnote / anchored drawing (bail otherwise);
                            // fresh-one-page ONLY (the T > fresh cutoff is untested — the
                            // §7 Word-COM probe). Latin scope.
                            // The discriminator (row_chain, below) was PINNED by a
                            // 152-arm Word-COM probe (REPORT_S1024_keepNext_table_
                            // chain_probe): Word whole-moves iff the table carries an
                            // internal row-chain (every non-final row's leftmost
                            // cell first paragraph keepNext=1). Table height / row
                            // count / remainder / direct-vs-style keepNext were ALL
                            // falsified. Census: golden 0 / JP 0 chain candidates →
                            // the change set is {target, technical__002c1ffa}, both
                            // verified (target 0.9463→0.9563 +15/0-new; the other
                            // byte-safe). Default ON, opt-out OXI_S1024_DISABLE.
                            let s1024 = std::env::var("OXI_S1024_DISABLE").is_err()
                                && (!self.doc_body_has_real_cjk
                                || std::env::var("OXI_CJK_TABLE_KEEP_NEXT").ok().as_deref() == Some("1"))
                                && !para.style.page_break_before
                                && next_table.style.position.is_none()
                                && next_table.rows.iter().all(|row| {
                                    row.cells.iter().all(|cell| {
                                        cell.blocks.iter().all(|b| match b {
                                            Block::Table(_) => false, // no nested table
                                            Block::Paragraph(p) => {
                                                p.runs.iter().all(|r| r.footnote_ref.is_none())
                                            }
                                            _ => true,
                                        })
                                    })
                                });
                            if s1024 {
                                let this_h = self.estimate_para_height(
                                    para,
                                    self.s1211c_floor_body_width(para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                                    grid_pitch,
                                    None,
                                    false,
                                    None,
                                    None,
                                );
                                let remaining = start_y + content_height - cursor.cursor_y;
                                // S1024 discriminator (REPORT_S1024_keepNext_table_chain_probe,
                                // 152-arm Word-COM probe): Word whole-moves heading+table iff
                                // the table carries an INTERNAL ROW-CHAIN — every NON-FINAL
                                // row's LEFTMOST cell's FIRST paragraph has keepNext=1. Form 22
                                // (rows 0-12 all keepNext) → chain → whole-move; wp56 "Table"
                                // (all cell paras keepNext=0) → no chain → Word splits. Table
                                // height / row count / remainder / direct-vs-style keepNext were
                                // ALL falsified by the probe. NOT the S970 terminal-cell rule
                                // (that is the last cell's terminal paragraph; this is the
                                // leftmost cell's FIRST paragraph of each non-final row).
                                let row_chain = next_table.rows.len() >= 2
                                    && next_table.rows[..next_table.rows.len() - 1].iter().all(
                                        |row| {
                                            row.cells
                                                .first()
                                                .and_then(|cell| {
                                                    cell.blocks.iter().find_map(|block| match block
                                                    {
                                                        Block::Paragraph(p) => {
                                                            Some(p.style.keep_next)
                                                        }
                                                        _ => None,
                                                    })
                                                })
                                                .unwrap_or(false)
                                        },
                                    );
                                // ACTUAL probe: lay the table out on a FRESH page in
                                // throwaway state (&self is immutable — no mutation of the
                                // committed layout). pages.is_empty() ⇒ the table fits one
                                // page; the max element bottom is its extent.
                                let mut probe_pages: Vec<LayoutPage> = Vec::new();
                                let mut probe_elements: Vec<LayoutElement> = Vec::new();
                                let mut probe_cursor = LayoutCursor::new(start_y);
                                let probe_els = self.layout_table(
                                    next_table,
                                    start_x,
                                    &mut probe_cursor,
                                    content_width,
                                    grid_pitch,
                                    page.grid_char_pitch,
                                    page.grid_char_cw_ratio,
                                    start_y,
                                    content_height,
                                    page.size.width,
                                    page.size.height,
                                    &mut probe_pages,
                                    &mut probe_elements,
                                    Some(block_idx + 1),
                                    page,
                                    false,
                                    None,
                                    None,
                                    0.0,
                                    0.0,
                                    false,
                                    None,
                                );
                                let fresh_one_page = probe_pages.is_empty();
                                let e = probe_els
                                    .iter()
                                    .chain(probe_elements.iter())
                                    .map(|el| el.y + el.height)
                                    .fold(start_y, f32::max)
                                    - start_y;
                                // current_splits: after the heading is placed the table
                                // overflows the current page. both_fit_fresh: heading +
                                // table fit one fresh page.
                                let current_splits =
                                    cursor.cursor_y + this_h + e > start_y + content_height + 0.5;
                                let both_fit_fresh = this_h + e <= content_height + 0.5;
                                // S1038 (2026-07-29, opt-out OXI_S1038_DISABLE):
                                // keepNext with a TABLE follower means the heading must
                                // sit with the table's FIRST ROW - not with the whole
                                // table (S1024's whole-move, which needs the table to fit
                                // one fresh page). policies__0021ede1's "6 Definitions"
                                // heads a 13-row definition table whose first row is
                                // ~210pt; Word leaves ~86pt of p2 empty and pushes the
                                // heading, because the row cannot follow it. Oxi kept the
                                // heading and started the table on the next page - a plain
                                // keepNext violation that no pairwise check sees, since
                                // the "follower" it measured was the whole table.
                                // Row height from the same throwaway probe (never an
                                // estimate: S970's estimate_table_row_natural_h is 4x off).
                                // S1248 (2026-08-28, default ON, opt-out
                                // OXI_S1248_DISABLE): what the heading has to sit
                                // with is the table's leading keepNext ROW-CHAIN,
                                // not merely its first row. S1038 measured one row
                                // and S1024 measures the whole table only when
                                // EVERY non-final row is chained; a table whose
                                // chain stops part-way falls between them.
                                // DERIVED, `_pb_kntbl2` KN arm (Word PDF truth, 21
                                // filler counts): caption keepNext, rows 0..2
                                // keepNext, rows 3..4 plain. Word keeps the whole
                                // thing on p1 up to fill 53 and moves caption AND
                                // rows 0..2 to p2 at fill 54 -- the filler count at
                                // which row 3 stops fitting. Row 2 keeps with row 3,
                                // row 1 with row 2, row 0 with row 1, the caption
                                // with row 0: one link fails and the chain travels
                                // whole. Measuring only the first row put the move at
                                // fill 61, leaving the caption orphaned for 7 arms.
                                // The chain is the leading run of keepNext rows PLUS
                                // the row it keeps with; with no leading keepNext row
                                // this is 1 row = S1038 unchanged.
                                let s1248_chain_rows = if std::env::var("OXI_S1248_DISABLE").is_ok()
                                {
                                    1
                                } else {
                                    let kn = |ri: usize| -> bool {
                                        next_table
                                            .rows
                                            .get(ri)
                                            .and_then(|r| r.cells.first())
                                            .and_then(|c| {
                                                c.blocks.iter().find_map(|b| match b {
                                                    Block::Paragraph(p) => Some(p.style.keep_next),
                                                    _ => None,
                                                })
                                            })
                                            .unwrap_or(false)
                                    };
                                    let mut k = 0usize;
                                    while k < next_table.rows.len() && kn(k) {
                                        k += 1;
                                    }
                                    (k + 1).min(next_table.rows.len().max(1))
                                };
                                // S1522 (2026-09-24): the pair is (whole chain height,
                                // height up to the first text line of the chain's LAST
                                // row). See `s1038_need_h` below for which one Word asks.
                                let (first_row_h, s1522_first_line_h) = if std::env::var("OXI_S1038_DISABLE").is_ok()
                                    || (self.doc_body_has_real_cjk
                                        && std::env::var("OXI_CJK_TABLE_KEEP_NEXT").ok().as_deref() != Some("1"))
                                    || next_table.rows.is_empty()
                                {
                                    (0.0, 0.0)
                                } else {
                                    let mut one = next_table.clone();
                                    one.rows.truncate(s1248_chain_rows);
                                    let mut rp: Vec<LayoutPage> = Vec::new();
                                    let mut re: Vec<LayoutElement> = Vec::new();
                                    let mut rc = LayoutCursor::new(start_y);
                                    let rels = self.layout_table(
                                        &one,
                                        start_x,
                                        &mut rc,
                                        content_width,
                                        grid_pitch,
                                        page.grid_char_pitch,
                                        page.grid_char_cw_ratio,
                                        start_y,
                                        content_height,
                                        page.size.width,
                                        page.size.height,
                                        &mut rp,
                                        &mut re,
                                        Some(block_idx + 1),
                                        page,
                                        false,
                                        None,
                                        None,
                                        0.0,
                                        0.0,
                                        false,
                                        None,
                                    );
                                    if rp.is_empty() {
                                        let whole = rels.iter()
                                            .chain(re.iter())
                                            .map(|el| el.y + el.height)
                                            .fold(start_y, f32::max)
                                            - start_y;
                                        // First text line of the chain's last row: per
                                        // cell the lowest-bottom text element, then the
                                        // tallest cell (the row splits per cell, each
                                        // cell contributes its first line).
                                        let last = s1248_chain_rows.max(1) - 1;
                                        let mut per_cell: std::collections::BTreeMap<usize, f32> =
                                            std::collections::BTreeMap::new();
                                        for el in rels.iter().chain(re.iter()) {
                                            if el.cell_row_index != Some(last) {
                                                continue;
                                            }
                                            if !matches!(el.content, LayoutContent::Text { .. }) {
                                                continue;
                                            }
                                            let b = el.y + el.height;
                                            let slot = per_cell.entry(el.cell_col_index.unwrap_or(0)).or_insert(b);
                                            if b < *slot {
                                                *slot = b;
                                            }
                                        }
                                        let first_line = per_cell.values().cloned().fold(0.0f32, f32::max);
                                        let first_line = if first_line > start_y { first_line - start_y } else { whole };
                                        (whole, first_line.min(whole))
                                    } else {
                                        (0.0, 0.0)
                                    }
                                };
                                // S1127 (2026-08-15, SHIPPED default-ON with S1126, opt-out
                // OXI_S1127_DISABLE): the heading's OWN
                // space-before is missing from `this_h` on the TABLE-follower
                // path. estimate_para_height drops a style-defined space_before,
                // and the paragraph-follower branch adds it back (S709/S1087 at
                // the `let this_h = if s709` site) — this branch never did, so a
                // heading is measured ~one gap too short against `remaining`.
                // technical__00549a8f p25: Heading2 sb=216tw=10.8, heading 15.8,
                // first row 17.2 → Word needs 43.8 of the 34.6 left and pushes;
                // Oxi compares 33.0 ≤ 34.6 and keeps the heading with the table
                // starting on p26 = an orphaned keepNext heading. Same reset key
                // as S1087 (has_direct_before), same exact/atLeast exclusion.
                let s1127_sb = if std::env::var("OXI_S1127_DISABLE").is_err()
                    && !para.style.has_direct_before
                    && para.style.line_spacing_rule.as_deref() != Some("exact")
                    && para.style.line_spacing_rule.as_deref() != Some("atLeast")
                {
                    if let (Some(bl), Some(pitch)) = (para.style.before_lines, grid_pitch) {
                        bl / 100.0 * s1480_lines_unit(page.doc_grid_no_type, pitch)
                    } else {
                        para.style.space_before.unwrap_or(0.0)
                    }
                } else {
                    0.0
                };
                let this_h = this_h + s1127_sb;
                // ...and the `!fresh_one_page` gate goes with it: whether the WHOLE
                // table would fit a fresh page is S1024's whole-move question, not
                // this one. 00549a8f's table is small (207.7pt, fresh1=true) and
                // has no keepNext row-chain, so S1024 leaves it alone; the first
                // row still cannot follow the heading, and Word still pushes.
                // S1521 (2026-09-24, opt-out OXI_S1038_SPLIT_DISABLE): S1038 asks
                // whether the WHOLE first row (or leading keepNext row-chain)
                // fits under the heading. That is Word's rule only when that
                // row CANNOT break: policies__0021ede1, where S1038 was derived,
                // has cantSplit on every row. legal__003b5088 has none, and Word
                // keeps `Division 3` + `[Heading inserted` on p27 with the first
                // TWO lines of an 8-line (107pt) row 1 under them, breaking the
                // row across the page. Faithful slice, 22 arms (rem 0..120 x
                // cantSplit on/off, Word PDF): without cantSplit the heading
                // stays on the page at every rem, down to a single row-1 line
                // beneath it; with cantSplit injected it moves until the whole
                // row fits (rem >= 96). So the whole-row orphan test applies
                // only when a row of the measured chain is cantSplit; otherwise
                // the ordinary row-split path places the table.
                // S1522 (2026-09-24, opt-out OXI_S1038_SPLIT_DISABLE = old whole-chain
                // test): the 785 gate on S1521 lost legal__0010437a (rows 0-2 keepNext,
                // row 0 tblHeader) and technical__00549a8f (row 0 tblHeader, one line),
                // where Word pushes the heading although no row is cantSplit. Faithful
                // slice, 44 arms (rem 0..120 x {as-is, cantSplit, tblHeader, keepNext}
                // on row 1, Word PDF): as-is keeps the heading with ONE line of row 1;
                // cantSplit and tblHeader move it until the whole row fits (rem 96);
                // keepNext moves it until the whole row AND the first line of row 2 fit
                // (rem 108). So the height Word needs under the heading is: every
                // keepNext row of the chain whole, then the row they keep with -- whole
                // when it is cantSplit / tblHeader, else its first text line only.
                let s1038_last = next_table.rows.get(s1248_chain_rows.max(1) - 1);
                let s1038_locked_row = std::env::var_os("OXI_S1038_SPLIT_DISABLE").is_some()
                    || s1038_last.map_or(true, |r| r.cant_split || r.header);
                let s1038_need_h = if s1038_locked_row { first_row_h } else { s1522_first_line_h };
                let s1038_row_orphan = std::env::var("OXI_S1038_DISABLE").is_err()
                    && (!self.doc_body_has_real_cjk
                                || std::env::var("OXI_CJK_TABLE_KEEP_NEXT").ok().as_deref() == Some("1"))
                                    && (!fresh_one_page
                                        || std::env::var("OXI_S1127_DISABLE").is_err())
                                    // The heading's estimate includes space that
                                    // can fall below the last rendered line. Its
                                    // own estimated overflow does not guarantee a break.
                                    && s1038_need_h > 0.0
                                    && this_h + s1038_need_h > remaining + 0.5
                                    && this_h + s1038_need_h <= content_height + 0.5;
                                // S1577 (2026-09-26, default ON for Latin documents, opt-out
                                // OXI_S1577_DISABLE): the follower-placement probe (a table laid
                                // after the heading whose first page carries no data row moves
                                // wholly, so the keepNext heading must go with it) is not a
                                // CJK-only rule. reports__0079718f p109/110: 「Table 3.4: Budgeted
                                // departmental statement…」 (TableHeading keepNext) then a table
                                // whose first row does not fit the 42pt left -- Word starts p110
                                // with both; Oxi left the heading at the p109 bottom.
                                let s1577 = !self.doc_body_has_real_cjk
                                    && std::env::var_os("OXI_S1577_DISABLE").is_none();
                                let table_follower_moves = if (s1577
                                    || (std::env::var_os("OXI_CJK_TABLE_FOLLOWER_PLACEMENT").is_some()
                                        && self.doc_body_has_real_cjk)) && current_splits
                                    && this_h <= remaining && remaining < content_height - 0.5
                                {
                                    let mut pp = Vec::new();
                                    // S1577c: the heading will be on this page, so the table
                                    // does not start on an empty page. Without a seed element
                                    // the row loop's `has_content` guard was false and row 0,
                                    // which Word and the real layout push whole, was split.
                                    let mut pe: Vec<LayoutElement> = if s1577 {
                                        elements.last().cloned().into_iter().collect()
                                    } else { Vec::new() };
                                    // S1577b: start the probe where the table really starts -- the
                                    // pending collapsed gap above the heading (the previous
                                    // paragraph's space-after is applied only when the heading is
                                    // laid out, so `cursor` still sits above it) and the heading's
                                    // own space-after, which the table arm adds before the table.
                                    // 0079718f p109: 20 empty paragraphs end with after=12; the
                                    // heading (2 x 11.5, after 1) then starts at 665.3 (Word Info(6)
                                    // 665.5), the table at 689.3 and its 37.35pt first row passes
                                    // the 718.6 bottom. From the bare cursor (653.3) the probe put
                                    // the row at 676.3 and saw it fit.
                                    let s1577_lead = if s1577 {
                                        let sb = para.style.space_before.unwrap_or(0.0);
                                        (prev_space_after.max(sb) - s1127_sb).max(0.0)
                                            + para.style.space_after.unwrap_or(0.0)
                                    } else { 0.0 };
                                    let mut pc = LayoutCursor::new(cursor.cursor_y + this_h + s1577_lead);
                                    if std::env::var_os("OXI_DBG_KN635").is_some() {
                                        eprintln!("[S1577] cy={:.2} this_h={:.2} lead={:.2} prev_sa={:.2} sb={:?} sa={:?}",
                                            cursor.cursor_y, this_h, s1577_lead, prev_space_after, para.style.space_before, para.style.space_after);
                                    }
                                    let _ = self.layout_table(next_table, start_x, &mut pc, content_width,
                                        grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                                        start_y, content_height, page.size.width, page.size.height,
                                        &mut pp, &mut pe, Some(block_idx + 1), page,
                                        false, None, None, 0.0, 0.0, false, None);
                                    let header_rows = next_table.rows.iter().take_while(|row| row.header).count();
                                    // S1577d: for Latin documents a repeated header row left on the
                                    // page keeps the heading too. technical__00549a8f p17: 「DMIS ID
                                    // Name」 + tblHeader row 0 stay (Word COM page 17), row 1 goes
                                    // to p18; counting only data rows sent the heading to p18.
                                    let first_data_row = if s1577 { 0 }
                                        else if header_rows < next_table.rows.len() { header_rows } else { 0 };
                                    pp.first().map_or(false, |first| {
                                        !first.elements.iter().any(|el| match &el.content {
                                            LayoutContent::Text { text, .. } => !text.trim().is_empty()
                                                && el.cell_row_index.map_or(false, |row| row >= first_data_row),
                                            _ => false,
                                        })
                                    })
                                } else { false };
                                if std::env::var("OXI_DBG_KN635").is_ok() {
                                    let t: String = para
                                        .runs
                                        .iter()
                                        .flat_map(|r| r.text.chars())
                                        .take(16)
                                        .collect();
                                    eprintln!("[KN635-TBL] {:?} this_h={:.1} tbl_e={:.1} rem={:.1} row_chain={} chain_rows={} fresh1={} splits={} bothfit={} push={}",
                                        t, this_h, e, remaining, row_chain,
                                        next_table.rows.len().saturating_sub(1), fresh_one_page,
                                        current_splits, both_fit_fresh,
                                        row_chain && fresh_one_page && current_splits && both_fit_fresh && this_h <= remaining);
                                    eprintln!(
                                        "[KN635-TBL2] first_row_h={:.1} first_line_h={:.1} locked={} need_h={:.1} s1038={}",
                                        first_row_h, s1522_first_line_h, s1038_locked_row, s1038_need_h, s1038_row_orphan
                                    );
                                }
                                if (row_chain
                                    && fresh_one_page
                                    && current_splits
                                    && both_fit_fresh
                                    && this_h <= remaining)
                                    || s1038_row_orphan || table_follower_moves
                                {
                                    // Same S963b transitive back-pull + push as the image
                                    // arm (the push is follower-agnostic).
                                    let s963b = !self.doc_body_has_real_cjk
                                        && std::env::var("OXI_S802B_DISABLE").is_err()
                                        && std::env::var("OXI_S963_DISABLE").is_err();
                                    let mut pull_from = block_idx;
                                    if s963b
                                        && !(num_columns > 1 && current_column + 1 < num_columns)
                                    {
                                        while pull_from > 0 {
                                            if let Some(Block::Paragraph(pp)) =
                                                page.blocks.get(pull_from - 1)
                                            {
                                                if pp.style.keep_next
                                                    && block_page_indices.get(pull_from - 1)
                                                        == Some(&current_page_idx)
                                                {
                                                    pull_from -= 1;
                                                    continue;
                                                }
                                            }
                                            break;
                                        }
                                    }
                                    let pulled: Vec<LayoutElement> = if pull_from < block_idx {
                                        let y0 = block_y_positions[pull_from] - 0.1;
                                        let (keep, moved): (
                                            Vec<LayoutElement>,
                                            Vec<LayoutElement>,
                                        ) = elements.drain(..).partition(|e| {
                                            if (std::env::var("OXI_KEEP_CHAIN_SOURCE").is_ok()
                || std::env::var("OXI_S1474_DISABLE").is_err()) {
                                                if let Some(index) = e.paragraph_index {
                                                    return index < pull_from || index >= block_idx;
                                                }
                                            }
                                            e.y < y0
                                        });
                                        elements = keep;
                                        moved
                                    } else {
                                        Vec::new()
                                    };
                                    let chain_end_old = cursor.cursor_y;
                                    if num_columns > 1 && current_column + 1 < num_columns {
                                        current_column += 1;
                                        start_x = col_x_positions[current_column];
                                        content_width = col_widths[current_column];
                                        cursor.set(col_band_top);
                                    } else {
                                        dbg_page_push(pages.len(), 0);
                                        pages.push(LayoutPage {
                                            width: page.size.width,
                                            height: page.size.height,
                                            elements: std::mem::take(&mut elements),
                                        });
                                        if let Some(g) = s755_geom.as_ref() {
                                            start_y = g.top(pages.len() + 1);
                                            content_height = g.ch(pages.len() + 1);
                                        }
                                        cursor.set(start_y);
                                        if std::env::var("OXI_S963_DISABLE").is_err() {
                                            current_column = 0;
                                            start_x = col_x_positions[0];
                                            content_width = col_widths[0];
                                            lm2_cells = 0;
                                            current_page_idx += 1;
                                            footnote_reserve_current = 0.0;
                                            footnote_ids_current_page.clear();
                                            s900_fold(
                                                &mut footnote_reserve_current,
                                                &mut footnote_ids_current_page,
                                                &mut s900_pending_deferred,
                                                current_page_idx,
                                            );
                                            commit_para_footnotes(
                                                &mut footnote_reserve_current,
                                                &mut footnote_ids_current_page,
                                                current_page_idx,
                                                block_idx,
                                            );
                                            if !pulled.is_empty() {
                                                let min_y = pulled
                                                    .iter()
                                                    .map(|e| e.y)
                                                    .fold(f32::MAX, f32::min);
                                                let dy = cursor.cursor_y - min_y;
                                                if std::env::var("OXI_DBG_KN635").is_ok() {
                                                    eprintln!("[S963B-TBL] pull {} blocks ({} els) dy={:.1}",
                                                        block_idx - pull_from, pulled.len(), dy);
                                                }
                                                for mut e in pulled {
                                                    e.y += dy;
                                                    if let LayoutContent::TableBorder {
                                                        ref mut y1,
                                                        ref mut y2,
                                                        ..
                                                    } = e.content
                                                    {
                                                        *y1 += dy;
                                                        *y2 += dy;
                                                    }
                                                    elements.push(e);
                                                }
                                                cursor.set(chain_end_old + dy);
                                                for bi in pull_from..block_idx {
                                                    if let Some(p) = block_page_indices.get_mut(bi)
                                                    {
                                                        *p = current_page_idx;
                                                    }
                                                    if let Some(y) = block_y_positions.get_mut(bi) {
                                                        *y += dy;
                                                    }
                                                }
                                            }
                                            *block_page_indices.last_mut().unwrap() =
                                                current_page_idx;
                                            *block_y_positions.last_mut().unwrap() =
                                                cursor.cursor_y;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Multi-column pre-check: advance column if paragraph won't fit.
                    // S723 (2026-07-03): DISABLED by default. This pre-S637 leftover
                    // whole-MOVED any paragraph that didn't fit the remaining column
                    // space to the next column — but Word line-SPLITS a normal
                    // paragraph across a column boundary (newspaper flow), which
                    // S637's per-line flow inside layout_paragraph already does.
                    // keepLines paragraphs (the ones Word DOES keep together across a
                    // column boundary) are pre-moved by the keepLines block above
                    // (3374). Firing this on NORMAL paragraphs made each column hold
                    // ~1 fewer paragraph → every paragraph cascaded one column forward
                    // → +1 page delta at each page crossing (probe2col +15, probe3col
                    // +21). Retained behind OXI_S723_DISABLE for A/B. Only kyotei /
                    // albaluna have num_columns>1; the single-column corpus is
                    // byte-identical (num_columns==1 → block never runs).
                    if num_columns > 1 && std::env::var("OXI_S723_DISABLE").is_ok() {
                        let est_h = self.estimate_para_height(
                            para,
                            self.s1211c_floor_body_width(para, content_width, page.grid_char_pitch, page.grid_char_cw_ratio),
                            grid_pitch,
                            None,
                            false,
                            None,
                            None,
                        );
                        let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                        if est_h > remaining && est_h <= effective_content_h {
                            if std::env::var("OXI_DBG_COL").is_ok() {
                                eprintln!("[COL] precheck block_idx={} col {}->{} (est_h={:.1} rem={:.1}) page={}", block_idx, current_column, if current_column+1<num_columns {current_column+1} else {0}, est_h, remaining, current_page_idx);
                            }
                            if current_column + 1 < num_columns {
                                current_column += 1;
                                start_x = col_x_positions[current_column];
                                content_width = col_widths[current_column];
                                cursor.set(col_band_top);
                            } else {
                                dbg_page_push(pages.len(), 0);
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: std::mem::take(&mut elements),
                                });
                                if let Some(g) = s755_geom.as_ref() {
                                    start_y = g.top(pages.len() + 1);
                                    content_height = g.ch(pages.len() + 1);
                                }
                                cursor.set(start_y);
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                                lm2_cells = 0;
                                current_page_idx += 1;
                                footnote_reserve_current = 0.0;
                                footnote_ids_current_page.clear();
                                s900_fold(
                                    &mut footnote_reserve_current,
                                    &mut footnote_ids_current_page,
                                    &mut s900_pending_deferred,
                                    current_page_idx,
                                );
                                commit_para_footnotes(
                                    &mut footnote_reserve_current,
                                    &mut footnote_ids_current_page,
                                    current_page_idx,
                                    block_idx,
                                );
                            }
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        }
                    }

                    let pages_before = pages.len();
                    // S758: does this paragraph START inside a wrapSquare band on
                    // the current page? Right-side bands only (v1): pass
                    // (band_bottom, width_reduction). Reduction = the part of the
                    // column the float + its 9pt distL eat; requires a real
                    // horizontal overlap so decorative off-column floats stay
                    // inert.
                    let (s758_para_band, s758_two_seg, wrap_advance) = self.body_paragraph_wrap_bands(
                        para, page, &s758_bands, current_page_idx,
                        cursor.cursor_y, start_x, content_width,
                    );
                    cursor.advance(wrap_advance);
                    // S1515 (2026-09-21, default ON, opt-out OXI_S1515_DISABLE): a
                    // paragraph that resumes BELOW a square-wrapped float starts at
                    // the band's bottom; the previous paragraph's after-spacing is
                    // absorbed by that move. technical__01242a0a p1: 'List of
                    // figure' (after=8) hosts a 104.75pt wrapSquare group at +29.9;
                    // Word sets 'Fig. 1' at 207.0 = band bottom 206.9, Oxi 214.6.
                    if wrap_advance > 0.0 && std::env::var_os("OXI_S1515_DISABLE").is_none() {
                        prev_space_after = (prev_space_after - wrap_advance).max(0.0);
                    }
                    if std::env::var("OXI_DBG773").is_ok() {
                        for s in &para.shapes {
                            eprintln!(
                                "[S773probe] blk={} shape {} w={:.1} h={:.1} pos={:?}",
                                block_idx,
                                s.shape_type,
                                s.width,
                                s.height,
                                s.position.as_ref().map(|p| (p.x, p.y))
                            );
                        }
                        for tb in &page.text_boxes {
                            if tb.anchor_block_index == block_idx {
                                eprintln!("[S773probe] blk={} tb w={:.1} h={:.1} pos={:?} wrap={:?} blocks={}",
                                    block_idx, tb.width, tb.height,
                                    tb.position.as_ref().map(|p| (p.x, p.y)), tb.wrap_type, tb.blocks.len());
                            }
                        }
                    }
                    if std::env::var("OXI_DBG758").is_ok() {
                        if let Some((bot, red, sh)) = s758_para_band {
                            let txt: String = para
                                .runs
                                .iter()
                                .flat_map(|r| r.text.chars())
                                .take(20)
                                .collect();
                            eprintln!("[B758] pg={} cur_y={:.1} bot={:.1} red={:.1} sh={:.1} ind_l={:?} ind_r={:?} txt={:?}",
                                current_page_idx, cursor.cursor_y, bot, red, sh,
                                para.style.indent_left, para.style.indent_right, txt);
                        }
                    }
                    // Round 29: pass the per-page effective content height so the
                    // paragraph's internal line-by-line page-break logic accounts
                    // for the footnote area below. Multi-page paragraphs with
                    // footnoteRefs get the same reservation on each spanned page
                    // (slight under-use on continuation pages, acceptable).
                    // Always pass lm2_cells: LM2 uses it for grid tracking,
                    // LM0 single-spacing uses it for cross-paragraph cumulative round.
                    let lm2_param = Some(&mut lm2_cells);

                    // COM-confirmed (2026-04-16, 683f p2 + minimal repro):
                    // content paragraph gets +0.5pt extra advance when adjacent to a RUN
                    // of ≥2 consecutive empty paragraphs, provided the paragraph on the
                    // far side of the empty run is a content paragraph (NOT a Table).
                    // 683f p1 exception: empty run after a table (P11 table → P12/P13 empty
                    // → P14 content) — Word does NOT +0.5 here.
                    let adjacent_to_empty_run = LayoutEngine::body_adjacent_to_empty_run(para, page, block_idx);
                    // S603: is the next sibling block a table? (page-bottom full-cell rule)
                    let next_block_is_table =
                        matches!(page.blocks.get(block_idx + 1), Some(Block::Table(_)));

                    // Step 0: bucket for per-page fn refs actually rendered by this
                    // paragraph. Used by the post-layout reserve-correction (Step 1).
                    let mut para_fn_refs_per_page: Vec<Vec<u32>> = Vec::new();
                    let mut s900_para_deferred: Vec<u32> = Vec::new();
                    let s903_next_borders_body: Option<&ParagraphBorders> =
                        page.blocks.get(block_idx + 1).and_then(|b| match b {
                            Block::Paragraph(p) => p.style.borders.as_ref(),
                            _ => None,
                        });
                    // S676: a pending drop-cap indents THIS paragraph's body to the right
                    // of the floated cap (start shifted, available width reduced).
                    let dc_indent = pending_dropcap.take().unwrap_or(0.0);
                    let dbg_para_start_y = cursor.cursor_y;
                    // S1249: the space the PREVIOUS paragraph left behind, kept
                    // before `prev_space_after` is overwritten with this one's.
                    // The chain back-pull below needs it: the pulled chain's
                    // trailing space is not part of the chain's own extent.
                    let s1249_prev_sa = prev_space_after;
                    let dbg_para_start_pages = pages.len();
                    if std::env::var("OXI_DBG_COL").is_ok() && num_columns > 1 {
                        eprintln!("[COL] enter block_idx={} col={} cursor_y={:.1} band_top={:.1} pages={}",
                            block_idx, current_column, cursor.cursor_y, col_band_top, pages.len());
                    }
                    let (mut para_elements, sa, final_col) = self.layout_paragraph(
                        para,
                        start_x + dc_indent,
                        &mut cursor,
                        content_width - dc_indent,
                        effective_content_h,
                        start_y,
                        page,
                        &mut pages,
                        &mut elements,
                        grid_pitch,
                        prev_para_style_id.as_deref(),
                        prev_contextual_spacing,
                        prev_autospacing_numid.as_deref(),
                        prev_borders.as_ref(),
                        prev_keep_next,
                        false,
                        prev_space_after,
                        Some(block_idx),
                        lm2_param,
                        Some(&mut mult_cumul_raw),
                        adjacent_to_empty_run,
                        next_block_is_table,
                        Some(&mut para_fn_refs_per_page),
                        // R7.53: first-line lenient — add back this para's
                        // own fn reserve delta so line 0 fits if it would
                        // without the para's footnotes.
                        delta_if_current,
                        // S168: per-fn heights for per-line lenient calculation.
                        &para_fn_heights_map,
                        // S637: multi-column column-flow state (num_columns==1 off
                        // the heterogeneous path → no-op, byte-identical).
                        num_columns,
                        current_column,
                        &col_x_positions,
                        col_band_top,                                        // S749
                        false,                                               // S691: body context
                        footer_tight,                                        // S726
                        s755_geom.as_ref(),                                  // S755
                        s758_para_band,                                      // S758
                        s758_two_seg,                                        // S-TWOSEG
                        (footnote_reserve_current + delta_if_current) > 0.0, // S835
                        footnote_reserve_current,                            // S900
                        Some(&mut s900_para_deferred),                       // S900
                        s903_next_borders_body,                              // S903
                        s916_split,                                          // S916
                        Some((&s758_bands, current_page_idx)),
                    );
                    prev_space_after = sa;
                    // S1497: whatever the host paragraph's own lines did, the
                    // content after it resumes below the band (the S734
                    // reservation used to guarantee this by advancing first).
                    if let Some(&(off, h)) = s1497_mid.get(&block_idx) {
                        if let Some(&(fp, fy)) = s734_flow_pos.get(&block_idx).or(s1089_flow_pos.get(&block_idx)) {
                            let bottom = fy + off + h;
                            if fp == current_page_idx && bottom > cursor.cursor_y && bottom < start_y + content_height {
                                if std::env::var_os("OXI_S1500_DISABLE").is_none() {
                                    s1500_unpushed = Some((block_idx + 1, current_page_idx, cursor.cursor_y));
                                }
                                cursor.set(bottom);
                            }
                        }
                    }
                    S1497_BAND.with(|c| c.set(None));
                    // S1461: drop the cursor below a page-anchored
                    // wrapTopAndBottom shape hosted by this block.
                    if let Some(&bot) = s1461_tb_shapes.get(&block_idx) {
                        if bot > cursor.cursor_y && bot < start_y + content_height {
                            cursor.set(bot);
                        }
                    }
                    // S1459: drop the cursor below a page-anchored
                    // wrapTopAndBottom box hosted by this block.
                    if let Some(&bot) = s1459_page_boxes.get(&block_idx) {
                        if bot > cursor.cursor_y && bot < start_y + content_height {
                            cursor.set(bot);
                        }
                    }
                    if std::env::var("OXI_DBG_PARA").is_ok() {
                        let txt: String = para
                            .runs
                            .iter()
                            .flat_map(|r| r.text.chars())
                            .take(22)
                            .collect();
                        let (mut ymin, mut ymax) = (f32::INFINITY, f32::NEG_INFINITY);
                        for e in &para_elements {
                            if matches!(e.content, LayoutContent::Text { .. }) {
                                ymin = ymin.min(e.y);
                                ymax = ymax.max(e.y);
                            }
                        }
                        eprintln!("[PARA] blk={} start_y={:.1} end_cur={:.1} pages+={} el_y=[{:.1},{:.1}] n={} txt={:?}",
                            block_idx, dbg_para_start_y, cursor.cursor_y, pages.len()-dbg_para_start_pages,
                            ymin, ymax, para_elements.len(), txt);
                    }
                    let s1324_col_before = current_column;
                    // S637: layout_paragraph may have advanced the column (flowing
                    // an overflowing line into the next column on the same page)
                    // or page-pushed (resetting to column 0). Sync the loop state.
                    if num_columns > 1 && final_col != current_column {
                        current_column = final_col;
                        start_x = col_x_positions[current_column];
                        content_width = col_widths[current_column];
                    }
                    // S1324 (2026-09-05, default ON, opt-out OXI_S1324_DISABLE): a
                    // keepNext chain stranded at the foot of the column its follower
                    // just LEFT (the S1323 whole-move into the next column) follows
                    // it to that column's top -- Word's keepNext unit crosses a
                    // column boundary as one, exactly as it crosses a page (S960).
                    // reports__167853 p28: ●自立支援事業 (keepNext+keepLines) stayed
                    // at the foot of column 0 while its 4-line paragraph opened
                    // column 1; Word's column 1 begins with the heading (20/22 lines
                    // after balancing, Oxi 21/21 with the heading stranded).
                    if num_columns > 1
                        && pages.len() == pages_before
                        && final_col == s1324_col_before + 1
                        && !para.style.page_break_before
                        && std::env::var("OXI_S1324_DISABLE").is_err()
                    {
                        let new_x = col_x_positions[final_col];
                        let old_x = col_x_positions[s1324_col_before];
                        // Whole-move only: a split (S790) keeps lines in the old
                        // column and the chain stays with them.
                        let on_new = !para_elements.is_empty()
                            && para_elements.iter().all(|e| e.x >= new_x - 0.5);
                        let in_old_col = |pi: usize| {
                            elements.iter().any(|e| {
                                e.paragraph_index == Some(pi) && e.x < new_x - 0.5
                            })
                        };
                        let mut pull_from = block_idx;
                        if on_new {
                            while pull_from > 0 {
                                if let Some(Block::Paragraph(pp)) = page.blocks.get(pull_from - 1) {
                                    if pp.style.keep_next
                                        && block_page_indices.get(pull_from - 1) == Some(&current_page_idx)
                                        && in_old_col(pull_from - 1)
                                    {
                                        pull_from -= 1;
                                        continue;
                                    }
                                }
                                break;
                            }
                        }
                        if pull_from < block_idx {
                            let pulled_here: Vec<usize> = (pull_from..block_idx).collect();
                            let mine = |e: &LayoutElement| {
                                e.paragraph_index.is_some_and(|pi| pulled_here.contains(&pi))
                            };
                            let chain_min_y = elements
                                .iter()
                                .filter(|e| mine(e))
                                .map(|e| e.y)
                                .fold(f32::INFINITY, f32::min);
                            // Same contract as S960: plain text only, nothing unowned
                            // in the vacated band, and the old column keeps a body.
                            let region_clean = chain_min_y.is_finite()
                                && !elements.iter().any(|e| {
                                    let m = mine(e);
                                    (e.x < new_x - 0.5 && e.y >= chain_min_y - 0.1 && !m)
                                        || (m && !matches!(e.content, LayoutContent::Text { .. }))
                                });
                            let old_keeps_body = elements.iter().any(|e| {
                                e.x < new_x - 0.5 && e.paragraph_index.is_some() && !mine(e)
                            });
                            let s1324_gap = s1249_prev_sa.max(para.style.space_before.unwrap_or(0.0));
                            let chain_advance = dbg_para_start_y - chain_min_y + s1324_gap;
                            let fits = chain_advance > 0.0
                                && cursor.cursor_y + chain_advance <= start_y + effective_content_h + 0.05;
                            if region_clean && old_keeps_body && fits {
                                let mut pulled: Vec<LayoutElement> = Vec::new();
                                elements.retain(|e| {
                                    if mine(e) {
                                        pulled.push(e.clone());
                                        false
                                    } else {
                                        true
                                    }
                                });
                                let dx = new_x - old_x;
                                let dy = col_band_top - chain_min_y;
                                for e in pulled.iter_mut() {
                                    e.x += dx;
                                    e.y += dy;
                                }
                                for e in para_elements.iter_mut() {
                                    e.y += chain_advance;
                                }
                                pulled.extend(std::mem::take(&mut para_elements));
                                para_elements = pulled;
                                cursor.advance(chain_advance);
                                for pi in pulled_here.iter() {
                                    if let Some(y) = block_y_positions.get_mut(*pi) {
                                        *y += dy;
                                    }
                                }
                                if std::env::var("OXI_DBG_COL").is_ok() {
                                    eprintln!(
                                        "[COL] S1324 pull {} block(s) into col {} dy={:.1} advance={:.1}",
                                        block_idx - pull_from, final_col, dy, chain_advance
                                    );
                                }
                            }
                        }
                    }
                    // ── S960 (default ON, opt-out OXI_S960_DISABLE): pull a
                    // stranded keepNext chain onto the page its follower
                    // actually reached.
                    //
                    // S802B already does this, but it lives inside `if do_push`
                    // — it only fires when the PREDICTION pushed. A chain whose
                    // pairwise checks all pass, yet overflows during real
                    // layout, leaves its heads behind: policies__00148f8d's TOC
                    // kept «Part 10» + «Division 1» at the p8 bottom while
                    // «140.» flowed to p9 (both checks fit: 21.8+24.0 <= 62.6
                    // and 27.0+12.0 <= 40.5). Word keeps all three together.
                    //
                    // Word's rule, MEASURED (tools/metrics/_pb_kchain_gen.py,
                    // 68 cases over N ∈ {2,3,5,8,60} × body lines × widowControl
                    // × start position, page capacity calibrated empirically):
                    // the whole contiguous keepNext run moves as ONE unit — no
                    // maximal-suffix split at any N — and it stops only when
                    // the unit cannot fit a fresh page (N=60 fills the fresh
                    // page to capacity and splits naturally) or would blank the
                    // page it came from. So the cutoff is actual geometry, not
                    // an estimate, and no page is ever left empty.
                    let s960_added = pages.len() - pages_before;
                    // S970 v2: the table recorded just before this paragraph keeps
                    // it (its terminal paragraph is keepNext). If the paragraph
                    // whole-moved to a fresh page, the table must follow it — pull
                    // the already-emitted table elements onto this page, ahead of
                    // the paragraph. Same actual-geometry contract as S960: no page
                    // is pushed here, the follower's own layout already performed
                    // the transition, and the table's element range is moved intact
                    // (borders and shading have no paragraph_index, so a range move
                    // is the only correct one).
                    if let Some((tbl_blk, tbl_page, e0, e1, tbl_top)) = s970_pending.take() {
                        // The follower must have landed on the page AFTER the
                        // table's. It gets there two ways — a natural push inside
                        // layout_paragraph (s960_added == 1, current_page_idx not yet
                        // advanced) or a pre-layout push by the block loop
                        // (s960_added == 0, current_page_idx already advanced) — and
                        // legal__0010437a takes the second. One expression covers both.
                        if current_page_idx + s960_added == tbl_page + 1
                            && tbl_blk + 1 == block_idx
                            && !para.style.page_break_before
                        {
                            let old_idx = tbl_page;
                            let on_old = pages[old_idx]
                                .elements
                                .iter()
                                .any(|e| e.paragraph_index == Some(block_idx));
                            let on_new = para_elements
                                .iter()
                                .any(|e| e.paragraph_index == Some(block_idx));
                            let range_ok = e1 <= pages[old_idx].elements.len() && e0 < e1;
                            // Never blank the page the table came from.
                            let old_keeps_body = range_ok
                                && pages[old_idx].elements.iter().enumerate().any(|(i, e)| {
                                    (i < e0 || i >= e1) && e.paragraph_index.is_some()
                                });
                            let (new_top, new_ch) = s755_geom
                                .as_ref()
                                .map_or((start_y, effective_content_h), |g| {
                                    (g.top(tbl_page + 2), g.ch(tbl_page + 2))
                                });
                            // The table's ACTUAL extent on the old page. S960 could
                            // use `dbg_para_start_y - chain_top` because its follower
                            // was still on the old page when that cursor was sampled;
                            // here the follower may already have been pre-pushed, so
                            // that difference is meaningless (measured -35.6 on
                            // legal__0010437a). Measure the emitted elements instead.
                            let tbl_advance = if range_ok {
                                pages[old_idx].elements[e0..e1]
                                    .iter()
                                    .map(|e| e.y + e.height)
                                    .fold(f32::NEG_INFINITY, f32::max)
                                    - tbl_top
                            } else {
                                0.0
                            };
                            let fits_fresh = tbl_advance > 0.0
                                && cursor.cursor_y + tbl_advance <= new_top + new_ch + 0.05;
                            if !on_old && on_new && range_ok && old_keeps_body && fits_fresh {
                                let mut pulled: Vec<LayoutElement> =
                                    pages[old_idx].elements.drain(e0..e1).collect();
                                let dy = new_top - tbl_top;
                                for e in pulled.iter_mut() {
                                    e.y += dy;
                                    if let LayoutContent::TableBorder {
                                        ref mut y1,
                                        ref mut y2,
                                        ..
                                    } = e.content
                                    {
                                        *y1 += dy;
                                        *y2 += dy;
                                    }
                                }
                                for e in para_elements.iter_mut() {
                                    e.y += tbl_advance;
                                }
                                pulled.extend(std::mem::take(&mut para_elements));
                                para_elements = pulled;
                                cursor.advance(tbl_advance);
                                if let Some(pi) = block_page_indices.get_mut(tbl_blk) {
                                    *pi = tbl_page + 1;
                                }
                                if let Some(y) = block_y_positions.get_mut(tbl_blk) {
                                    *y += dy;
                                }
                                if std::env::var("OXI_DBG_KN635").is_ok() {
                                    eprintln!(
                                        "[S970] pull table blk={} els={}..{} \
tbl_top={:.1} advance={:.1} new_top={:.1}",
                                        tbl_blk, e0, e1, tbl_top, tbl_advance, new_top
                                    );
                                }
                            }
                        }
                    }
                    if std::env::var("OXI_S960_DISABLE").is_err()
                        && !self.doc_body_has_real_cjk
                        && (num_columns == 1
                            || (s1324_col_before + 1 == num_columns && final_col == 0))
                        && s960_added == 1
                        && !para.style.page_break_before
                        && std::env::var("OXI_S802").is_err()
                    {
                        let old_idx = pages.len() - 1;
                        // A whole-move is the only shape this may touch: an S916
                        // n−2/2 split or an S790 widow split leaves some of this
                        // paragraph's elements on the old page, and pulling the
                        // heads across such a split would break those rules.
                        let on_old = pages[old_idx]
                            .elements
                            .iter()
                            .any(|e| e.paragraph_index == Some(block_idx));
                        let on_new = para_elements
                            .iter()
                            .any(|e| e.paragraph_index == Some(block_idx));
                        if !on_old && on_new {
                            let mut pull_from = block_idx;
                            while pull_from > 0 {
                                match page.blocks.get(pull_from - 1) {
                                    Some(Block::Paragraph(p))
                                        if p.style.keep_next
                                            && !p.style.page_break_before
                                            && p.style.borders.is_none()
                                            && !p.runs.iter().any(|r| r.footnote_ref.is_some())
                                            && block_page_indices.get(pull_from - 1)
                                                == Some(&current_page_idx) =>
                                    {
                                        pull_from -= 1;
                                    }
                                    _ => break,
                                }
                            }
                            // v1 scope: two or more heads. A single keepNext head
                            // is what the S635/S709/S914/S925/S934/S948 pair rules
                            // already predict; only the TRANSITIVE case (the chain
                            // that no pairwise check can see) is new here.
                            // S1176 (2026-08-20, opt-out OXI_S1176_DISABLE): a
                            // SINGLE head is admitted too — the pair rules predict
                            // from ESTIMATES, and when the real layout then
                            // whole-moves the follower (`!on_old && on_new`), the
                            // predicted-fit head is stranded at the page bottom in
                            // exactly the way the _pb_kchain derivation forbids
                            // (the keepNext unit moves as one, at every N — N=1 is
                            // the degenerate case). legal__000ad039's «Authors
                            // Cited» (direct keepNext, after=720) stays at the
                            // p13 bottom while Word starts p14 with it.
                            let chain_len = block_idx - pull_from;
                            let s1176_single =
                                chain_len == 1 && std::env::var("OXI_S1176_DISABLE").is_err();
                            let pulled_here: Vec<usize> = (pull_from..block_idx).collect();
                            let chain_min_y = pages[old_idx]
                                .elements
                                .iter()
                                .filter(|e| {
                                    e.paragraph_index
                                        .is_some_and(|pi| pulled_here.contains(&pi))
                                })
                                .map(|e| e.y)
                                .fold(f32::INFINITY, f32::min);
                            // Anything unowned sitting in the vacated band (a
                            // float, a paragraph border, a shape) would be left
                            // behind by a paragraph_index drain — bail instead.
                            // The chain must also be PLAIN TEXT: an image or a
                            // box rect inside it may be positioned by machinery
                            // that keys off the block's page (float bands,
                            // anchors), which a bare element move would desync.
                            let source_column_x = col_x_positions[s1324_col_before];
                            let region_clean = (chain_len >= 2 || s1176_single)
                                && chain_min_y.is_finite()
                                && !pages[old_idx].elements.iter().any(|e| {
                                    let mine = e
                                        .paragraph_index
                                        .is_some_and(|pi| pulled_here.contains(&pi));
                                    (e.y >= chain_min_y - 0.1
                                        && (num_columns == 1 || e.x >= source_column_x - 0.5)
                                        && !mine)
                                        || (mine && num_columns > 1 && e.x < source_column_x - 0.5)
                                        || (mine
                                            && !matches!(e.content, LayoutContent::Text { .. }))
                                });
                            // Never blank the page the chain came from (Word's
                            // own cutoff: the N=60 probe splits rather than push).
                            let old_keeps_body = pages[old_idx].elements.iter().any(|e| {
                                e.paragraph_index
                                    .is_some_and(|pi| !pulled_here.contains(&pi))
                            });
                            // S1249 (default ON, opt-out OXI_S1249_DISABLE): `dbg_para_start_y` is the
                            // cursor at this paragraph's START, i.e. the chain's last
                            // element bottom -- the collapsed gap BETWEEN the chain and
                            // this paragraph is applied as this paragraph's leading and
                            // is therefore missing from the advance. On the old page that
                            // leading was dropped (the paragraph had moved to a page top);
                            // once the chain is pulled over, the paragraph is interior
                            // again and the gap has to come back. legal__000ad039: the
                            // "Authors Cited" heading carries w:after=720 (36pt) and Oxi
                            // shifted its follower by the heading's 13.8pt line alone.
                            let s1249_gap = if std::env::var("OXI_S1249_DISABLE").is_err() {
                                s1249_prev_sa.max(para.style.space_before.unwrap_or(0.0))
                            } else {
                                0.0
                            };
                            let chain_advance = dbg_para_start_y - chain_min_y + s1249_gap;
                            let (new_top, new_ch) = s755_geom
                                .as_ref()
                                .map_or((start_y, effective_content_h), |g| {
                                    (g.top(pages.len() + 1), g.ch(pages.len() + 1))
                                });
                            let fits_fresh = chain_advance > 0.0
                                && cursor.cursor_y + chain_advance <= new_top + new_ch + 0.05;
                            if region_clean && old_keeps_body && fits_fresh {
                                let mut pulled: Vec<LayoutElement> = Vec::new();
                                pages[old_idx].elements.retain(|e| {
                                    if e.paragraph_index
                                        .is_some_and(|pi| pulled_here.contains(&pi))
                                    {
                                        pulled.push(e.clone());
                                        false
                                    } else {
                                        true
                                    }
                                });
                                let chain_dy = new_top - chain_min_y;
                                for e in pulled.iter_mut() {
                                    e.y += chain_dy;
                                    if num_columns > 1 {
                                        e.x += col_x_positions[0] - source_column_x;
                                    }
                                }
                                for e in para_elements.iter_mut() {
                                    e.y += chain_advance;
                                }
                                pulled.extend(std::mem::take(&mut para_elements));
                                para_elements = pulled;
                                cursor.advance(chain_advance);
                                for pi in pulled_here.iter() {
                                    if let Some(p) = block_page_indices.get_mut(*pi) {
                                        *p = current_page_idx + 1;
                                    }
                                    if let Some(y) = block_y_positions.get_mut(*pi) {
                                        *y += chain_dy;
                                    }
                                }
                                if std::env::var("OXI_DBG_KN635").is_ok() {
                                    eprintln!(
                                        "[KCHAIN-ACTUAL] pull_from={} block_idx={} \
old_page={} chain_advance={:.1} chain_min_y={:.1} new_top={:.1} fresh_bottom={:.1}",
                                        pull_from,
                                        block_idx,
                                        old_idx,
                                        chain_advance,
                                        chain_min_y,
                                        new_top,
                                        new_top + new_ch
                                    );
                                }
                            }
                        }
                    }
                    elements.extend(para_elements);
                    if std::env::var("OXI_FN_PROBE").is_ok() && !para_fn_refs_per_page.is_empty() {
                        let any_ref = para_fn_refs_per_page.iter().any(|v| !v.is_empty());
                        if any_ref {
                            eprintln!(
                                "[FN_LINE_REFS] block_idx={} per_page={:?}",
                                block_idx, para_fn_refs_per_page
                            );
                        }
                    }

                    // Track page/column breaks that happened inside layout_paragraph
                    let pages_added = pages.len() - pages_before;
                    if std::env::var("OXI_DBG_COL").is_ok() && num_columns > 1 {
                        let txt: String = para
                            .runs
                            .iter()
                            .flat_map(|r| r.text.chars())
                            .take(12)
                            .collect();
                        eprintln!("[COL] para block_idx={} col={} pages_added={} page={} cursor_y={:.1} cw={:.1} txt={:?}",
                            block_idx, current_column, pages_added, current_page_idx, cursor.cursor_y, content_width, txt);
                    }
                    // Step 1 partial: attribute per-line fn refs to the page each
                    // line actually rendered on. para_fn_refs_per_page[i] holds
                    // the ids on the (start_page + i)-th page. Always runs —
                    // pages_added==0 has 1 bucket == current page's slot.
                    // S793 (2026-07-12): align buckets to the paragraph's FINAL
                    // page — pre-line-loop pushes inside layout_paragraph (the
                    // list-marker pre-break, band pushes) add pages WITHOUT
                    // opening a bucket, so `start + offset` mis-attributed a
                    // whole-moved paragraph's refs to the PREVIOUS page
                    // (nyserda: the sole footnote rendered on p17 while its
                    // reference line sits on p18; Word puts both on p18). The
                    // last bucket is always the final page; line-loop pushes
                    // open buckets 1:1, so end-alignment maps every bucket to
                    // the page its lines actually rendered on.
                    let start_page_for_para = if std::env::var("OXI_S793_DISABLE").is_ok() {
                        current_page_idx
                    } else {
                        (current_page_idx + pages_added)
                            .saturating_sub(para_fn_refs_per_page.len().saturating_sub(1))
                    };
                    for (offset, refs) in para_fn_refs_per_page.iter().enumerate() {
                        let page_i = start_page_for_para + offset;
                        while page_fn_refs.len() <= page_i {
                            page_fn_refs.push(Vec::new());
                        }
                        for id in refs {
                            if !page_fn_refs[page_i].contains(id) {
                                page_fn_refs[page_i].push(*id);
                            }
                        }
                    }
                    if pages_added > 0 {
                        // S637: column state is now handled INSIDE layout_paragraph
                        // (it flows overflow into the next column when one exists,
                        // and only page-pushes — resetting to column 0 — when all
                        // columns are exhausted). The `final_col` sync right after
                        // the call already updated current_column/start_x/content_width.
                        // The former half-baked column-advance here (which advanced
                        // AFTER layout_paragraph had already placed overflow on a new
                        // page's column 0) is removed; only the page-index advance
                        // and footnote bookkeeping remain.
                        current_page_idx += pages_added;
                        // Update block_page_index: if the paragraph moved entirely
                        // to the new page (no elements left on the old page), update
                        // the index so footnote rendering assigns to the correct page.
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        if std::env::var("OXI_FN_PROBE").is_ok() {
                            eprintln!("[FN_MID_BREAK] block_idx={} pages_added={} now_page={} reserve_before_clear={:.1} ids={:?}",
                                block_idx, pages_added, current_page_idx,
                                footnote_reserve_current, footnote_ids_current_page);
                        }
                        // 2026-05-05 Track A (Session 55+): pre-layout commit added
                        // ALL of this paragraph's fn refs to OLD page reserve. After
                        // mid-break, fn markers on lines that landed on NEW page
                        // actually render there. Reset reserve and re-commit only
                        // refs from FINAL spanned page (current_page_idx after
                        // pages_added increment). Without this, NEW page's body
                        // overflow check sees reserve=0 and over-packs body into
                        // fn area, silently dropping the fns (b837 p5 cascade).
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        s900_fold(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            &mut s900_pending_deferred,
                            current_page_idx,
                        );
                        // S829 (2026-07-13, opt-out OXI_S829_DISABLE): re-commit
                        // the refs the AREA mapping above assigns to the FINAL
                        // page, using the SAME saturating arithmetic. The literal
                        // get(pages_added) assumed one bucket per spanned page;
                        // a break path that doesn't push a per-line bucket
                        // (nyserda block 82: 1 bucket, pages_added=1) returned
                        // None → the new page carried ZERO reserve while the fn
                        // area still rendered there → the body packed into the
                        // fn area (Word widow-pushes the following paragraph).
                        // For the normal shape (buckets == pages_added+1) the
                        // last bucket maps to the final page — byte-identical.
                        let s829 = std::env::var("OXI_S829_DISABLE").is_err();
                        let s829_start = current_page_idx
                            .saturating_sub(para_fn_refs_per_page.len().saturating_sub(1));
                        let final_refs: Vec<u32> = if s829 {
                            para_fn_refs_per_page
                                .iter()
                                .enumerate()
                                .filter(|(o, _)| s829_start + o == current_page_idx)
                                .flat_map(|(_, refs)| refs.iter().copied())
                                .collect()
                        } else {
                            para_fn_refs_per_page
                                .get(pages_added)
                                .map(|v| v.clone())
                                .unwrap_or_default()
                        };
                        for id in &final_refs {
                            if !footnote_ids_current_page.contains(id) {
                                if footnote_ids_current_page.is_empty() {
                                    // S160: see estimate-path comment near line 1934.
                                    // S596b: no-docGrid separator = one footnote line.
                                    footnote_reserve_current += footnote_sep_alloc(*id);
                                }
                                footnote_ids_current_page.push(*id);
                                footnote_reserve_current += estimate_footnote_h(*id);
                            }
                        }
                    } else {
                        // R7.53 (2026-05-13): non-spanning case — paragraph
                        // stayed on the current page. Pre-commit was deferred
                        // (see comment near mod.rs:1924). Now commit this
                        // para's footnotes that actually rendered on the
                        // start page (para_fn_refs_per_page[0]).
                        if let Some(start_refs) = para_fn_refs_per_page.first() {
                            for id in start_refs {
                                if !footnote_ids_current_page.contains(id) {
                                    if footnote_ids_current_page.is_empty() {
                                        // S160: see estimate-path comment near line 1934.
                                        // S596b: no-docGrid separator = one footnote line.
                                        footnote_reserve_current += footnote_sep_alloc(*id);
                                    }
                                    footnote_ids_current_page.push(*id);
                                    footnote_reserve_current += estimate_footnote_h(*id);
                                }
                            }
                        }
                    }
                    // S900: deferred notes belong to the NEXT page's area —
                    // attribute their ids there and queue their reserve for the
                    // page-transition fold (the target page's body must keep
                    // room for them; 81e80 Word p3 opens its area with the
                    // deferred 16/17/18 ahead of p3's own refs).
                    if !s900_para_deferred.is_empty() && std::env::var("OXI_S900_DISABLE").is_err()
                    {
                        let target = current_page_idx + 1;
                        while page_fn_refs.len() <= target {
                            page_fn_refs.push(Vec::new());
                        }
                        for (k, id) in s900_para_deferred.iter().enumerate() {
                            if !page_fn_refs[target].contains(id) {
                                page_fn_refs[target].push(*id);
                            }
                            let mut h = estimate_footnote_h(*id);
                            if k == 0 {
                                h += footnote_sep_alloc(*id);
                            }
                            s900_pending_deferred.push((target, *id, h));
                        }
                        if std::env::var("OXI_DBG900").is_ok() {
                            eprintln!(
                                "[S900] blk={} deferred {:?} -> page {}",
                                block_idx, s900_para_deferred, target
                            );
                        }
                    }
                    // Round 30: render shapes attached to this paragraph (e.g.
                    // bracketPair preset frame around the date block in
                    // b837808d0555). The shape's anchor reference uses the
                    // paragraph's start Y position; pos.y is the offset from
                    // the paragraph start in points.
                    let para_anchor_y =
                        block_y_positions.get(block_idx).copied().unwrap_or(start_y);
                    for shape in &para.shapes {
                        if let Some(ref pos) = shape.position {
                            // S1460 (2026-09-17, default ON, opt-out
                            // OXI_S1460_DISABLE): resolve the shape's DECLARED
                            // anchors instead of assuming column/paragraph.
                            // `vertical_shape_origin` already implements the
                            // whole table and falls through to
                            // (margin.left + pos.x, anchor_y + pos.y) for the
                            // text/column/paragraph cases, so every shape that
                            // does not declare page/margin is unchanged.
                            // golden parttime's title hosts a VML box declaring
                            // mso-position-vertical-relative:page with
                            // margin-top 107.05 -- Word draws it at the PAGE's
                            // 107.05, the hardcoded form put it at
                            // 34.85 + 107.05 = 141.9.
                            let (sx, sy) = if std::env::var("OXI_S1460_DISABLE").is_err() {
                                LayoutEngine::vertical_shape_origin(page, shape, para_anchor_y)
                            } else {
                                (page.margin.left + pos.x, para_anchor_y + pos.y)
                            };
                            let content = shape_fill_boxrect(shape).unwrap_or_else(|| {
                                LayoutContent::PresetShape {
                                    shape_type: shape.shape_type.clone(),
                                    stroke_color: shape.stroke_color.clone(),
                                    stroke_width: shape.stroke_width.unwrap_or(0.75),
                                    flip_h: shape.flip_h,
                                    flip_v: shape.flip_v,
                                    arrow_head: shape.arrow_head,
                                    arrow_tail: shape.arrow_tail,
                                }
                            });
                            elements.push(LayoutElement::new(
                                sx,
                                sy,
                                shape.width,
                                shape.height,
                                content,
                            ));
                        }
                    }

                    // page_break_after: render the (typically empty) paragraph
                    // on the current page, then force a new page for the NEXT
                    // block. Used for the inline-br-in-empty-paragraph pattern;
                    // see `project_empty_br_para_stub.md`.
                    if para.style.page_break_after && !elements.is_empty() {
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        if let Some(g) = s755_geom.as_ref() {
                            start_y = g.top(pages.len() + 1);
                            content_height = g.ch(pages.len() + 1);
                        }
                        cursor.set(start_y);
                        // The preceding paragraph's trailing gap belongs to the
                        // page it ended on, including when a table follows.
                        // S1550 (2026-09-25, default ON, opt-out OXI_S1550_DISABLE; was
                        // the sleeping opt-in OXI_PAGE_BREAK_SPACING). reports__0079718f
                        // (compat 15): after every break-only paragraph (after 0 / 160 /
                        // 200) the next block sits at the body top (133.2-134.6) in
                        // Word's PDF whether it is a paragraph or a table; Oxi added
                        // the 10pt (8pt for 160) before a TABLE only (143.6 / 141.6) —
                        // the paragraph path already drops it at the page top.
                        if std::env::var("OXI_PAGE_BREAK_SPACING").is_ok()
                            || std::env::var("OXI_S1550_DISABLE").is_err()
                        {
                            prev_space_after = 0.0;
                        }
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        current_page_idx += 1;
                        lm2_cells = 0;
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        s900_fold(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            &mut s900_pending_deferred,
                            current_page_idx,
                        );
                    }

                    prev_para_style_id = para.style.style_id.clone();
                    prev_contextual_spacing = para.style.contextual_spacing;
                    // S1516 (2026-09-21): the previous LIST id is tracked for every
                    // list item, tagged with whether its after-spacing was auto
                    // ("|auto") or explicit ("|plain"); see paragraph_spacing_before.
                    prev_autospacing_numid = if std::env::var_os("OXI_S1516_DISABLE").is_none() {
                        para.style.num_id.clone().map(|n| format!("{}|{}", n, if para.style.after_autospacing { "auto" } else { "plain" }))
                    } else if para.style.after_autospacing {
                        para.style.num_id.clone()
                    } else {
                        None
                    };
                    prev_borders = para.style.borders.clone();
                    prev_keep_next = para.style.keep_next; // S739
                }
                Block::Table(table) => {
                    // COM-confirmed: prev paragraph's space_after is always added before table
                    if std::env::var("OXI_DBG_TBLSTART").is_ok() {
                        eprintln!("[TBLARM] blk={} cur={:.2} prev_space_after={:.2}", block_idx, cursor.cursor_y, prev_space_after);
                    }
                    cursor.advance(prev_space_after);
                    prev_space_after = 0.0;

                    let is_floating = table.style.position.is_some();
                    let mut saved_cursor_y = cursor.cursor_y;
                    let mut committed_float_fit = None;
                    let mut move_float_to_next_page = false;
                    // S1468 (2026-09-18, default ON, opt-out OXI_S1468_DISABLE):
                    // promotes the OXI_FLOAT_TABLE_REFLOW checkpoint. A floating
                    // table whose remainder cannot sit in the gap above the page's
                    // own flow content moves on as a unit instead of being packed
                    // in. policies__00602e8a is eight "Provision:" tables, each a
                    // `tblpPr vertAnchor="page" tblpY="2431"` float: Word puts
                    // table 2168's row 0 on page 4 (y 122.25..173.25) and rows
                    // 1..3 on page 5 (from 114.75) while page 4 already carries
                    // the previous table's tail at 328.5..500.2. Oxi packed the
                    // remaining rows onto page 4 and came out 8 pages against
                    // Word's 9 (46 paragraphs at -1). With the flag the document
                    // goes 0.5340 -> 0.7087 and all four EN benchmark failures
                    // match Word's page count.
                    let float_reflow_enabled = std::env::var_os("OXI_FLOAT_TABLE_REFLOW").is_some()
                        || std::env::var_os("OXI_S1468_DISABLE").is_none();
                    if !is_floating && (std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() || float_reflow_enabled) {
                        previous_table_probe = Some((block_idx, cursor.cursor_y, pages.len(), start_y, content_height, s755_geom));
                        previous_table_probe_elements = elements.clone();
                    }

                    if is_floating && (std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() || float_reflow_enabled) {
                        let nominal_top = page.margin.top;
                        let nominal_anchor = saved_cursor_y - (start_y - nominal_top).max(0.0);
                        let mut probe_table = table.clone();
                        probe_table.style.position = None;
                        let mut probe_cursor = LayoutCursor::new(nominal_anchor);
                        let mut probe_pages = Vec::new();
                        let mut probe_pending = Vec::new();
                        let probe_tail = self.layout_table(
                            &probe_table, start_x, &mut probe_cursor, content_width,
                            grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                            nominal_top, start_y + content_height - nominal_top,
                            page.size.width, page.size.height,
                            &mut probe_pages, &mut probe_pending, Some(block_idx), page,
                            false, None, None, 0.0, 0.0, false, None,
                        );
                        let first_elements: Vec<&LayoutElement> = if let Some(first) = probe_pages.first() {
                            first.elements.iter().collect()
                        } else {
                            probe_pending.iter().chain(probe_tail.iter()).collect()
                        };
                        let first_bottom = first_elements.iter().map(|e| match &e.content {
                            LayoutContent::TableBorder { y1, y2, .. } => y1.max(*y2),
                            _ => e.y + e.height,
                        }).fold(nominal_anchor, f32::max);
                        let mut actual_cursor = LayoutCursor::new(saved_cursor_y);
                        let mut actual_pages = Vec::new();
                        let mut actual_pending = Vec::new();
                        let actual_tail = self.layout_table(
                            &probe_table, start_x, &mut actual_cursor, content_width,
                            grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                            start_y, content_height, page.size.width, page.size.height,
                            &mut actual_pages, &mut actual_pending, Some(block_idx), page,
                            false, None, None, 0.0, 0.0, false, None,
                        );
                        let actual_first: Vec<&LayoutElement> = if let Some(first) = actual_pages.first() {
                            first.elements.iter().collect()
                        } else {
                            actual_pending.iter().chain(actual_tail.iter()).collect()
                        };
                        let actual_bottom = actual_first.iter().map(|e| match &e.content {
                            LayoutContent::TableBorder { y1, y2, .. } => y1.max(*y2),
                            _ => e.y + e.height,
                        }).fold(saved_cursor_y, f32::max);
                        let actual_prefix = actual_bottom - saved_cursor_y;
                        // S1478 (2026-09-19, default ON, opt-out
                        // OXI_S1478_DISABLE): Word breaks a float BETWEEN rows,
                        // never inside one. policies__00602e8a: float 2 is
                        // anchored 19.9pt above the page-5 bottom and float 3
                        // 32.1pt above page 6's, while their first rows are
                        // 85.9 and 68.9pt -- Word starts both on the next page
                        // (PDF rules 114.9.. on p6 and 132.0.. on p7, with
                        // nothing after the previous float on the page before).
                        // Float 1 keeps its 244.2pt prefix against a 68.9pt
                        // first row and stays put, so the test is per-ROW, not
                        // "does the whole float fit".
                        let row0_height = if std::env::var("OXI_S1478_DISABLE").is_err()
                            && !probe_table.rows.is_empty()
                            && !actual_pages.is_empty()
                        {
                            let mut row0_table = probe_table.clone();
                            row0_table.rows.truncate(1);
                            let mut r0_cursor = LayoutCursor::new(nominal_anchor);
                            let mut r0_pages = Vec::new();
                            let mut r0_pending = Vec::new();
                            let r0_tail = self.layout_table(
                                &row0_table, start_x, &mut r0_cursor, content_width,
                                grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                                nominal_top, start_y + content_height - nominal_top,
                                page.size.width, page.size.height,
                                &mut r0_pages, &mut r0_pending, Some(block_idx), page,
                                false, None, None, 0.0, 0.0, false, None,
                            );
                            let r0_elems: Vec<&LayoutElement> = if let Some(f) = r0_pages.first() {
                                f.elements.iter().collect()
                            } else {
                                r0_pending.iter().chain(r0_tail.iter()).collect()
                            };
                            Some(r0_elems.iter().map(|e| match &e.content {
                                LayoutContent::TableBorder { y1, y2, .. } => y1.max(*y2),
                                _ => e.y + e.height,
                            }).fold(nominal_anchor, f32::max) - nominal_anchor)
                        } else {
                            None
                        };
                        if let Some(h0) = row0_height {
                            if h0 > 0.0 && actual_prefix + 0.5 < h0 {
                                move_float_to_next_page = true;
                            }
                        }
                        if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                            eprintln!("[FLOAT-ROW0] block={} actual_prefix={:.3} row0_height={:?} move={}",
                                block_idx, actual_prefix, row0_height, move_float_to_next_page);
                        }
                        // S1476 (2026-09-18, default ON, opt-out OXI_S1476_DISABLE):
                        // when the float BREAKS at its real anchor, the band it
                        // excludes from the page's own flow is what it actually
                        // fits here, not its whole height. policies__00602e8a
                        // page 4: the float's row 0 occupies 121.55..365.79
                        // (actual_prefix 244.24) and Word resumes the previous
                        // table's tail at 365.8 -- measured off the Word PDF's
                        // rules. The old branch reserved the float's FULL height
                        // (311.59), pushing that tail to 433.1 and freeing the
                        // 190..430 band, so rows 1..3 packed onto page 4 too.
                        let s1476 = std::env::var("OXI_S1476_DISABLE").is_err();
                        let placement_prefix = if probe_pages.is_empty()
                            && !(s1476 && !actual_pages.is_empty())
                        {
                            first_bottom - nominal_anchor
                        } else {
                            actual_prefix
                        };
                        if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                            eprintln!("[FLOAT-ACTUAL] block={} first_bottom={:.3} prefix_height={:.3} breaks={} placement_prefix={:.3}",
                                block_idx, actual_bottom, actual_prefix, actual_pages.len(), placement_prefix);
                        }
                        if let (Some((prev_idx, entry_y, entry_page, entry_top, entry_height, geometry)), Some(pos)) = (previous_table_probe, table.style.position.as_ref()) {
                            if pos.v_anchor.as_deref() == Some("page") && num_columns == 1 {
                                if let Some(Block::Table(previous)) = page.blocks.get(prev_idx) {
                                    let mut baseline_pages: Vec<LayoutPage> = (0..entry_page).map(|_| LayoutPage {
                                        width: page.size.width, height: page.size.height, elements: Vec::new(),
                                    }).collect();
                                    let mut baseline_pending = previous_table_probe_elements.clone();
                                    let mut baseline_cursor = LayoutCursor::new(entry_y);
                                    let _ = self.layout_table(
                                        previous, start_x, &mut baseline_cursor, content_width,
                                        grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                                        entry_top, entry_height, page.size.width, page.size.height,
                                        &mut baseline_pages, &mut baseline_pending, Some(prev_idx), page,
                                        false, None, None, 0.0, 0.0, false, geometry.as_ref(),
                                    );
                                    if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                                        eprintln!("[FLOAT-CONTEXT] previous={} entry_page={} entry_y={:.3} entry_top={:.3} entry_height={:.3} target_page={} baseline_end_page={} baseline_cursor={:.3} geometry={:?}",
                                            prev_idx, entry_page + 1, entry_y, entry_top, entry_height,
                                            current_page_idx + 1, baseline_pages.len() + 1, baseline_cursor.cursor_y, geometry);
                                    }
                                    let exclusion_bottom = pos.y + placement_prefix;
                                    let target_page = current_page_idx + 1;
                                    let mut replay_geometry = geometry.unwrap_or(S755Geom {
                                        first_even: first_logical % 2 == 0,
                                        first: (entry_top, entry_height), odd: (entry_top, entry_height),
                                        even: (entry_top, entry_height), page_override: None,
                                    });
                                    let bottom = replay_geometry.top(target_page) + replay_geometry.ch(target_page);
                                    replay_geometry.page_override = Some((target_page, exclusion_bottom, bottom - exclusion_bottom));
                                    let mut replay_pages: Vec<LayoutPage> = (0..entry_page).map(|_| LayoutPage {
                                        width: page.size.width, height: page.size.height, elements: Vec::new(),
                                    }).collect();
                                    let mut replay_pending = previous_table_probe_elements.clone();
                                    let mut replay_cursor = LayoutCursor::new(if entry_page + 1 == target_page { entry_y.max(exclusion_bottom) } else { entry_y });
                                    let mut replay_tail = self.layout_table(
                                        previous, start_x, &mut replay_cursor, content_width,
                                        grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                                        entry_top, entry_height, page.size.width, page.size.height,
                                        &mut replay_pages, &mut replay_pending, Some(prev_idx), page,
                                        false, None, None, 0.0, 0.0, false, Some(&replay_geometry),
                                    );
                                    if replay_pages.len() + 1 == target_page {
                                        let flow_anchor = replay_cursor.cursor_y + saved_cursor_y - baseline_cursor.cursor_y;
                                        let mut positioned_cursor = LayoutCursor::new(pos.y);
                                        let mut positioned_pages: Vec<LayoutPage> = (0..target_page - 1).map(|_| LayoutPage {
                                            width: page.size.width, height: page.size.height, elements: Vec::new(),
                                        }).collect();
                                        let mut positioned_pending = Vec::new();
                                        let positioned_tail = self.layout_table_with_fit(
                                            table, start_x, &mut positioned_cursor, content_width,
                                            grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                                            start_y, content_height, page.size.width, page.size.height,
                                            &mut positioned_pages, &mut positioned_pending, Some(block_idx), page,
                                            false, None, None, 0.0, 0.0, false, s755_geom.as_ref(), Some(flow_anchor - pos.y),
                                        );
                                        let entries = positioned_pages.iter().enumerate().flat_map(|(pi, pg)| {
                                            pg.elements.iter().filter_map(move |e| {
                                                if let LayoutContent::Text { text, .. } = &e.content {
                                                    if !text.is_empty() { return Some((pi + 1, e.cell_row_index, e.cell_paragraph_index, e.y)); }
                                                }
                                                None
                                            })
                                        }).chain(positioned_tail.iter().filter_map(|e| {
                                            if let LayoutContent::Text { text, .. } = &e.content {
                                                if !text.is_empty() { return Some((positioned_pages.len() + 1, e.cell_row_index, e.cell_paragraph_index, e.y)); }
                                            }
                                            None
                                        })).collect::<Vec<_>>();
                                        if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                                            eprintln!("[FLOAT-POSITION] block={} anchor={:.3} end_page={} end_cursor={:.3} text={:?}",
                                                block_idx, flow_anchor, positioned_pages.len() + 1, positioned_cursor.cursor_y, entries);
                                        }
                                    }
                                    if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                                        eprintln!("[FLOAT-REFLOW] block={} previous={} exclusion_bottom={:.3} end_page={} end_cursor={:.3} tail_elements={}",
                                            block_idx, prev_idx, exclusion_bottom, replay_pages.len() + 1, replay_cursor.cursor_y, replay_tail.len());
                                    }
                                    let separator: Vec<LayoutElement> = elements.iter()
                                        .filter(|e| e.paragraph_index == Some(prev_idx + 1)).cloned().collect();
                                    let can_commit_separator = float_reflow_enabled && prev_idx + 2 == block_idx
                                        && page.footnotes.is_empty() && !separator.is_empty()
                                        && separator.iter().all(|e| matches!(&e.content,
                                            LayoutContent::Text { text, .. } if text.is_empty()));
                                    // A preceding table that ends above this band (or on an
                                    // earlier page) does not move. Its following paragraph
                                    // still wraps below the floating fragment and determines
                                    // the capacity available at the float's text anchor.
                                    let previous_ends_before_band = baseline_pages.len() + 1 < target_page
                                        || (baseline_pages.len() + 1 == target_page
                                            && baseline_cursor.cursor_y <= pos.y + 0.025);
                                    if can_commit_separator && previous_ends_before_band {
                                        let separator_top = separator.iter().map(|e| e.y)
                                            .fold(f32::INFINITY, f32::min);
                                        if separator_top < exclusion_bottom && saved_cursor_y > pos.y {
                                            let delta = exclusion_bottom - separator_top;
                                            for e in elements.iter_mut().filter(|e| e.paragraph_index == Some(prev_idx + 1)) {
                                                e.y += delta;
                                            }
                                            saved_cursor_y += delta;
                                            cursor.set(saved_cursor_y);
                                            committed_float_fit = Some(saved_cursor_y - pos.y);
                                        }
                                    } else if can_commit_separator && baseline_pages.len() + 1 == target_page {
                                        let new_anchor = replay_cursor.cursor_y + saved_cursor_y - baseline_cursor.cursor_y;
                                        // The placement capacity is measured from the
                                        // section's nominal top, before a header displaces
                                        // the body. Translate that fitted fragment by the
                                        // page-relative float offset. The actual fragment
                                        // still supplies the exclusion for the preceding
                                        // table; these two measurements serve different
                                        // purposes when a header reduces the body's height.
                                        let nominal_fit_anchor = first_bottom + pos.y - nominal_top;
                                        if replay_pages.len() + 1 == target_page
                                            && nominal_fit_anchor <= bottom && new_anchor <= bottom {
                                            let delta = replay_cursor.cursor_y - baseline_cursor.cursor_y;
                                            replay_tail.extend(separator.into_iter().map(|mut e| { e.y += delta; e }));
                                            pages.truncate(entry_page);
                                            pages.extend(replay_pages.into_iter().skip(entry_page));
                                            elements = replay_pending;
                                            elements.extend(replay_tail);
                                            saved_cursor_y = new_anchor;
                                            cursor.set(new_anchor);
                                            committed_float_fit = Some(new_anchor - pos.y);
                                        } else {
                                            move_float_to_next_page = true;
                                        }
                                    }

                                }
                            }
                        }
                        if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                            eprintln!("[FLOAT-FLOW] block={} page={} actual_anchor={:.3} nominal_anchor={:.3} first_bottom={:.3} prefix_height={:.3} breaks={} final_cursor={:.3}",
                                block_idx, current_page_idx + 1, saved_cursor_y, nominal_anchor,
                                first_bottom, first_bottom - nominal_anchor, probe_pages.len(), probe_cursor.cursor_y);
                        }
                    }

                    if move_float_to_next_page {
                        // S1569 (2026-09-26, default ON, opt-out OXI_S1569_DISABLE): the
                        // keepNext paragraphs directly above a float that starts on the
                        // next page go with it (the S802B chain back-pull, for the
                        // S1478 whole-move). policies__1a7a3fec p27/28: 「例示」
                        // (keepNext) then a text-anchored float whose header row does
                        // not fit -- Word starts p28 with 「例示」 and the table; Oxi left
                        // the heading alone at the p27 bottom.
                        let mut pull_from = block_idx;
                        if std::env::var_os("OXI_S1569_DISABLE").is_none() && num_columns == 1 {
                            while pull_from > 0 {
                                match page.blocks.get(pull_from - 1) {
                                    Some(Block::Paragraph(pp))
                                        if pp.style.keep_next
                                            && block_page_indices.get(pull_from - 1)
                                                == Some(&current_page_idx) =>
                                    {
                                        pull_from -= 1;
                                    }
                                    _ => break,
                                }
                            }
                        }
                        let pulled: Vec<LayoutElement> = if pull_from < block_idx {
                            let y0 = block_y_positions[pull_from] - 0.1;
                            let (keep, moved): (Vec<LayoutElement>, Vec<LayoutElement>) =
                                elements.drain(..).partition(|e| match e.paragraph_index {
                                    Some(index) => index < pull_from || index >= block_idx,
                                    None => e.y < y0,
                                });
                            elements = keep;
                            moved
                        } else {
                            Vec::new()
                        };
                        let chain_end_old = cursor.cursor_y;
                        pages.push(LayoutPage {
                            width: page.size.width, height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        current_page_idx += 1;
                        if let Some(g) = s755_geom.as_ref() {
                            start_y = g.top(current_page_idx + 1);
                            content_height = g.ch(current_page_idx + 1);
                        }
                        saved_cursor_y = start_y;
                        cursor.set(start_y);
                        if !pulled.is_empty() {
                            let y_first = block_y_positions[pull_from];
                            let dy = start_y - y_first;
                            for mut e in pulled {
                                e.y += dy;
                                if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = e.content {
                                    *y1 += dy;
                                    *y2 += dy;
                                }
                                elements.push(e);
                            }
                            for bi in pull_from..block_idx {
                                if let Some(p) = block_page_indices.get_mut(bi) {
                                    *p = current_page_idx;
                                }
                                if let Some(y) = block_y_positions.get_mut(bi) {
                                    *y += dy;
                                }
                            }
                            saved_cursor_y = chain_end_old + dy;
                            cursor.set(saved_cursor_y);
                        }
                    }

                    // Floating table (tblpPr): position relative to anchor
                    let mut candidate_y_top: f32 = 0.0;
                    let mut is_body_floating: bool = false;
                    if let Some(ref pos) = table.style.position {
                        candidate_y_top = match pos.v_anchor.as_deref() {
                            Some("page") => pos.y,
                            Some("margin") => start_y + pos.y,
                            _ => cursor.cursor_y + pos.y, // "text": offset from anchor para bottom
                        };
                        // S991 (2026-07-23, default ON, opt-out OXI_S991_DISABLE):
                        // a floating table with w:tblpYSpec="bottom" and vertAnchor
                        // ∈ {absent, margin} is BOTTOM-ALIGNED to the page content
                        // bottom — the rows that fit in avail = CBOT − anchor_bottom
                        // end exactly at CBOT (row0 = CBOT − Σ fitting-row heights),
                        // and the remainder split to the next page top via the
                        // existing row-split. DERIVED (_pb_tblpybottom_gen.py, Word
                        // COM): a small table (60pt) sits at CBOT−60 anchor-
                        // independently; policies' 7-row 540pt table (anchor@325,
                        // avail 444.9) fits 5 rows bottom-aligned to CBOT (row0=405.0
                        // ≈769.9−365.3), 2 rows split to p2. vertAnchor="text"
                        // (forms/0013892c) is NOT bottom-aligned — it renders near
                        // the anchor, so it keeps the "text" arm above. SCOPE:
                        // tblpYSpec="bottom" occurs in 0 golden + 0 JP docs (census)
                        // → byte-identical everywhere else by construction; only
                        // policies__003496577 fires. Its Table 3-1 renders ~60pt too
                        // high in Oxi (row4 y=567.6 vs Word 631.4), tipping row5
                        // 'Climatic shell' onto p10 where Word has it on p11.
                        if pos.y_spec.as_deref() == Some("bottom")
                            && !matches!(pos.v_anchor.as_deref(), Some("text") | Some("page"))
                            && std::env::var("OXI_S991_DISABLE").is_err()
                        {
                            let anchor_bottom = candidate_y_top;
                            let cb = start_y + content_height;
                            let avail = cb - anchor_bottom;
                            let cw_est =
                                self.resolve_table_col_widths_n(table, content_width, false);
                            let dp = table.style.default_cell_margins.as_ref();
                            let (pl, pr, pt, pb) = (
                                dp.and_then(|m| m.left).unwrap_or(5.4),
                                dp.and_then(|m| m.right).unwrap_or(5.4),
                                dp.and_then(|m| m.top).unwrap_or(0.0),
                                dp.and_then(|m| m.bottom).unwrap_or(0.0),
                            );
                            let mut fit_h: f32 = 0.0;
                            for row in &table.rows {
                                let nat = self.estimate_table_row_natural_h(
                                    row,
                                    &cw_est,
                                    pl,
                                    pr,
                                    pt,
                                    pb,
                                    table,
                                    page.grid_line_pitch,
                                    page.grid_char_pitch,
                                    None,
                                );
                                let h = row
                                    .height
                                    .map(|th| match row.height_rule.as_deref() {
                                        Some("exact") => th,
                                        _ => (th + self.rowbox2_trh_bw(table, row)).max(nat),
                                    })
                                    .unwrap_or(nat);
                                if fit_h + h > avail + 0.5 {
                                    break;
                                }
                                fit_h += h;
                            }
                            if fit_h > 0.5 {
                                candidate_y_top = cb - fit_h;
                            }
                        }
                        cursor.set(candidate_y_top);
                        // R7.60 body-floating eligibility: vertAnchor="page" AND
                        // table positioned below top margin (in body region).
                        is_body_floating = pos.v_anchor.as_deref() == Some("page")
                            && candidate_y_top > start_y + 0.1;
                    }
                    // Modern page-anchored and negative text-relative floats start on
                    // a fresh page when the anchor page cannot hold them. Oversized
                    // tables then split from that fresh page; legacy modes retain
                    // their fragment-capacity rule.
                    let modern_float_pagination = self.compat_mode_explicit && self.compat_mode >= 15;
                    if is_floating
                        && table.style.position.as_ref().map_or(false, |p| {
                            (p.v_anchor.as_deref() == Some("page") && modern_float_pagination)
                                || (p.v_anchor.as_deref() != Some("page")
                                    && (self.keep_floating_tables_together || p.y < -0.5))
                        })
                        && std::env::var("OXI_S878_DISABLE").is_err()
                    {
                        let cb = start_y + content_height;
                        let cw_est = self.resolve_table_col_widths_n(table, content_width, false);
                        let dp = table.style.default_cell_margins.as_ref();
                        let (pl, pr, pt, pb) = (
                            dp.and_then(|m| m.left).unwrap_or(5.4),
                            dp.and_then(|m| m.right).unwrap_or(5.4),
                            dp.and_then(|m| m.top).unwrap_or(0.0),
                            dp.and_then(|m| m.bottom).unwrap_or(0.0),
                        );
                        let mut est: f32 = 0.0;
                        for row in &table.rows {
                            let nat = self.estimate_table_row_natural_h(
                                row,
                                &cw_est,
                                pl,
                                pr,
                                pt,
                                pb,
                                table,
                                page.grid_line_pitch,
                                page.grid_char_pitch,
                                None,
                            );
                            let h = row
                                .height
                                .map(|th| match row.height_rule.as_deref() {
                                    Some("exact") => th,
                                    _ => (th + self.rowbox2_trh_bw(table, row)).max(nat),
                                })
                                .unwrap_or(nat);
                            est += h;
                        }
                        let fit_origin = if modern_float_pagination && table.style.position.as_ref()
                            .is_some_and(|p| p.v_anchor.as_deref() == Some("page")) {
                            // Reserve the table's height at its text anchor. Its
                            // absolute drawing position has a separate fragment
                            // capacity check in layout_table_with_fit below.
                            saved_cursor_y
                        } else { candidate_y_top };
                        let overflows = fit_origin + est > cb + 0.5
                            && saved_cursor_y > start_y + 0.5;
                        let mut move_whole = self.keep_floating_tables_together || modern_float_pagination;
                        if overflows && !move_whole && est <= content_height {
                            // A fragment can fit at its shifted visual origin but fail
                            // to reserve the same height at the unshifted text anchor.
                            // Measure the actual split rather than estimating a line count.
                            let mut trial_cursor = LayoutCursor::new(candidate_y_top);
                            let mut trial_pages: Vec<_> = pages.iter().map(|p| LayoutPage {
                                width: p.width, height: p.height, elements: p.elements.clone(),
                            }).collect();
                            let mut trial_elements = elements.clone();
                            let entry_pages = trial_pages.len();
                            let entry_elements = trial_elements.len();
                            let _ = self.layout_table_with_fit(
                                table, start_x, &mut trial_cursor, content_width,
                                grid_pitch, page.grid_char_pitch, page.grid_char_cw_ratio,
                                start_y, content_height, page.size.width, page.size.height,
                                &mut trial_pages, &mut trial_elements, Some(block_idx), page,
                                false, None, None, 0.0, footnote_reserve_current,
                                !footnote_ids_current_page.is_empty(), s755_geom.as_ref(),
                                committed_float_fit,
                            );
                            if let Some(first_page) = trial_pages.get(entry_pages) {
                                let fragment_bottom = first_page.elements.iter().skip(entry_elements)
                                    .map(|e| e.y + e.height).fold(candidate_y_top, f32::max);
                                move_whole = saved_cursor_y + fragment_bottom - candidate_y_top > cb + 0.5;
                            }
                        }
                        if overflows && move_whole {
                            dbg_page_push(pages.len(), 0);
                            pages.push(LayoutPage {
                                width: page.size.width,
                                height: page.size.height,
                                elements: std::mem::take(&mut elements),
                            });
                            if let Some(g) = s755_geom.as_ref() {
                                start_y = g.top(pages.len() + 1);
                                content_height = g.ch(pages.len() + 1);
                            }
                            current_page_idx += 1;
                            lm2_cells = 0;
                            footnote_reserve_current = 0.0;
                            footnote_ids_current_page.clear();
                            s900_fold(
                                &mut footnote_reserve_current,
                                &mut footnote_ids_current_page,
                                &mut s900_pending_deferred,
                                current_page_idx,
                            );
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            cursor.set(start_y);
                            saved_cursor_y = start_y;
                            committed_float_fit = None;
                            if let Some(ref pos) = table.style.position {
                                candidate_y_top = match pos.v_anchor.as_deref() {
                                    Some("page") => pos.y,
                                    Some("margin") => start_y + pos.y,
                                    _ => cursor.cursor_y + pos.y, // "text"
                                };
                            }
                            cursor.set(candidate_y_top);
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        }
                    }
                    // A splittable page-positioned float uses the remaining anchor-page
                    // capacity even when its visible origin is above the anchor.
                    // S1547 (2026-09-25, default ON, opt-out OXI_S1547_DISABLE): only
                    // when the float spans the flow column. A float that leaves a
                    // side lane (>= 18.5pt net of the wrap distance, the S1195
                    // floor) never competes with the flow for vertical space, so
                    // its split budget is the page from its OWN top.
                    // educational__1b2cea60: right-lane 評価規準 table (tblpX
                    // 10491, 465pt wide beside a 430pt flow table) anchored at
                    // 498.2 with tblpY 45.4 — Word draws all 307pt at 45.4..352.6
                    // on the anchor page; the anchor-capacity budget split it at
                    // 230.8 and pushed a page per lesson (+8).
                    if !self.keep_floating_tables_together
                        && committed_float_fit.is_none()
                        && table.style.position.as_ref().map_or(false, |p|
                            p.v_anchor.as_deref() == Some("page"))
                        && !(std::env::var("OXI_S1547_DISABLE").is_err() && {
                            let tw: f32 = table.grid_columns.iter().sum();
                            let tp = table.style.position.as_ref();
                            let band_x0 = match tp {
                                Some(tp) => match tp.h_align.as_deref() {
                                    Some(ha) => {
                                        let (rl, rw) = match tp.h_anchor.as_deref() {
                                            Some("page") => (0.0, page.size.width),
                                            _ => (start_x, content_width),
                                        };
                                        match ha {
                                            "center" => rl + (rw - tw) * 0.5,
                                            "right" => rl + rw - tw,
                                            _ => rl,
                                        }
                                    }
                                    None => match tp.h_anchor.as_deref() {
                                        Some("page") => tp.x,
                                        _ => start_x + tp.x,
                                    },
                                },
                                None => start_x,
                            };
                            let (dl, dr) = tp.map_or((0.0, 0.0), |tp| (tp.left_from_text, tp.right_from_text));
                            let left_lane = band_x0 - dl - start_x;
                            let right_lane = start_x + content_width - (band_x0 + tw + dr);
                            // The lane must carry REAL text, not just empties: the
                            // S1031 usable-band floor (41.5pt Latin / 100pt CJK),
                            // not the S1195 empty-paragraph floor (18.5). The
                            // floating_table_keep fixture (flag0_page_lines6/12,
                            // Word-measured) leaves a 32.9pt right lane and Word
                            // still splits the float at the anchor page's capacity.
                            let s1547_lane_min = if std::env::var("OXI_S1031_DISABLE").is_err()
                                && !self.doc_body_has_real_cjk
                            {
                                41.5
                            } else {
                                100.0
                            };
                            let lane = left_lane.max(right_lane) >= s1547_lane_min;
                            if std::env::var("OXI_DEBUG_FLOAT_FLOW").is_ok() {
                                eprintln!("[FLOAT-S1547] block={} band_x0={:.2} tw={:.2} left_lane={:.2} right_lane={:.2} lane_min={:.1} lane={}",
                                    block_idx, band_x0, tw, left_lane, right_lane, s1547_lane_min, lane);
                            }
                            lane
                        }) {
                        // Legacy floating tables retain their text-anchor capacity
                        // when drawn lower on the page. Modern pagination limits
                        // ordinary anchor capacity to the physical body bottom.
                        // Keep Some(0) distinct from no independent floating area.
                        let adjustment = saved_cursor_y - candidate_y_top;
                        committed_float_fit = Some(if modern_float_pagination {
                            adjustment.max(0.0)
                        } else { adjustment });
                    }
                    let pages_before = pages.len();
                    // S740 (2026-07-04, default ON, opt-out OXI_S740_DISABLE):
                    // footnote refs INSIDE table cells reserve footnote-area
                    // height per ROW (Word shrinks the body area on the page of
                    // the referencing row; probeqfncell packed 20 rows + 20
                    // notes as if the notes took no space → −1×7). The render
                    // side already collected cell refs (collect_footnote_refs);
                    // the RESERVATION side only scanned Block::Paragraph runs.
                    // Per-row (ids, height) precomputed here; layout_table
                    // shrinks its row-fit page_bottom by the running reserve
                    // and returns per-page ids for the footnote-area renderer.
                    // 0 corpus docs carry footnote refs inside tables (scanned)
                    // → None everywhere → byte-identical by construction.
                    let mut s740_row_fn: Vec<(Vec<u32>, f32, Vec<f32>)> = Vec::new();
                    let mut s740_any = false;
                    // S1578 (2026-09-26, default ON, opt-out OXI_S1578_DISABLE): a
                    // FLOATING table's cell footnotes reserve the note area too.
                    // reports__0079718f p150: ref 24 sits in a tblpPr table
                    // («Investigations reported to the ARC»); Word draws note 24
                    // at the p150 bottom (separator 689.5) and so moves the next
                    // floating table («Year | Performance measures») to p151.
                    // Oxi reserved nothing and started that table at 674.5.
                    if !page.footnotes.is_empty()
                        && std::env::var("OXI_S740_DISABLE").is_err()
                        && (!is_floating || std::env::var_os("OXI_S1578_DISABLE").is_none())
                    {
                        fn cell_fn_ids(blocks: &[Block], out: &mut Vec<u32>) {
                            for b in blocks {
                                match b {
                                    Block::Paragraph(p) => {
                                        for r in &p.runs {
                                            if let Some(id) = r.footnote_ref {
                                                if !out.contains(&id) {
                                                    out.push(id);
                                                }
                                            }
                                        }
                                    }
                                    Block::Table(t) => {
                                        for row in &t.rows {
                                            for cell in &row.cells {
                                                cell_fn_ids(&cell.blocks, out);
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                        for row in &table.rows {
                            let mut ids: Vec<u32> = Vec::new();
                            for cell in &row.cells {
                                cell_fn_ids(&cell.blocks, &mut ids);
                            }
                            let hs: Vec<f32> = ids.iter().map(|id| estimate_footnote_h(*id)).collect();
                            let h: f32 = hs.iter().sum();
                            if !ids.is_empty() {
                                s740_any = true;
                            }
                            s740_row_fn.push((ids, h, hs));
                        }
                    }
                    // S727-derived: in a TYPED docGrid the footnote separator
                    // consumes NO grid slot (probefn render-truth: body 607.6 →
                    // note1 613.1) — and probeqfncell render-truth confirms the
                    // note area starts right below the last body row with no
                    // slot-sized gap. Reserve the separator only for no-grid /
                    // no-type pages (where it occupies a footnote line, S596b).
                    let s740_sep = if s740_any
                        && !(page.grid_line_pitch.is_some() && !page.doc_grid_no_type)
                    {
                        let first = s740_row_fn
                            .iter()
                            .find_map(|(ids, _, _)| ids.first().copied())
                            .unwrap_or(1);
                        footnote_sep_alloc(first)
                    } else {
                        0.0
                    };
                    let mut s740_fn_pages: Vec<Vec<u32>> = Vec::new();
                    let s970_pages_before_tbl = pages.len();
                    if std::env::var("OXI_DBG_TBLSTART").is_ok() {
                        eprintln!("[TBLARM2] blk={} cur={:.2} before layout_table_with_fit", block_idx, cursor.cursor_y);
                    }
                    let mut table_elements = self.layout_table_with_fit(
                        table,
                        start_x,
                        &mut cursor,
                        content_width,
                        grid_pitch,
                        page.grid_char_pitch,
                        page.grid_char_cw_ratio,
                        start_y,
                        content_height,
                        page.size.width,
                        page.size.height,
                        &mut pages,
                        &mut elements,
                        Some(block_idx),
                        page,
                        false,
                        if s740_any { Some(&s740_row_fn) } else { None },
                        if s740_any {
                            Some(&mut s740_fn_pages)
                        } else {
                            None
                        },
                        s740_sep,
                        footnote_reserve_current,
                        !footnote_ids_current_page.is_empty(),
                        s755_geom.as_ref(),
                        committed_float_fit,
                    );
                    // S1593 (2026-09-29, default ON, opt-out OXI_S1593_DISABLE): a
                    // table that splits inside a multi-column section continues in
                    // the NEXT COLUMN, not on the next page. Word's truth for blind-G
                    // policies__1e87d3e6 p14 (2 columns, linesAndChars): the 予防接種
                    // table fills the left column to y 756.75 and its row 14 opens the
                    // right column of the SAME page (x 311, y 84.75); Oxi pushed a page
                    // and left the right column to the notes, W35/O37. layout_table
                    // splits at page height from the page top, so each continuation
                    // segment already has a column's geometry whenever that column
                    // starts at the page top: segment j goes to column slot
                    // current_column + j (page = slot / ncol, column = slot % ncol),
                    // shifted by the column offset. Scope: equal column widths
                    // matching the table's box, no per-row footnotes (their page
                    // attribution is per pushed page), no per-page geometry, and the
                    // band on the current page must begin at the page top when a
                    // segment lands in a later column of that page.
                    let mut s1593_applied = false;
                    if num_columns > 1
                        && pages.len() > s970_pages_before_tbl
                        && std::env::var_os("OXI_S1593_DISABLE").is_none()
                        && !s740_any
                        && col_widths.iter().all(|w| (*w - content_width).abs() < 0.1)
                        && col_x_positions.len() == num_columns
                        && {
                            // Per-page geometry: a segment laid out for page A may
                            // only move to page B when both have the same body box.
                            let c0 = current_column;
                            let n_cont = pages.len() - s970_pages_before_tbl - 1;
                            (0..=n_cont).all(|j| {
                                let laid = s970_pages_before_tbl + 1 + j; // 0-based page it was laid out on
                                let target = s970_pages_before_tbl + (c0 + j + 1) / num_columns;
                                match s755_geom.as_ref() {
                                    None => true,
                                    Some(g) => (g.top(laid + 1) - g.top(target + 1)).abs() < 0.5
                                        && (g.ch(laid + 1) - g.ch(target + 1)).abs() < 0.5,
                                }
                            })
                        }
                    {
                        let c0 = current_column;
                        let n_cont = pages.len() - s970_pages_before_tbl - 1;
                        let n_seg = n_cont + 1; // continuation pages + the tail
                        let lands_on_current_page = c0 + 1 < num_columns;
                        let band_at_top = (col_band_top - start_y).abs() < 0.5;
                        if !lands_on_current_page || band_at_top {
                            let mut segs: Vec<Vec<LayoutElement>> = pages
                                .drain(s970_pages_before_tbl + 1..)
                                .map(|p| p.elements)
                                .collect();
                            segs.push(std::mem::take(&mut table_elements));
                            debug_assert_eq!(segs.len(), n_seg);
                            let first_page = s970_pages_before_tbl;
                            let mut tail_col = c0;
                            let mut tail_poff = 0usize;
                            for (j, mut seg) in segs.into_iter().enumerate() {
                                let slot = c0 + j + 1;
                                let poff = slot / num_columns;
                                let col = slot % num_columns;
                                let dx = col_x_positions[col] - col_x_positions[c0];
                                for e in &mut seg {
                                    e.x += dx;
                                    if let LayoutContent::TableBorder { x1, x2, .. } = &mut e.content {
                                        *x1 += dx;
                                        *x2 += dx;
                                    }
                                }
                                if j + 1 == n_seg {
                                    tail_col = col;
                                    tail_poff = poff;
                                    table_elements = seg;
                                } else {
                                    while pages.len() <= first_page + poff {
                                        dbg_page_push(pages.len(), 0);
                                        pages.push(LayoutPage {
                                            width: page.size.width,
                                            height: page.size.height,
                                            elements: Vec::new(),
                                        });
                                    }
                                    pages[first_page + poff].elements.extend(seg);
                                }
                            }
                            // Pages up to the tail's page stay pushed; the tail's own
                            // page is the page still being built.
                            while pages.len() > first_page + tail_poff {
                                let last = pages.pop().expect("page to reopen");
                                let mut reopened = last.elements;
                                reopened.extend(std::mem::take(&mut elements));
                                elements = reopened;
                            }
                            while pages.len() < first_page + tail_poff {
                                dbg_page_push(pages.len(), 0);
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: Vec::new(),
                                });
                            }
                            current_column = tail_col;
                            start_x = col_x_positions[current_column];
                            if tail_poff > 0 {
                                col_band_top = start_y;
                            }
                            s1593_applied = true;
                            if std::env::var_os("OXI_DBG_COL").is_some() {
                                eprintln!("[COL] S1593 table blk={} from col {} over {} segment(s) -> col {} page+{}",
                                    block_idx, c0, n_seg, tail_col, tail_poff);
                            }
                        }
                    }
                    if !s1593_applied && num_columns > 1 && current_column > 0
                        && pages.len() > s970_pages_before_tbl
                        && std::env::var("OXI_TABLE_CONTINUATION_ORIGIN_DISABLE").is_err()
                        && col_widths.iter().all(|w| (*w - content_width).abs() < 0.1)
                    {
                        let dx = col_x_positions[0] - start_x;
                        let shift = |e: &mut LayoutElement| {
                            e.x += dx;
                            if let LayoutContent::TableBorder { x1, x2, .. } = &mut e.content {
                                *x1 += dx;
                                *x2 += dx;
                            }
                        };
                        for pg in pages.iter_mut().skip(s970_pages_before_tbl + 1) {
                            for e in &mut pg.elements { shift(e); }
                        }
                        for e in &mut table_elements { shift(e); }
                    }
                    // S740: merge the table's per-page note ids into page_fn_refs
                    // (footnote-area render) + roll the LAST page's notes into the
                    // body's running reserve so following paragraphs fit correctly.
                    if pages.len() > s970_pages_before_tbl {
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        s900_fold(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            &mut s900_pending_deferred,
                            current_page_idx + pages.len() - s970_pages_before_tbl,
                        );
                    }
                    if s740_any && !s740_fn_pages.is_empty() {
                        s740_attributed_tables.insert(block_idx);
                        for (off, ids) in s740_fn_pages.iter().enumerate() {
                            let page_i = current_page_idx + off;
                            while page_fn_refs.len() <= page_i {
                                page_fn_refs.push(Vec::new());
                            }
                            for id in ids {
                                if !page_fn_refs[page_i].contains(id) {
                                    page_fn_refs[page_i].push(*id);
                                }
                            }
                        }
                        if let Some(last_ids) = s740_fn_pages.last() {
                            for id in last_ids {
                                if !footnote_ids_current_page.contains(id) {
                                    if footnote_ids_current_page.is_empty() {
                                        footnote_reserve_current += footnote_sep_alloc(*id);
                                    }
                                    footnote_ids_current_page.push(*id);
                                    footnote_reserve_current += estimate_footnote_h(*id);
                                }
                            }
                        }
                    }
                    // S1452 (2026-09-17, default ON, opt-out OXI_S1452_DISABLE): the
                    // table's BOTTOM border takes vertical room in the flow, so the next
                    // block starts that much lower. COM via the Word PDF's rule lines
                    // (tools/metrics/_pb_tblgap_gen.py, tests/fixtures/tblgap, 8 arms):
                    // table / empty paragraph / table measures bottom-rule to top-rule
                    // 16.44 with a 1.5pt border and 15.48 with a 0.5pt one on a 15pt grid
                    // (14.04 / 13.08 without the grid) = the paragraph's own height plus
                    // the border width. Two ADJACENT tables share one rule and get no
                    // addend. forms__008a2f3e's whole page-1 drift was this single 1.44pt.
                    // S1580 (2026-09-27, default ON, opt-out OXI_S1580_DISABLE): a table
                    // whose foot S1191 already advanced inside layout_table (no
                    // tblBorders, no insideH) must not get the bottom rule a second
                    // time here. reports__003862302b p2 table 1 (tcBorders only, last
                    // row bottom sz12): Word's next block starts at the rule + 1.44
                    // (PDF 383.21 -> 384.65 -> empty 11.51 -> 396.16), Oxi at the rule +
                    // 3.0, and every later line on the page sat 1.45 low.
                    let s1580_foot_done = std::env::var_os("OXI_S1580_DISABLE").is_none()
                        && self.s1191_on()
                        && self.s1191_table_needs_foot(table);
                    if !is_floating
                        && !s1580_foot_done
                        && std::env::var_os("OXI_S1452_DISABLE").is_none()
                        && !matches!(page.blocks.get(block_idx + 1), Some(Block::Table(_)))
                    {
                        // Scope: only a CELL-bordered table. A table-level `tblBorders`
                        // table already carries the gap in Oxi (fixture sz12_p1_g1: Oxi's
                        // bottom-to-top gap is 16.50 against Word's 16.44) — its own start
                        // is what sits 1.5pt low, so adding here double-counts and cost
                        // 008a2f3e and 9e4d04b4 a PASS each.
                        let bw = if table.style.border {
                            // S1536 (2026-09-24, OPT-IN OXI_S1536=1 since 2026-09-25):
                            // on a linesAndChars grid Bug A never shifts the table's start
                            // (bug_a_enabled = grid_char_pitch.is_none()), so nothing
                            // stands in for the bottom edge and the flow below the table
                            // resumes a border width too high. Word PDF rules
                            // (tools/metrics/_pb_tblgap_para_gen.py, 12 arms: grid
                            // {linesAndChars 323, lines 300, none} x sz {4, 12} x
                            // {empty, empty+text} between two table-bordered tables):
                            // the table starts AT the cursor in every arm and the
                            // block after it starts bottom-rule + width (lc sz4: empty
                            // 16.70 = 16.15 + 0.55, text 32.90; sz12 17.78 / 33.86).
                            // Oxi lc gave 15.8 / 32.3 / 16.3 / 32.3. policies__1d77cba8
                            // p10-11: +0.4 / +0.7 / +0.7 at three such junctions
                            // tipped 【表２の２】 row 1's first line into p11. lines/none
                            // keep the Bug A start shift as their stand-in (their gaps
                            // already match: 15.50 / 16.58 / 13.82 / 14.78).
                            // Demoted to opt-in 2026-09-25: the s1537 785 gate lost
                            // reference__0ea3ec86480140c2 (43 -> 45 pages: the +0.5
                            // after a table on p3 tips a last line that Word keeps,
                            // then an empty page) and technical__9e4d04b448f84674
                            // (4 -> 6), while policies__1d77cba8 gained. The probe
                            // measured the gap below a table; the page-bottom fit of
                            // the line that follows is a second question the rule
                            // does not answer yet. Re-derive with both before ON.
                            // S1588 (2026-09-27): S1536 is back ON by default for EXPLICIT
                            // compat 15 (the probe's mode; policies__1d77cba8 is compat 15 and
                            // passes with it). The two documents that failed the 2026-09-25
                            // gate are compat 14 (reference__0ea3ec86: Word adds ~0.2 after the
                            // table, not the rule width) and compat 11 (technical__9e4d04b4:
                            // the +0.5 pushed each following paragraph onto the next grid line,
                            // 16pt by p4, where Word keeps it). Opt-out OXI_S1536_DISABLE;
                            // OXI_S1536 still forces it on everywhere.
                            let s1536_on = std::env::var_os("OXI_S1536").is_some()
                                || (std::env::var_os("OXI_S1536_DISABLE").is_none()
                                    && self.compat_mode >= 15
                                    && self.compat_mode_explicit);
                            if page.grid_char_pitch.is_some() && s1536_on {
                                table.style.bottom_border.as_ref().map_or(
                                    table.style.border_width.unwrap_or(0.5),
                                    |d| {
                                        if d.style == "none" || d.style == "nil" {
                                            0.0
                                        } else if d.style == "double" {
                                            d.width * 3.0
                                        } else {
                                            d.width
                                        }
                                    },
                                )
                            } else {
                                0.0
                            }
                        } else {
                            table.rows.last().map_or(0.0, |r| {
                                r.cells
                                    .iter()
                                    .filter_map(|c| c.borders.as_ref())
                                    .map(|b| {
                                        b.bottom.as_ref().map_or(0.0, |d| {
                                            if d.style == "double" { d.width * 3.0 } else { d.width }
                                        })
                                    })
                                    .fold(0.0_f32, f32::max)
                            })
                        };
                        if bw > 0.0 {
                            cursor.advance(bw);
                        }
                    }
                    let candidate_y_bottom = cursor.cursor_y;
                    // The wrapping boundary includes the table's configured clearance.
                    let float_text_bottom = candidate_y_bottom
                        + table.style.position.as_ref().map_or(0.0, |p| p.bottom_from_text);

                    // R7.60: for body-position vertAnchor=page floating tables,
                    // check overlap with previously-placed floating tables on the
                    // current page. If overlap, push table elements to next page.
                    let mut target_page = current_page_idx;
                    if is_body_floating && is_floating {
                        while floating_tables_per_page
                            .get(target_page)
                            .map_or(false, |ranges| {
                                ranges.iter().any(|(t, b)| {
                                    !(candidate_y_bottom <= *t || candidate_y_top >= *b)
                                })
                            })
                        {
                            target_page += 1;
                        }
                    }

                    if target_page > current_page_idx {
                        // Finalize current page; advance to target page.
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        current_page_idx += 1;
                        while current_page_idx < target_page {
                            dbg_page_push(pages.len(), 0);
                            pages.push(LayoutPage {
                                width: page.size.width,
                                height: page.size.height,
                                elements: Vec::new(),
                            });
                            current_page_idx += 1;
                        }
                    }
                    while floating_tables_per_page.len() <= current_page_idx {
                        floating_tables_per_page.push(Vec::new());
                    }
                    let s970_elem_start = elements.len();
                    elements.extend(table_elements);
                    let mut s1615_extra: f32 = 0.0;
                    // S1615 (2026-09-30, default ON, opt-out OXI_S1615_DISABLE): a
                    // floating table anchored to text with a tblpYSpec (top / center /
                    // bottom -- not a numeric tblpY) sits after the whole preceding
                    // paragraph, and that paragraph's LAST line is set again BELOW the
                    // table; the flow resumes after it. `_pb_ftbl_yspec_min_gen.py`
                    // (self-authored, Calibri 11): inline / tblpY=0 note 89.25, table
                    // 108, After 160.5; YSpec bottom/top/center note 160.5 (below the
                    // table), table 107.25, After 177.75 (+1 line); a two-line note stays
                    // at 89.25 with the table after both lines and After +1 line; an
                    // empty note behaves as the one-line one. `_pb_ftbl_ybottom_gen.py`
                    // (blind-G JA forms__02157d72): the ※ note 447.75 below the 緊急連絡先
                    // table, next paragraph +12.75 against the inline arm.
                    if std::env::var_os("OXI_S1615_DISABLE").is_none()
                        && pages.len() == s970_pages_before_tbl
                        && block_idx > 0
                        && matches!(page.blocks.get(block_idx - 1), Some(Block::Paragraph(_)))
                        && table.style.position.as_ref().map_or(false, |p| {
                            p.v_anchor.as_deref() == Some("text") && p.y_spec.is_some()
                        })
                    {
                        let prev = block_idx - 1;
                        let last_top = elements[..s970_elem_start]
                            .iter()
                            .filter(|e| e.paragraph_index == Some(prev)
                                && matches!(e.content, LayoutContent::Text { .. }))
                            .map(|e| e.y)
                            .fold(f32::NEG_INFINITY, f32::max);
                        if last_top.is_finite() {
                            let line_h = elements[..s970_elem_start]
                                .iter()
                                .filter(|e| e.paragraph_index == Some(prev)
                                    && matches!(e.content, LayoutContent::Text { .. })
                                    && (e.y - last_top).abs() < 0.5)
                                .map(|e| e.height)
                                .fold(0.0f32, f32::max);
                            let dy = cursor.cursor_y - last_top;
                            if dy > 0.0 && line_h > 0.0 {
                                for e in elements[..s970_elem_start].iter_mut() {
                                    if e.paragraph_index == Some(prev) && e.y >= last_top - 0.5 {
                                        e.y += dy;
                                    }
                                }
                                if std::env::var_os("OXI_DBG_S1615").is_some() {
                                    eprintln!("[S1615] para {} last line {:.2} -> {:.2} (+{:.2})",
                                        prev, last_top, last_top + dy, line_h);
                                }
                                s1615_extra = line_h;
                            }
                        }
                    }
                    // S970 v2: remember this table's element range when it is a
                    // candidate. `pages.len() == pages_before_tbl` is the actual
                    // one-page predicate — a table that spanned pages pushed at
                    // least one, and Word does not whole-move a splitting table.
                    s970_pending = None;
                    if std::env::var("OXI_S970_DISABLE").is_err()
                        && !self.doc_body_has_real_cjk
                        && num_columns == 1
                        && !is_floating
                        && pages.len() == s970_pages_before_tbl
                        && !elements.is_empty()
                    {
                        // The terminal paragraph is the LAST one in document order
                        // (last row, last cell). Widening this to "any cell of the
                        // last row" is wrong: technical__002c1ffa has a table whose
                        // keepNext sits on a NON-terminal cell.
                        let terminal_kn = table
                            .rows
                            .last()
                            .and_then(|r| r.cells.last())
                            .and_then(|c| {
                                c.blocks.iter().rev().find_map(|b| match b {
                                    Block::Paragraph(pp) => Some(pp.style.keep_next),
                                    _ => None,
                                })
                            })
                            .unwrap_or(false);
                        // v1 scope: plain tables only. A cell footnote would need its
                        // S740 ids and reserve moved between pages, and an anchored
                        // drawing carries page state in another registry.
                        let plain = !table.rows.iter().any(|r| {
                            r.cells.iter().any(|c| {
                                c.blocks.iter().any(|b| match b {
                                    Block::Paragraph(pp) => {
                                        pp.runs.iter().any(|rn| rn.footnote_ref.is_some())
                                    }
                                    _ => true,
                                })
                            })
                        });
                        if terminal_kn
                            && plain
                            && matches!(page.blocks.get(block_idx + 1), Some(Block::Paragraph(_)))
                        {
                            let top = elements[s970_elem_start..]
                                .iter()
                                .map(|e| e.y)
                                .fold(f32::INFINITY, f32::min);
                            if top.is_finite() {
                                s970_pending = Some((
                                    block_idx,
                                    current_page_idx,
                                    s970_elem_start,
                                    elements.len(),
                                    top,
                                ));
                            }
                        }
                    }
                    if is_body_floating {
                        floating_tables_per_page[current_page_idx]
                            .push((candidate_y_top, candidate_y_bottom));
                    }

                    if is_floating {
                        // R7.76 (Session 61): wrap-below mechanism for vertAnchor=text
                        // floating tables.
                        // Two sub-cases per Session 60 [[session60-word-floating-table-wrap-mechanism]]:
                        //   (a) pages_added > 0 (table spilled to new page) — cursor_y =
                        //       saved_cursor_y is wrong because saved was on the OLD page.
                        //       Body must follow to the page where the table ended and
                        //       wrap below it. This was the missing case in R7.75 v3.
                        //   (b) pages_added == 0 + wide table — Session 60's same-page
                        //       wrap-below case (R7.75 v3 implementation).
                        // Spatial gate `(pos_x_zero || h_anchor_page)` retained from v3
                        // — excludes ed025c's tblpX!=0 horz=margin floating tables.
                        let pages_added = pages.len() - pages_before;
                        let table_w_pt: f32 = table.grid_columns.iter().sum();
                        let v_anchor_text = table
                            .style
                            .position
                            .as_ref()
                            .map_or(false, |p| p.v_anchor.as_deref() == Some("text"));
                        let pos_x_zero = table
                            .style
                            .position
                            .as_ref()
                            .map_or(false, |p| p.x.abs() < 0.5);
                        let h_anchor_page = table
                            .style
                            .position
                            .as_ref()
                            .map_or(false, |p| p.h_anchor.as_deref() == Some("page"));
                        // S686 (2026-06-28, opt-out OXI_S686_DISABLE): a vertAnchor="text"
                        // full-width float anchored to the MARGIN fills the text column,
                        // so the following body must wrap BELOW it. tokyoshugyo's パワハラ
                        // box (tblpX=250tw≈12.5pt, horz=margin, width≈content) was excluded
                        // by the (pos_x_zero||h_anchor_page) gate → Oxi overlapped the
                        // following body (【第１２条】 heading + commentary) INSIDE the box
                        // → 4 extra lines packed onto p16 → 【参考】 and 28 paragraphs all
                        // shifted -1 page (the gate's −1 cascade origin). ed025c's
                        // OFF-column floats (horzAnchor absent → None, or "page", tblpX
                        // 32-100pt, overflow the column) keep h_anchor!=margin → unaffected.
                        let h_anchor_margin = std::env::var("OXI_S686_DISABLE").is_err()
                            && table
                                .style
                                .position
                                .as_ref()
                                .map_or(false, |p| p.h_anchor.as_deref() == Some("margin"));
                        let wide_table = table_w_pt > content_width - 30.0;
                        // S856 (2026-07-15, default ON, opt-out OXI_S856_DISABLE): a
                        // CENTERED vertAnchor="text" float that FITS within the content
                        // width but leaves too little room for body text on either side
                        // (usable side < 50pt) forces the body BELOW it, like a
                        // full-width float. `wide_table` misses it (float 410.85 <
                        // content-30) and the S758 side-wrap gate needs a wide side
                        // (≥100pt), so neither fired → Oxi overlapped the following
                        // body (policies__0009e9db Pathways float: 410.85pt on 451.3pt
                        // content, 11.2pt usable side → body wi92 placed INSIDE the
                        // float region → -338pt under-reserve). Frozen-corpus scan: NO
                        // doc has a centered float that fits-but-narrow (the centered
                        // canaries are overflow [1ec1/kyotei/459f05] or tiny [ukhmrc])
                        // → byte-identical by construction.
                        let centered_narrow = std::env::var("OXI_S856_DISABLE").is_err()
                            && v_anchor_text
                            && table.style.position.as_ref().map_or(false, |p| {
                                p.h_align.as_deref() == Some("center")
                                    && table_w_pt <= content_width + 0.5
                                    && (content_width - table_w_pt) / 2.0
                                        - p.left_from_text.max(p.right_from_text)
                                        < 50.0
                            });
                        // S857 (2026-07-15, default ON, opt-out OXI_S857_DISABLE): a
                        // WIDE (overflow) vertAnchor="page" body-floating table also
                        // forces the following body BELOW it — Word flows the body
                        // around a full-width page-anchored float (policies__0009e9db
                        // Cognition SEN table: 471pt overflow at tblpY=111.65 on page
                        // 2; Word's empties wi94-98 SKIP the float region y112-373 and
                        // resume at y370.5, but Oxi packed them through it → wi108 held
                        // on page 2 instead of page 3). Frozen-corpus scan: the ONLY
                        // wide body-page-float doc is 459f05 (PASS canary, 2 overflow
                        // page-floats) — VERIFIED pagination byte-identical (PASS
                        // 1.0000 {0:88}) AND SSIM byte-identical (+0.0000, only empty
                        // paragraphs shift, no pixels) → Word flows 459f05's body below
                        // its floats too. policies__0009e9db: 0.987 FAIL → 1.000 PASS.
                        // A table whose top is above the body can still overlap its text.
                        let page_float_wide = std::env::var("OXI_S857_DISABLE").is_err()
                            && table.style.position.as_ref().map_or(false, |p| {
                                p.v_anchor.as_deref() == Some("page")
                            })
                            // A continuation's bottom belongs to a later page than the
                            // saved anchor. Comparing those y coordinates can restore
                            // an obsolete flow cursor instead of following the table.
                            && (pages_added > 0 || float_text_bottom > saved_cursor_y + 0.1)
                            && wide_table;
                        // S1549 (2026-09-25, default ON, opt-out OXI_S1549_DISABLE):
                        // in a multi-column section a page-anchored float pushes
                        // only the flow of the COLUMNS it horizontally overlaps
                        // (the S1497 rule for drawings, applied to floating
                        // tables). educational__1b2cea60 p11 (cols=2): the 465pt
                        // 評価規準 float sits over column 2 (x 524.55..989.5);
                        // Word starts column 1's 430pt flow table at the page top
                        // beside it (rules 45.4/62.4/70.9..), Oxi put it at 417.6
                        // below the float and paid a page.
                        let page_float_wide = page_float_wide
                            && !(num_columns > 1
                                && std::env::var("OXI_S1549_DISABLE").is_err()
                                && {
                                    let tp = table.style.position.as_ref();
                                    let fx0 = match tp {
                                        Some(tp) => match tp.h_align.as_deref() {
                                            Some(ha) => {
                                                let (rl, rw) = match tp.h_anchor.as_deref() {
                                                    Some("page") => (0.0, page.size.width),
                                                    _ => (start_x, content_width),
                                                };
                                                match ha {
                                                    "center" => rl + (rw - table_w_pt) * 0.5,
                                                    "right" => rl + rw - table_w_pt,
                                                    _ => rl,
                                                }
                                            }
                                            None => match tp.h_anchor.as_deref() {
                                                Some("page") => tp.x,
                                                _ => start_x + tp.x,
                                            },
                                        },
                                        None => start_x,
                                    };
                                    let (dl, dr) = tp.map_or((0.0, 0.0), |tp| (tp.left_from_text, tp.right_from_text));
                                    let (fl, fr) = (fx0 - dl, fx0 + table_w_pt + dr);
                                    let no_overlap = fr <= start_x + 0.5 || fl >= start_x + content_width - 0.5;
                                    if std::env::var("OXI_DBG_FLOAT").is_ok() {
                                        eprintln!("[FLOAT-S1549] blk={} col=[{:.2}..{:.2}] float=[{:.2}..{:.2}] no_overlap={}",
                                            block_idx, start_x, start_x + content_width, fl, fr, no_overlap);
                                    }
                                    no_overlap
                                });
                        // S864: an edge-aligned float leaving <100pt beside it has
                        // no usable text band; Word flows following body below it.
                        // S1031 (2026-07-29, default ON, opt-out OXI_S1031_DISABLE):
                        // the S864-A lane minimum is ~44pt, not 100pt. Word DOES
                        // wrap the following body into a narrow side lane — the
                        // 100pt cutoff sent lane-wrapping docs below the float.
                        //   MEASURED: probe s1031_pbcollapse V5/V6 (lane 47.4pt) —
                        //   Word puts the following VISIBLE text at x0=477.46, i.e.
                        //   INSIDE the lane beside a float ending at ~470, at the
                        //   anchor's y (92.42), NOT below it; reports__00156ad9
                        //   (lane 49.5pt) — its post-float empties flow at the
                        //   anchor (Word 5 pages; forcing them below made 2 phantom
                        //   pages) and its visible footnotes land at 713-766 only
                        //   because 30 empties already carried the cursor past the
                        //   float bottom 692.9; administrative__0001ce58 (lane
                        //   40.25pt) — Word DOES push its visible bullets below
                        //   (y 614.44 > table bottom 614.1). So the cutoff lies in
                        //   (40.25, 47.4]; 44.0 is mid-interval. Oxi has no lane
                        //   narrowing for this float class, so ≥44 falls to the
                        //   "floats don't advance text flow" branch — the same
                        //   approximation Word's lane produces for empty content.
                        let lane_min = if std::env::var("OXI_S1031_DISABLE").is_err()
                            && !self.doc_body_has_real_cjk
                        {
                            41.5
                        } else {
                            100.0
                        };
                        let edge_narrow = s864_part("A")
                            && v_anchor_text
                            && pos_x_zero
                            && table_w_pt <= content_width + 0.5
                            && table.style.position.as_ref().map_or(false, |p| {
                                content_width - table_w_pt - p.left_from_text.max(p.right_from_text)
                                    < lane_min
                            });
                        let needs_wrap_below = (v_anchor_text
                            && wide_table
                            && (pos_x_zero || h_anchor_page || h_anchor_margin))
                            || centered_narrow
                            || page_float_wide
                            || edge_narrow;
                        if std::env::var("OXI_DBG_FLOAT").is_ok() {
                            eprintln!("[FLOAT] blk={} pg={} saved_y={:.1} cand_y=[{:.1}..{:.1}] w={:.1} x={:?} h_align={:?} h_anchor={:?} pages_added={} wide={} wrap_below={}",
                                block_idx, current_page_idx, saved_cursor_y, candidate_y_top, candidate_y_bottom,
                                table_w_pt, table.style.position.as_ref().map(|p| p.x),
                                table.style.position.as_ref().and_then(|p| p.h_align.as_deref()),
                                table.style.position.as_ref().and_then(|p| p.h_anchor.as_deref()),
                                pages_added, wide_table, needs_wrap_below);
                        }

                        // Keeping a floating table intact does not exempt a table
                        // covering the text column from excluding following body text.
                        let mut table_side_bounds = None;
                        let covers_body = table.style.position.as_ref().is_some_and(|pos| {
                            let (left, width) = if pos.h_anchor.as_deref() == Some("page") {
                                (0.0, page.size.width)
                            } else { (start_x, content_width) };
                            let x = match pos.h_align.as_deref() {
                                Some("center") => left + (width - table_w_pt) / 2.0,
                                Some("right") => left + width - table_w_pt,
                                Some(_) => left,
                                None => left + pos.x,
                            };
                            let padding = table.rows.first().and_then(|r| r.cells.first())
                                .and_then(|c| c.margins.as_ref()).and_then(|m| m.left)
                                .or_else(|| table.style.default_cell_margins.as_ref().and_then(|m| m.left))
                                .unwrap_or(5.4);
                            let outer_left = if pos.h_align.is_none() { x - padding } else { x };
                            let left_gap = pos.left_from_text.max(0.5);
                            let right_gap = pos.right_from_text.max(0.5);
                            let left_room = (outer_left - left_gap - start_x).max(0.0);
                            let right_room = (start_x + content_width - outer_left - table_w_pt
                                - right_gap).max(0.0);
                            table_side_bounds = Some((outer_left - left_gap,
                                outer_left + table_w_pt + right_gap));
                            left_room.max(right_room) < 18.75
                        });
                        let needs_wrap_below = needs_wrap_below
                            || (self.keep_floating_tables_together && covers_body);
                        if self.keep_floating_tables_together && !covers_body {
                            if let Some((left, right)) = table_side_bounds {
                                s758_bands.push((current_page_idx, candidate_y_top, float_text_bottom,
                                    left, right, false, BodyWrapPolicy::FLOATING_TABLE));
                            }
                            cursor.set(saved_cursor_y);
                        } else if needs_wrap_below && pages_added > 0 {
                            current_page_idx += pages_added;
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            cursor.set(float_text_bottom);
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        } else if needs_wrap_below
                            && std::env::var("OXI_S638_DISABLE").is_err()
                            && (candidate_y_top - saved_cursor_y) > 6.0
                            && ((candidate_y_bottom - candidate_y_top) > 250.0
                                // ★S638 height>250 gate DROPPED by default
                                // (2026-07-07, ROWBOX2 bundle; opt-out
                                // OXI_S638_HGATE=1 restores the strict gate).
                                // The gate excluded 2ea81a to protect its SSIM
                                // under the OLD compensating geometry; under
                                // ROWBOX2 the gap-flow is what Word does there
                                // (level 12: Word keeps the anchor-side para in
                                // the gap and resumes pi27 at the float bottom
                                // 772.4; the S469 path put the anchor para below
                                // the float = +16). 2ea81a: PASS 1.0 all-zero +
                                // SSIM +0.0149 under gap-flow. The gap>6 gate
                                // above stays (excludes kyotei/3a4f/ed025c
                                // anchor-adjacent floats).
                                || std::env::var("OXI_S638_HGATE").is_err())
                        {
                            // S638 (kyotei): the float leaves a GAP above it
                            // (>250pt height = a FULL-PAGE form float, the kyotei
                            // case; excludes 2ea81a's shorter multi-float forms whose
                            // body should keep the original wrap-below — gap-flow
                            // regressed 2ea81a SSIM −0.0139).
                            // [saved_cursor_y, candidate_y_top]; the immediately-
                            // following short body (the form header label) flows
                            // INTO that gap, then later body SKIPS the float region.
                            // Record the region; the block-loop snap bumps any block
                            // landing inside it to candidate_y_bottom. Gated on a
                            // real gap (>6pt) so it only fires for anchor+tblpY floats
                            // (kyotei tblpY=13.4) not zero-offset ones (3a4f/ed025c).
                            cursor.set(saved_cursor_y);
                            let (s1241_fx0, s1241_fx1) = {
                                let px = table.style.position.as_ref().map_or(0.0, |p| p.x);
                                let base = match table
                                    .style
                                    .position
                                    .as_ref()
                                    .and_then(|p| p.h_anchor.as_deref())
                                {
                                    Some("page") => 0.0,
                                    Some("margin") => page.margin.left,
                                    _ => start_x,
                                };
                                (base + px, base + px + table_w_pt)
                            };
                            // S1509 (2026-09-20, default ON, opt-out OXI_S1509_DISABLE):
                            // the lane beside this float, placed the way S1195 places
                            // its band (an ALIGN-positioned float resolves against its
                            // anchor rect). legal__07b25a's cover: a centred 378pt
                            // float (lanes 44.85pt net) anchored 21.8pt below eight
                            // empty paragraphs -- Word flows the empties in the lane
                            // (Info6 417..599 beside the table at 441..614); the
                            // region bump sent them below it and onto a second page.
                            let s1509_lane = {
                                let tpos = table.style.position.as_ref();
                                let x0 = match tpos {
                                    Some(tp) => match tp.h_align.as_deref() {
                                        Some(ha) => {
                                            let (rl, rw) = match tp.h_anchor.as_deref() {
                                                Some("page") => (0.0, page.size.width),
                                                _ => (start_x, content_width),
                                            };
                                            match ha {
                                                "center" => rl + (rw - table_w_pt) * 0.5,
                                                "right" => rl + rw - table_w_pt,
                                                _ => rl,
                                            }
                                        }
                                        None => s1241_fx0,
                                    },
                                    None => s1241_fx0,
                                };
                                let (dl, dr) = tpos.map_or((0.0, 0.0), |tp| (tp.left_from_text, tp.right_from_text));
                                (x0 - dl - start_x).max(start_x + content_width - (x0 + table_w_pt + dr))
                            };
                            text_float_region = Some((
                                candidate_y_top,
                                float_text_bottom,
                                current_page_idx,
                                s1241_fx0,
                                s1241_fx1,
                                s1509_lane,
                            ));
                        } else if needs_wrap_below {
                            // S469: the cursor advances below the table so body
                            // TEXT wraps under it, but objects anchored to the
                            // following paragraph keep the natural (pre-wrap) Y.
                            // Record the advance so block_y_positions can undo it.
                            if s469_enabled {
                                anchor_flow_offset += float_text_bottom - saved_cursor_y;
                            }
                            // S1195 (2026-08-22, default ON, opt-out
                            // `OXI_S1195_DISABLE`): keep the LANE open when
                            // one exists and the next block is an empty paragraph.
                            // Word flows those empties beside the float; only real
                            // content drops below it. The lane floor is measured
                            // (`_pb_floatlane2_gen.py`, Arial 11, 1pt steps): the
                            // free side NET of the tblpPr wrap distance flips the
                            // paragraph from below-the-float to in-the-lane between
                            // 17.9 and 18.9pt for leftFromText 142/284tw, and
                            // between 19 and 20 for 0 — so 18.5 sits in every
                            // bracket. ed025cbecffb's page-6 float leaves 27.97pt.
                            let s1195_lane = std::env::var("OXI_S1195_DISABLE").is_err() && {
                                let tpos = table.style.position.as_ref();
                                // Same placement the S758 band uses below: an
                                // ALIGN-positioned float (tblpXSpec) resolves
                                // against its anchor rect, an OFFSET one is tblpX.
                                let band_x0 = match tpos {
                                    Some(tp) => match tp.h_align.as_deref() {
                                        Some(ha) => {
                                            let (rl, rw) = match tp.h_anchor.as_deref() {
                                                Some("page") => (0.0, page.size.width),
                                                _ => (start_x, content_width),
                                            };
                                            match ha {
                                                "center" => rl + (rw - table_w_pt) * 0.5,
                                                "right" => rl + rw - table_w_pt,
                                                _ => rl,
                                            }
                                        }
                                        None => match tp.h_anchor.as_deref() {
                                            Some("page") => tp.x,
                                            _ => start_x + tp.x,
                                        },
                                    },
                                    None => start_x,
                                };
                                let (dl, dr) = tpos.map_or((0.0, 0.0), |tp| {
                                    (tp.left_from_text, tp.right_from_text)
                                });
                                let left_lane = band_x0 - dl - start_x;
                                let right_lane =
                                    start_x + content_width - (band_x0 + table_w_pt + dr);
                                left_lane.max(right_lane) >= 18.5
                            } && matches!(page.blocks.get(block_idx + 1),
                                Some(Block::Paragraph(p))
                                    if fits_blank_float_lane(p));
                            if s1195_lane {
                                cursor.set(saved_cursor_y);
                                float_lane_below =
                                    Some((float_text_bottom, current_page_idx));
                            } else {
                                // S1489 v2 (2026-09-19, default ON, opt-out
                                // OXI_S1489_DISABLE): the EMPTY paragraph laid out
                                // just before this float, whose line box the float's
                                // top cuts, is re-issued below the float. golden
                                // parttime p2: the exact-11 empty at 273.5..284.5
                                // against the band top 281.2 -- Word starts
                                // 「２ 年次…」 at table bottom 387.05 + 11.
                                let mut s1489_extra = 0.0f32;
                                if std::env::var_os("OXI_S1489_DISABLE").is_none()
                                    && block_idx > 0
                                    && matches!(page.blocks.get(block_idx - 1),
                                        Some(Block::Paragraph(p)) if p.runs.iter().all(|r| r.text.is_empty())
                                            || (self.keep_floating_tables_together && covers_body))
                                {
                                    for e in elements.iter_mut().filter(|e| e.paragraph_index == Some(block_idx - 1)) {
                                        if e.y < candidate_y_top - 0.1 && e.y + e.height > candidate_y_top + 0.1 {
                                            s1489_extra = s1489_extra.max(e.height);
                                            e.y = float_text_bottom;
                                        }
                                    }
                                }
                                cursor.set(float_text_bottom + s1489_extra);
                            }
                        } else {
                            // Original behavior: floating tables don't advance text flow
                            // S772 (2026-07-10, opt-out OXI_S772_DISABLE): when a
                            // NARROW vertAnchor="text" float anchored near the page
                            // bottom does not fit, layout_table pushes page(s)
                            // internally (widow/no-fit whole-push) and the float
                            // lands on the NEW page — but this branch restored the
                            // cursor to saved_cursor_y on the OLD page without
                            // advancing current_page_idx, stranding the following
                            // body (uk_hmrc_checklist: the Types-of-Student-Loan box
                            // + the "9" number box each spawned a page, question 9's
                            // text then overflowed onto yet another page → 4 pages
                            // vs Word's 2). Word moves the float AND its anchor to
                            // the next page together (hmrc p2: box top-right,
                            // question 9 flows beside it). Corpus blast radius:
                            // ZERO docs hit v_anchor_text && !needs_wrap_below &&
                            // pages_added>0 (OXI_DBG_FLOAT sweep over all 12 tblpPr
                            // docs — every pages_added>0 case is wrap_below=true;
                            // 459f05's is vertAnchor=page) → byte-identical by
                            // construction.
                            let s772_pushed = std::env::var("OXI_S772_DISABLE").is_err()
                                && v_anchor_text
                                && pages_added > 0;
                            // S758b: a NARROW vertAnchor="text" float leaves room
                            // beside it — Word wraps the following text NEXT TO the
                            // table (probexfloattbl: body lines narrowed to x1≈347
                            // beside the tblpXSpec=right box [354.7..524.7]). Push a
                            // side-wrap band over the table's rect; the S758
                            // paragraph machinery does the narrowing/rebreak. Wide
                            // floats keep the wrap-below paths above.
                            // ALIGN-positioned floats (tblpXSpec — the probe uses
                            // "right") ship since S758b. OFFSET-positioned (tblpX,
                            // h_align None) floats join per the probeqtbloffset
                            // Word truth (an in-column offset float IS wrapped
                            // beside) — see the offset arm below.
                            if std::env::var("OXI_S758_DISABLE").is_err()
                                && v_anchor_text
                                && !wide_table
                                && table_w_pt > 6.0
                            {
                                let band_x0 = if let Some(ref tpos) = table.style.position {
                                    if let Some(ref ha) = tpos.h_align {
                                        let (rl, rw) = match tpos.h_anchor.as_deref() {
                                            Some("page") => (0.0, page.size.width),
                                            _ => (start_x, content_width),
                                        };
                                        match ha.as_str() {
                                            "center" => rl + (rw - table_w_pt) * 0.5,
                                            "right" => rl + rw - table_w_pt,
                                            _ => rl,
                                        }
                                    } else {
                                        match tpos.h_anchor.as_deref() {
                                            Some("page") => tpos.x,
                                            _ => start_x + tpos.x,
                                        }
                                    }
                                } else {
                                    start_x
                                };
                                // Column containment: Word wraps text beside a float
                                // that sits INSIDE the text column (probexfloattbl
                                // x[354.7..524.7] = flush to content-right). ed025c's
                                // OFF-column floats (large tblpX, overflowing the
                                // column) get no wrap in Word — banding them
                                // regressed its word_png −0.0019.
                                let is_align = table
                                    .style
                                    .position
                                    .as_ref()
                                    .map_or(false, |tp| tp.h_align.is_some());
                                // OFFSET-arm gates (S772b, 2026-07-10): the pinned
                                // probeqtbloffset truth is BOTH-SIDES two-segment
                                // wrap; the band machinery is single-segment
                                // (wider side). Band an offset float only when the
                                // NARROWER free side is negligible (≤24pt — hmrc's
                                // Types box right gap ≈0 / number-box left gap ≈0),
                                // so the single-side model is exact. Mid-column
                                // offset floats (both sides wide) stay unbanded
                                // until the two-segment flow lands. Dist for the
                                // offset arm = the tblpPr left/rightFromText
                                // (ECMA default 0 — hmrc has no attrs and Word
                                // wraps flush to the box edge; the align arm keeps
                                // its shipped 9.0 calibration).
                                let (dl, dr) = if is_align {
                                    (9.0, 9.0)
                                } else {
                                    table.style.position.as_ref().map_or((0.0, 0.0), |tp| {
                                        (tp.left_from_text, tp.right_from_text)
                                    })
                                };
                                let left_gap = (band_x0 - dl - start_x).max(0.0);
                                let right_gap = (start_x + content_width
                                    - (band_x0 + table_w_pt + dr))
                                    .max(0.0);
                                // Offset-arm scope: (a) the NARROW side ≤24pt so the
                                // single-side model is exact (probeqtbloffset's true
                                // both-sides two-segment case stays unbanded), AND
                                // (b) the WIDE side ≥100pt — a real host column.
                                // ed025c #0 leaves only a 33.75pt left sliver
                                // (indented paras don't fit it; Word flows them
                                // BELOW the float, a mechanism the band lacks) —
                                // banding it exploded its notes into 30pt-floor
                                // strips. hmrc: Types box 240.65 / number boxes
                                // ~490 → banded.
                                // S-TWOSEG: the case this gate was holding open a
                                // place for. When BOTH free sides are wide enough
                                // to hold text, the float is banded too and the
                                // paragraph flows through the pair of strips.
                                // ed025c and hmrc are untouched by it: their floats
                                // run to (or past) one column edge, so the narrow
                                // side is ~0 and the pair never forms.
                                let two_seg_ok = std::env::var("OXI_TWOSEG_DISABLE").is_err()
                                    && left_gap.min(right_gap) >= 30.0;
                                let offset_ok = std::env::var("OXI_S772_DISABLE").is_err()
                                    && ((left_gap.min(right_gap) <= 24.0
                                        && left_gap.max(right_gap) >= 100.0)
                                        || two_seg_ok);
                                if std::env::var("OXI_DBG_TWOSEG").is_ok() {
                                    eprintln!(
                                        "[TWOSEG] band_x0={:.1} w={:.1} left_gap={:.1} right_gap={:.1} two_seg_ok={} is_align={}",
                                        band_x0, table_w_pt, left_gap, right_gap, two_seg_ok, is_align
                                    );
                                }
                                if band_x0 >= start_x - 1.0
                                    && band_x0 + table_w_pt <= start_x + content_width + 6.0
                                    && (is_align || offset_ok)
                                {
                                    // S772: when the float was pushed to a new page,
                                    // the band lives there, from the page top.
                                    let (band_pg, band_top) = if s772_pushed {
                                        (current_page_idx + pages_added, start_y)
                                    } else {
                                        (current_page_idx, candidate_y_top)
                                    };
                                    s758_bands.push((
                                        band_pg,
                                        band_top,
                                        candidate_y_bottom,
                                        band_x0 - dl,
                                        band_x0 + table_w_pt + dr,
                                        false, BodyWrapPolicy::OBJECT,
                                    ));
                                }
                            }
                            if s772_pushed {
                                current_page_idx += pages_added;
                                *block_page_indices.last_mut().unwrap() = current_page_idx;
                                if let Some(g) = s755_geom.as_ref() {
                                    start_y = g.top(pages.len() + 1);
                                    content_height = g.ch(pages.len() + 1);
                                }
                                if num_columns > 1 {
                                    current_column = 0;
                                    start_x = col_x_positions[0];
                                    content_width = col_widths[0];
                                }
                                lm2_cells = 0;
                                footnote_reserve_current = 0.0;
                                footnote_ids_current_page.clear();
                                s900_fold(
                                    &mut footnote_reserve_current,
                                    &mut footnote_ids_current_page,
                                    &mut s900_pending_deferred,
                                    current_page_idx,
                                );
                                cursor.set(start_y);
                                *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                            } else {
                                cursor.set(saved_cursor_y);
                            }
                        }
                    }
                    // S1615: the preceding paragraph's last line now sits below the
                    // table, so the flow resumes one line further down.
                    if s1615_extra > 0.0 {
                        cursor.advance(s1615_extra);
                    } else {
                        let pages_added = pages.len() - pages_before;
                        if pages_added > 0 {
                            current_page_idx += pages_added;
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                            if num_columns > 1 {
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                            }
                        }
                    }
                    prev_para_style_id = None;
                    prev_borders = None; // S658: a table breaks border-merge adjacency
                    prev_autospacing_numid = None; // S931: and list adjacency
                    prev_keep_next = false; // S739
                    prev_space_after = 0.0;
                }
                Block::Image(img) => {
                    // S1056: the image-only host paragraph led with
                    // `<w:br w:type="page"/>` — the image starts a new page. Mirrors
                    // the paragraph `pageBreakBefore` handler's page transition.
                    if img.page_break_before
                        && !elements.is_empty()
                        && std::env::var("OXI_S1056_DISABLE").is_err()
                    {
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        if let Some(g) = s755_geom.as_ref() {
                            start_y = g.top(pages.len() + 1);
                            content_height = g.ch(pages.len() + 1);
                        }
                        cursor.set(start_y);
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        lm2_cells = 0;
                        current_page_idx += 1;
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        s900_fold(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            &mut s900_pending_deferred,
                            current_page_idx,
                        );
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        // S1293b: see the overflow path below -- the START index
                        // drives float anchors and has to move with the block.
                        if std::env::var("OXI_S1293_DISABLE").is_err() {
                            *block_start_page_indices.last_mut().unwrap() = current_page_idx;
                        }
                    }
                    // S549 (2026-06-12, opt-out OXI_S549_DISABLE): in a docGrid
                    // lines section the image-only paragraph's line occupies a
                    // WHOLE number of grid cells — ceil(extent/pitch)×pitch.
                    // COM repro (_s549_img_grid.py, pitch 18): extent 185→198
                    // (11 cells), 100→108, 90→90, 36→36 (exact multiples pass
                    // through); docGrid none → extent EXACTLY (S537 model
                    // unchanged). Live 3a4f "/" figures (extent 185 → Word
                    // block 198) were leaving every downstream para 13pt high
                    // → the last 3 Phase-1 delta=-1 boundary paras.
                    // NOTE (2026-07-21): S549's whole-cell rounding was derived on a
                    // TYPED docGrid ("docGrid none → extent EXACTLY"), and a NO-TYPE
                    // grid reaches it only because S571-refine makes grid_line_pitch
                    // Some for a custom (≠360) pitch. Disabling the rounding for
                    // no-type grids was TESTED and is WRONG: it fixes the
                    // policies__00148f8d p45 figure but breaks its p81/p109/p110
                    // figures (0.9957 → 0.9941), so Word does round most of them.
                    // The p45 residual (~5.4pt) needs a real inline-image line-box
                    // probe, not a blanket scope change.
                    let img_line =
                        self.s971_image_line_h(img, content_width, page.grid_line_pitch, true);
                    // S1101 (2026-08-08, default ON, opt-out OXI_S1101_DISABLE):
                    // a NO-TYPE docGrid does NOT round the image-only paragraph
                    // to whole cells — the extent is used EXACTLY, which is what
                    // S549's own derivation says for "docGrid none". A no-type
                    // grid only reaches the rounding path because S571-refine
                    // makes grid_line_pitch Some for a custom (non-360) pitch.
                    // ★The NOTE above ("TESTED and is WRONG … Word does round
                    // most of them") was decided on the pagination score alone;
                    // measuring Word GEOMETRY over all 20 figures of
                    // policies__00148f8d (no-type linePitch 326 = 16.3pt) shows
                    // the opposite — in 13 of them Oxi's gap to the next line is
                    // exactly Word's gap PLUS the snap padding
                    // (pad = ceil(h/16.3)*16.3 − h), i.e. only Oxi rounds:
                    //   p45  h=182.20 pad=13.40 | Word gap 0.05  Oxi gap 13.40
                    //   p55  h=150.15 pad=12.85 | Word gap 4.02  Oxi gap 16.85
                    //   p49  h=104.75 pad= 9.35 | Word gap 6.27  Oxi gap 15.35
                    // With the skip, 9 of 10 sampled figures land within 0.4pt of
                    // Word's own y/bottom/gap. legal__001410a8's vector figure
                    // agrees too (Word gap ~4.28, skip 3.00, round 16.25).
                    let s1101_no_type_exact = page.doc_grid_no_type
                        && (!self.doc_body_has_real_cjk
                            || std::env::var_os("OXI_CJK_INLINE_NATURAL").is_some())
                        && std::env::var("OXI_S1101_DISABLE").is_err();
                    let img_adv = match page.grid_line_pitch {
                        Some(p)
                            if p > 0.1
                                && !s1101_no_type_exact
                                && std::env::var("OXI_S549_DISABLE").is_err() =>
                        {
                            (img_line / p).ceil() * p
                        }
                        _ => img_line,
                    };
                    // S965 (2026-07-21, opt-out OXI_S965_DISABLE): an image-only
                    // paragraph is a real paragraph, so its spacing collapses with
                    // its neighbours' exactly like any other — max(prev.after,
                    // own.before) above, and its own.after is what the NEXT
                    // paragraph collapses against. Before this the arm applied
                    // nothing above and let the PREVIOUS paragraph's after leak
                    // across the image onto the next one (a third behaviour that is
                    // neither). Word truth on the three specimens: the image
                    // paragraph resolves to after=0 in policies__00148f8d
                    // (`Graphics`) and after=10pt in legal__00089377 /
                    // reports__000e8acd (style-less → docDefaults after=200), and
                    // legal p2 measures 14.95pt MORE space around the figure in
                    // Word than Oxi placed. Suppressed at a fresh region top, where
                    // a body paragraph's space_before drops too.
                    let s965 = std::env::var("OXI_S965_DISABLE").is_err();
                    // S1183 (2026-08-21, opt-out OXI_S1183_DISABLE): HTML
                    // autospacing on an IMAGE host resolves to the derived flat
                    // amounts — 6.75 for a LIST host, 14.0 (S901) for a plain
                    // one — and auto OVERRIDES the explicit w:before (S675).
                    // DERIVED (_pb_imgnum 20-arm matrix): numaut_img36/60 =
                    // img + 6.75 at any size/multiplier; aut_img36 = img + 14.6
                    // (14.0 within the COM quantum); Oxi applied the explicit
                    // 5pt + nothing. Same-list consecutive image hosts keep the
                    // full 6.75 (measured) — the S931 text suppression does NOT
                    // extend to images, and this arm already resets
                    // prev_autospacing_numid below. The after side mirrors the
                    // S675/S901 text rule. Latin scope (matrix measured on
                    // Latin; JP byte-identical by construction).
                    let s1183 = s965
                        && !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S1183_DISABLE").is_err();
                    let s1183_host = img.host_paragraph.as_deref().map(|h| &h.style);
                    let s1183_sb = match (s1183, s1183_host) {
                        (true, Some(st)) if st.before_autospacing => {
                            Some(if st.num_id.is_some() { 6.75 } else { 14.0 })
                        }
                        _ => None,
                    };
                    let s1183_sa = match (s1183, s1183_host) {
                        (true, Some(st)) if st.after_autospacing => Some(14.0),
                        _ => None,
                    };
                    let mut img_before = if s965 {
                        prev_space_after.max(s1183_sb.unwrap_or(img.paragraph_space_before))
                    } else {
                        0.0
                    };
                    let image_section_spacing = block_idx.checked_sub(1)
                        .and_then(|i| page.blocks.get(i))
                        .and_then(|block| match block {
                            Block::Paragraph(marker)
                                if marker.style.continuous_section_break
                                    && marker.style.page_section_break
                                    && marker.runs.iter().all(|r| r.text.is_empty()) =>
                            {
                                let suppressed = s1183_host.map_or(false, |host|
                                    host.contextual_spacing && host.style_id == marker.style.style_id);
                                Some(if suppressed { 0.0 } else {
                                    (img_before - marker.style.space_after.unwrap_or(0.0)).max(0.0)
                                })
                            }
                            _ => None,
                        });
                    if let Some(spacing) = image_section_spacing {
                        img_before = spacing;
                    }
                    let image_section_top_spacing = if s1183_host.map_or(false, |host| host.before_autospacing) {
                        0.0
                    } else {
                        image_section_spacing.unwrap_or(0.0)
                    };
                    if cursor.cursor_y <= start_y + 0.1 {
                        img_before = image_section_top_spacing;
                    }
                    // S1418b (2026-09-23, opt-out OXI_IMAGE_INK_FIT_DISABLE):
                    // an inline image line is PLACED on the grid-rounded advance
                    // (S1418) but is judged against the page bottom by its own
                    // INK, not by that rounded advance. MEASURED on a faithful
                    // one-image slice of policies__1db396de (453.5x218.0pt on an
                    // 18pt grid, body bottom 771.0): sweeping a snapToGrid=0
                    // exact spacer 455..495pt in 2pt steps, Word keeps the image
                    // on page 1 up to spacer 473 -- where it draws at
                    // 551.34..769.90, i.e. ink bottom 769.90 <= 771.0 -- and
                    // moves it at 475. The rounded advance would put 473 at
                    // 543.9 + 234.0 = 777.9 > 771.0 and move it a page early,
                    // which is exactly what Oxi did: Word fits three of the five
                    // figures on its page 29, Oxi only two, and the whole tail
                    // of the document shifted by one page.
                    // The grid leading around the image splits evenly: on the
                    // 18pt grid above, advance 234.0 holds a 218.56 image drawn
                    // at 7.44 below the line top and ending 226.0 below it, and
                    // (234.0 + 218.56) / 2 = 226.28 reproduces that. Judging by
                    // the bare ink (218) instead is 8pt too generous and let
                    // tokyoshugyo pull three paragraphs onto page 74.
                    let img_fit = if std::env::var_os("OXI_IMAGE_INK_FIT_DISABLE").is_none() {
                        img_adv - (img_adv - img_line).max(0.0) * 0.5
                    } else {
                        img_adv
                    };
                    if cursor.cursor_y + img_before + img_fit > start_y + content_height {
                        img_before = image_section_top_spacing;
                        if num_columns > 1 && current_column + 1 < num_columns {
                            current_column += 1;
                            start_x = col_x_positions[current_column];
                            content_width = col_widths[current_column];
                            cursor.set(col_band_top);
                        } else {
                            dbg_page_push(pages.len(), 0);
                            pages.push(LayoutPage {
                                width: page.size.width,
                                height: page.size.height,
                                elements: std::mem::take(&mut elements),
                            });
                            if let Some(g) = s755_geom.as_ref() {
                                start_y = g.top(pages.len() + 1);
                                content_height = g.ch(pages.len() + 1);
                            }
                            cursor.set(start_y);
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            lm2_cells = 0;
                            current_page_idx += 1;
                        }
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        // S1293b: the block's START moved too. S1123 resolves a
                        // float's anchor against `block_start_page_indices`, and
                        // leaving it on the pre-push page splits a text box from
                        // its own frame -- the frame goes to the new page, its
                        // text stays on the old one (legal__02f84965 boxes 3 and
                        // 4: an image-ONLY page 8 and 10, which the gate counts
                        // as blank sheets). An image that did not fit genuinely
                        // starts on the page it was moved to.
                        if std::env::var("OXI_S1293_DISABLE").is_err() {
                            *block_start_page_indices.last_mut().unwrap() = current_page_idx;
                        }
                    }
                    if img_before > 0.0 {
                        cursor.advance(img_before);
                        *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                    }
                    // S1181 v2: an image-only paragraph's top PAINTS on the
                    // 96dpi pixel like any other paragraph under a no-type
                    // docGrid (walk group C/F: term exact, boundary rounded).
                    // Visual track only — the fit already happened on the
                    // exact cursor above. Re-sync after the advance.
                    let s1181_img_unsnap = if page.doc_grid_no_type
                        && !self.doc_body_has_real_cjk
                        && std::env::var("OXI_S1181").is_ok()
                    {
                        let exact = cursor.visual_y;
                        let snapped = (exact / 0.75).round() * 0.75;
                        cursor.advance_split(0.0, snapped - exact);
                        exact - snapped
                    } else {
                        0.0
                    };
                    // S1321 (2026-09-05, default ON, opt-out OXI_S1321_DISABLE): an
                    // inline object in an EXACT-spaced line is painted with its
                    // BOTTOM on the host line's baseline, overflowing upward.
                    // MEASURED (`_pb_exactimg_pos_gen.py`, 7 arms, Word's PDF): the
                    // picture rect's bottom equals the host baseline at every
                    // distance from the page bottom (182.06 / 594.5 / 650.75 /
                    // 688.25 / 725.75 -- no clamping to the page), and legal's two
                    // boxes end at 755.6 / 739.0 = their host lines' baselines.
                    // The baseline of an exact line sits at
                    // top + ascent + (exact - natural) / 2 (9pt ＭＳ Ｐ明朝 in a
                    // 13pt line: 746.2 + 9.4 = 755.6; 12pt ＭＳ 明朝 in 16pt:
                    // +11.9). The text box that shares the block follows the
                    // shifted top through `block_y_positions`.
                    let s1321_top = match img.host_exact_line {
                        Some(exact) if exact > 0.0 && std::env::var("OXI_S1321_DISABLE").is_err() => {
                            let (fs, m) = match img.host_paragraph.as_deref() {
                                Some(host) => {
                                    let rpr_ref = host.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                    let fs = rpr_ref.font_size.unwrap_or(self.default_font_size);
                                    (fs, self.metrics_for_para_mark(&rpr_ref, &host.style))
                                }
                                None => {
                                    let rpr_ref = RunStyle::default();
                                    let fs = self.default_font_size;
                                    (fs, self.metrics_for_para_mark(&rpr_ref, &ParagraphStyle::default()))
                                }
                            };
                            let asc = m.word_ascent_pt(fs);
                            let desc = m.word_descent_pt(fs);
                            let baseline_off = (exact + asc - desc) * 0.5;
                            Some(cursor.visual_y + baseline_off - img.height)
                        }
                        _ => None,
                    };
                    let img_y = s1321_top.unwrap_or(cursor.visual_y + if !self.doc_body_has_real_cjk { img.effect_extent_t.max(0.0) } else { 0.0 });
                    if s1321_top.is_some() {
                        if let Some(v) = block_y_positions.last_mut() {
                            *v = img_y;
                        }
                    }
                    let img_x = if !self.doc_body_has_real_cjk
                        && std::env::var("OXI_BODY_IMAGE_ALIGNMENT_DISABLE").is_err() {
                        img.host_paragraph.as_deref().map_or(start_x, |host| {
                            let left = host.style.indent_left.unwrap_or(0.0)
                                + host.style.indent_first_line.unwrap_or(0.0);
                            let right = host.style.indent_right.unwrap_or(0.0);
                            let remaining = (content_width - left - right - img.width).max(0.0);
                            start_x + left + match host.alignment {
                                Alignment::Center => remaining * 0.5,
                                Alignment::Right => remaining,
                                _ => 0.0,
                            }
                        })
                    } else { start_x };
                    elements.push(LayoutElement::new(
                        img_x,
                        img_y,
                        img.width,
                        img.height,
                        LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                            crop: img
                                .crop
                                .as_ref()
                                .map(|c| (c.top, c.right, c.bottom, c.left)),
                        },
                    ));
                    if std::env::var_os("OXI_DBG_IMGADV").is_some() {
                        eprintln!("[IMGADV] cy={:.2} img_h={:.2} img_w={:.2} adv={:.2} before={:.2} after={:.2} host_exact={:?}",
                            cursor.cursor_y, img.height, img.width, img_adv, img.paragraph_space_before, img.paragraph_space_after, img.host_exact_line);
                    }
                    cursor.advance(img_adv);
                    if s1181_img_unsnap != 0.0 {
                        cursor.advance_split(0.0, s1181_img_unsnap);
                    }
                    prev_para_style_id = None;
                    // S1566 (2026-09-26, default ON, opt-out OXI_S1566_DISABLE): the
                    // image paragraph's OWN contextualSpacing is what the next
                    // paragraph's S874 collapse sees, not the one from the paragraph
                    // before the image. administrative__00433283 p3: «1. Director of
                    // Finance» (List Paragraph, contextualSpacing) / an inline EMF
                    // paragraph (Normal) / «Mr. Wells made a motion» (Normal): the stale
                    // prev_ctx=true took the (true,false) arm and dropped the 10pt
                    // after (Word Info(6) 570.0, Oxi 560.17), so the page ran 10pt
                    // high and a double-spaced bullet line Word pushes to p4 stayed.
                    if std::env::var_os("OXI_S1566_DISABLE").is_none() {
                        prev_contextual_spacing = img
                            .host_paragraph
                            .as_ref()
                            .map_or(false, |h| h.style.contextual_spacing);
                    }
                    prev_borders = None; // S658: an image breaks border-merge adjacency
                    prev_autospacing_numid = None; // S931: and list adjacency
                    prev_keep_next = false; // S739
                                            // S961 (2026-07-21, HELD OPT-IN OXI_S961=1, default OFF):
                                            // this arm resets every OTHER paragraph carry (style id,
                                            // borders, autospacing, keepNext) but leaves space_after, so
                                            // the paragraph BEFORE an image leaks its after-spacing onto
                                            // the paragraph AFTER it — the Table arm (6680) zeroes it.
                                            // Zeroing it here IS right for policies__00148f8d p45 (its
                                            // caption sits 8pt low without it) but WRONG for
                                            // legal__00089377 and reports__000e8acd (both PASS 1.0 →
                                            // 0.98 with it). Neither behaviour is Word's: an image-only
                                            // paragraph is a real paragraph, so Word collapses
                                            // max(prev.after, imagePara.before) above it and
                                            // max(imagePara.after, next.before) below it — and S537
                                            // discards the image paragraph's own spacing, so the IR
                                            // cannot express either boundary. The real fix carries that
                                            // paragraph's space_before/space_after onto ir::Image and
                                            // collapses at both ends; until then neither approximation
                                            // may ship.
                    if std::env::var("OXI_S961").is_ok() {
                        prev_space_after = 0.0;
                    } else if s965 {
                        // S965: the image paragraph's OWN after is what the next
                        // paragraph collapses against — not the one before it.
                        // S1183: an after-autospacing host hands the flat auto
                        // amount to that collapse instead of its explicit after.
                        prev_space_after = s1183_sa.unwrap_or(img.paragraph_space_after);
                    }
                    // S1293: the drawing's host paragraph ended with a page break
                    // AFTER the drawing. The paragraph is gone (its mark shared the
                    // drawing's line in Word and added no height), so its break
                    // rides on the image and fires HERE -- after the drawing is on
                    // the page, which is the whole point. Mirrors the paragraph
                    // `page_break_after` handler above.
                    if img.page_break_after
                        && !elements.is_empty()
                        && std::env::var("OXI_S1293_DISABLE").is_err()
                    {
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        if let Some(g) = s755_geom.as_ref() {
                            start_y = g.top(pages.len() + 1);
                            content_height = g.ch(pages.len() + 1);
                        }
                        cursor.set(start_y);
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        current_page_idx += 1;
                        lm2_cells = 0;
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        s900_fold(
                            &mut footnote_reserve_current,
                            &mut footnote_ids_current_page,
                            &mut s900_pending_deferred,
                            current_page_idx,
                        );
                    }
                }
                Block::UnsupportedElement(_) => {
                    // Skip unsupported elements in layout
                }
                Block::Math(math_block) => {
                    // Phase 3: emit positioned LayoutElements for math primitives.
                    // Fraction/Sup/Sub/SubSup render stacked; other primitives
                    // fall back to flat text for now.
                    // S1613 (2026-09-30, default ON, opt-out OXI_S1613_DISABLE): a
                    // display equation takes its host paragraph's size and spacing.
                    // `_pb_dispmath_h_gen.py` (blind-G EN educational__005f2e39 slice,
                    // docDefaults after 200): Word M1->M2 x 26.64 / 36.72, a/b 36.36 /
                    // 46.32, the document's t= 60.60 / 70.56 and S= 56.28 / 66.36 with
                    // after 0 / 200 -- the paragraph's 10pt after always applies; Oxi
                    // gave both arms the same height (it dropped the paragraph). The
                    // x arm (no run size) is 11pt, the document default: Cambria
                    // Math's natural line 12.90 vs Word 12.84; Oxi's fixed 10.5 gave
                    // 12.31.
                    let s1613_host = match math_block {
                        crate::ir::MathBlock::Display { host: Some(h), .. } => Some(h.as_ref()),
                        _ => None,
                    };
                    let math_font_size: f32 = s1613_host
                        .and_then(|h| h.ppr_rpr.as_ref().and_then(|r| r.font_size)
                            .or_else(|| h.default_run_style.as_ref().and_then(|r| r.font_size)))
                        .unwrap_or(10.5);
                    if let Some(h) = s1613_host {
                        cursor.advance(h.space_before.unwrap_or(0.0));
                    }
                    // S524 (coverage, 2026-06-09): apply the display equation's jc
                    // (default Center) — Word CENTERS display math (oMathPara) at the
                    // page center; Oxi previously hard-coded the left margin. PDF-confirmed
                    // on a/b, x^2, x_i, sqrt(x), nested (all Word-centered at page mid).
                    // Compute the bbox width first, then position by jc.
                    let content_w =
                        (page.size.width - page.margin.left - page.margin.right).max(0.0);
                    let bbox_pre =
                        crate::layout::math::layout_math_block(math_block, math_font_size);
                    let math_jc = match math_block {
                        crate::ir::MathBlock::Display { jc, .. } => *jc,
                        _ => crate::ir::MathAlignment::Left,
                    };
                    let x = match math_jc {
                        crate::ir::MathAlignment::Center
                        | crate::ir::MathAlignment::CenterGroup => {
                            page.margin.left + ((content_w - bbox_pre.advance) * 0.5).max(0.0)
                        }
                        crate::ir::MathAlignment::Right => {
                            page.margin.left + (content_w - bbox_pre.advance).max(0.0)
                        }
                        crate::ir::MathAlignment::Left => page.margin.left,
                    };
                    let (math_elems, bbox) = crate::layout::math::emit_math_block(
                        math_block,
                        x,
                        cursor.cursor_y,
                        math_font_size,
                    );
                    if !math_elems.is_empty() {
                        // S652 (coverage, 2026-06-24): reserve the equation
                        // paragraph's vertical advance from the ACTUAL emitted
                        // glyph geometry, not bbox.height(). The layout bbox
                        // over-estimates the rendered extent (emit_nary's
                        // descent = op_size + sub.height() double-counts the
                        // full operator height below the baseline; leaf glyph
                        // boxes are a loose 0.8em/0.4em), so the old
                        // `bbox.height().max(fs*1.2)+fs*0.3` over-reserved by
                        // +1.7pt (rad) to +16pt (n-ary sum) vs Word — pixel-
                        // confirmed by tools/metrics/mixedh_lineplace.py. Word
                        // reserves max(ink_height + ~1.4pt leading, math line
                        // height). Glyph BASELINES render correctly (mathH
                        // matches Word within ±0.7pt), and emit_text_at sets a
                        // text element's top = baseline − 0.8·fs with h = 1.2·fs,
                        // so baseline = y + h·2/3 is recoverable per element.
                        // Take a tight cap-ascent above the topmost baseline and
                        // a small descent below the bottommost; non-text
                        // elements (fraction bar, radical rule, box rect, matrix
                        // lines) are already tight so use their raw [y, y+h].
                        // Display math is absent from the whole gate corpus
                        // (0/2391 docx) → pure coverage, zero gate risk.
                        // Opt-out OXI_S652_DISABLE.
                        // Constants calibrated against Word (mixedh_lineplace.py
                        // 7-structure pixel sweep, _s529_sweep.py): cap-ascent
                        // 0.60·fs above the topmost baseline, 0.05·fs descent
                        // below the bottommost, +1.5pt leading, floored at the
                        // math line height 1.14·fs. Residual ≤ ~2pt (glyph-class
                        // x-height vs cap variation + radical overbar), vs the
                        // old +1.7..+16pt over-reservation.
                        let advance = if std::env::var("OXI_S652_DISABLE").is_ok() {
                            bbox.height().max(math_font_size * 1.2) + math_font_size * 0.3
                        } else {
                            let asc = 0.60_f32;
                            let desc = 0.05_f32;
                            let lead = 1.5_f32;
                            let floor = 1.14_f32;
                            let mut ink_top = f32::INFINITY;
                            let mut ink_bot = f32::NEG_INFINITY;
                            for e in &math_elems {
                                let (lo, hi) = match &e.content {
                                    LayoutContent::Text { text, .. } => {
                                        let fs = e.height / 1.2;
                                        let baseline = e.y + e.height * (2.0 / 3.0);
                                        // An INTEGRAL sign (∫∮∬∭∮… U+222B–2233) is the
                                        // one math glyph that curls ~0.3em BELOW the
                                        // baseline, so its tight descent under-counts
                                        // and the next line would overlap it (the ∫
                                        // "None"/overlap case). Use its raw box. Every
                                        // other glyph — including an ENLARGED √ or ∑,
                                        // which are tall above the baseline but shallow
                                        // below — keeps the tight cap-ascent/descent
                                        // (the box over-reserves them, e.g. radfrac).
                                        let is_integral = text
                                            .chars()
                                            .any(|c| ('\u{222B}'..='\u{2233}').contains(&c));
                                        if is_integral {
                                            (e.y, e.y + e.height)
                                        } else {
                                            (baseline - asc * fs, baseline + desc * fs)
                                        }
                                    }
                                    _ => (e.y, e.y + e.height),
                                };
                                if lo < ink_top {
                                    ink_top = lo;
                                }
                                if hi > ink_bot {
                                    ink_bot = hi;
                                }
                            }
                            if ink_bot > ink_top {
                                {
                                    // S1260 (2026-08-29, default ON, opt-out
                                    // OXI_S1260_DISABLE): the floor on a
                                    // single-line equation is the FACE's own
                                    // natural line height, not the calibrated
                                    // 1.14. Cambria Math measures 1.172363em
                                    // (S1258 read it off the file), so at 10.5pt
                                    // the floor is 12.31 where the constant gave
                                    // 11.97. Word's own figure, recovered from
                                    // the `_pb_eqgrid` sweep by inverting the
                                    // S1259 snap, lies in (12, 18]: the `plain`
                                    // arm takes 1 cell on an 18pt grid, 1 on a
                                    // 24pt grid but **2** on a 12pt grid, which
                                    // is only consistent with a natural just
                                    // OVER 12 -- and 11.97 sits 0.03 under, the
                                    // single miss of that sweep. Only equations
                                    // whose ink is smaller than the floor move.
                                    let f = if std::env::var("OXI_S1260_DISABLE").is_err() {
                                        self.registry
                                            .get("Cambria Math")
                                            .natural_line_height_hhea(math_font_size)
                                    } else {
                                        math_font_size * floor
                                    };
                                    (ink_bot - ink_top + lead).max(f)
                                }
                            } else {
                                bbox.height().max(math_font_size * 1.2) + math_font_size * 0.3
                            }
                        };
                        elements.extend(math_elems);
                        // S1259 (2026-08-29, default ON, opt-out
                        // OXI_S1259_DISABLE): inside a TYPED docGrid a DISPLAY
                        // equation occupies a WHOLE NUMBER of grid cells.
                        // WORD TRUTH (`tools/metrics/_pb_eqgrid_{gen,read}.py`,
                        // 18 arms = 6 equation heights x 3 line pitches, advance
                        // read body-line to body-line out of Word's PDF): every
                        // single arm is integral --
                        //   pitch 360 (18pt)  plain 2  frac 3  nary 3  deep 4
                        //   pitch 240 (12pt)  plain 4  frac 4  nary 5  deep 6
                        //   pitch 480 (24pt)  plain 2  frac 2  nary 3  deep 3
                        // and the count TRACKS the height, so it is a snap and
                        // not a constant. `ceil(oxi_natural / cell)` reproduces
                        // Word in 17 of the 18 (the miss is `240/plain`, whose
                        // natural lands 0.03pt under an exact 3-cell boundary --
                        // an accuracy residual in the natural, not in this rule).
                        // WITNESS probeomml_equations: Word gives all 7 of its
                        // equations exactly 3.000 cells (54.00 on an 18.00 grid)
                        // where Oxi spent the raw 2.865 (51.57), so the body
                        // after each one crept up 2.44pt and the page ended up
                        // holding 51 lines against Word's 48.
                        let advance = match grid_pitch {
                            Some(cell)
                                if cell > 0.0
                                    && std::env::var("OXI_S1259_DISABLE").is_err() =>
                            {
                                (advance / cell).ceil() * cell
                            }
                            _ => advance,
                        };
                        cursor.advance(advance);
                    }
                    if let Some(h) = s1613_host {
                        cursor.advance(h.space_after.unwrap_or(0.0));
                    }
                }
            }
            // S1294: now that the block is placed, ask whether its section
            // opened a page, and pad if the restart would repeat the previous
            // page's parity. The blank is pushed BEFORE the in-flight elements,
            // which have not been pushed yet, so it lands between the two with
            // no reflow -- the section's content is already at a page top.
            if let Some(n) = s1294_restart {
                // S1421 (2026-09-16, default ON, opt-out OXI_CONTINUOUS_SECTION_ORIGIN_DISABLE):
                // the checkpoint's opt-in promoted -- a continuous section's
                // page-number restart is judged on the page its first block
                // actually starts on (a paragraph may begin on the preceding
                // page). reference__13e1b7fca0030560: 0.0131 (pcd +2) -> 0.977
                // (pcd 0). Env gates: ja 188 same, golden 185 same, en 291 same.
                let section_origin = std::env::var_os("OXI_CONTINUOUS_SECTION_ORIGIN_DISABLE").is_none();
                // A paragraph may start on the preceding page and continue onto
                // another. Numbering belongs to its first occupied flow page.
                let origin = if section_origin {
                    pages.iter().enumerate().skip(s1294_pages_before)
                        .find(|(_, p)| p.elements.iter().any(|e| e.paragraph_index == Some(block_idx)))
                        .map_or(pages.len(), |(i, _)| i)
                } else { pages.len() };
                let began_page = s1294_at_top_before || origin > s1294_pages_before;
                if began_page {
                    let here = origin as i64; // the in-flight page's index
                    let prev_logical = logical_base + here - 1;
                    let mut numbered_origin = origin;
                    if page.even_odd_hf
                        && here > 0
                        && (n as i64).rem_euclid(2) == prev_logical.rem_euclid(2)
                    {
                        dbg_page_push(pages.len(), 0);
                        pages.insert(origin, LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: Vec::new(),
                        });
                        current_page_idx += 1;
                        numbered_origin += 1;
                        if let Some(v) = block_page_indices.last_mut() {
                            if section_origin {
                                if *v >= origin { *v += 1; }
                            } else { *v = current_page_idx; }
                        }
                        if let Some(v) = block_start_page_indices.last_mut() {
                            if section_origin {
                                if *v >= origin { *v += 1; }
                            } else { *v = current_page_idx; }
                        }
                    }
                    if section_origin {
                        if let Some(v) = block_start_page_indices.last_mut() { *v = numbered_origin; }
                    }
                    logical_base = n as i64 - numbered_origin as i64;
                }
            }
            // S560: record the deepest column-bottom reached on this page so a
            // following column-section (heterogeneous path) starts below it.
            if heterogeneous {
                section_max_y = section_max_y.max(cursor.cursor_y);
            }
        }

        // S1352 (2026-09-11, default ON, opt-out OXI_COL_TRAILING_BALANCE_DISABLE):
        // a column run closed by a section that holds NOTHING still balances.
        // The balance above fires when a block belonging to the next run comes
        // past; when that run is empty no such block ever comes, and the run
        // was left filled down the left column instead. Word balances it:
        // `end_8lines_emptyafter` splits 4/4 where this engine had all 8 left.
        //
        // A run that is closed by no section at all is NOT balanced — Word
        // fills it — which is why this asks for a later run to exist rather
        // than simply balancing whatever is left (`end_8lines_nothingafter`).
        if num_columns == 2
            && active_run_idx + 1 < col_runs.len()
            && col_runs[active_run_idx + 1].1 != num_columns
            && std::env::var("OXI_S750_DISABLE").is_err()
            && std::env::var("OXI_TEXT_BALANCE_DISABLE").is_err()
            && std::env::var("OXI_COL_TRAILING_BALANCE_DISABLE").is_err()
        {
            let allocated_bottom = column_search.observe(
                ir_index, allocation_start, page.blocks.len(), page, current_page_idx,
                col_band_top, start_y + content_height, &col_widths, &elements, pending_section_gap);
            if let Some(bottom) = allocated_bottom.or_else(|| LayoutEngine::rebalance_text_columns(
                &mut elements,
                col_band_top,
                &col_x_positions,
                &page.blocks,
                if std::env::var("OXI_COLUMN_COMPAT_BALANCE_DISABLE").is_err() { pending_section_gap } else { 0.0 },
                self.compat_mode,
                if page.doc_grid_no_type { None } else { page.grid_line_pitch },
            )) {
                cursor.set(bottom);
                section_max_y = section_max_y.max(bottom);
            }
        }

        // S743 (2026-07-04, default ON, opt-out OXI_S743_DISABLE): ENDNOTE
        // bodies flow after the last body block (Word renders them at the
        // document end, subject to normal page breaks — probexendnote: Word 4
        // pages with notes 1-26 on p3 + xxvii..xxx on p4; Oxi dropped them
        // entirely = the gate-masked content loss found by the probe hunt).
        // Marker = lowerRoman (Word's default endnote numFmt), overwriting the
        // empty endnoteRef first run (the footnote-marker pattern). Separator =
        // Word's default 144pt line (S479) after a one-line gap. 0 corpus docs
        // reference endnotes (scanned) → byte-identical by construction.
        if !page.endnotes.is_empty() && std::env::var("OXI_S743_DISABLE").is_err() {
            fn to_lower_roman(mut n: u32) -> String {
                if n == 0 {
                    return "0".into();
                }
                let vals: &[(u32, &str)] = &[
                    (1000, "m"),
                    (900, "cm"),
                    (500, "d"),
                    (400, "cd"),
                    (100, "c"),
                    (90, "xc"),
                    (50, "l"),
                    (40, "xl"),
                    (10, "x"),
                    (9, "ix"),
                    (5, "v"),
                    (4, "iv"),
                    (1, "i"),
                ];
                let mut out = String::new();
                for &(v, s) in vals {
                    while n >= v {
                        out.push_str(s);
                        n -= v;
                    }
                }
                out
            }
            let sep_gap = grid_pitch.unwrap_or(14.0).max(6.0);
            if cursor.cursor_y + sep_gap * 2.0 > start_y + content_height {
                dbg_page_push(pages.len(), 0);
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(&mut elements),
                });
                if let Some(g) = s755_geom.as_ref() {
                    start_y = g.top(pages.len() + 1);
                    content_height = g.ch(pages.len() + 1);
                }
                cursor.set(start_y);
                current_page_idx += 1;
            }
            // S1257 (2026-08-29, default ON, opt-out OXI_S1257_DISABLE): the
            // CONTINUATION separator paragraph, which Word lays out at the top
            // of every continuation page of the note area.
            // WORD TRUTH `educational__001217ec`: its body pages start at
            // 87.37 and its endnote CONTINUATION pages at 99.13 -- 11.76pt
            // lower, one empty Normal (Cambria 10) line. The paragraph is
            // `<w:endnote w:type="continuationSeparator" w:id="0">`, which the
            // parser used to discard outright. Reserving it moves the page
            // boundaries onto Word's (Oxi fitted one note too many per page).
            let en_contsep_h: f32 = if std::env::var("OXI_S1257_DISABLE").is_err() {
                page.endnotes
                    .iter()
                    .find(|n| n.number == u32::MAX - 2)
                    .and_then(|note| {
                        note.blocks.iter().find_map(|b| match b {
                            Block::Paragraph(p) => Some(p),
                            _ => None,
                        })
                    })
                    .map(|p| {
                        // The S833 `special_h` convention: a TEXT-EMPTY special
                        // paragraph sizes through the DEFAULT PARAGRAPH STYLE's
                        // run props, and the estimate drops style-level spacing.
                        let mut h = self.estimate_para_height(
                            p,
                            content_width,
                            None,
                            None,
                            false,
                            None,
                            None,
                        );
                        if p.runs.iter().all(|r| r.text.trim().is_empty()) {
                            if let Some(drs) = p.style.default_run_style.as_ref() {
                                if let Some(fs) = drs.font_size {
                                    let m = self.metrics_for(drs, &p.style);
                                    let line = m.natural_line_height_hhea(fs);
                                    if line > 0.0 {
                                        h = line;
                                    }
                                }
                            }
                        }
                        if h <= 0.0 {
                            // ★A BARE `<w:p/>` has no runs and no pPr, so the
                            // estimate yields nothing and S833's
                            // default_run_style fallback has nothing to read.
                            // It is still a LINE: size it through the paragraph
                            // mark the way any empty paragraph is sized.
                            let rs = RunStyle::default();
                            let fs = self.resolve_font_size(&rs, &p.style);
                            let line = self
                                .metrics_for(&rs, &p.style)
                                .natural_line_height_hhea(fs);
                            if line > 0.0 {
                                h = line;
                            }
                        }
                        if !p.style.has_direct_spacing {
                            h += p.style.space_before.unwrap_or(0.0)
                                + p.style.space_after.unwrap_or(0.0);
                        }
                        h
                    })
                    .unwrap_or(0.0)
            } else {
                0.0
            };
            if std::env::var("OXI_DBG_CONTSEP").is_ok() {
                eprintln!("[CONTSEP] h={:.2} n_endnotes={} numbers={:?}",
                    en_contsep_h, page.endnotes.len(),
                    page.endnotes.iter().map(|n| n.number).rev().take(3).collect::<Vec<_>>());
            }
            // S1256 (2026-08-29, default ON, opt-out OXI_S1256_DISABLE): the
            // LAST body paragraph's after-spacing, which the block loop never
            // adds because nothing follows it. The endnote block does follow it.
            // WORD TRUTH `educational__001217ec`: its `References` heading is
            // Heading1 = `spacing before=520 after=440 line=440 atLeast`, i.e.
            // 22pt after. Word puts the heading line at 105.91 (h 25.3, the
            // Arial natural beating the 22pt atLeast) and the first endnote at
            // 166.45; Oxi ended the heading at 130.30 and started the notes at
            // 144.30 = 130.30 + the 14pt separator gap alone. Adding the 22
            // gives 166.30 -- 0.15pt from Word.
            if std::env::var("OXI_S1256_DISABLE").is_err() {
                if let Some(Block::Paragraph(last)) = page.blocks.last() {
                    let sa = last.style.space_after.unwrap_or(0.0);
                    if sa > 0.0 {
                        cursor.advance(sa);
                    }
                }
            }
            // S1257: draw the rule only when the document DECLARES one. The
            // separator note of `educational__001217ec` is a bare empty
            // paragraph, and Word draws nothing (p16 `get_drawings()` = 0);
            // `probexendnote_endnotes` declares `<w:separator/>` and keeps its
            // rule. The GAP is spent either way -- the paragraph is laid out in
            // both documents.
            if self.endnote_sep_line || std::env::var("OXI_S1257_DISABLE").is_ok() {
                elements.push(LayoutElement::new(
                    start_x,
                    cursor.cursor_y + sep_gap * 0.5,
                    144.0_f32.min(content_width),
                    1.0,
                    LayoutContent::BoxRect {
                        fill: Some("#000000".to_string()),
                        stroke_color: None,
                        stroke_width: 0.0,
                        corner_radius: 0.0,
                    },
                ));
            }
            cursor.advance(sep_gap);
            let empty_fn_h_en = std::collections::HashMap::new();
            // S1257: the sentinel-numbered special paragraphs are carriage,
            // not notes -- they must not be rendered or take a marker seq.
            // S1603 (2026-09-29, default ON, opt-out OXI_S1603_DISABLE): the gap
            // BETWEEN two endnotes is max(previous note's space after, next
            // note's space before), not their sum. `_pb_notespace_gen.py`
            // (blind-G EN forms__005d851e, compat 15, every note before/after
            // edited together): 6/6 -> +6.0, 12/3 -> +12, 3/12 -> +12,
            // 0/12 -> +12, 12/0 -> +12 on a 13.4 line; the first note keeps
            // its whole before (+6 / +12 / +3 / +0 / +12 below the separator).
            // Oxi summed them (+12 at 6/6): four notes too many on p4, W4/O5.
            let s1603 = std::env::var_os("OXI_S1603_DISABLE").is_none();
            let mut en_pending_after: Option<f32> = None;
            for (en_seq, note) in page
                .endnotes
                .iter()
                .filter(|n| n.number < u32::MAX - 2)
                .enumerate()
            {
                let mut first_para = true;
                let mut s1603_note_first = true;
                let mut s1603_note_last_sa: Option<f32> = None;
                for block in &note.blocks {
                    if let Block::Paragraph(para) = block {
                        let para_to_render: Paragraph = if first_para {
                            let mut p = para.clone();
                            // Marker = the SEQUENCE position (the footnote-path
                            // `seq` convention) — note.number is the raw XML id,
                            // which is offset by the separator/continuation
                            // entries (probexendnote ids start at 2 -> the
                            // number-based marker rendered xxviii for 注27).
                            // S1254 (2026-08-29, default ON, opt-out
                            // OXI_S1254_DISABLE): the SECTION's declared endnote
                            // format wins. ECMA-376 defaults endnotes to
                            // lowerRoman, which is what S743 hard-coded when no
                            // corpus document referenced endnotes;
                            // `educational__001217ec` declares
                            // `<w:endnotePr><w:numFmt w:val="decimal"/>` on its
                            // sectPr and Word stamps `17 / 35 / 51` where Oxi
                            // stamped `xvii / xxxv / li`.
                            let n = en_seq as u32 + 1;
                            let prefix = match page.endnote_number_format.as_deref() {
                                Some(f)
                                    if std::env::var("OXI_S1254_DISABLE").is_err()
                                        && !f.eq_ignore_ascii_case("lowerRoman") =>
                                {
                                    crate::parser::numbering::format_number(n, f)
                                }
                                _ => to_lower_roman(n),
                            };
                            if let Some(first_run) = p.runs.first_mut() {
                                if first_run.text.is_empty() {
                                    first_run.text = prefix.clone();
                                } else {
                                    first_run.text = format!("{}{}", prefix, first_run.text);
                                }
                            }
                            first_para = false;
                            p
                        } else {
                            para.clone()
                        };
                        let mut para_to_render = para_to_render;
                        if s1603 && std::mem::replace(&mut s1603_note_first, false) {
                            if let Some(prev_sa) = en_pending_after.take() {
                                let sb = para_to_render.style.space_before.unwrap_or(0.0);
                                para_to_render.style.space_before = Some((sb - prev_sa).max(0.0));
                            }
                        }
                        let (en_elements, _, _) = self.layout_paragraph(
                            &para_to_render,
                            start_x,
                            &mut cursor,
                            content_width,
                            // S1257: a continuation page of the note area starts
                            // BELOW the continuation separator, and ends at the
                            // same bottom margin -- so the top moves down and the
                            // height shrinks by the same amount.
                            content_height - en_contsep_h,
                            start_y + en_contsep_h,
                            page,
                            &mut pages,
                            &mut elements,
                            grid_pitch,
                            None,
                            false,
                            None,
                            None,
                            false,
                            false,
                            0.0,
                            None,
                            None,
                            None,
                            false,
                            false,
                            None,
                            0.0,
                            &empty_fn_h_en,
                            1,
                            0,
                            &[],
                            0.0, // S749: band top unused (1-col)
                            false,
                            false,
                            None,  // S755
                            None,  // S758
                            None,  // S-TWOSEG
                            false, // S835
                            0.0,
                            None,  // S900
                            None,  // S903
                            false, // S916
                            None,
                        );
                        elements.extend(en_elements);
                        // S1255 (2026-08-29, default ON, opt-out
                        // OXI_S1255_DISABLE): an endnote paragraph's OWN
                        // before/after spacing. S743 called `layout_paragraph`
                        // and advanced the cursor by the line heights alone, so
                        // every endnote sat directly under the one above.
                        // WORD TRUTH `educational__001217ec` (61 endnotes, style
                        // FootnoteText = `spacing after=120 line=270 atLeast`):
                        // the pitch WITHIN a note is 13.56 and Oxi already
                        // matches it at 13.50 (the atLeast line resolves), but
                        // the pitch BETWEEN notes is Word 19.44 against Oxi
                        // 13.50 -- the missing 5.94 is the style's 120tw = 6pt.
                        // Over 61 notes that is ~362pt = the page Word has and
                        // Oxi does not (Word 19 pages, Oxi 18).
                        if std::env::var("OXI_S1255_DISABLE").is_err() {
                            let sa = para_to_render.style.space_after.unwrap_or(0.0);
                            if sa > 0.0 {
                                // ★`advance`, not `cursor_y += `: the Cursor
                                // carries a pagination track AND a drawing
                                // track, and touching only the first made the
                                // page BREAKS see the spacing while the ink did
                                // not — the page count moved to Word's 19 while
                                // every note still sat at the 13.50 pitch.
                                cursor.advance(sa);
                            }
                            s1603_note_last_sa = Some(sa);
                        }
                    }
                }
                en_pending_after = s1603_note_last_sa;
            }
            let _ = current_page_idx;
        }

        // Final page
        dbg_page_push(pages.len(), 0);
        pages.push(LayoutPage {
            width: page.size.width,
            height: page.size.height,
            elements,
        });

        // R7.61 (Day 36 part 8, 2026-05-14): post-paginate sweep — move
        // vMerge="restart" cell content that overflowed past page_bottom to
        // the next page. Each marked element is shifted so its Y in the next
        // page mirrors the same offset past page_top. This rectifies Oxi's
        // page assignment for a1d6 row 13 ※２/※３ (Word p4, Oxi visually on
        // p3 past page bottom) without disturbing other docs because the
        // marker is only set for vMerge=restart cell text past page_bottom
        // — body text and other cells are not marked. Scope-verified across
        // 55-doc Phase 1 baseline: only a1d6 has 2+ vMerge restart overflow
        // entries on the same page.
        {
            let page_top = page.margin.top;
            let page_bottom_sweep = page_top + content_height;
            for i in 0..pages.len() {
                let take = std::mem::take(&mut pages[i].elements);
                let (keep, overflow): (Vec<_>, Vec<_>) = take
                    .into_iter()
                    .partition(|e| !e.vmerge_restart_overflow_to_next_page && e.vmerge_destination_page.is_none());
                pages[i].elements = keep;
                if overflow.is_empty() {
                    continue;
                }
                let shift = page_bottom_sweep - page_top;
                for mut elem in overflow {
                    let next_idx = elem.vmerge_destination_page.take().unwrap_or(i + 1);
                    let offset = next_idx as f32 - i as f32;
                    while next_idx >= pages.len() {
                        dbg_page_push(pages.len(), 0);
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: Vec::new(),
                        });
                    }
                    elem.y -= shift * offset as f32;
                    elem.vmerge_restart_overflow_to_next_page = false;
                    pages[next_idx].elements.push(elem);
                }
            }
        }

        // Layout text boxes and add to the correct layout page
        // The current_page_idx tracking tells us which layout page each anchor block ended up on
        //
        // S478: Word draws floating objects in ascending wp:anchor relativeHeight
        // order (highest = drawn last = on top). Oxi previously emitted text boxes
        // in parse order, so an opaque (white-filled) callout with a HIGH
        // relativeHeight that should hide an overlapping lower-relativeHeight box's
        // content was drawn FIRST → the other box's text painted over it (bled
        // through). 2ea81a p2: callout 予納する納税者名義 (relHeight 251801088, white
        // fill) overlaps the 留意事項 box content (relHeight 251694592); Word draws
        // the callout on top (opaque), Oxi drew it under. Fix: emit text boxes in
        // ascending relativeHeight (stable for ties = parse order). Anchor Y /
        // pagination untouched (render order only) = Phase-1 safe.
        // Default ON, opt-out OXI_S478_DISABLE.
        //
        // KNOWN UNTESTED EDGE (zero instances in the current corpus, so not
        // handled here — confirmed by the S478 13-doc structural audit):
        //   standalone text-bearing VML (relativeHeight=0): all corpus VML is
        //       an mc:Fallback alternate (never rendered — Oxi takes mc:Choice
        //       DrawingML, which always carries a relativeHeight). A future
        //       standalone VML text box would parse relHeight=0 and be forced to
        //       the bottom of every overlap; give it a doc-order-derived key then.
        let s478_zorder = std::env::var("OXI_S478_DISABLE").is_err();
        // S765 (2026-07-08, default ON, opt-out OXI_S765_DISABLE): draw
        // floating IMAGES and TEXTBOXES in ONE z-order pass keyed by
        // (behind_doc first, then relativeHeight ascending) — Word draws all
        // floating objects in one relativeHeight order, but Oxi historically
        // drew ALL textboxes then ALL images, so a lower-relHeight image
        // always covered a higher-relHeight textbox. uk_local_spending's
        // cover: the title textbox (relHeight 251662848, WHITE text on the
        // purple banner) sits ABOVE two purple picture graphics (relHeight
        // 251660800/251661824) — Oxi drew the pictures last → they buried the
        // title + subtitle + date text (the whole cover text vanished).
        // Verified language-independent (framework's giant crest = same).
        let s765 = s478_zorder && std::env::var("OXI_S765_DISABLE").is_err();
        // Build a unified float list: (behind_doc, relative_height, kind, idx).
        enum FloatKind {
            Tb,
            Img,
        }
        let mut floats: Vec<(bool, u32, FloatKind, usize)> = Vec::new();
        for (i, tb) in page.text_boxes.iter().enumerate() {
            floats.push((tb.behind_doc, tb.relative_height, FloatKind::Tb, i));
        }
        for (i, img) in page.floating_images.iter().enumerate() {
            floats.push((img.behind_doc, img.relative_height, FloatKind::Img, i));
        }
        if s765 {
            // behind_doc=true (behind text) first, then by relativeHeight asc.
            floats.sort_by_key(|&(bd, rh, _, _)| (!bd, rh));
        } else {
            // Legacy order: all textboxes (relHeight-sorted) then all images.
            floats.sort_by_key(|&(_, rh, ref k, i)| {
                let tb_first = matches!(k, FloatKind::Tb);
                (
                    !tb_first,
                    if s478_zorder && tb_first { rh } else { 0 },
                    i as u32,
                )
            });
        }
        // S918 (2026-07-18, default ON, opt-out OXI_S918_DISABLE): sorting a
        // behindDoc float before other floats is not enough. Body elements
        // were already emitted, so appending still painted an opaque behindDoc
        // image over body text. Insert behindDoc elements before the existing
        // body layer while preserving their stable relativeHeight order.
        let s918_behind_body = s765 && std::env::var("OXI_S918_DISABLE").is_err();
        let mut behind_insert_counts = vec![0usize; pages.len()];
        for (_, _, kind, fi) in &floats {
            let behind_doc = match kind {
                FloatKind::Tb => page.text_boxes[*fi].behind_doc,
                FloatKind::Img => page.floating_images[*fi].behind_doc,
            };
            match kind {
                FloatKind::Tb => {
                    let tbi = *fi;
                    let text_box = &page.text_boxes[tbi];
                    // S1123: anchors resolve on the block's START page.
                    let s1123 = std::env::var("OXI_S1123_DISABLE").is_err();
                    let target_page = if s1123 {
                        block_start_page_indices
                            .get(text_box.anchor_block_index)
                            .copied()
                            .unwrap_or(0)
                    } else {
                        block_page_indices
                            .get(text_box.anchor_block_index)
                            .copied()
                            .unwrap_or(0)
                    };
                    // S1270 NOT APPLIED HERE — see `TextBox::host_break_after`.
                    // Moving the box back a page on that flag alone is WRONG:
                    // the anchor's recorded y is the new page's TOP, so the box
                    // lands at y=71.0 over the previous page's own content
                    // (legal__02f84965dccfe4db p4: 140 box glyphs at y=79.6..160.6
                    // OVERLAPPING 26 body elements, and the page count stays 11
                    // because the body still breaks where it did). The page and
                    // the y have to be corrected together, and the way to get
                    // both is to make the inline drawing RESERVE ITS HEIGHT in
                    // the line — then the cursor advances past the box before the
                    // break fires and everything follows. See the ship note.
                    let _ = text_box.host_break_after;
                    if std::env::var("OXI_DEBUG_TB").is_ok() {
                        let (rx, ry) =
                            self.resolve_textbox_position(text_box, page, &block_y_positions, &block_col_x);
                        let anchor_in_range = text_box.anchor_block_index < block_y_positions.len();
                        let preview: String = text_box
                            .blocks
                            .iter()
                            .filter_map(|b| match b {
                                Block::Paragraph(p) => Some(
                                    p.runs
                                        .iter()
                                        .flat_map(|r| r.text.chars())
                                        .take(8)
                                        .collect::<String>(),
                                ),
                                _ => None,
                            })
                            .find(|s| !s.is_empty())
                            .unwrap_or_default();
                        eprintln!("[TB] tbi={} anchor={} in_range={} tgt_page={} resolved=({:.1},{:.1}) wh=({:.0},{:.0}) vrel={:?} text={:?}",
                            tbi, text_box.anchor_block_index, anchor_in_range, target_page, rx, ry,
                            text_box.width, text_box.height,
                            text_box.position.as_ref().and_then(|p| p.v_relative.clone()),
                            preview);
                    }
                    let mut tb_elements = self.layout_text_box(text_box, page, &block_y_positions, &block_col_x);
                    // S1089: keep the band where it was RESERVED — the anchor
                    // paragraph has since moved below it, so the block-relative
                    // resolve would double-shift (the S734 image contract).
                    let mut target_page = target_page;
                    if let Some(&(fp, fy)) = s1089_flow_pos.get(&text_box.anchor_block_index) {
                        let off = text_box
                            .position
                            .as_ref()
                            .filter(|p| p.v_relative.as_deref() == Some("paragraph"))
                            .map(|p| p.y.max(0.0));
                        if text_box.wrap_type == Some(crate::ir::WrapType::TopAndBottom) {
                            if let Some(off) = off {
                                let (_, ry) =
                                    self.resolve_textbox_position(text_box, page, &block_y_positions, &block_col_x);
                                let dy = (fy + off) - ry;
                                if dy.abs() > 0.001 {
                                    for e in tb_elements.iter_mut() {
                                        e.y += dy;
                                        if let LayoutContent::TableBorder { y1, y2, .. } =
                                            &mut e.content
                                        {
                                            *y1 += dy;
                                            *y2 += dy;
                                        }
                                    }
                                }
                                target_page = fp;
                            }
                        }
                    }
                    if let Some(lp) = pages.get_mut(target_page) {
                        if s918_behind_body && behind_doc {
                            let insert_at = behind_insert_counts[target_page];
                            let inserted = tb_elements.len();
                            lp.elements.splice(insert_at..insert_at, tb_elements);
                            behind_insert_counts[target_page] += inserted;
                        } else {
                            lp.elements.extend(tb_elements);
                        }
                    }
                }
                FloatKind::Img => {
                    let img = &page.floating_images[*fi];
                    if let Some(ref _pos) = img.position {
                        let (mut abs_x, mut abs_y) = self.resolve_floating_image_position(
                            img,
                            page,
                            &block_y_positions,
                            page.margin.top,
                        );
                        // S1123: see the textbox arm.
                        let mut target_page = if std::env::var("OXI_S1123_DISABLE").is_err() {
                            block_start_page_indices
                                .get(img.anchor_block_index)
                                .copied()
                                .unwrap_or(0)
                        } else {
                            block_page_indices
                                .get(img.anchor_block_index)
                                .copied()
                                .unwrap_or(0)
                        };
                        if let Some(&(fp, fy)) = s734_flow_pos.get(&img.anchor_block_index) {
                            if img.wrap_type == Some(crate::ir::WrapType::TopAndBottom) {
                                target_page = fp;
                                abs_y = fy + img.position.as_ref().map_or(0.0, |p| p.y.max(0.0));
                            }
                        }
                        if crate::layout::s1467_float_column_flow()
                            && _pos.h_relative.as_deref() == Some("column")
                            && _pos.h_align.is_none()
                        {
                            abs_x += block_col_x.get(img.anchor_block_index).copied().unwrap_or(page.margin.left) - page.margin.left;
                        }
                        let el = LayoutElement::new(
                            abs_x,
                            abs_y,
                            img.width,
                            img.height,
                            LayoutContent::Image {
                                data: img.data.clone(),
                                content_type: img.content_type.clone(),
                                crop: img
                                    .crop
                                    .as_ref()
                                    .map(|c| (c.top, c.right, c.bottom, c.left)),
                            },
                        );
                        if let Some(lp) = pages.get_mut(target_page) {
                            if s918_behind_body && behind_doc {
                                let insert_at = behind_insert_counts[target_page];
                                lp.elements.insert(insert_at, el);
                                behind_insert_counts[target_page] += 1;
                            } else {
                                lp.elements.push(el);
                            }
                        } else if !pages.is_empty() {
                            let last = pages.len() - 1;
                            let insert_at = behind_insert_counts[last];
                            let lp = &mut pages[last];
                            if s918_behind_body && behind_doc {
                                lp.elements.insert(insert_at, el);
                                behind_insert_counts[last] += 1;
                            } else {
                                lp.elements.push(el);
                            }
                        }
                    } else if let Some(lp) = pages.last_mut() {
                        lp.elements.push(LayoutElement::new(
                            start_x,
                            0.0,
                            img.width,
                            img.height,
                            LayoutContent::Image {
                                data: img.data.clone(),
                                content_type: img.content_type.clone(),
                                crop: img
                                    .crop
                                    .as_ref()
                                    .map(|c| (c.top, c.right, c.bottom, c.left)),
                            },
                        ));
                    }
                }
            }
        }

        // Layout header/footer on each layout page
        // Header y = headerDistance (from page top edge), default 36pt (0.5in)
        // Footer y = pageHeight - footerDistance - footerContentHeight
        // S1174: ingest this section's trailing paragraphs so the end-state
        // registry is complete (the next section's page 1 and any trailing
        // page's emit fallback resolve against it).
        if S1174_ACTIVE.with(|c| c.get()) {
            while s1174_ingested < page.blocks.len() {
                LayoutEngine::s1174_ingest_block(&page.blocks[s1174_ingested], false);
                s1174_ingested += 1;
            }
        }
        let header_y = page.header_distance.unwrap_or(36.0);
        let footer_dist = page.footer_distance.unwrap_or(36.0);
        let hdr_x = page.margin.left;
        // S690 (2026-06-29): headers/footers span the FULL page text width, NOT a
        // body column width. `content_width` is the mutable body cursor that ends at
        // `col_widths[last_column]` after the body loop — for a MULTI-COLUMN section
        // (e.g. the bidi 2-col albalunaTaidan) that is a single column (150pt), so a
        // right-aligned header title was placed against the column edge (x≈142) instead
        // of the page right margin (x≈315). Headers always use the full text width.
        // For SINGLE-column docs col_widths[0] == total_content_width, so byte-identical.
        // Opt-out OXI_S690_DISABLE.
        let hdr_width = if std::env::var("OXI_S690_DISABLE").is_err() {
            total_content_width
        } else {
            content_width
        };
        for (page_idx, lp) in pages.iter_mut().enumerate() {
            // WATERMARK: one element per page, centered on the MARGIN box
            // (mso-position-horizontal/vertical-relative:margin, the Word
            // watermark idiom), inserted at index 0 = painted BEHIND the body
            // (VML z-index < 0). Parser scope: only headers carrying a
            // v:textpath string produce Page.watermark (JP corpus: 0 docs).
            if let Some(wm) = &page.watermark {
                let mx0 = page.margin.left;
                let mx1 = page.size.width - page.margin.right;
                let my0 = page.margin.top;
                let my1 = page.size.height - page.margin.bottom;
                let cx = (mx0 + mx1) / 2.0;
                let cy = (my0 + my1) / 2.0;
                lp.elements.insert(
                    0,
                    LayoutElement::new(
                        cx - wm.width / 2.0,
                        cy - wm.height / 2.0,
                        wm.width,
                        wm.height,
                        LayoutContent::WatermarkText {
                            text: wm.text.clone(),
                            color: wm.color.clone(),
                            rotation: wm.rotation,
                            font_family: wm.font_family.clone(),
                        },
                    ),
                );
            }
            // S755: per-page header/footer variant selection — page 1 of a
            // titlePg section renders the FIRST-type blocks, even physical
            // pages render the EVEN-type blocks (blank when the flag is set
            // but no reference exists), everything else the default.
            let s755_pno = page_idx + 1;
            // S1553: the header/footer set of the merged continuous section in
            // which this physical page BEGINS (its lowest block index).
            let s1553_run = if std::env::var_os("OXI_S1553_DISABLE").is_none()
                && page.header_runs.len() > 1
            {
                lp.elements.iter().filter_map(|e| e.paragraph_index).min()
                    .and_then(|b| page.header_runs.iter().rposition(|r| r.block_start <= b))
                    .map(|i| &page.header_runs[i])
            } else {
                None
            };
            let (hs_hdr, hs_hdr_first, hs_hdr_even, hs_ftr, hs_ftr_first, hs_ftr_even, hs_title_pg, hs_even_odd):
                (&[Block], &[Block], &[Block], &[Block], &[Block], &[Block], bool, bool) = match s1553_run {
                Some(r) => (&r.header, &r.header_first, &r.header_even, &r.footer, &r.footer_first, &r.footer_even, r.title_pg, r.even_odd_hf),
                None => (&page.header, &page.header_first, &page.header_even, &page.footer, &page.footer_first, &page.footer_even, page.title_pg, page.even_odd_hf),
            };
            let hdr_blocks: &[Block] = if s755_on && hs_title_pg && s755_pno == 1 {
                hs_hdr_first
            } else if s755_on && hs_even_odd && (first_logical as usize + page_idx) % 2 == 0 {
                hs_hdr_even
            } else {
                hs_hdr
            };
            let ftr_blocks: &[Block] = if s755_on && hs_title_pg && s755_pno == 1 {
                hs_ftr_first
            } else if s755_on && hs_even_odd && (first_logical as usize + page_idx) % 2 == 0 {
                hs_ftr_even
            } else {
                hs_ftr
            };
            // S1174: draw the page's re-resolved STYLEREF text. Pages that
            // began mid-paragraph carry no snapshot — use the nearest earlier
            // page's map, else the section's end state.
            let s1174_hdr_sub: Vec<Block>;
            let s1174_ftr_sub: Vec<Block>;
            let (hdr_blocks, ftr_blocks): (&[Block], &[Block]) = if s1174_have_ref {
                let map = (1..=s755_pno)
                    .rev()
                    .find_map(|k| s1174_snapshots.get(&k))
                    .cloned()
                    .unwrap_or_else(LayoutEngine::s1174_map);
                s1174_hdr_sub = LayoutEngine::s1174_substitute(hdr_blocks, &map);
                s1174_ftr_sub = LayoutEngine::s1174_substitute(ftr_blocks, &map);
                (&s1174_hdr_sub, &s1174_ftr_sub)
            } else {
                (hdr_blocks, ftr_blocks)
            };
            if !hdr_blocks.is_empty() {
                let mut cy = LayoutCursor::new(header_y);
                // S1105: the previous block was a TEXT-bearing paragraph, so any
                // inline image that follows was already merged into that
                // paragraph's lines by S1HDR (s755_header_bottom) — do not give
                // it a line of its own here.
                let mut s1105_prev_text = false;
                // S1268b: the top of the block laid out before this one. A
                // floating drawing is hoisted OUT of its host paragraph into a
                // Block::Image that follows it, so the paragraph a
                // relativeFrom="paragraph" anchor references is the previous
                // block, whose top the cursor has already left.
                let mut prev_block_top = header_y;
                let mut header_float_indices: Vec<(u32, usize)> = Vec::new();
                let header_text_flow = !self.doc_body_has_real_cjk
                    && std::env::var("OXI_HEADER_TEXT_FLOW").is_ok();
                let mut previous_header_para: Option<&Paragraph> = None;
                for block in hdr_blocks {
                    if let Block::Paragraph(para) = block {
                        if LayoutEngine::is_floating_header_frame(para) {
                            let (elements, _, _) = self.layout_header_frame(para, page, cy.cursor_y);
                            lp.elements.extend(elements);
                            continue;
                        }
                    }
                    let previous = previous_header_para;
                    previous_header_para = match block {
                        Block::Paragraph(p) if header_text_flow => Some(p),
                        _ => None,
                    };
                    let this_block_top = cy.cursor_y;
                    if let Block::Paragraph(para) = block {
                        s1105_prev_text = para.runs.iter().any(|r| !r.text.is_empty());
                    }
                    if let Block::Paragraph(para) = block {
                        let empty_fn_h_hdr = std::collections::HashMap::new();
                        let (hdr_elements, _, _) = self.layout_paragraph(
                            para,
                            hdr_x,
                            &mut cy,
                            hdr_width,
                            page.size.height,
                            header_y,
                            page,
                            &mut Vec::new(),
                            &mut Vec::new(),
                            grid_pitch,
                            previous.and_then(|p| p.style.style_id.as_deref()),
                            previous.is_some_and(|p| p.style.contextual_spacing),
                            previous.filter(|p| p.style.after_autospacing).and_then(|p| p.style.num_id.as_deref()),
                            None,
                            false,
                            false,
                            previous.and_then(|p| p.style.space_after).unwrap_or(0.0),
                            None,
                            None,
                            None,
                            false,
                            false,
                            None,
                            0.0,
                            &empty_fn_h_hdr,
                            1,
                            0,
                            &[],
                            0.0,   // S749: band top unused (1-col)
                            true,  // S691: header context
                            false, // S726: header bottom differs
                            None,  // S755
                            None,  // S758
                            None,  // S-TWOSEG
                            false, // S835
                            0.0,
                            None,  // S900
                            None,  // S903
                            false, // S916
                            None,
                        );
                        lp.elements.extend(hdr_elements);
                    } else if let Block::Image(img) = block {
                        // S759 (2026-07-09): a FLOATING header image (wp:anchor,
                        // page-relative — uk_health_form's Ofsted logo) draws at
                        // its absolute page position on every page. Inline header
                        // images (position None) keep the existing height-only
                        // behavior (s755_header_bottom). Opt-out OXI_HDRFLOAT_DISABLE.
                        if img.position.is_none()
                            && !s1105_prev_text
                            && std::env::var("OXI_S1105_DISABLE").is_err()
                        {
                            // S1105 (2026-08-08, opt-out OXI_S1105_DISABLE): an
                            // INLINE header image is PAINTED, not just measured.
                            // S742 taught s755_header_bottom to add `img.height`
                            // (probeqhdrimg 0.7818 → PASS) but nothing ever drew
                            // it: the probe's 113×85pt crest is absent from the
                            // dump, and only the text paragraph below it renders.
                            // S759's arm handles the FLOATING case; this is the
                            // inline one. Images that FOLLOW a text-bearing
                            // paragraph are skipped — S1HDR already folded them
                            // into that paragraph's line height, so drawing them
                            // here would double-count the advance.
                            let image_x = if std::env::var("OXI_HEADER_IMAGE_LEADING").ok().as_deref() == Some("1") {
                                img.host_paragraph.as_ref().map_or(hdr_x, |host| {
                                    let left = host.style.indent_left.unwrap_or(0.0)
                                        + host.style.indent_first_line.unwrap_or(0.0);
                                    let right = host.style.indent_right.unwrap_or(0.0);
                                    let remaining = (hdr_width - left - right - img.width).max(0.0);
                                    hdr_x + left + match host.alignment {
                                        Alignment::Center => remaining * 0.5,
                                        Alignment::Right => remaining,
                                        _ => 0.0,
                                    }
                                })
                            } else { hdr_x };
                            lp.elements.push(LayoutElement::new(
                                image_x,
                                cy.cursor_y,
                                img.width,
                                img.height,
                                LayoutContent::Image {
                                    data: img.data.clone(),
                                    content_type: img.content_type.clone(),
                                    crop: img
                                        .crop
                                        .as_ref()
                                        .map(|c| (c.top, c.right, c.bottom, c.left)),
                                },
                            ));
                            cy.advance(img.height + self.header_inline_image_leading(img));
                        } else if img.position.is_some()
                            && std::env::var("OXI_HDRFLOAT_DISABLE").is_err()
                        {
                            // S1268b (2026-09-01, default ON, opt-out
                            // OXI_S1268B_DISABLE): a header float anchored
                            // relativeFrom="paragraph" references the HEADER's
                            // paragraph, not the body's top margin. The header
                            // has its own block list, so block_y_positions never
                            // carries the anchor and the old `&[]` fell through
                            // to page.margin.top -- 29.45pt too low on an
                            // 851tw header. Word's PDF, three full-page header
                            // backgrounds, image top = header_distance + offset,
                            // exact to 0.00pt in all three:
                            //   20f1ad3e 42.55-42.70 = -0.15 (Word -0.15)
                            //   05a78ecd 42.55-43.60 = -1.05 (Word -1.05)
                            //   1076f12a 42.55-47.20 = -4.65 (Word -4.65)
                            // The page-edge clamp S1268 removes used to slide
                            // these back to y=0, which LOOKED right for a
                            // full-page background and hid the real error.
                            let hdr_anchor_y = if std::env::var("OXI_S1268B_DISABLE").is_err() {
                                prev_block_top
                            } else {
                                page.margin.top
                            };
                            let (ax, ay) =
                                self.resolve_floating_image_position(img, page, &[], hdr_anchor_y);
                            let (mut paint_x, mut paint_y, mut paint_w, mut paint_h) =
                                (ax, ay, img.width, img.height);
                            let mut paint_crop = img.crop;
                            if std::env::var("OXI_POSITIONED_OLE").ok().as_deref() == Some("1") {
                                if let Some(crop) = paint_crop.as_mut() {
                                    // Negative crop values add empty space around the source.
                                    // Keep that space in flow, but paint only the source area.
                                    let width_scale = (100.0 - crop.left - crop.right).max(0.001);
                                    let height_scale = (100.0 - crop.top - crop.bottom).max(0.001);
                                    paint_x += img.width * (-crop.left).max(0.0) / width_scale;
                                    paint_y += img.height * (-crop.top).max(0.0) / height_scale;
                                    paint_w *= (100.0 - crop.left.max(0.0) - crop.right.max(0.0)).max(0.0) / width_scale;
                                    paint_h *= (100.0 - crop.top.max(0.0) - crop.bottom.max(0.0)).max(0.0) / height_scale;
                                    crop.left = crop.left.max(0.0); crop.right = crop.right.max(0.0);
                                    crop.top = crop.top.max(0.0); crop.bottom = crop.bottom.max(0.0);
                                }
                            }
                            if std::env::var("OXI_HEADER_FLOAT_BANDS").ok().as_deref() == Some("1") {
                                header_float_indices.push((img.relative_height, lp.elements.len()));
                            }
                            lp.elements.push(LayoutElement::new(
                                paint_x,
                                paint_y,
                                paint_w,
                                paint_h,
                                LayoutContent::Image {
                                    data: img.data.clone(),
                                    content_type: img.content_type.clone(),
                                    crop: paint_crop
                                        .as_ref()
                                        .map(|c| (c.top, c.right, c.bottom, c.left)),
                                },
                            ));
                        }
                    } else if let Block::Table(tbl) = block {
                        // S1104: the header's paint loop gets the same Table arm as
                        // the footer's (S731 taught s755_header_bottom to MEASURE a
                        // header table; nothing painted it). Render-only.
                        // ★FLOATING tables (tblpPr) are EXCLUDED: they do not advance
                        // the header/footer cursor — Word floats them beside the text
                        // (ja/reference/071997db's header stamp is
                        // `vertAnchor=text horzAnchor=margin tblpXSpec=right`, drawn at
                        // x=402 with the title still on the first line at x=70.9).
                        // Painting one through the inline path would push every
                        // following header paragraph down by the table's height.
                        if std::env::var("OXI_S1104_DISABLE").is_err() && tbl.style.position.is_none() {
                            let mut dummy_pages = Vec::new();
                            let mut dummy_elems = Vec::new();
                            let tbl_elements = self.layout_table(
                                tbl,
                                hdr_x,
                                &mut cy,
                                hdr_width,
                                grid_pitch,
                                None,
                                None,
                                header_y,
                                99999.0,
                                page.size.width,
                                99999.0,
                                &mut dummy_pages,
                                &mut dummy_elems,
                                None,
                                page,
                                false,
                                None,
                                None,
                                0.0,
                                0.0,
                                false,
                                None,
                            );
                            lp.elements.extend(tbl_elements);
                        }
                    }
                    if std::env::var("OXI_HEADER_FLOAT_BANDS").ok().as_deref() != Some("1")
                        || !matches!(block, Block::Image(image) if image.position.is_some())
                    {
                        prev_block_top = this_block_top;
                    }
                }
                // Order floating images by their explicit stacking levels while
                // retaining the existing positions of non-image paint operations.
                let slots: Vec<usize> = header_float_indices.iter().map(|(_, index)| *index).collect();
                header_float_indices.sort_by_key(|(level, _)| *level);
                let paints: Vec<LayoutElement> = header_float_indices.iter()
                    .map(|(_, index)| lp.elements[*index].clone()).collect();
                for (index, paint) in slots.into_iter().zip(paints) {
                    lp.elements[index] = paint;
                }
            }
            if !ftr_blocks.is_empty() {
                // Estimate footer content height first.
                // Day 33 part 18: skip framePr paragraphs (floating frames) —
                // they're positioned independently of inline flow, so they
                // should not shift footer_top.
                let image_flow = ftr_blocks.iter().any(|b| matches!(b, Block::Paragraph(p) if p.runs.iter().any(|r| r.style.inline_object_image.is_some())));
                let mut measured_previous = None;
                let mut measured_after = 0.0;
                let mut footer_h: f32 = 0.0;
                for block in ftr_blocks {
                    if let Block::Paragraph(para) = block {
                        if para.style.frame_pr.is_some() {
                            continue;
                        }
                        if image_flow {
                            let (height, after) = self.footer_paragraph_flow_height(para, measured_previous, measured_after, page, hdr_width);
                            footer_h += height;
                            measured_previous = Some(para);
                            measured_after = after;
                        } else {
                        footer_h += self.estimate_para_height(
                            para, hdr_width, grid_pitch, None, false, None, None,
                        );
                        }
                    }
                }
                if image_flow { footer_h += measured_after; }
                // S1104: this paint-side `footer_h` is Paragraph-only, so a footer
                // holding a TABLE placed its content far too low (the table's height
                // never entered footer_top). S868 already computes the full stack —
                // including the ROWBOX2 table term — inside s755_footer_geom, so take
                // the top from that reservation instead of re-deriving it here.
                // Scoped to footers that actually contain a table: a paragraph-only
                // footer keeps the historical expression byte-for-byte.
                let s1104_has_tbl = std::env::var("OXI_S1104_DISABLE").is_err()
                    && ftr_blocks
                    .iter()
                    .any(|b| matches!(b, Block::Table(t) if t.style.position.is_none()));
                let footer_top = if s1104_has_tbl {
                    let (fr, _) = self.s755_footer_geom(ftr_blocks, page);
                    page.size.height - fr
                } else {
                    page.size.height - footer_dist - footer_h
                };
                let mut cy = LayoutCursor::new(footer_top);
                let mut previous_para: Option<&Paragraph> = None;
                let mut previous_after = 0.0;
                for block in ftr_blocks {
                    if let Block::Paragraph(para) = block {
                        let empty_fn_h_ftr = std::collections::HashMap::new();
                        let (ftr_elements, after, _) = self.layout_paragraph(
                            para,
                            hdr_x,
                            &mut cy,
                            hdr_width,
                            page.size.height,
                            footer_top,
                            page,
                            &mut Vec::new(),
                            &mut Vec::new(),
                            grid_pitch,
                            previous_para.and_then(|p| p.style.style_id.as_deref()),
                            previous_para.is_some_and(|p| p.style.contextual_spacing),
                            previous_para.filter(|p| p.style.after_autospacing).and_then(|p| p.style.num_id.as_deref()),
                            None,
                            false,
                            false,
                            previous_after,
                            None,
                            None,
                            None,
                            false,
                            false,
                            None,
                            0.0,
                            &empty_fn_h_ftr,
                            1,
                            0,
                            &[],
                            0.0,   // S749: band top unused (1-col)
                            true,  // S691: footer context
                            false, // S726: footer bottom differs
                            None,  // S755
                            None,  // S758
                            None,  // S-TWOSEG
                            false, // S835
                            0.0,
                            None,  // S900
                            None,  // S903
                            false, // S916
                            None,
                        );
                        lp.elements.extend(ftr_elements);
                        if image_flow { previous_para = Some(para); previous_after = after; }
                    } else if let Block::Table(tbl) = block {
                        if image_flow { cy.advance(previous_after); previous_after = 0.0; previous_para = None; }
                        // S1104 (2026-08-08, default ON, opt-out OXI_S1104_DISABLE):
                        // a TABLE in the footer is PAINTED, not just measured. S868
                        // taught `s755_footer_geom` to count a footer table's height
                        // (and S731 did the same for headers), but BOTH paint loops
                        // stayed `if let Block::Paragraph` — so a footer whose text
                        // lives in a table rendered COMPLETELY EMPTY.
                        // reference__0061531a: every one of its 67 pages lost the
                        // whole footer (Word draws 5 spans/page, Oxi drew 0).
                        // Render-only: the height is already reserved by S868, so
                        // cursor_y / pagination are untouched. A huge content_height
                        // keeps the table from trying to split inside the footer.
                        if std::env::var("OXI_S1104_DISABLE").is_err() && tbl.style.position.is_none()
                        {
                            let mut dummy_pages = Vec::new();
                            let mut dummy_elems = Vec::new();
                            let tbl_elements = self.layout_table(
                                tbl,
                                hdr_x,
                                &mut cy,
                                hdr_width,
                                grid_pitch,
                                None,
                                None,
                                footer_top,
                                99999.0,
                                page.size.width,
                                99999.0,
                                &mut dummy_pages,
                                &mut dummy_elems,
                                None,
                                page,
                                false,
                                None,
                                None,
                                0.0,
                                0.0,
                                false,
                                None,
                            );
                            lp.elements.extend(tbl_elements);
                        }
                    }
                }
            }

            // Render footnotes for this layout page (Round 29, 2026-04-08).
            // Word places footnotes at the bottom of the page where their reference
            // appears, above the footer (or above the bottom margin if no footer).
            // We:
            //   1. Scan blocks belonging to this layout page for footnoteReference runs
            //      (paragraph runs and recursively into table cells)
            //   2. Look up each referenced footnote by id in `page.footnotes`
            //   3. Render a separator + each footnote paragraph at the footnote area top
            // For now we do NOT shrink body content_height to reserve footnote space —
            // body fitting drift remains a known limitation handled in a later round.
            if !page.footnotes.is_empty() {
                fn collect_footnote_refs(blocks: &[Block], out: &mut Vec<u32>) {
                    for b in blocks {
                        match b {
                            Block::Paragraph(p) => {
                                for r in &p.runs {
                                    if let Some(id) = r.footnote_ref {
                                        if !out.contains(&id) {
                                            out.push(id);
                                        }
                                    }
                                }
                            }
                            Block::Table(t) => {
                                for row in &t.rows {
                                    for cell in &row.cells {
                                        collect_footnote_refs(&cell.blocks, out);
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }

                // Step 1 partial (2026-04-22): per-line paragraph fn refs
                // (accurate across mid-para page breaks) + table cell fn refs
                // via block-level attribution (tables not yet instrumented to
                // return per-line data).
                let mut referenced_ids: Vec<u32> = Vec::new();
                if let Some(p_ids) = page_fn_refs.get(page_idx) {
                    for id in p_ids {
                        if !referenced_ids.contains(id) {
                            referenced_ids.push(*id);
                        }
                    }
                }
                for (i, b) in page.blocks.iter().enumerate() {
                    if block_page_indices.get(i).copied().unwrap_or(0) == page_idx {
                        if let Block::Table(_) = b {
                            // S740: tables whose cell footnotes were attributed
                            // per-page already live in page_fn_refs — skip the
                            // coarse "all refs on the block's page" fallback.
                            if !s740_attributed_tables.contains(&i) {
                                collect_footnote_refs(std::slice::from_ref(b), &mut referenced_ids);
                            }
                        }
                    }
                }

                if !referenced_ids.is_empty() {
                    // Resolve referenced footnotes (preserve order, dedup already done).
                    let notes: Vec<&Footnote> = referenced_ids
                        .iter()
                        .filter_map(|id| page.footnotes.iter().find(|n| n.number == *id))
                        .collect();

                    if !notes.is_empty() {
                        // SG0RAW footnote scope-out: the whole footnote-area
                        // placement (grid_snap_para estimates + the note
                        // layout_paragraph render calls) uses the footnote
                        // conventions, not the body sg0 raw-natural.
                        let _fng = FnLayoutGuard::new();
                        // Footnote area bottom: just above the footer, or at the
                        // bottom margin if no footer is present.
                        // S896 (2026-07-17, default ON, opt-out OXI_S896_DISABLE):
                        // an INK-FREE footer doesn't lower the footnote area either
                        // — the S894 exemption mirrored onto this legacy footer
                        // recompute (which predates s755_footer_geom and never saw
                        // S806/S894). legal__00081e80: 3 blank footers → this path
                        // put fn_bot at 792−39.6−15.9−4 = 732.5 vs Word's 720.6 =
                        // pageH − bottom margin EXACT (rt.pdf: notes 591.22 +
                        // 14×9.24 = 720.6; margin 1440tw). Latin scope with S894.
                        let s896_blank_footer = !page.footer.is_empty()
                            && !self.doc_body_has_real_cjk
                            && std::env::var("OXI_S896_DISABLE").is_err()
                            && !page.footer.iter().any(|b| match b {
                                Block::Paragraph(p) => {
                                    p.runs.iter().any(|r| !r.text.trim().is_empty())
                                }
                                Block::Table(_) | Block::Image(_) => true,
                                _ => false,
                            });
                        let footer_image_flow = ftr_blocks.iter().any(|b| matches!(b, Block::Paragraph(p) if p.runs.iter().any(|r| r.style.inline_object_image.is_some())));
                        // Footnotes share the body's usable bottom. A footer below
                        // the bottom margin does not move the footnote stack down;
                        // a footer intruding into the body reserves its full flow.
                        let footnote_bottom = if std::env::var_os("OXI_FOOTNOTE_BODY_BOTTOM_DISABLE").is_none() {
                            page.size.height - self.s755_footer_geom(ftr_blocks, page).0
                        } else if footer_image_flow {
                            page.size.height - self.s755_footer_geom(ftr_blocks, page).0
                        } else if !page.footer.is_empty() && !s896_blank_footer {
                            // Recompute footer top here (mirrors lines 832-839 above).
                            // Day 33 part 18: skip framePr paragraphs (floating frames).
                            let mut footer_h: f32 = 0.0;
                            for block in ftr_blocks {
                                if let Block::Paragraph(para) = block {
                                    if para.style.frame_pr.is_some() {
                                        continue;
                                    }
                                    footer_h += self.estimate_para_height(
                                        para, hdr_width, grid_pitch, None, false, None, None,
                                    );
                                }
                            }
                            page.size.height - footer_dist - footer_h - 4.0
                        } else {
                            page.size.height - page.margin.bottom
                        };

                        let footnote_bottom = footnote_bottom - legacy_notice_height();

                        // Find the last body element Y on this page to avoid overlap.
                        // S596 (2026-06-17): exclude FOOTER-region elements (those at or
                        // below footnote_bottom — e.g. the footer page-number). For a
                        // no-docGrid doc (bunkacontract) the footer "N" glyph sits in
                        // lp.elements at y≈757 (below footnote_bottom 753) and was taken
                        // as body_bottom_y=774.8 → the footnote area (top ≈ footnote_bottom)
                        // always overlapped it → fit=0, footnotes never rendered. Grid
                        // docs (b837/kojin) were unaffected (their footer is not in
                        // lp.elements here / body is already above the area), so this is
                        // a no-op for them. Opt-out OXI_S596_DISABLE.
                        let body_bottom_y = if std::env::var("OXI_S596_DISABLE").is_ok() {
                            lp.elements
                                .iter()
                                .map(|e| e.y + e.height)
                                .fold(0.0_f32, f32::max)
                        } else {
                            lp.elements
                                .iter()
                                .filter(|e| e.y < footnote_bottom)
                                .map(|e| {
                                    let height = if !self.doc_body_has_real_cjk
                                        && std::env::var("OXI_FN_PLACEMENT_FIT_DISABLE").is_err()
                                    {
                                        e.content_fit_height.unwrap_or(e.height)
                                    } else {
                                        e.height
                                    };
                                    e.y + height
                                })
                                .fold(0.0_f32, f32::max)
                        };

                        // Calculate footnote heights — grid-snap per paragraph to
                        // match actual render (layout_paragraph stacks lines at
                        // grid pitch when grid_pitch is Some). estimate_para_height
                        // returns natural height which diverges from the render.
                        // COM-derived 2026-04-20 from 6 minimal repros: render uses
                        // grid_pitch × line_count per paragraph.
                        // S828 (2026-07-13, opt-out OXI_S828_DISABLE): a NO-TYPE
                        // docGrid does not snap footnote lines — the ESTIMATE
                        // already excluded it (fn_est_gp, S727 "No-type grids
                        // don't snap") but the PLACEMENT (this fn) and the note
                        // RENDER below kept the raw pitch → est 11.5/line vs
                        // place/render 14.95/line (nyserda pitch=299): the body
                        // under-reserved, the area over-computed, [FN_PLACE]
                        // fit=0 → footnote bodies silently DROPPED. Word
                        // render-truth (nyserda p18 fn 1): line box 11.4 = hhea
                        // natural, NOT the pitch. est == place == render.
                        // Corpus scan: no-type grid + footnoteReference = the 4
                        // EN docs only → JP byte-identical by construction.
                        let s828 = std::env::var("OXI_S828_DISABLE").is_err();
                        let fn_gp = if s828 && page.doc_grid_no_type {
                            None
                        } else {
                            grid_pitch
                        };
                        let grid_snap_para = |p: &Paragraph| -> (f32, usize) {
                            // Natural estimated height (may include space_before/after)
                            let nat = self
                                .estimate_para_height(p, hdr_width, fn_gp, None, false, None, None);
                            let mut line_para = p.clone();
                            line_para.style.space_before = Some(0.0);
                            line_para.style.space_after = Some(0.0);
                            line_para.style.before_lines = None;
                            line_para.style.after_lines = None;
                            let line_height = self.estimate_para_height(
                                &line_para, hdr_width, fn_gp, None, false, None, None);
                            let paragraph_spacing = nat - line_height;
                            // Per-line natural height (used to derive line_count).
                            // S828(b): the first run is the SUPERSCRIPT ref mark
                            // (auto-shrunk 2/3 by resolve_font_size) — keying the
                            // per-line height off it under-sizes line_nat →
                            // line_count inflates (nyserda fn 1: round(nat/8.6)=2
                            // for a 1-line URL). Use the first non-superscript
                            // text run for fs/metrics (Latin scope; JP typed-grid
                            // footnote estimates keep their calibration).
                            let fs_run = if s828 && !self.doc_body_has_real_cjk {
                                p.runs
                                    .iter()
                                    .find(|r| {
                                        !r.text.trim().is_empty()
                                            && !matches!(
                                                r.style.vertical_align,
                                                Some(VerticalAlign::Superscript)
                                                    | Some(VerticalAlign::Subscript)
                                            )
                                    })
                                    .or_else(|| p.runs.first())
                            } else {
                                p.runs.first()
                            };
                            let line_fs = self.resolve_font_size(
                                fs_run.map(|r| &r.style).unwrap_or(&RunStyle::default()),
                                &p.style,
                            );
                            let metrics = fs_run
                                .map(|r| self.metrics_for_text(&r.text, &r.style, &p.style))
                                .unwrap_or_else(|| {
                                    let rpr = p.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                    self.metrics_for_para_mark(&rpr, &p.style)
                                });
                            let line_nat = metrics.word_line_height_no_grid(line_fs).max(0.01);
                            let line_count = ((line_height / line_nat).round() as usize).max(1);
                            // Only grid-snap if the paragraph opts in (snapToGrid default=true).
                            // b837's FootnoteText style has snapToGrid=0 → Word uses natural.
                            // S808 render mirror: Latin fn lines = hhea natural.
                            // S810: auto-rule lines only (exact keeps its box).
                            let s808_line = if !self.doc_body_has_real_cjk
                                && std::env::var("OXI_S808_DISABLE").is_err()
                                && matches!(
                                    p.style.line_spacing_rule.as_deref(),
                                    None | Some("auto")
                                ) {
                                metrics.natural_line_height_hhea(line_fs).max(line_nat)
                            } else {
                                line_nat
                            };
                            let height = if let Some(pitch) = fn_gp {
                                if pitch > 0.0 && p.style.snap_to_grid {
                                    line_count as f32 * pitch
                                } else {
                                    line_count as f32 * s808_line
                                }
                            } else {
                                line_count as f32 * s808_line
                            };
                            (height + paragraph_spacing, line_count)
                        };
                        let mut note_heights: Vec<f32> = Vec::new();
                        let s804_r = std::env::var("OXI_S804_DISABLE").is_err();
                        for note in &notes {
                            let mut nh: f32 = 0.0;
                            // S804 render mirror: the fit/area math must include
                            // the same style spacing the reservation now counts
                            // (Fix C estimate==render invariant; without it the
                            // rendered notes overflowed fn_bot by the spacing).
                            let mut prev_sa: Option<f32> = None;
                            let mut s807_first = true;
                            for nb in &note.blocks {
                                if let Block::Paragraph(p) = nb {
                                    let (h, _) = grid_snap_para(p);
                                    nh += h;
                                    // S807 render mirror (estimate==render).
                                    // S810: exact-rule box clamps the raise.
                                    if s807_first {
                                        s807_first = false;
                                        // S807 retired to opt-in (see estimate site).
                                        if !self.doc_body_has_real_cjk
                                            && std::env::var("OXI_S807").is_ok()
                                            && p.style.line_spacing_rule.as_deref() != Some("exact")
                                        {
                                            let rs = p
                                                .runs
                                                .iter()
                                                .find(|r| !r.text.trim().is_empty())
                                                .map(|r| &r.style)
                                                .cloned()
                                                .unwrap_or_default();
                                            let fs = self.resolve_font_size(&rs, &p.style);
                                            nh += (0.35 * fs * 2.0).round() / 2.0;
                                        }
                                    }
                                    if s804_r && self.footnote_twip_spacing_supported(&p.style) {
                                        nh += self.footnote_twip_spacing_correction(
                                            &p.style, &mut prev_sa);
                                    } else if s804_r && !p.style.has_direct_spacing {
                                        let sb = p.style.space_before.unwrap_or(0.0);
                                        let sa = p.style.space_after.unwrap_or(0.0);
                                        if let Some(prev) = prev_sa {
                                            nh += prev.max(sb);
                                        }
                                        prev_sa = Some(sa);
                                        // S810 strip (see estimate_footnote_h).
                                        if !matches!(
                                            p.style.line_spacing_rule.as_deref(),
                                            None | Some("auto")
                                        ) {
                                            nh -= sb + sa;
                                        }
                                    } else {
                                        prev_sa = Some(0.0);
                                    }
                                }
                            }
                            if s804_r {
                                if let Some(last) = prev_sa {
                                    nh += last;
                                }
                            }
                            note_heights.push(nh);
                        }
                        // Word anchors the LAST footnote line's BOTTOM to
                        // page_h - margin.bottom (derived 2026-04-20 from 6 minimal
                        // repros). Only INNER lines stack at grid pitch; the last
                        // line's height is natural (word_line_height_no_grid).
                        // Subtract (grid_pitch - natural_last) once to compensate.
                        let last_line_adjust: f32 = if let (Some(pitch), Some(last_note)) =
                            (grid_pitch, notes.last())
                        {
                            if pitch > 0.0 {
                                if let Some(Block::Paragraph(last_p)) = last_note.blocks.last() {
                                    // Only applicable when the last paragraph grid-snaps.
                                    // snapToGrid=0 footnote lines are already natural-height.
                                    if !last_p.style.snap_to_grid {
                                        0.0
                                    } else {
                                        let text_run = last_p
                                            .runs
                                            .iter()
                                            .rev()
                                            .find(|r| !r.text.is_empty())
                                            .or_else(|| last_p.runs.first());
                                        let fs = self.resolve_font_size(
                                            text_run
                                                .map(|r| &r.style)
                                                .unwrap_or(&RunStyle::default()),
                                            &last_p.style,
                                        );
                                        let metrics = text_run
                                            .map(|r| {
                                                self.metrics_for_text(
                                                    &r.text,
                                                    &r.style,
                                                    &last_p.style,
                                                )
                                            })
                                            .unwrap_or_else(|| {
                                                let rpr = last_p
                                                    .style
                                                    .ppr_rpr
                                                    .as_ref()
                                                    .cloned()
                                                    .unwrap_or_default();
                                                self.metrics_for_para_mark(&rpr, &last_p.style)
                                            });
                                        let natural_last = metrics.word_line_height_no_grid(fs);
                                        // Word centers the last footnote line in the grid-pitch
                                        // overshoot: adjustment = (pitch - natural) / 2 — confirmed
                                        // 2026-04-20 via 6 minimal repros (fn top at page_h -
                                        // margin.bottom - 14.85pt = pitch - 3.15 for 9pt MS Mincho).
                                        ((pitch - natural_last) * 0.5).max(0.0)
                                    }
                                } else {
                                    0.0
                                }
                            } else {
                                0.0
                            }
                        } else {
                            0.0
                        };

                        let separator_h_pre: f32 = 2.0;
                        let separator_pad_pre: f32 = 4.0;
                        // Determine how many notes fit: add notes one by one from the
                        // bottom; stop when area_top would overlap body content.
                        let mut total_h: f32 = 0.0;
                        let mut fit_count = notes.len();
                        for i in 0..notes.len() {
                            let candidate =
                                total_h + note_heights[i] + separator_h_pre + separator_pad_pre;
                            let candidate_top = footnote_bottom - candidate;
                            if candidate_top < body_bottom_y + 2.0 {
                                // This note doesn't fit; truncate here
                                fit_count = i;
                                break;
                            }
                            total_h += note_heights[i];
                        }
                        if std::env::var("OXI_FN_PROBE").is_ok() {
                            eprintln!("[FN_PLACE] page_idx={} n_req={} fit={} body_bot={:.1} fn_bot={:.1} total_h={:.1} area_top={:.1} gap={:.1} heights={:?}",
                                page_idx, notes.len(), fit_count, body_bottom_y, footnote_bottom, total_h,
                                footnote_bottom - total_h - separator_pad_pre - separator_h_pre,
                                (footnote_bottom - total_h - separator_pad_pre - separator_h_pre) - body_bottom_y,
                                note_heights);
                        }
                        let notes: Vec<&Footnote> = notes[..fit_count].to_vec();
                        // Separator: short horizontal line above the footnotes.
                        let separator_h: f32 = 2.0;
                        let separator_pad: f32 = 4.0;
                        let area_top = footnote_bottom - total_h - separator_pad - separator_h
                            + last_line_adjust;

                        // Draw the footnote separator line. Word's default for a
                        // bare <w:separator/> is a fixed 2-inch (144pt) line at the
                        // left margin — NOT a fraction of content width (S479,
                        // pixel-confirmed on b837 p1: 144.0pt at x0=71pt). Cap at
                        // content width for narrow columns. 1pt thick, black.
                        // Default ON, opt-out OXI_S479_DISABLE.
                        let sep_w = if std::env::var("OXI_S479_DISABLE").is_ok() {
                            hdr_width * 0.33
                        } else {
                            144.0_f32.min(hdr_width)
                        };
                        lp.elements.push(LayoutElement::new(
                            hdr_x,
                            area_top,
                            sep_w,
                            1.0,
                            LayoutContent::BoxRect {
                                fill: Some("#000000".to_string()),
                                stroke_color: None,
                                stroke_width: 0.0,
                                corner_radius: 0.0,
                            },
                        ));

                        // Lay out each footnote's body paragraphs from area_top
                        // downward. CRITICAL: pass a huge content_height so the
                        // page-break logic inside layout_paragraph never fires —
                        // otherwise overflow would push a fake "page" and reset
                        // cy back to footnote_page_top, causing all footnotes to
                        // stack at the same Y (visible as overlapping notes).
                        let mut cy = LayoutCursor::new(area_top + separator_h + separator_pad);
                        let footnote_page_top = cy.cursor_y;
                        let footnote_page_height_huge = 1e6_f32;
                        for note in &notes {
                            // Round 29: section-local sequential number (Word
                            // displays footnotes as 1,2,3... regardless of OOXML
                            // ids). page.footnotes is sorted by id; the seq is
                            // the index + 1.
                            let seq = page
                                .footnotes
                                .iter()
                                .position(|n| n.number == note.number)
                                .map(|p| (p as u32) + 1)
                                .unwrap_or(note.number);
                            let mut first_para = true;
                            for nb in &note.blocks {
                                if let Block::Paragraph(para) = nb {
                                    // Prefix the FIRST paragraph of each note
                                    // with the seq number to identify it
                                    // visually. Use a clone to keep the IR
                                    // immutable.
                                    let para_to_render: Paragraph = if first_para {
                                        let mut p = para.clone();
                                        // Round 29: just the seq number, NO trailing space.
                                        // Word's footnote body has its own leading space run
                                        // (which renders as the separator between marker and
                                        // text). Adding another space here yields a double
                                        // space "1  震災..." which compresses content area.
                                        let prefix = format!("{}", seq);
                                        if let Some(first_run) = p.runs.first_mut() {
                                            // First run is usually <w:footnoteRef/> with empty
                                            // text. OVERWRITE it with the seq, don't prepend.
                                            if first_run.text.is_empty() {
                                                first_run.text = prefix.clone();
                                            } else {
                                                first_run.text =
                                                    format!("{}{}", prefix, first_run.text);
                                            }
                                        } else {
                                            // Empty paragraph: insert a run with just the prefix
                                            p.runs.push(Run {
                                                text: prefix,
                                                style: RunStyle::default(),
                                                url: None,
                                                footnote_ref: None,
                                                endnote_ref: None,
                                                comment_range_start: Vec::new(),
                                                comment_range_end: Vec::new(),
                                                comment_references: Vec::new(),
                                                tracked_change: None,
                                                rpr_change: None,
                                                ruby: None,
                                                bookmark_name: None,
                                                is_math: false,
                                                field_type: None,
                                                has_last_rendered_page_break: false,
                                            });
                                        }
                                        first_para = false;
                                        p
                                    } else {
                                        para.clone()
                                    };
                                    // Round 29: use total_content_width (full
                                    // body width) explicitly. content_width may
                                    // have been mutated by the body loop column
                                    // switching state and the residual value
                                    // can be smaller than the full body area.
                                    let footnote_width =
                                        page.size.width - page.margin.left - page.margin.right;
                                    let empty_fn_h_note = std::collections::HashMap::new();
                                    let (note_elements, _, _) = self.layout_paragraph(
                                        &para_to_render,
                                        page.margin.left,
                                        &mut cy,
                                        footnote_width,
                                        footnote_page_height_huge,
                                        footnote_page_top,
                                        page,
                                        &mut Vec::new(),
                                        &mut Vec::new(),
                                        // S828: no-type grids render footnote lines at
                                        // natural hhea, not the pitch (est==place==render).
                                        fn_gp,
                                        None,
                                        false,
                                        None,
                                        None,
                                        false,
                                        false,
                                        0.0,
                                        None,
                                        None,
                                        None,
                                        false,
                                        false,
                                        None,
                                        0.0,
                                        &empty_fn_h_note,
                                        1,
                                        0,
                                        &[],
                                        0.0,   // S749: band top unused (1-col)
                                        false, // S691: footnote context
                                        false, // S726
                                        None,  // S755
                                        None,  // S758
                                        None,  // S-TWOSEG
                                        false, // S835
                                        0.0,
                                        None,  // S900
                                        None,  // S903
                                        false, // S916
                                        None,
                                    );
                                    lp.elements.extend(note_elements);
                                }
                            }
                        }
                    }
                }
            }

            // Render shapes (e.g. bracketPair) positioned relative to anchor paragraph
            for shape in &page.shapes {
                if let Some(ref pos) = shape.position {
                    // Get anchor paragraph's Y position and page index
                    let anchor_y = block_y_positions
                        .get(shape.anchor_block_index)
                        .copied()
                        .unwrap_or(start_y);
                    let anchor_page = block_page_indices
                        .get(shape.anchor_block_index)
                        .copied()
                        .unwrap_or(0);

                    // Only render on the correct page
                    if anchor_page == page_idx {
                        // h_relative="column": x = margin_left + offset
                        // v_relative="paragraph": y = anchor_paragraph_y + offset
                        let sx = start_x + pos.x;
                        let sy = anchor_y + pos.y;
                        let content = shape_fill_boxrect(shape).unwrap_or_else(|| {
                            LayoutContent::PresetShape {
                                shape_type: shape.shape_type.clone(),
                                stroke_color: shape.stroke_color.clone(),
                                stroke_width: shape.stroke_width.unwrap_or(0.75),
                                flip_h: shape.flip_h,
                                flip_v: shape.flip_v,
                                arrow_head: shape.arrow_head,
                                arrow_tail: shape.arrow_tail,
                            }
                        });
                        lp.elements.push(LayoutElement::new(
                            sx,
                            sy,
                            shape.width,
                            shape.height,
                            content,
                        ));
                    }
                }
            }
        }

        // S1073: hand this section's trailing space-after to the next one.
        // Taken from the IR, not from the running `prev_space_after`: that value
        // is zeroed by the table / image / push arms, whereas Word collapses the
        // next section's first paragraph against this section's LAST PARAGRAPH
        // (uk_local_spending's Annex II boundary ends in an empty Heading2 with
        // after=120tw and Word's excess there is 12 - 6 = 6pt).
        let s1073_tail = page
            .blocks
            .iter()
            .rev()
            .find_map(|b| match b {
                Block::Paragraph(p) => Some(p.style.space_after.unwrap_or(0.0)),
                _ => None,
            })
            .unwrap_or(prev_space_after);
        if std::env::var("OXI_DBG1073").is_ok() {
            eprintln!(
                "[S1073-STORE] running={:.3} ir_tail={:.3}",
                prev_space_after, s1073_tail
            );
        }
        S1073_PREV_SECTION_AFTER.with(|c| c.set(s1073_tail));

        // S1294: hand the caller the logical number of the LAST page emitted.
        // `logical_base` has already absorbed any mid-page restart, so this is
        // the same walk S912's post-pass does, just carried forward live.
        if !pages.is_empty() {
            *logical = (logical_base + pages.len() as i64 - 1).max(0) as u32;
        }

        pages
    }
}
