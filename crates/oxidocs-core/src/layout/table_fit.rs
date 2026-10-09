// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `LayoutEngine::layout_table_with_fit_pass` -- moved out of `layout/mod.rs` so that it is its own
//! codegen unit (see tools/metrics/split_layout_mod.py). Behaviour-preserving.

use super::*;

/// `layout_table_with_fit_pass` as a method of its own type: rustc puts a method's code in the
/// codegen unit of its self type's module, so this (not the file move alone)
/// is what gives the giant its own unit. Deref keeps `self.x` meaning the engine.
pub(super) struct TableFitLayouter<'a>(pub(super) &'a LayoutEngine);

impl<'a> std::ops::Deref for TableFitLayouter<'a> {
    type Target = LayoutEngine;
    fn deref(&self) -> &LayoutEngine {
        self.0
    }
}

impl<'a> TableFitLayouter<'a> {
    pub(super) fn layout_table_with_fit_pass(
        &self,
        table: &Table,
        start_x: f32,
        cursor: &mut LayoutCursor,
        content_width: f32,
        grid_pitch: Option<f32>,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
        mut page_top: f32,
        mut content_height: f32,
        page_width: f32,
        page_height: f32,
        pages: &mut Vec<LayoutPage>,
        current_elements: &mut Vec<LayoutElement>,
        block_idx: Option<usize>,
        page: &Page,
        is_nested: bool,
        // S740 (2026-07-04): footnote refs INSIDE table cells. Per-row
        // (ids, reserve_height) precomputed by the body caller; None (all other
        // call sites / tables without cell footnotes) = byte-identical.
        row_footnotes: Option<&[(Vec<u32>, f32, Vec<f32>)]>,
        // Per-page (offset from the table's entry page) footnote ids placed by
        // this table's rows — the caller merges into page_fn_refs so the
        // footnote-area renderer draws each note on the page of its row.
        mut fn_pages_out: Option<&mut Vec<Vec<u32>>>,
        // Separator allocation for a page's FIRST note + the entry page's
        // already-committed body reserve (page 0 subtracts it too).
        fn_sep: f32,
        fn_entry_reserve: f32,
        fn_entry_has_notes: bool,
        page_geometry: Option<&S755Geom>,
        flow_fit_offset: Option<f32>,
        float_replay: Option<&CellFloatReplay>,
    ) -> Vec<LayoutElement> {
        let flow_entry_page = pages.len();
        let _s1429_guard = TableLayoutGuard::new();
        self.s1612_fill_above(table, block_idx, page);
        if std::env::var("OXI_DBG_TBLSTART").is_ok() {
            let head: String = table.rows.first().and_then(|r| r.cells.first()).map(|c| c.blocks.iter().filter_map(|b| match b {
                Block::Paragraph(p) => Some(p.runs.iter().flat_map(|r| r.text.chars()).take(12).collect::<String>()),
                _ => None }).next().unwrap_or_default()).unwrap_or_default();
            eprintln!("[TBLSTART] pages={} cursor_y={:.2} visual_y={:.2} fit_offset={:?} nested={} head={:?}",
                pages.len(), cursor.cursor_y, cursor.visual_y, flow_fit_offset, is_nested, head);
        }
        let page_geometry = page_geometry.filter(|_| {
            std::env::var("OXI_TABLE_PAGE_GEOMETRY_DISABLE").is_err()
                // Repeated heading rows also depend on continuation sizing.
                // Keep their current geometry until header replay and that
                // sizing can be corrected together.
                && (!table.rows.iter().any(|row| row.header)
                    || std::env::var_os("OXI_HEADER_TABLE_GEOMETRY_DISABLE").is_none())
        });
        let mut elements = Vec::new();
        // S740 running state: reserve on the CURRENT page + page-offset tracking.
        let mut s740_reserve: f32 = fn_entry_reserve;
        let mut s740_page_has_notes: bool = fn_entry_has_notes;
        let mut s740_pages_len: usize = pages.len();
        let mut s740_fn_pages: Vec<Vec<u32>> = vec![Vec::new()];
        let mut s740_pending_commit: Option<usize> = None;
        // S1527 (2026-09-24): note ids a row SPLIT already placed on the page of
        // their referencing line (see the split site); the row-end commit skips
        // them so they are neither re-listed nor re-reserved on the next page.
        let mut s1527_early: Vec<u32> = Vec::new();
        // S1527: the table's entry page; `s740_fn_pages[k]` is the page k pages
        // after it, so a continuation page can be addressed while the row is
        // still being split (before the next row start syncs the list).
        let s740_entry_pages: usize = pages.len();

        // Resolve column widths from grid_columns, cell widths, or equal split
        // S1003: thread is_nested so a DIRECT BODY autofit table waterfills all
        // columns while a NESTED overflow keeps the last-column-only clamp.
        let col_widths = self.resolve_table_col_widths_n(table, content_width, is_nested);
        let table_width: f32 = col_widths.iter().sum();

        // Table positioning: tblpPr horizontal or inline alignment
        // S1158 (2026-08-17, default ON, opt-out OXI_S1158_DISABLE): a FLOATING
        // table absorbs the leading cell margin exactly like a tblInd one. S621
        // established the absorption but only along the non-positioned path
        // below, so a `w:tblpPr` table keeps its BORDER on the margin and pushes
        // its text in by cellMar. `_pb_tblanchor_gen.py`, 10 arms x compat
        // {11, 15}, Word PDF, page margin 85.05:
        //   compat 11  float cellMar 0/108/200/400 -> border 84.86 / 79.70 /
        //              75.14 / 65.06 = margin - cellMar, text 85.10 = the margin
        //              float tblpX=567             -> border 108.02, text 113.42
        //   compat 15  every arm                   -> border 85.34, no absorption
        // Oxi gives 85.05 for all six compat-11 float arms. tokyoshugyo (compat
        // 11, horzAnchor=margin, no tblpX, default cellMar) is exactly arm 1:
        // Word 79.70/85.10 vs Oxi 85.05/90.50, which costs its p17 block one
        // character per line and hands p18-19 the +21pt.
        // ★Also measured, NOT yet changed: with compat 11 and NO tblInd element
        // at all Word does NOT absorb (border stays on the margin) while Oxi
        // does -- S621 reads `indent.map_or(true, ..)`, i.e. it treats ABSENT as
        // zero. Word's rule is "tblInd PRESENT (any value) or floating". Fixing
        // that is a second, separate change: the ind567 arm shows the absorption
        // for a non-zero tblInd already arrives through another path, so the two
        // must not be stacked blind.
        // Gate: probe 6/6 float arms match Word, Phase 1 95/96 with zero per-doc
        // change, all 238 SSIM sentinel documents byte-identical (no compat-<=14
        // floating table among them), and tokyoshugyo -- scored against its own
        // Word PDF -- 0.8575 -> 0.8606 with p18 +0.1639, p19 +0.0559, p17 +0.0415
        // and nothing regressed. Its p18-19 +21pt departure is gone.
        let s1158_float_absorb = std::env::var("OXI_S1158_DISABLE").is_err()
            && self.compat_mode <= 14
            && table.style.position.is_some();
        // The shift is the cell margin ALONE -- no border/2. Word's tblpX=567 arm
        // lands at 113.42 - 5.40 = 108.02 and the non-floating tblInd=567 arm at
        // the same 108.02, so both paths absorb exactly cellMar; adding half the
        // border stroke put every float arm 0.25pt left of Word.
        let s1158_shift = if s1158_float_absorb {
            table
                .style
                .default_cell_margins
                .as_ref()
                .and_then(|m| m.left)
                .unwrap_or(4.95)
        } else {
            0.0
        };
        let table_x = s1158_shift.mul_add(-1.0, if let Some(ref pos) = table.style.position {
            if let Some(ref h_align) = pos.h_align {
                let (ref_left, ref_width) = match pos.h_anchor.as_deref() {
                    Some("page") => (0.0, page_width),
                    _ => (start_x, content_width), // "margin" or "text"
                };
                match h_align.as_str() {
                    "center" => ref_left + (ref_width - table_width) / 2.0,
                    "right" => ref_left + ref_width - table_width,
                    _ => ref_left,
                }
            } else {
                match pos.h_anchor.as_deref() {
                    Some("page") => pos.x,
                    _ => start_x + pos.x,
                }
            }
        } else {
            // COM-confirmed (2026-04-13, gen2_052): Word positions the table border
            // at margin - padding - border/2. The cell text then starts at
            // border_x + padding = margin - border/2, matching Word's COM output.
            // COM-confirmed (2026-04-13, gen2_052): Word positions the left-aligned
            // table border at margin - padding - border/2. Only apply when no
            // explicit indent is set (indent=0 means default positioning).
            let pad_l_default = table
                .style
                .default_cell_margins
                .as_ref()
                .and_then(|m| m.left)
                .unwrap_or(4.95);
            // S494b tblInd cell-margin absorption (env-gated OFF, opt-in OXI_S494B_TBLIND_ENABLE).
            // The leading-edge spec is COM-confirmed by repros (tblind_multi/cellmar/noborder/
            // layout/gridbefore/nested: a TOP-LEVEL table's leading cell text lands at
            // margin + tblInd, absorbing the cell left margin). But applying it as a whole-table
            // table_x shift is NET-NEGATIVE on the corpus per the per-glyph gate: 04b88e +0.0168
            // and 34140b +0.0112, yet 15076df −0.0223 AND 2ea81a −0.0170 (both bottom-N). The
            // nested-scope (is_nested below) was necessary but NOT sufficient — per-element
            // localization showed 15076df's residual regression is its TOP-LEVEL multi-col table
            // (tbl0): the absorption improves its MEAN offset (+0.93→+0.33) but more glyphs land
            // FARTHER from Word (482 vs 321) — i.e. a uniform +0.93 has better pixel overlap than
            // the variance-spread +0.33, so Word absorbs LESS than the full cell margin for
            // content that isn't at the literal leading edge. The whole-table shift over-applies
            // it. Kept OFF until the absorption is modeled per-cell (leading-edge only), not as a
            // table_x translate. The is_nested gate is retained (nested tables never absorb).
            // Always legacy: the tblInd absorption is now applied PER-CELL (only the
            // leading-edge column cell shifts left by its margin), NOT as a table_x translate.
            // See the cell loop below (OXI_S494B_TBLIND_ENABLE). A whole-table table_x shift
            // moved the BORDERS too, which regressed border-visible docs (15076df).
            // S621 (2026-06-19): the leading-cell margin absorption (table border
            // outsets left by cellMar so the cell CONTENT aligns with the text margin)
            // is a Word **compatibilityMode ≤ 14** behavior (S496-confirmed), NOT a
            // border-visibility one. The old gate used `!explicit_borders` as a proxy,
            // which WRONGLY excluded gen2 (mode 14, VISIBLE borders, no tblInd): word_png
            // measured all 5 table borders uniformly +5.28pt right of Word = exactly the
            // default cell left margin. FIX: gate on compat_mode ≤ 14 && indent≈0 (incl.
            // tblInd absent = None), and shift the WHOLE table (border_offset > 0 →
            // table_x -= cellMar+border/2). Mode ≥ 15 still does NOT shift (3a4f etc.).
            // Opt-out OXI_S621_DISABLE restores the old !explicit_borders gate.
            let border_offset = {
                let border_w = table.style.border_width.unwrap_or(0.5);
                // S1160 (2026-08-17, opt-out OXI_S1160_DISABLE): the absorption
                // needs tblInd to be PRESENT, not merely zero-or-absent. Word
                // (probe _pb_tblanchor, compat 11) leaves a table with NO tblInd
                // anywhere sitting on the margin, while Oxi pulled it left by the
                // whole cell margin (20pt on the cellMar=400 arm). S621 read
                // absent as zero because the style-inherited tblInd was being
                // dropped at parse time -- gen2's TableGrid does carry
                // <w:tblInd w:w="0"/>, so with that now inherited the two halves
                // agree and gen2 keeps its absorption.
                // (S1162 ATTEMPTED + REVERTED 2026-08-17: absorbing here for ANY
                // tblInd value double-shifts, because the non-zero case already
                // absorbs through S496's per-cell `lead_absorb` below. The real
                // defect is the WIDTH, see S1163 there.)
                let indent_zero = if std::env::var("OXI_S1160_DISABLE").is_err() {
                    table.style.indent.is_some_and(|v| v.abs() < 0.01)
                } else {
                    table.style.indent.map_or(true, |v| v.abs() < 0.01)
                };
                // S1239 (2026-08-27, default ON, opt-out OXI_S1239_DISABLE): a
                // NEGATIVE tblInd absorbs the cell margin like the zero (S621)
                // and positive (S496) cases — the legacy law is ONE formula,
                // grid_x = margin + tblInd − cellMar.left, whole table with
                // borders. COM probe (_pb_tblneg, no-compat, style cellMar
                // 108): ind −162 → 21.72, −500 → 4.92, 0 → 29.88, +162 →
                // 38.04, and −162 with a DIRECT cellMar 72 absorbs its own
                // 3.6 (23.52) — all = margin + ind − cellMar ± 0.1. The cm15
                // twin absorbs nothing on any arm. Witness technical__0009d767
                // (compat undeclared, tblInd −162, TableGrid cellMar 108):
                // both its tables sat 5.76pt right of Word. The positive case
                // stays per-cell (S496 lead_absorb) — no overlap, negative
                // never enters lead_absorb.
                // UNDECLARED compat parses as (15, explicit=false) but lays out
                // as legacy (the probe's no-compat doc absorbs on every arm) —
                // the witness doc has no compatSetting at all. The S621 zero
                // case keeps its explicit ≤14 gate for now (extending it to
                // undeclared docs needs its own corpus gate — probe arm ind0
                // says Word absorbs there too; recorded as follow-up).
                let s1239_negative = std::env::var("OXI_S1239_DISABLE").is_err()
                    && table.style.indent.is_some_and(|v| v < -0.1)
                    && (self.compat_mode <= 14 || !self.compat_mode_explicit);
                let absorb = if std::env::var("OXI_S621_DISABLE").is_err() {
                    (indent_zero && !self.table_modern_compat()) || s1239_negative
                } else {
                    matches!(table.style.indent, Some(v) if v.abs() < 0.01)
                        && !table.style.explicit_borders
                };
                if absorb {
                    // S1161 (2026-08-17, opt-out OXI_S1161_DISABLE): the shift
                    // is the cell margin ALONE. Word puts the absorbed table's
                    // border at anchor + offset - cellMar exactly, on both the
                    // floating and the tblInd path (_pb_tblanchor compat 11:
                    // tblpX=567 and tblInd=567 both land at 113.42 - 5.40 =
                    // 108.02, and Oxi's own ind567 arm already matched at
                    // 108.00). The extra half-stroke put plain_ind0 at 79.40
                    // against Word's 79.70 -- the only arm of the ten still
                    // outside the 0.05 stroke convention.
                    if std::env::var("OXI_S1161_DISABLE").is_err() {
                        pad_l_default
                    } else {
                        pad_l_default + border_w / 2.0
                    }
                } else {
                    0.0
                }
            };
            match table.style.alignment.as_deref() {
                Some("center") => start_x + (content_width - table_width) / 2.0,
                Some("right") => start_x + content_width - table_width,
                _ => start_x + table.style.indent.unwrap_or(0.0) - border_offset,
            }
        });

        // Default cell padding from table style or OOXML default
        // COM-measured 2026-03-29: L/R=4.95pt (99tw), T/B=0pt
        let default_pad = &table.style.default_cell_margins;
        let default_pad_l = default_pad.as_ref().and_then(|m| m.left).unwrap_or(4.95);
        let default_pad_r = default_pad.as_ref().and_then(|m| m.right).unwrap_or(4.95);
        let default_pad_t = default_pad.as_ref().and_then(|m| m.top).unwrap_or(0.0);
        let default_pad_b = default_pad.as_ref().and_then(|m| m.bottom).unwrap_or(0.0);

        // Table cell grid snap: Word snaps table ROW HEIGHTS to grid pitch
        // regardless of `adjustLineHeightInTable`. COM-measured 04b88e7e0b25
        // (which DOES set adjustLineHeightInTable) still has Word rendering
        // rows at linePitch * ceil(content/pitch) — 18.5pt for linePitch=360tw.
        // The flag affects intra-cell line-height behavior (see line_height_inner)
        // but NOT the row-height grid-snap.
        let table_grid_pitch: Option<f32> = grid_pitch;

        // COM-confirmed (2026-04-09): top border displaces table content downward
        // by the border width. cell_top_y = table_start_y + top_border_width.
        // Measured: 1row_outer4 marker_y=72.0, cell_y=97.5 → offset=0.5pt=top_bw.
        // S138 (2026-05-20): Bug A from S56 — this per-table top_bw add was
        // 1 of 2 causes of tokumei row drift.
        // S148 (2026-05-21) H9 DEFAULT ON (S151): BugA correct for type="lines"
        // docs (04b88e/d77a/34140b9c/b35/683ffc) but wrong for type="linesAndChars"
        // (tokumei/29dc6e). Apply BugA only for non-linesAndChars docs.
        // S242 (2026-05-23): removed OXI_LEGACY_BUGA_ALWAYS legacy env-var
        // fallback during hardening pass. OXI_BUG_A_REVERT preserved as
        // research toggle (binary opt-out for diagnostic purposes).
        // S1618 (2026-09-30, default ON, opt-out OXI_S1618_DISABLE): Bug A's
        // border width belongs UNDER the table, not above it. Word draws a
        // `lines`-grid table's top border AT the cursor (tokyoshugyo slice 649.3
        // vs cursor 649.25; policies__074da728 619.54, Oxi without Bug A 619.45,
        // with it 620.2), and starts the next block at the bottom rule's LOWER
        // edge: `_pb_gridbottom_tbl_gen.py` T_51700 (sz 6 borders) Word PDF
        // rules 619.18/636.70/670.78 = Oxi's 618.75/636.20/670.35 + 0.43 (the
        // same offset top and bottom), while «※雇用期間» sits 0.68 lower than
        // Oxi's (Oxi starts it at the rule's top edge). Retiring Bug A alone
        // lost those 0.75 for everything below a table (S1617 was a wrong
        // reading of that loss), so the width now advances after the last row.
        // Scope: a TYPED grid. Without one (EN reports__0013bcb8, docGrid with no
        // type) Word draws the top rule at the cursor too (p3 71.04, Oxi 71.10)
        // but the first row grows by the width (next rule 91.46 = Bug A's 91.48),
        // so there the cursor keeps Bug A's advance.
        // S1620 (2026-10-01, default ON, opt-out OXI_S1620_DISABLE): a CJK document
        // without a typed grid takes S1618 too. JA forms__000af17f (docGrid with no
        // type): Word's top rule 208.46 against Oxi's cursor 208.18 (+0.28, the rule's
        // top edge on the cursor) and the next rules 226.70 / 262.22 / 311.69 against
        // Oxi without Bug A 226.69 / 262.19 / 311.64 -- the first row does NOT grow
        // there, unlike the Latin no-type table in reports__0013bcb8.
        let s1620_cjk = std::env::var_os("OXI_S1620_DISABLE").is_none()
            && self.doc_body_has_real_cjk
            && (grid_pitch.is_none() || page.doc_grid_no_type);
        let s1618_on = std::env::var_os("OXI_S1618_DISABLE").is_none()
            && ((grid_pitch.is_some() && !page.doc_grid_no_type) || s1620_cjk);
        let bug_a_enabled = if std::env::var("OXI_BUG_A_REVERT").is_ok() || s1618_on {
            false
        } else {
            // Default (S151): apply only for non-linesAndChars docs
            grid_char_pitch.is_none()
        };
        // In the non-grid Latin row model the first row already carries
        // its top edge. Reserve the independent bottom edge at the end,
        // rather than compensating for it with an extra leading border.
        let separate_outer_edges = !self.doc_body_has_real_cjk
            && grid_pitch.is_none() && grid_char_pitch.is_none()
            && self.s1188_on() && table.style.border;
        if bug_a_enabled && table.style.border && !separate_outer_edges {
            let top_bw = table.style.border_width.unwrap_or(0.4);
            cursor.advance(top_bw);
        }
        let s1618_foot = if s1618_on
            && std::env::var("OXI_BUG_A_REVERT").is_err()
            && grid_char_pitch.is_none()
            && table.style.border
            && !separate_outer_edges
        {
            table.style.border_width.unwrap_or(0.4)
        } else {
            0.0
        };

        let num_rows = table.rows.len();
        let dump_table = std::env::var("OXI_DUMP_TABLE").is_ok();
        // A rendered row must fit the page. Half-point slack can keep an
        // entire overflow line; retain only floating-point roundoff tolerance.
        // S1469 (2026-09-18, default ON, opt-out OXI_S1469_DISABLE): promotes the
        // OXI_TABLE_BOTTOM_FIT checkpoint. A row is split only when
        // `row_bottom > page_bottom + row_fit_epsilon`, so a 0.5pt tolerance let a
        // row hang off the page instead of splitting. educational__003299ba p4:
        // the six-cell row whose fourth cell holds two paragraphs starts at 742.94
        // with 26.96 left under a content bottom of 769.90 and needs 27.36, so it
        // overruns by 0.40 -- inside the old tolerance. Word splits it, keeping
        // '1 aged 9-12' on page 4 and '1 aged 4-8' on page 5; Oxi kept both and
        // every later paragraph sat one page early. With the tolerance at 0.001
        // the document goes 0.9963 -> 1.0.
        // S1470 (2026-09-18, default ON, opt-out OXI_S1470_DISABLE): S1469's tight
        // tolerance is for tables in the FLOW. A FLOATING table (w:tblpPr) keeps
        // the old 0.5: policies__00602e8a is eight page-anchored "Provision:"
        // floats, and Word pushes their rows on whole where the tight tolerance
        // makes Oxi split them -- the document drops 0.7087 -> 0.5340 and loses a
        // page. Measured by toggling the two flags against each other:
        // S1469 off = 0.7087 pcd 0, S1469 on = 0.5340 pcd -1, with S1468 on in
        // both cases.
        let s1470_float_keeps_slack = table.style.position.is_some()
            && std::env::var("OXI_S1470_DISABLE").is_err();
        let row_fit_epsilon = if !s1470_float_keeps_slack
            && (std::env::var("OXI_TABLE_BOTTOM_FIT").is_ok()
                || std::env::var("OXI_S1469_DISABLE").is_err())
        {
            0.001
        } else {
            0.5
        };
        // S463 (2026-05-31): whole-table CJK check for the Latin-border-overhead
        // gate below. Cell-level "no CJK" mis-fired on numeric/Latin cells inside
        // CJK forms (459f05/34140b −0.12) — a row's height is the max over its
        // cells, so inflating one Latin cell in a mixed table over-grows the row.
        // Scope to tables that are ENTIRELY Latin (the gen2 English template
        // family) so mixed CJK tables are untouched.
        let table_is_latin = !table.rows.iter().any(|r| {
            r.cells.iter().any(|c| {
                c.blocks.iter().any(|b| {
                    if let Block::Paragraph(p) = b {
                        p.runs
                            .iter()
                            .any(|run| run.text.chars().any(kinsoku::is_cjk))
                    } else {
                        false
                    }
                })
            })
        });
        // S487: cell-anchored floating text boxes must render ON TOP of the whole
        // table (Word z-orders floating drawings in front of the table grid). Adding
        // them inline during the cell loop puts later rows' borders on top of the
        // box's white fill — the table grid lines then show THROUGH the callout. Defer
        // them to a separate vec and append AFTER the row loop so they paint last.
        let mut deferred_cell_textboxes: Vec<LayoutElement> = Vec::new();
        // S728 (2026-07-03): w:tblHeader repeat. Word re-draws the table's
        // LEADING header row(s) at the top of every page the table continues
        // onto (probethdr render-truth: 項目/内容 at y=72.9 on p2-p5). The
        // parsed TableRow.header flag was never consumed in layout. Mechanism:
        // capture the header rows' EMITTED elements (borders/shading/text —
        // post row-height correction) at first layout; on a mid-table page
        // push, replay clones shifted to the new page top and advance the
        // cursor past them. Corpus-safe: fires only when a tblHeader table
        // SPANS pages — the only 2 corpus tblHeader docs (ailitguide, b35123)
        // have non-spanning tables (verified) → byte-identical.
        let s728_on = std::env::var("OXI_S728_DISABLE").is_err();
        let mut s728_hdr_elems: Vec<LayoutElement> = Vec::new();
        let mut s728_hdr_h: f32 = 0.0;
        let mut s728_capture_done = false;
        // S1587: the page (pages.len()) the first header row was captured on.
        let mut s728_capture_page: Option<usize> = None;
        // S1083 (2026-08-06, default ON, opt-out OXI_S1083_DISABLE):
        // (row_idx, entry cursor) for the rows laid
        // out on the CURRENT page, so a page push can pull a keepNext row-chain
        // over with the row that triggered it. Cleared at every page push.
        let mut s1083_row_start: Vec<(usize, f32)> = Vec::new();
        // S1428 (2026-09-16, default ON, opt-out OXI_S1428_DISABLE): promoted
        // from the OXI_CJK_HEADER_ROW_CHAIN opt-in. `_pb_keepnext_hdr_gen.py`
        // (tests/fixtures/keepnext_hdr): a tblHeader row whose first data row
        // moves whole (binding atLeast, room < trH) moves with it -- plain arm
        // N=35 row1 p2 y=57.75 while the same row without tblHeader stays at
        // 705.75 -- and a keepNext heading before the table follows (heading
        // p2 y=59.25). policies__07543a6b 3.2.4 / 3.2.7 / 3.2.8.
        let cjk_header_chain = self.doc_body_has_real_cjk
            && (std::env::var_os("OXI_S1428_DISABLE").is_none()
                || std::env::var("OXI_CJK_HEADER_ROW_CHAIN").is_ok());
        let s1083_on = std::env::var("OXI_S1083_DISABLE").is_err()
            && (!self.doc_body_has_real_cjk || cjk_header_chain);
        // S1579 (2026-09-26, default ON, opt-out OXI_S1579_DISABLE): the header
        // chain is not a CJK rule. reports__0079718f p181: a tblHeader row
        // («Year | Performance measures | Expected performance results») fits at
        // the page bottom, its first data row does not; Word starts p182 with
        // both, Oxi left the header row alone on p181.
        // Compat-mode gated: `_pb_hdrchain_latin_gen.py` (Calibri 11, tblHeader row
        // + 4-line data rows, N swept over the page bottom, Word COM) -- compat 15
        // moves the header with a data row that goes whole (cantSplit, or not one
        // line fits) in every variant (bordered / borderless / 3-cell /
        // one spanning cell); compat 14 and 12 leave the header alone on the page
        // and repeat it (row 1 at 87.0 = header 72.75 + 14.25). technical__00549a8f
        // (compat 14) keeps 「Responsible TMA Organization」 on p17 exactly so.
        let header_chain = cjk_header_chain
            || (!self.doc_body_has_real_cjk
                && self.compat_mode >= 15
                && std::env::var_os("OXI_S1579_DISABLE").is_none());
        // A row "keeps with the next row" when its LEFTMOST cell's FIRST
        // paragraph declares keepNext (the S1024 row-chain predicate).
        let s1083_kn = |ri: usize| -> bool {
            table
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
        let mut s728_hdr_rows_seen: usize = 0;
        // S1192b: a vMerge span's outstanding height, per cell column. Recorded
        // at the restart row from the SAME `pad_t + content_h + pad_b` the emit
        // pass measures for a normal cell (the estimator's synthetic-row
        // measurement came out 11.25pt/line where emit renders 12.21), then
        // drawn down by each row the span passes through. The span's LAST row
        // has to cover whatever is left — Word's rule, `_pb_vmergedist_gen.py`.
        let mut s1192_pending: Vec<(usize, f32)> = Vec::new();
        // Keep the page coordinate scale stable when continuation space changes.
        let vmerge_coordinate_stride = content_height;
        let mut vmerge_absolute_ends: std::collections::HashMap<usize, f32> = std::collections::HashMap::new();
        let mut vmerge_text_flows: Vec<MergedCellTextFlow> = Vec::new();
        // Outer rules can survive a direct override of only the side/inside rules.
        let inherited_outer_rules = !self.doc_body_has_real_cjk && !table.style.border
            && std::env::var("OXI_TABLE_OUTER_RULES_DISABLE").is_err();
        let prepare_outer = |edge: &Option<BorderDef>| {
            if !inherited_outer_rules { return None; }
            edge.clone().map(|mut edge| {
                if edge.style == "nil" { edge.style = "none".into(); }
                if edge.color.is_none() { edge.color = Some("000000".into()); }
                edge
            })
        };
        let inherited_top_rule = prepare_outer(&table.style.top_border);
        let inherited_bottom_rule = prepare_outer(&table.style.bottom_border);
        // S1621: how far this row's drawn top edge sits above its box on a
        // continuation page (set by the table restart below, 0 otherwise).
        let mut s1621_lift: f32;
        for (row_idx, row) in table.rows.iter().enumerate() {
            // Resolve vertical defaults once for this row, before both height
            // estimation and emission. Explicit cell margins still win.
            let (row_default_pad_t, row_default_pad_b) =
                LayoutEngine::row_vertical_padding_defaults(row, default_pad_t, default_pad_b);
            let inherited_outer_page_before_row = pages.len();
            let mut row_declared_top_edges = Vec::new();
            s1621_lift = 0.0;
            // S740: page-transition bookkeeping + commit of the PREVIOUS row's
            // footnote reserve. On a page push the new page starts with zero
            // table-note reserve; the previous row's notes are committed to the
            // page the CURRENT row begins on (v1 approximation for split rows).
            if pages.len() != s740_pages_len {
                while s740_fn_pages.len() < pages.len() - s740_entry_pages + 1 {
                    s740_fn_pages.push(Vec::new());
                }
                s740_pages_len = pages.len();
                s740_reserve = 0.0;
                s740_page_has_notes = false;
            }
            if let Some(rf) = row_footnotes {
                if let Some(prev) = s740_pending_commit.take() {
                    let (ids_all, _h_all, hs_all) = &rf[prev];
                    // S1527: leave out the ids the split already placed.
                    let ids: Vec<u32> = ids_all.iter().copied().filter(|id| !s1527_early.contains(id)).collect();
                    let h: f32 = ids_all.iter().zip(hs_all.iter()).filter(|(id, _)| ids.contains(id)).map(|(_, h)| *h).sum();
                    if !ids.is_empty() {
                        if !s740_page_has_notes {
                            s740_reserve += fn_sep;
                            s740_page_has_notes = true;
                        }
                        s740_reserve += h;
                        if let Some(last) = s740_fn_pages.last_mut() {
                            for id in &ids {
                                if !last.contains(id) {
                                    last.push(*id);
                                }
                            }
                        }
                    }
                }
            }
            let mut row_height: f32 = 0.0;
            // Session 79c: visual_row_h = max cell content_h with emit-equivalent
            // line-height formula (grid-snapped when adjustLineHeightInTable). Used
            // ONLY for vAlign=center offset, NOT for row_height (page break logic
            // preserves the natural pre-pass to avoid 3a4f9f cascade — see
            // session79_adjust_lh_in_table_mixed_cell_valign_falsified.md).
            let mut visual_row_h: f32 = 0.0;
            let mut row_float_positions: std::collections::HashMap<usize, Vec<Option<(f32, f32)>>> =
                std::collections::HashMap::new();
            // S666 (2026-06-25): does this row have CELL-LEVEL horizontal borders
            // (tcBorders top/bottom) while the table has NO table-level insideH?
            // Word renders a cell-bordered row's height = content + the inside border
            // (~0.5pt for sz=4) — a border-box overhead. Oxi adds this for table-level
            // tblBorders insideH (via has_inside_h / pad_t) but MISSES it for cell-level
            // tcBorders → the 様式 form tables (tokumei 08_* T4, d4d126/de6e32/6514/
            // a1d6e4/31420 pages 2-7) render ~0.5pt/row too short, accumulating down the
            // page. Repro border_box/bb_repro: Word renders table-insideH AND cell-tcBorder
            // tables BOTH at 14.5/row (=14.0 content + 0.5 border); Oxi gives 14.0 for the
            // cell-border table. Corrects S477's "CJK line-height" misattribution (the
            // per-line height MATCHES Word at 14.0 — the drift is the missed border-box).
            let mut row_cell_hborder = false;
            // ROWBOX2: max cell-level horizontal border width seen in the row
            // (for the border-box overhead when the table has no insideH).
            let mut row_cell_hborder_w: f32 = 0.0;
            // S1065 (2026-08-03): the row's EFFECTIVE vertical cell margin
            // (max over non-merged cells of resolved top+bottom, direct cell
            // tcMar winning over the table-style default). Used as the extra
            // floor on BINDING atLeast rows (see the _ => atLeast branch).
            let mut row_eff_vmar: f32 = 0.0;
            // S503 (2026-06-08): centering-only row height using the ACTUAL GDI render
            // line-height (line_height_inner ~13.5) instead of the estimate's
            // word_line_height_table_cell (~12.625). visual_row_h under-counts when the
            // two diverge, so vAlign=center cells (e.g. vc_2cell_auto col0, and col0
            // generally — it is centered before later/taller cells' actual height is
            // known) center too HIGH. Tracked as a diff vs visual_row_h (same wrap, same
            // pad/border/nested — only the cell line-height differs) and fed into
            // effective_row_h ONLY for the v_offset centering, gated by OXI_S503_ENABLE
            // (default OFF until corpus-gated). Pagination row_height is untouched.
            //
            // S503 STATUS (2026-06-08): VALIDATED + zero-regression, kept OPT-IN.
            // Fixes vc_2cell_auto col0 (−1.0pt→0.0). Confirms S499's e3c545 −0.0974
            // was the SHARED pagination estimate, NOT centering (this centering-only
            // path is e3c545-safe: SSIM +0.0000). HOWEVER no current-corpus impact: it
            // only fires when snap_in_cell=FALSE (no docGrid/snap_to_grid), but the
            // bottom-N/tokumei docs all have docGrid → snap_in_cell=TRUE → their estimate
            // ALREADY uses line_height_inner → center_extra=0. The real db9ca/tokumei
            // cell-Y errors are a DIFFERENT mechanism (NOT this col0-before-taller-cell
            // ordering). Opt-in so a future no-docGrid vAlign=center doc gets the fix.
            let s503_enable = std::env::var("OXI_S503_ENABLE").is_ok();
            // S1423 (2026-09-16, default ON, opt-out OXI_S1423_DISABLE): a cantSplit
            // row's page fit is measured with its RENDER height in CJK documents
            // too. The Latin-only gate left the pagination estimate (25.75 for a
            // 2-line MS Mincho 10.5 row) deciding against a 27.23 render:
            // `_pb_rowfit_gen.py` (tests/fixtures/rowfit, Word COM, bottom 785.2)
            // keeps the row at line-1 top 757.75 (last line bottom 784.75) and
            // moves it whole at 758.75 (785.75); Oxi kept it through 759.7.
            // Word's rule = the row's LAST LINE box must end above the content
            // bottom; the bottom border and cell padding may hang past it.
            let measure_cant_split_fit = row.cant_split
                && (!self.doc_body_has_real_cjk || std::env::var_os("OXI_S1423_DISABLE").is_none())
                && row.height_rule.as_deref() != Some("exact");
            let mut center_row_h: f32 = 0.0;
            let mut kept_first_paragraph_height: f32 = 0.0;
            let mut first_cell_line_fit = f32::INFINITY;
            let mut row_has_multiple_text_lines = row.cells.iter().any(|cell|
                cell.blocks.iter().filter(|b| matches!(b, Block::Paragraph(_))).count() > 1);

            let row_entry_cursor_y = cursor.cursor_y;
            // S1083 (2026-08-06, default ON, opt-out OXI_S1083_DISABLE):
            // remember where each row of THIS page
            // started, so a page push can walk back over a keepNext row-chain.
            s1083_row_start.push((row_idx, row_entry_cursor_y));

            // S361 (2026-05-27, FALSIFIED): hypothesized that trHeight rows
            // should NOT grid-snap the cell line (b5f706e9 row 1 Word cellH=17pt
            // for a 9pt header looked un-snapped). Env-gated test FALSIFIED:
            // OXI_S361_TRHEIGHT_NO_LINE_SNAP=1 → Phase 2 0.9603→0.9205 (-0.0398)
            // AND Phase 1 53/55→51/55. Same S349 trap: Word Cell.Height reports
            // the LOGICAL trHeight (17pt), NOT the visual rendered extent — Word
            // DOES snap the line to 18pt; the row visually is 18pt. So the
            // +1.0pt cluster is NOT from line grid-snap. Most trHeight rows need
            // the snap (it's correct). Gate kept OFF; default unchanged.
            let row_line_pitch: Option<f32> = if row.height.is_some()
                && std::env::var("OXI_S361_TRHEIGHT_NO_LINE_SNAP").is_ok()
            {
                None
            } else {
                table_grid_pitch
            };

            // First pass: calculate row height
            let mut grid_idx = row.grid_before as usize;
            for (replay_cell_idx, cell) in row.cells.iter().enumerate() {
                // S666: cell-level horizontal border detection (see row_cell_hborder above).
                if !table.style.has_inside_h {
                    if let Some(b) = &cell.borders {
                        if b.top.is_some() || b.bottom.is_some() {
                            row_cell_hborder = true;
                            for d in [&b.top, &b.bottom].into_iter().flatten() {
                                if d.width > row_cell_hborder_w {
                                    row_cell_hborder_w = d.width;
                                }
                            }
                        }
                    }
                }
                let span = cell.grid_span.max(1) as usize;
                // vMerge="continue" cells don't contribute to row height
                // (their content is part of the vMerge="restart" cell above).
                // vMerge="restart" cells also don't contribute: Word distributes
                // the restart cell's content across the entire vMerge span, so
                // the row's own height comes from non-merged cells in the same row.
                if cell.v_merge.as_deref() == Some("continue")
                    || cell.v_merge.as_deref() == Some("")
                    || cell.v_merge.as_deref() == Some("restart")
                {
                    grid_idx += span;
                    continue;
                }
                // S1493 (2026-09-20): a cell whose gridSpan runs past the table's grid
                // (reference__009644b180d1bc56) must not index past col_widths.
                let cell_w: f32 = col_widths[grid_idx.min(col_widths.len())..(grid_idx + span).min(col_widths.len())].iter().sum();
                let _pad_l = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.left)
                    .unwrap_or(default_pad_l);
                let _pad_r = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.right)
                    .unwrap_or(default_pad_r);
                let mut pad_t = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.top)
                    .unwrap_or(row_default_pad_t);
                let pad_b = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.bottom)
                    .unwrap_or(row_default_pad_b);
                // S1575 (2026-09-26, default ON, opt-out OXI_S1575_DISABLE): a cell's
                // top/bottom margin is ROW-wide -- every cell of the row takes the
                // row's largest. reference__009644b1: only the label cells carry tcMar
                // 100/100; Word starts the content cell's first line on the label's
                // baseline (PDF 256.13 both) and the 3-line row is 5 + 43.92 + 5,
                // Oxi's content cell had no margin (43.9) and page 2 ran ~50pt short.
                let (pad_t, pad_b) = if std::env::var_os("OXI_S1575_DISABLE").is_none() {
                    let mt = row.cells.iter().map(|c| c.margins.as_ref().and_then(|m| m.top).unwrap_or(row_default_pad_t)).fold(pad_t, f32::max);
                    let mb = row.cells.iter().map(|c| c.margins.as_ref().and_then(|m| m.bottom).unwrap_or(row_default_pad_b)).fold(pad_b, f32::max);
                    (mt, mb)
                } else { (pad_t, pad_b) };
                #[allow(unused_mut)]
                let mut pad_t = pad_t;
                // S1065: accumulate the EFFECTIVE vertical cell margin (resolved
                // direct-else-table-style top+bottom) for the binding-atLeast
                // row floor. NOTE pad_t is mutated below by the ROWBOX2 border
                // pad, so capture the margin-only value here before that.
                let cell_eff_vmar = pad_t + pad_b;
                if cell_eff_vmar > row_eff_vmar {
                    row_eff_vmar = cell_eff_vmar;
                }
                // Round 30: implicit border padding (matches second pass)
                // S359 (2026-05-27): test confirmed Round 30 is load-bearing
                // (OXI_S359_NO_ROUND30=1 caused -0.0186 corpus regression).
                // S386 (2026-05-27): hypothesis "bug_a + Round30 double-count the
                // top border on row 0" FALSIFIED. OXI_S386_NO_DOUBLE_BORDER=1
                // (suppress Round30 on row 0 when bug_a fired) → corpus 0.9603
                // → 0.9521 (-0.0082, pass 18→17), and b5f706 barely moved
                // (0.9715→0.9707) because the iou_yrange_adj median absorbs
                // uniform per-table shifts. Round30 on row 0 is load-bearing.
                if self.rowbox2_pad_on() {
                    // ROWBOX2: generalized Round30 — bw pads the content top
                    // ADDITIVELY with explicit cellMar, incl. cell tcBorders.
                    // S870: Latin docs use the ROW's rule (see the helper).
                    pad_t += self.rowbox2_border_pad_row(table, row_idx, cell);
                } else if pad_t == 0.0 && table.style.border {
                    pad_t = table.style.border_width.unwrap_or(0.4);
                }
                // COM-confirmed (2026-04-09, 10 minimal repros + 3 real docs):
                // Each row's height includes its BOTTOM-EDGE border:
                //   - Non-last rows: bottom edge = insideH width (0 if no insideH)
                //   - Last row: bottom edge = outer bottom border (0 if none)
                // Top/side borders do NOT add to row height.
                // OOXML default single border sz=4 = 0.5pt (4/8).
                let bw =
                    table
                        .style
                        .border_width
                        .unwrap_or(if table.style.border { 0.5 } else { 0.0 });
                let is_last = row_idx + 1 == num_rows;
                let _border_overhead = if is_last {
                    if table.style.border {
                        bw
                    } else {
                        0.0
                    }
                } else if table.style.has_inside_h {
                    if std::env::var("OXI_S921_DISABLE").is_err() {
                        table
                            .style
                            .inside_horizontal_border
                            .as_ref()
                            .filter(|b| b.style != "none")
                            .map(|b| b.width)
                            .unwrap_or(bw)
                    } else {
                        bw
                    }
                } else {
                    0.0
                };
                // For line-wrapping estimation, use cell_w (not inner_w after padding)
                // Word allows text to extend into cell margins for wrapping purposes.
                // For line-wrapping estimation, use cell_w (not inner_w after padding).
                // S562 (2026-06-14): the roudoujoken r7 (5)裁量 wrap IS a cellMar-budget
                // issue (count_cell_lines CCL: cum to る = 430.5 ≤ cell_w 432 → fits;
                // Word's budget cell_w − cellMar 426.8 → る wraps). But subtracting
                // cellMar here only fixes the ESTIMATE — the RENDER's cell wrap
                // (mod.rs:9640+) is the operative budget for pagination, and it is a
                // KNOWN doc-dependent discriminator problem (191cb uses cell_w-extend,
                // d77a/29dc6e use cell_w−cellMar; "No simple toggle works"). See memory.
                // OXI_CELLPAIR experiment (2026-07-02): universal SUBTRACT boundary
                // (Word always wraps cell text at cell_w - cellMar; 24-config derivation)
                // paired with the small-cap cell 約物 credit below. See char_budget_wall.
                // S713 (2026-07-02, default ON, opt-out OXI_S713_DISABLE): a LEGACY
                // (compat<=14) single-cell row in a table with an author-declared
                // tblCellMar wraps at cell_w - pads (Word render-truth tokyoshugyo p30
                // (注) cell: content [96.98, 517.42] inside borders [92.06, 522.22] =
                // both cellMars subtracted; Oxi wrapped to the border -> fit +1 char/
                // line -> the (注) para 2 lines vs Word 3 -> the 変形 -1x6 cascade).
                // This envelope (explicit tblCellMar) was excluded from EVERY prior
                // subtract gate (s531/s559/s585 all require !has_explicit_cellmar), so
                // it is orthogonal to the tuned cellMar discriminators. Corpus scope:
                // compat<=14 + explicit tblCellMar = tokyoshugyo/parttime/ohnoitaku
                // ONLY (3a4f/model are compat15; repro_tcmar_* have NO compatSetting,
                // which parse_compat_mode reads as 15 -> not fired, byte-identical).
                // Paired with the S421 oikomi arm below (same discriminator): the
                // narrowed budget forces wraps Word resolves by oikomi/oidashi.
                let s713_cellmar = std::env::var("OXI_S713_DISABLE").is_err()
                    && row.cells.len() == 1
                    && table.style.has_explicit_cellmar
                    && self.compat_mode <= 14;
                // S768 = the S585c 本体 wrap-margin fix for PURE-LATIN documents
                // (2026-07-08, default ON, opt-out OXI_S768_DISABLE). Oxi's default
                // cell wrap is the FULL cell_w (Word wraps at cell_w − cellMar on both
                // sides); the JP corpus calibrated around this over-wide wrap because
                // cell_w is over-computed and the full-width wrap COMPENSATES (the
                // S585c/S562 wall — narrowing it regresses JP SSIM, LATINCELLWRAP was
                // reverted for net −0.0069). But for a NO-CJK document (uk_local_
                // spending etc.) the column widths already MATCH Word exactly (measured:
                // char-width + col-width RULED OUT) so the ONLY error is the missing
                // margin subtract → Oxi packs ~1 char/line more → Annex table rows
                // under-wrap by ~20%. Scope to `!doc_body_has_cjk` (a doc-level flag
                // true for EVERY JP doc → byte-identical for the whole JP corpus by
                // construction; the per-para Latin scope was NOT clean because JP forms
                // have Latin-only numeric/date cells sharing this path). Mirrored at the
                // render wrap_base (Fix C estimate==render invariant). See
                // [[english_corpus_bug_mine]] / [[tokumei_form_family_ssim]].
                // ★HELD OPT-IN (OXI_S768=1, default OFF = byte-identical everywhere).
                // The doc-level strict-CJK scope RESOLVES the JP-safety block that
                // reverted LATINCELLWRAP (tracked 238-doc corpus byte-identical, net
                // +0.0000), but the ENGLISH corpus is ITSELF heterogeneous — the S562
                // wrap-budget wall persists within Latin: uk_local_spending +0.0018 /
                // health_form +0.0014 improve, BUT risk_assessment −0.0027 regresses
                // (no tblLayout/cellMar discriminator; risk_assessment is the CELLWORD
                // showcase doc calibrated at the FULL cell_w wrap). Ships default-ON
                // once a per-table discriminator (or CELLWORD co-calibration) is found.
                // S768 default-ON (2026-07-12, opt-out OXI_S768_DISABLE): with the
                // S803 footer-reserve fix in place the doc-level Latin cell-wrap
                // moves uklocalspending's Annex tables another page toward Word
                // (pcd -2 -> -1, neg-delta mass 551 -> 453). The held reason
                // (risk_assessment SSIM -0.0027, a 1-page PASS doc calibrated on
                // the full-cell_w compensation) is a render-only trade accepted
                // under the Phase-1-first directive; JP is byte-identical by
                // construction (doc-level real-CJK gate, 238-doc A/B +0.0000).
                let s768_latin_wrap =
                    std::env::var("OXI_S768_DISABLE").is_err() && !self.doc_body_has_real_cjk;
                // S1173 estimate side: the derived law supersedes the whole
                // allowlist (see `celllaw_inset`). Estimate and render must take
                // the same base or pagination and drawing disagree.
                let inner_w = if self.celllaw() {
                    LayoutEngine::celllaw_twips(
                        (cell_w
                            - self.celllaw_inset(table, cell, _pad_l, false)
                            - self.celllaw_inset(table, cell, _pad_r, true))
                        .max(0.0),
                    )
                } else if self.cellpair_active() || s713_cellmar || s768_latin_wrap {
                    (cell_w - _pad_l - _pad_r).max(0.0)
                } else {
                    cell_w.max(0.0)
                };
                let cell_float_flow = LayoutEngine::cell_float_enabled(cell) && !self.is_vert_writing_active(cell);
                let mut float_positions = vec![None; cell.blocks.len()];
                let mut next_float = cell.blocks.iter().enumerate().filter_map(|(i, block)| {
                    match block {
                        Block::Image(image) if cell_float_flow && image.position.is_some() => Some((i, image)),
                        _ => None,
                    }
                });
                let (mut cell_content_h, mut cell_content_h_visual, center_extra) = loop {
                let mut cell_content_h = pad_t;
                let mut float_tops = vec![0.0; cell.blocks.len()];
                // S751 (2026-07-05, default ON, opt-out OXI_S751_DISABLE): an
                // EMPTY cell with <w:hideMark/> contributes ZERO content height
                // (ECMA-376 17.4.22 — the end-of-cell mark is excluded from the
                // row height; the thin-spacer-row idiom). Word collapses an
                // auto row of empty hideMark cells to ~borders-only; Oxi gave
                // the empty paragraph a full line (probeqhidemk2 {+1:1}).
                // Cells with real text are untouched (the corpus's hideMark
                // rows are all trHeight-bound mixed rows -> byte-identical).
                let hidden_tail_para = cell.blocks.last().and_then(|block| match block {
                    Block::Paragraph(para) if LayoutEngine::hidden_cell_final_line(cell, cell.blocks.len() - 1, para) =>
                        Some(LayoutEngine::without_hidden_cell_after(para)),
                    _ => None,
                });
                let s1311_tail = self.s1311_hidemark_tail_pos(cell);
                let s751_hide_empty = cell.hide_mark
                    && std::env::var("OXI_S751_DISABLE").is_err()
                    && s1311_tail.is_none()
                    && cell.blocks.iter().all(|b| {
                        matches!(b, Block::Paragraph(p)
                        if p.runs.iter().all(|r| r.text.is_empty()))
                    });
                // Session 79c: parallel emit-equivalent content_h for visual_row_h
                let mut cell_content_h_visual = pad_t;
                // S503: extra height vs visual when using the render line-height (per-cell
                // sum of (render_para_h − estimate_para_h) over paragraphs). center cell
                // height = cell_content_h_visual + center_extra.
                let mut center_extra: f32 = 0.0;
                // S427 (2026-05-29): adjacent-paragraph spacing collapse inside a
                // cell. Word collapses sa(prev)+sb(cur) to max(sa,sb) — COM-confirmed
                // on 29dc6e tbl1 r2c2 (two empty paras, sa=sb=4.35pt exact-12:
                // para gap = 16.5pt = 12.0 + max, NOT 20.7 = 12.0 + sum). Mirrors the
                // body path collapse (mod.rs:3965). prev_sa carries the previous
                // paragraph's space_after; the credit min(prev_sa, cur_sb) is removed.
                let s427_collapse = std::env::var("OXI_S427_DISABLE").is_err();
                let mut prev_sa: Option<f32> = None;
                // S939: prev paragraph's (contextual_spacing, style_id) for the
                // in-cell layered collapse.
                let mut s939_prev: Option<(bool, Option<&str>)> = None;
                // S1075: (previous cell paragraph's after_autospacing, its numId)
                let mut s1075_prev: Option<(bool, Option<&str>)> = None;
                // Cell-autospace (OXI_CELLAS): first/last Paragraph positions for
                // container-edge suppression. See cell_effective_spacing.
                let first_para_pos = cell
                    .blocks
                    .iter()
                    .position(LayoutEngine::is_cell_spacing_paragraph);
                let last_para_pos = cell
                    .blocks
                    .iter()
                    .rposition(LayoutEngine::is_cell_spacing_paragraph);

                // Session 131 (2026-05-20): vertical writing — cell height
                // along the page-y axis equals the sum of vertical-text lengths
                // (chars × font_size), not the wrapped-horizontal line count.
                // Gated by OXI_VERT_WRITING env var.
                let vert_writing_active = self.is_vert_writing_active(cell);
                // S753 (2026-07-05, opt-out OXI_S753_DISABLE): a tbRlV cell does
                // NOT grow an auto row for its text — columns consume WIDTH, the
                // flow wraps into whatever row height the OTHER cells / trHeight
                // produce. Contribution = ONE line (first paragraph only). Full
                // derivation at s753_vert_cell_columns.
                let s753_vert = vert_writing_active && std::env::var("OXI_S753_DISABLE").is_err();
                let mut s753_first_done = false;
                // S716: the post-nested-table stub paragraph contributes no height.
                let s716_stub = self.nested_table_stub_pos(cell);
                for (block_pos, block) in cell.blocks.iter().enumerate() {
                    if s751_hide_empty {
                        break;
                    } // S751: empty hideMark cell = no content height
                    if Some(block_pos) == s716_stub || Some(block_pos) == s1311_tail {
                        continue;
                    }
                    match block {
                        Block::Paragraph(para) => {
                            let hidden_final_mark = !vert_writing_active
                                && LayoutEngine::hidden_cell_final_line(cell, block_pos, para);
                                                        let para = if hidden_final_mark { hidden_tail_para.as_ref().unwrap() } else { para };
                            let (mut para_h, mut para_h_visual, mut para_h_center) = if vert_writing_active {
                                let h = if s753_vert {
                                    if s753_first_done {
                                        0.0
                                    } else {
                                        s753_first_done = true;
                                        self.estimate_para_height(
                                            para,
                                            100_000.0,
                                            row_line_pitch,
                                            table.style.para_style.as_ref(),
                                            true,
                                            grid_char_pitch,
                                            grid_char_cw_ratio,
                                        )
                                    }
                                } else {
                                    self.vert_para_height(para)
                                };
                                (h, h, h)
                            } else {
                                let p1 = self.estimate_para_height(
                                    para,
                                    inner_w,
                                    row_line_pitch,
                                    table.style.para_style.as_ref(),
                                    true,
                                    grid_char_pitch,
                                    grid_char_cw_ratio,
                                );
                                let mut line_measure = cell_float::Measurement::default();
                                let p2 = self.estimate_para_height_inner(
                                    para, inner_w, row_line_pitch,
                                    table.style.para_style.as_ref(), true,
                                    grid_char_pitch, grid_char_cw_ratio,
                                    true, false, Some(&mut line_measure),
                                );
                                row_has_multiple_text_lines |= line_measure.heights.len() > 1;
                                if Some(block_pos) == first_para_pos {
                                    if let Some(&height) = line_measure.heights.first() {
                                        let top = self.cell_float_paragraph_top(
                                            para, table, row_line_pitch, pad_t, true,
                                            Some(block_pos) == last_para_pos, prev_sa,
                                            s939_prev, s1075_prev,
                                        );
                                        first_cell_line_fit = first_cell_line_fit.min(top + height);
                                    }
                                }
                                // S503: render-line-height variant for centering floor
                                // (opt-in; default OFF avoids the extra estimate call).
                                let p3 = if s503_enable || measure_cant_split_fit {
                                    self.estimate_para_height_emit_render(
                                        para,
                                        inner_w,
                                        row_line_pitch,
                                        table.style.para_style.as_ref(),
                                        true,
                                        grid_char_pitch,
                                        grid_char_cw_ratio,
                                    )
                                } else {
                                    p2
                                };
                                (p1, p2, p3)
                            };
                            if cell_float_flow {
                                let top = self.cell_float_paragraph_top(para, table, row_line_pitch,
                                    cell_content_h_visual - pad_t, Some(block_pos) == first_para_pos,
                                    Some(block_pos) == last_para_pos, prev_sa, s939_prev, s1075_prev);
                                let estimate_top = self.cell_float_paragraph_top(para, table, row_line_pitch,
                                    cell_content_h - pad_t, Some(block_pos) == first_para_pos,
                                    Some(block_pos) == last_para_pos, prev_sa, s939_prev, s1075_prev);
                                let render_top = self.cell_float_paragraph_top(para, table, row_line_pitch,
                                    cell_content_h_visual + center_extra - pad_t, Some(block_pos) == first_para_pos,
                                    Some(block_pos) == last_para_pos, prev_sa, s939_prev, s1075_prev);
                                // A float moved to a continuation keeps its anchor there.
                                // Otherwise removing its old obstacle pulls the anchor back,
                                // and the two page assignments oscillate on every replay.
                                let floor = float_replay.and_then(|r| r.origins.get(&(row_idx, replay_cell_idx, block_pos)))
                                    .copied().unwrap_or(0.0);
                                let gap = (floor - top).max(0.0);
                                let top = top + gap;
                                let estimate_top = estimate_top + gap;
                                let render_top = render_top + gap;
                                cell_content_h += gap;
                                cell_content_h_visual += gap;
                                float_tops[block_pos] = top;
                                let obstacles = LayoutEngine::cell_float_obstacles_at(cell, &float_positions);
                                para_h = self.measure_cell_float_para(para, inner_w, row_line_pitch,
                                    table.style.para_style.as_ref(), grid_char_pitch, grid_char_cw_ratio,
                                    false, false, estimate_top, &obstacles).0;
                                para_h_visual = self.measure_cell_float_para(para, inner_w, row_line_pitch,
                                    table.style.para_style.as_ref(), grid_char_pitch, grid_char_cw_ratio,
                                    true, false, top, &obstacles).0;
                                para_h_center = if s503_enable || measure_cant_split_fit {
                                    self.measure_cell_float_para(para, inner_w, row_line_pitch,
                                        table.style.para_style.as_ref(), grid_char_pitch, grid_char_cw_ratio,
                                        true, true, render_top, &obstacles).0
                                } else { para_h_visual };
                            }
                            if hidden_final_mark {
                                let credit = self.hidden_cell_final_mark_height(para,
                                    table.style.para_style.as_ref(), row_line_pitch);
                                para_h = (para_h - credit).max(0.0);
                                para_h_visual = (para_h_visual - credit).max(0.0);
                                para_h_center = (para_h_center - credit).max(0.0);
                            }
                            center_extra += para_h_center - para_h_visual;
                            // Day 33 part 17 (2026-05-10): subtract space_before for first
                            // paragraph in cell to match Word's behavior. Mirrors the
                            // suppression in layout_table cell loop at line ~5877. Without
                            // this, the row reserves extra height for borders even though
                            // the text is positioned correctly.
                            // S136 (2026-05-20): TR_V200-V203 + R1A re-measurement show
                            // Word DOES apply sb to first cell para (cell_para_y shifts
                            // 4.35pt when sb=87). Day 33 part 17 premise is wrong.
                            // S239 (2026-05-23): removed OXI_LEGACY_SB_SUPPRESS and
                            // OXI_SB_NO_SUPPRESS legacy env-var fallbacks during
                            // hardening pass. The `if sb_suppress_enabled` block
                            // was dead code (LEGACY var default false → block
                            // never executed). S151 default ON since 2026-05-21.
                            // S1432 (2026-09-16, default ON, opt-out OXI_S1432_DISABLE):
                            // promoted from the OXI_CJK_CELL_KEEP_LINES opt-in -- a
                            // keepLines first paragraph must fit the first fragment
                            // whole (policies__07543a6b p32 row 5: 6-line keepLines
                            // scenario cell, 69pt of room, Word moves the row whole).
                            if self.doc_body_has_real_cjk
                                && (std::env::var("OXI_CJK_CELL_KEEP_LINES").is_ok()
                                    || std::env::var_os("OXI_S1432_DISABLE").is_none())
                                && Some(block_pos) == first_para_pos
                                && para.style.keep_lines
                            {
                                kept_first_paragraph_height = kept_first_paragraph_height
                                    .max(cell_content_h_visual + para_h_visual);
                            }
                            cell_content_h += para_h;
                            cell_content_h_visual += para_h_visual;
                            if std::env::var("OXI_DBG_CELLPARA").is_ok() {
                                let head: String = para.runs.iter().flat_map(|r| r.text.chars()).take(10).collect();
                                eprintln!("[CELLPARA] para_h={:.3} visual={:.3} render={:.3} center_extra={:.3} cum={:.3} cum_visual={:.3} «{}»", para_h, para_h_visual, para_h_center, center_extra, cell_content_h, cell_content_h_visual, head);
                            }
                            // S753: vert-cell paragraph spacing maps to the WIDTH
                            // direction (each paragraph is a new column) — no height
                            // bookkeeping.
                            if !s753_vert {
                                // S427: collapse this paragraph's space_before against
                                // the previous paragraph's space_after.
                                let (cur_sb, cur_sa) = self.cell_para_spacing(
                                    para,
                                    table.style.para_style.as_ref(),
                                    row_line_pitch,
                                );
                                // Cell-autospace override: estimate_para_height added the
                                // explicit (cur_sb, cur_sa) to para_h above; replace it with
                                // the autospace-effective values (13.75 / edge-suppressed 0).
                                let s952_tbl =
                                    table.style.para_style.as_ref().map_or(false, |ts| {
                                        ts.before_autospacing || ts.after_autospacing
                                    });
                                let (cur_sb, cur_sa) = if para.style.before_autospacing
                                    || para.style.after_autospacing
                                    || para.style.contextual_spacing
                                    || s952_tbl
                                {
                                    let (eff_sb, eff_sa) = self.cell_effective_spacing(
                                        para,
                                        table.style.para_style.as_ref(),
                                        Some(block_pos) == first_para_pos,
                                        Some(block_pos) == last_para_pos,
                                        cur_sb,
                                        cur_sa,
                                    );
                                    cell_content_h += (eff_sb - cur_sb) + (eff_sa - cur_sa);
                                    cell_content_h_visual += (eff_sb - cur_sb) + (eff_sa - cur_sa);
                                    (eff_sb, eff_sa)
                                } else {
                                    (cur_sb, cur_sa)
                                };
                                if s427_collapse {
                                    if let Some(psa) = prev_sa {
                                        let credit = psa.min(cur_sb);
                                        cell_content_h -= credit;
                                        cell_content_h_visual -= credit;
                                    }
                                }
                                let s939 = self.s939_cell_ctx_credit(
                                    prev_sa,
                                    s939_prev.map_or(false, |p| p.0),
                                    s939_prev.and_then(|p| p.1),
                                    &para.style,
                                    cur_sb,
                                );
                                cell_content_h -= s939;
                                cell_content_h_visual -= s939;
                                let s1075 = self.s1075_cell_list_credit(
                                    prev_sa,
                                    s1075_prev.map_or(false, |p| p.0),
                                    s1075_prev.and_then(|p| p.1),
                                    &para.style,
                                    cur_sb,
                                );
                                cell_content_h -= s1075;
                                cell_content_h_visual -= s1075;
                                prev_sa = Some(cur_sa);
                                s939_prev = Some((
                                    para.style.contextual_spacing,
                                    para.style.style_id.as_deref(),
                                ));
                                s1075_prev = Some((
                                    para.style.after_autospacing,
                                    para.style.num_id.as_deref(),
                                ));
                            }
                        }
                        Block::Table(nested) => {
                            prev_sa = None;
                            s939_prev = None;
                            s1075_prev = None;
                            // Estimate nested table height from rows
                            // COM-confirmed: nested table width = cell width - 2 × padding
                            let nested_w = (inner_w).max(0.0);
                            // S1068 (2026-08-05, opt-out OXI_S1068_DISABLE): estimate each
                            // nested cell paragraph at ITS OWN resolved column width instead
                            // of the unconditional `nested_w / 2.0`, which assumes every
                            // nested table has exactly two equal columns. A ONE-column
                            // nested table was therefore wrapped at HALF its real width in
                            // the pre-pass: educational__002a301d's Tech Tip box counted 5
                            // lines where the render emits 2, and because the correction
                            // pass only ever GROWS max_actual_cell_h the 48pt phantom stayed
                            // as the outer row's floor. Mirrors the outer walk's own
                            // grid_span / vMerge column arithmetic and the enclosing
                            // `inner_w` cell-margin convention.
                            let s1068_cols = if std::env::var("OXI_S1068_DISABLE").is_err() {
                                Some(self.resolve_table_col_widths_n(nested, nested_w, true))
                            } else {
                                None
                            };
                            let n_def_pad = &nested.style.default_cell_margins;
                            let n_def_l = n_def_pad.as_ref().and_then(|m| m.left).unwrap_or(4.95);
                            let n_def_r = n_def_pad.as_ref().and_then(|m| m.right).unwrap_or(4.95);
                            let n_sub_pad =
                                self.cellpair_active() || s713_cellmar || s768_latin_wrap;
                            for nr in &nested.rows {
                                let mut nr_h = 0.0_f32;
                                let mut n_grid_idx = 0usize;
                                for nc in &nr.cells {
                                    let n_span = nc.grid_span.max(1) as usize;
                                    let np_w = match &s1068_cols {
                                        Some(cw) if n_grid_idx + n_span <= cw.len() => {
                                            let w: f32 =
                                                cw[n_grid_idx..n_grid_idx + n_span].iter().sum();
                                            if n_sub_pad {
                                                let pl = nc
                                                    .margins
                                                    .as_ref()
                                                    .and_then(|m| m.left)
                                                    .unwrap_or(n_def_l);
                                                let pr = nc
                                                    .margins
                                                    .as_ref()
                                                    .and_then(|m| m.right)
                                                    .unwrap_or(n_def_r);
                                                (w - pl - pr).max(0.0)
                                            } else {
                                                w.max(0.0)
                                            }
                                        }
                                        _ => nested_w / 2.0,
                                    };
                                    n_grid_idx += n_span;
                                    let mut nc_h = 0.0_f32;
                                    for nb in &nc.blocks {
                                        if let Block::Paragraph(np) = nb {
                                            nc_h += self.estimate_para_height(
                                                np,
                                                np_w,
                                                table_grid_pitch,
                                                nested.style.para_style.as_ref(),
                                                true,
                                                grid_char_pitch,
                                                grid_char_cw_ratio,
                                            );
                                        }
                                    }
                                    nr_h = nr_h.max(nc_h);
                                }
                                if let Some(h) = nr.height {
                                    match nr.height_rule.as_deref() {
                                        Some("exact") => {
                                            nr_h = h;
                                        }
                                        // ROWBOX2: binding atLeast = trH + bw (border-box)
                                        Some("atLeast") => {
                                            nr_h = nr_h.max(h + self.rowbox2_trh_bw(nested, nr));
                                        }
                                        _ => {}
                                    }
                                }
                                cell_content_h += nr_h;
                                cell_content_h_visual += nr_h;
                            }
                        }
                        Block::Image(img) => {
                            // S331 (2026-05-26): account for inline drawing
                            // height in cell. Pairs with parser fix at
                            // parser/ooxml.rs:5190 (forwards pr.inline_images
                            // to cell.blocks). Without this, cell height
                            // calculation ignores the drawing → cell renders
                            // shorter than Word → downstream content cascades
                            // to wrong page. Gated by parser-side env so this
                            // arm only matches when fix is active.
                            // S715 (2026-07-02, default ON, opt-out OXI_S715_DISABLE):
                            // a typed-grid cell IMAGE line snaps to whole grid cells
                            // like a text line (Word render-truth tokyoshugyo p34
                            // calendar: 289.5pt VML image cell → Word table bottom
                            // border at 596.11 ≈ image_top + 17 cells; Oxi resumed
                            // the body at image_top + 289.5 exactly, 1 grid line
                            // early → the 変形 p35/36 −1s).
                            let img_line = self.s971_image_line_h(img, 1.0e6, row_line_pitch, false);
                            let img_h_eff = if std::env::var("OXI_S715_DISABLE").is_err() {
                                match row_line_pitch {
                                    Some(p) if p > 0.0 => (img_line / p).ceil() * p,
                                    _ => img_line,
                                }
                            } else {
                                img_line
                            };
                            // S1053: a page-relative float reserves nothing.
                            if !(LayoutEngine::s1053_cell_float_no_reserve(img) || (cell_float_flow && img.position.is_some())) {
                                let (sb, sa) = self.cell_image_spacing(
                                    img, table, row_line_pitch, &cell.blocks, block_pos,
                                    prev_sa, s939_prev, s1075_prev,
                                );
                                cell_content_h += img_h_eff + sb + sa;
                                cell_content_h_visual += img_h_eff + sb + sa;
                                if let Some(host) = img.host_paragraph.as_deref() {
                                    prev_sa = Some(sa);
                                    s939_prev = Some((host.style.contextual_spacing, host.style.style_id.as_deref()));
                                    s1075_prev = Some((host.style.after_autospacing, host.style.num_id.as_deref()));
                                }
                            }
                        }
                        Block::Math(math_block) => {
                            // S1244 (2026-08-27, default ON, opt-out
                            // OXI_S1244_DISABLE): cell equations reserve the
                            // body S652 advance. S331 has forwarded them into
                            // cell.blocks since 2026-05-26, but BOTH cell
                            // passes dropped them — educational__002a301d's
                            // polar-coordinate answers (27 m:f fractions, 61
                            // of 63 oMath in cells) vanished: missing ink AND
                            // −67pt per solutions block → Problem 2 pulled up
                            // a page (EN-250 census pcd +1).
                            if std::env::var("OXI_S1244_DISABLE").is_err() {
                                let mfs: f32 = 10.5;
                                let (me, mbb) = crate::layout::math::emit_math_block(
                                    math_block,
                                    0.0,
                                    0.0,
                                    mfs,
                                );
                                if !me.is_empty() {
                                    let adv = LayoutEngine::s1244_math_advance(&me, &mbb, mfs);
                                    cell_content_h += adv;
                                    cell_content_h_visual += adv;
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if let Some((image_index, image)) = next_float.next() {
                    let (x, y) = LayoutEngine::cell_float_position(image, &float_tops, inner_w);
                    let origin = float_replay.map_or(0.0, |r| r.origin(row_idx, replay_cell_idx, image));
                    float_positions[image_index] = Some((x, y + origin));
                    continue;
                }
                if cell_float_flow {
                    let bottom = LayoutEngine::cell_float_obstacles_at(cell, &float_positions)
                        .iter().map(|obstacle| obstacle.bottom).fold(0.0_f32, f32::max) + pad_t;
                    cell_content_h = cell_content_h.max(bottom);
                    cell_content_h_visual = cell_content_h_visual.max(bottom);
                }
                cell_content_h += pad_b;
                cell_content_h_visual += pad_b;
                break (cell_content_h, cell_content_h_visual, center_extra);
                };
                if cell_float_flow { row_float_positions.insert(grid_idx, float_positions); }
                // COM-confirmed (2026-04-13, gen2_052): Word does NOT include
                // FULL inside-H border width in the row height calculation. The border
                // is drawn at the boundary between rows (overlapping). Including
                // full border_overhead caused 0.5pt/row cumulative drift (6 rows = 3pt).
                // cell_content_h += border_overhead;  // removed
                //
                // S375 (2026-05-27, FALSIFIED): S374 minimal repro showed Oxi rows
                // ~0.25pt SHORTER than Word per row; hypothesized half the shared
                // insideH border (0.25pt) per non-last row. Env-gated corpus test
                // CATASTROPHIC: Phase 2 0.9603→0.9270 (-0.0333), Phase 1 53→50.
                // Even HALF the border over-counts corpus-wide (gen2_052 found full
                // 0.5pt too much; half is still too much). The S374 -0.25pt is real
                // but repro-specific (that repro had no insideH so this gate didn't
                // even fire there) — NOT a corpus-wide pattern. Row overhead stays
                // at 0 (current behavior is corpus-correct). Gate kept OFF.
                if std::env::var("OXI_S375_HALF_INSIDEH").is_ok()
                    && !is_last
                    && table.style.has_inside_h
                {
                    cell_content_h += _border_overhead * 0.5;
                    cell_content_h_visual += _border_overhead * 0.5;
                }
                // S463 (2026-05-31, SHIPPED default-ON, opt-out OXI_S463_DISABLE):
                // the S375/2026-04-13 border-overhead dead-end was BLANKET (all
                // cells) — it regressed because CJK cells already over-snap
                // (b35123 +2pt/cell) so adding border height compounds. A
                // border-sweep minimal repro (tools/golden-test/repros/
                // gen2_lineheight, b0/b4/b8) shows Word DOES scale row pitch with
                // the inside-H border: Cambria 11pt single-line cell pitch =
                // 15.0(no border)/15.375(sz4=0.5pt)/15.75(sz8=1pt). Oxi already
                // adds ~0.16/row above the bare line, so the remaining deficit is
                // ~0.19/row = 0.375*border_width (block-calibrated: tbl_b4_sz22
                // OFF 91.31 -> Word 92.25 over 5 rows). This drives the gen2
                // English-template vertical drift (the dominant cause of their
                // ~0.81 SSIM vs OO/Libra 0.96). Discriminator (à la S455 is_cjk
                // scoping): apply ONLY to all-Latin tables in all-Latin documents.
                // CJK docs are excluded because there the (correct) overhead is
                // masked by a separate compensating error below the table, so
                // applying it regresses SSIM (gen2 JP family). Clean gate:
                // OFF 0.9098 -> ON 0.9119 (+0.0020), 30 improved / 0 regressed,
                // bottom-N flat, Phase-1 pagination unchanged.
                // Oxi OFF already adds ~+0.16/row above the bare line (15.16 vs
                // 15.0); Word wants 15.375 => remaining deficit ~0.19/row =
                // 0.375*border_width (block-calibrated on tbl_b4_sz22: OFF 91.31,
                // Word 92.25 => +0.94 over 5 rows). Scoped to all-Latin tables.
                // S940T regime (2026-07-19): the 0.375*bw was a CALIBRATED
                // proxy for the tombstone estimate's under-count (Word row =
                // hhea + border only — _pb_cellline_gen 42-config sweep).
                // Under the hhea estimate it double-counts (+0.188/row on
                // gen2_054), so it applies only in the tombstone regime.
                if std::env::var("OXI_S463_DISABLE").is_err()
                    && std::env::var("OXI_S940T_DISABLE").is_ok()
                    && table.style.has_inside_h
                    && table_is_latin
                    && !self.doc_body_has_cjk
                {
                    cell_content_h += _border_overhead * 0.375;
                    cell_content_h_visual += _border_overhead * 0.375;
                }

                if dump_table {
                    let ftext: String = cell
                        .blocks
                        .iter()
                        .filter_map(|b| {
                            if let Block::Paragraph(p) = b {
                                Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                            } else {
                                None
                            }
                        })
                        .collect::<Vec<_>>()
                        .join("|");
                    eprintln!("[CELL_DUMP] row={} span={} cell_w={:.2} inner_w={:.2} content_h={:.2} text={:?}",
                        row_idx, span, cell_w, inner_w, cell_content_h,
                        ftext.chars().take(24).collect::<String>());
                }
                row_height = row_height.max(cell_content_h);
                visual_row_h = visual_row_h.max(cell_content_h_visual);
                center_row_h = center_row_h.max(cell_content_h_visual + center_extra);
                grid_idx += span;
            }

            // S430 (2026-05-29, FALSIFIED — no code shipped): hypothesized that
            // row_height should contain the GRID-SNAPPED rendered content height
            // (visual_row_h, p2) instead of the natural estimate (p1), since the
            // render path already snaps cell lines when adjustLineHeightInTable
            // is set (mod.rs:6576) so a natural-sized row cannot contain its own
            // snapped content (b5f706: content_h=18 overflows row_h=17). Tested
            // env-gated `row_height = row_height.max(visual_row_h)` on the full
            // corpus: per-doc isolation showed ONLY b35123 moved — and it
            // CRATERED 0.9225→0.4453 (its cells already over-snap, +2pt each ×28,
            // so taller rows compound the over-height); b5f706 itself was FLAT
            // (its -9pt element_iou debt is per-LINE render height + matcher noise
            // per S417e, NOT row-container height); Phase 1 54→53. Confirms the
            // systemic finding's "inconsistent direction" (b35 cells too tall vs
            // b5f706 too short) — a single blanket grid-snap row rule cannot fix
            // both. adjustLineHeightInTable is near-universal (49/49 real docs)
            // so it is NOT a usable discriminator. Reverted; left as a tombstone
            // so this exact one-liner is not re-attempted.

            // Bug B Day 26 (Phase β step 1): row height snap removal.
            // COM-confirmed via R1-R6 ground truth (ffbd166): Word does NOT
            // grid-snap table row heights. All R1-R6 = natural sum/max.
            //
            // Day 19 (REVERTED, 2026-05-08) attempted same removal alone:
            // SSIM net -0.6412, 8 regressions. Diagnosis: cell line height
            // was still snapped (left at mod.rs:5144) → cell content
            // overflowed shrunk row.
            //
            // Day 26 plan: this step ALONE first, see specific regressions,
            // then proceed to step 2 (cell line snap gate by !in_table_cell).

            let mut minimum_row_height = None;
            // Apply trHeight constraint.
            // 2026-04-09 (COM re-verified, 0e7a contract sample table 1):
            //   <w:trHeight w:val="830"/> with NO w:hRule attribute →
            //   Word reports HeightRule = atLeast (1) and renders the row
            //   at exactly val (41.5pt = 830tw), not at content height.
            //   Treat missing hRule as atLeast to match Word behavior.
            if let Some(h) = row.height {
                match row.height_rule.as_deref() {
                    Some("exact") => {
                        // S1164 (2026-08-17, default ON, opt-out
                        // OXI_S1164_DISABLE): Word does NOT take an `exact` trHeight
                        // literally -- it adds the cell's BOTTOM margin.
                        // _pb_tblvert_gen.py, trHeight 1200tw = 60pt, compat 11,
                        // Word's own rules:
                        //   cellMar t/b   0/0    108/108  200/200  400/400
                        //   row height    60.00  65.40    69.96    80.04
                        //   addend         0.00   5.40     9.96    20.04
                        // and the asymmetric pair settles which margin it is:
                        //   top 400 / bottom 0  -> 60.00  (no addend)
                        //   top 0 / bottom 400  -> 80.04  (+20 = the BOTTOM one)
                        // Oxi returned the literal 60.00 in all six. Every other
                        // row shape in that probe already matches within 0.05
                        // (natural rows, uneven cells, atLeast above and below
                        // the natural need).
                        // Gate: probe 7/7 exact arms; Phase 1 95/96 with zero
                        // per-doc change; all 238 SSIM sentinel documents
                        // byte-identical (no corpus table pairs an exact
                        // trHeight with a non-zero bottom cell margin).
                        let s1164 = std::env::var("OXI_S1164_DISABLE").is_err();
                        row_height = if s1164 {
                            h + table
                                .style
                                .default_cell_margins
                                .as_ref()
                                .and_then(|m| m.bottom)
                                .unwrap_or(0.0)
                        } else {
                            h
                        };
                    }
                    // Default (None) or explicit "atLeast": atLeast semantics.
                    // S445 (2026-05-30, FALSIFIED — tombstone): 7ead52 has
                    // 860tw(43.0pt) atLeast rows whose VISUAL text-to-text pitch
                    // is 44.25pt (+1.25/row, accumulating -1.9 -> -11.65 over 8
                    // rows; cell_iou 0.79). Hypothesized Word renders binding
                    // atLeast+insideH rows taller than the logical trHeight
                    // (the prior 0e7a "renders at exactly val" note used
                    // Cell.Height = LOGICAL value, the S349/S361 trap) and that
                    // this was a universal systematic underestimate. Env-gated
                    // OXI_S445_ATLEAST_BUMP=1.25 (matches 7ead52 exactly) over
                    // the full corpus: CATASTROPHIC — net IoU 0.9692->0.9552,
                    // 19 docs DOWN / 5 up (31420af -0.3766, bd90b -0.176),
                    // Phase 1 54->52. The bump is DOC-SPECIFIC not universal:
                    // most docs render atLeast rows at ~the logical value; the
                    // +1.25 7ead52 needs overshoots them. Same wall as gen2_052
                    // / S374 / S375 (any per-row border/bump add regresses the
                    // corpus). 7ead52 is a trHeight-BINDING outlier; note its
                    // negative-drift neighbors (6514f2/d4d126/de6e32) are a
                    // DIFFERENT class (content-line-height, pitch 21>trH, b35
                    // class) — the "convergent negative-drift" assumption was
                    // false. Do NOT re-attempt a global atLeast bump.
                    _ => {
                        // ROWBOX2: a binding atLeast trHeight renders at
                        // trHeight + bw (_rowbox_sweep.py: 900tw=45.0 →
                        // Word 45.48; the bw pad sits on top of the
                        // trHeight box). Content already carries the bw
                        // via the generalized Round30 pad, so compare
                        // against trH + bw.
                        // (TRHCELL hypothesis FALSIFIED 2026-07-06: gating
                        // cell-tcBorders out of the trH+bw did NOT fix the
                        // kojin/2ea81a/bd90b00 flips — their rows are in
                        // table-bordered tables, and Word's OWN PDF renders
                        // their binding rows at trH+bw (kojin p3: 38.4/41.3/
                        // 49.56 vs trH 37.9/40.75/49.0; 2ea81a p1: 26.0 vs
                        // 25.5) — the flips are compensating-error exposures,
                        // not model errors. Cell tcBorders ≡ table insideH.)
                        let rb2_bw = self.rowbox2_trh_bw(table, row);
                        // Task T / S983 (2026-07-22) + S1065 (2026-08-03):
                        // Word adds the cell margins + border on top of a
                        // BINDING atLeast trHeight floor. Original S983
                        // narrow discriminator (fixed + insideH + 3 cells +
                        // trH in [20,24] + table-style cellMar>=5) shipped for
                        // forms__0020466f table 13 (Word = trHeight + 5 + 5 +
                        // 1), because a BLANKET "atLeast + table-style all
                        // cellMar" over-counted uklocalspending T5R3. But
                        // that blanket used the TABLE-STYLE default; the
                        // controlled probe (atleast_tcmar, 8 arms, Word PDF
                        // border truth) proves Word adds the EFFECTIVE cell
                        // margin — DIRECT cell tcMar wins, table-style is the
                        // fallback: c1_nom (no tcMar) = trH+bw 100.58, c1_57
                        // (direct 2.85x2) = 106.22, c1_100 (5x2) = 110.54,
                        // c1_toponly = trH+2.85+bw 103.46, c3_57 (3 cells,
                        // same tcMar) = 106.22, c1_400 (trH 20 binding) =
                        // 26.28, c1_small (trH<content) = content+tcMar+bw
                        // 18.84, c1_exact (hRule=exact) = 102.86 (separate
                        // rule). uklocalspending T5R3 declares DIRECT tcMar
                        // top/bottom = 0/0, so its EFFECTIVE vmar = 0 -> 51.5
                        // = unchanged, and the S983 over-count disappears.
                        // administrative__003381e4 table 1 (Job Details):
                        // binding trH 106.9 + direct vmar 5.7 + bw 0.5 =
                        // 113.1 = Word border truth 113.19 (Oxi was 107.4 ->
                        // body 6.2pt low -> heading to p2). The floor uses
                        // row_eff_vmar (max over non-merged cells); content-
                        // driven rows keep cell_content_h which already
                        // includes its own pad_t/pad_b.
                        // OXI_S1065_DISABLE is accepted as a synonym: S983's narrow rule was
                        // REPLACED by this one, so either name means "no vmar floor".
                        let s1065_vmar = if std::env::var("OXI_S983_DISABLE").is_err()
                            && std::env::var("OXI_S1065_DISABLE").is_err()
                        {
                            row_eff_vmar
                        } else {
                            0.0
                        };
                        // A minimum larger than the page fills the available
                        // page height; it does not create an off-page row box.
                        let minimum = (h + rb2_bw + s1065_vmar).min(content_height);
                        minimum_row_height = Some(minimum);
                        row_height = row_height.max(minimum);
                    }
                }
            }

            if row_height == 0.0 {
                let metrics = &*self.doc_default_metrics();
                row_height = self.line_height_inner(
                    self.default_font_size,
                    None,
                    None,
                    metrics,
                    true,
                    table_grid_pitch,
                    true,
                );
            }
            if dump_table {
                let trh = row.height.unwrap_or(0.0);
                let trh_rule = row.height_rule.as_deref().unwrap_or("(none)");
                eprintln!(
                    "[TBL_DUMP] row={} entry_cursor_y={:.3} row_height_pre={:.3} trHeight={:.3} rule={} n_cells={}",
                    row_idx, row_entry_cursor_y, row_height, trh, trh_rule, row.cells.len()
                );
            }
            // Page break check: if this row won't fit, push current page and reset
            // Allow break if there are elements from previous rows OR from before the table
            // A positioned table with an explicit flow-fit anchor owns a new
            // layout area. The surrounding story is preserved in current_elements,
            // but cannot force its first row to leave that area before splitting.
            let independent_float_start = flow_fit_offset.is_some()
                && row_idx == 0 && pages.len() == flow_entry_page;
            let has_content = !elements.is_empty()
                || (!independent_float_start && !current_elements.is_empty());
            // S740: shrink the page bottom by the notes committed by PRIOR rows
            // on this page. The row's OWN notes are excluded from its fit test:
            // Word render-truth (probeqfncell p1) keeps the referencing row and
            // SPILLS its footnote to the next page when the note area is full
            // (row12 stays, 表注12 renders on p2) — a line is never pushed by
            // its own note, only by the area accrued from earlier references.
            // S992 (R1, 2026-07-23, default ON, opt-out OXI_S992_DISABLE): H-ink
            // footer collision. An EMPTY atLeast row whole-pushed by a binding
            // footer may descend into the footer first line's ascent-leading
            // (the body limit is the footer glyph-INK, not the footer line-box
            // top). DERIVED _pb_r1_footer_collision_gen.py (Word COM+PDF):
            // max row-bottom Word keeps = footer_ink − ~0 (tracks the ink,
            // H-ink), the leading = (ascent − 0.71)×footer_fs. forms__001c51d6
            // row35 (empty atLeast, footer table 11pt) is rejected by 2.13pt =
            // exactly this leading. SCOPED to EMPTY atLeast rows only — the
            // global content-height version fit uk_local_spending's BODY TEXT
            // lines where Word pushes (the H-empty-row discriminator, report §8).
            // Latin-scoped (footer_ink_relief); s755_footer_geom only computed
            // for the rare empty-atLeast row.
            let s992_relief = if !self.doc_body_has_real_cjk
                && row.height.is_some()
                && row.height_rule.as_deref() != Some("exact")
                && row.cells.iter().all(|c| {
                    c.blocks.iter().all(|b| {
                        matches!(b, Block::Paragraph(p)
                        if p.runs.iter().all(|r| r.text.trim().is_empty()))
                    })
                }) {
                let (fr, _) = self.s755_footer_geom(&page.footer, page);
                if fr > page.margin.bottom + 0.05 {
                    self.footer_ink_relief(&page.footer)
                } else {
                    0.0
                }
            } else {
                0.0
            };
            let mut first_page_fit_offset = if pages.len() == flow_entry_page {
                // Capacity belongs to the text anchor; a lower visual origin
                // can put the table below the body margin without consuming
                // additional lines in the anchor's flow.
                flow_fit_offset.unwrap_or(0.0)
            } else { 0.0 };
            let mut page_bottom = page_top + content_height - s740_reserve + s992_relief - first_page_fit_offset;
            s740_pending_commit = Some(row_idx);
            // Test the whole rendered row against the remaining space without
            // changing the height used for cell alignment before rendering.
            let row_fit_height = if measure_cant_split_fit {
                row_height.max(center_row_h)
            } else {
                row_height
            };
            let s1191_foot = if std::env::var("OXI_TABLE_FOOT_FIT_DISABLE").is_err() {
                if separate_outer_edges {
                    self.table_fragment_bottom_width(table, Some(row))
                } else if row_idx + 1 == table.rows.len()
                    && self.s1191_on() && self.s1191_table_needs_foot(table) {
                    self.s1191_foot_bw(table)
                } else if std::env::var_os("OXI_S1581_DISABLE").is_none()
                    && self.s1191_on() && self.s1191_table_needs_foot(table) {
                    // S1581 (2026-09-27, default ON, opt-out OXI_S1581_DISABLE): a row
                    // that ENDS a page fragment closes it with its own bottom rule,
                    // which must fit too (`_pb_rowfoot_pdf_gen.py`: a row stays iff
                    // its bottom boundary + half the rule width <= the body bottom,
                    // boundaries measured at rule centres = the rule's lower edge in
                    // Oxi's rule-top convention). technical__014819 p4 row 11: rule
                    // top 704.25 + 15.5 + 0.5 = 720.25 > 720, Word sends it to p5.
                    // Only cell-bordered tables: S870 pads a row from the rule ABOVE
                    // it, so the row box never held its own bottom rule.
                    self.table_fragment_bottom_width(table, Some(row))
                } else { 0.0 }
            } else { 0.0 };
            let row_fit_height = row_fit_height + s1191_foot;
            // A terminating merged cell contributes to the atomic row fit
            // before pagination, as well as to its final border height.
            let row_fit_height = if row.cant_split {
                let bygrid = std::env::var_os("OXI_S1192G_DISABLE").is_none();
                vmerge_absolute_ends.iter().fold(row_fit_height, |height, (key, end)| {
                    let is_cont = |c: &TableCell| {
                        matches!(c.v_merge.as_deref(), Some("continue") | Some(""))
                    };
                    let here = LayoutEngine::s1192_cell_at(row, *key, bygrid).map_or(false, is_cont);
                    let next = table.rows.get(row_idx + 1)
                        .and_then(|r| LayoutEngine::s1192_cell_at(r, *key, bygrid))
                        .map_or(false, is_cont);
                    if here && !next {
                        height.max(end - (pages.len() as f32 * vmerge_coordinate_stride + cursor.visual_y))
                    } else { height }
                })
            } else { row_fit_height };
            let mut row_overflows = cursor.cursor_y + row_fit_height > page_bottom;
            // R7.47 (Day 34 part 16, 2026-05-13): row-level SOFT LRPB. When
            // ANY cell's FIRST paragraph carries `<w:lastRenderedPageBreak/>`
            // on its run[0], Word's saved render broke before this row.
            // Mirrors the body-paragraph LRPB SOFT rule at mod.rs:1888.
            // de6e / 29dc6e outliers (4 each) had LRPB-at-cell-start markers
            // that the table layout previously ignored.
            // Only a pass that trusts saved breaks may use this local hint.
            // A natural-layout pass must not consult saved cell page breaks:
            // those hints cannot validate the page count used to trust them.
            let cell_lrpb_enabled = !self.lrpb_count_distrust.get()
                || std::env::var_os("OXI_CELL_LRPB_TRUST_DISABLE").is_some();
            let row_has_lrpb_at_cell_start = cell_lrpb_enabled && row.cells.iter().any(|cell| {
                cell.blocks.first().map_or(false, |b| match b {
                    Block::Paragraph(p) => p
                        .runs
                        .first()
                        .map(|r| r.has_last_rendered_page_break)
                        .unwrap_or(false),
                    _ => false,
                })
            });
            let consumed_row = cursor.cursor_y - page_top;
            // R7.48 (2026-05-13): tighten R7.47 threshold from > 0.5 to > 0.85.
            // OXI_DUMP_ROW_LRPB traces showed Oxi cursor_y/content_height at the
            // firing point: de6e fires at 0.904, 29dc6e at 0.868 (correct PASSes),
            // a1d6 at 0.812 (stale LRPB — Word's current render doesn't break here).
            // 0.85 cleanly separates correct firings (page near full) from stale
            // hints fired around mid-page.
            let lrpb_threshold = content_height * 0.85;
            let lrpb_row_should_break = row_has_lrpb_at_cell_start
                && (self.doc_body_has_real_cjk || std::env::var("OXI_LATIN_ROW_CACHE_ENABLE").is_ok())
                && has_content
                && !row_overflows  // row would fit; LRPB hint says break anyway
                && consumed_row > lrpb_threshold
                // S904 (2026-07-17, opt-out OXI_S904_DISABLE): LATIN docs
                // respect a row-level SOFT LRPB only when the row's own fit
                // is MARGINAL (post-row room < 28pt, the S814-v2 family
                // constant) — a REAL saved break fires where the row barely
                // fits (ukhealthform E3: room-after-row 2.8pt, load-bearing
                // PASS), a STALE one replays with the page wide open
                // (0008ea8f checkbox table: room-after-row 82pt, Word's
                // current render keeps it → the whole {+1:13}). The blanket
                // Latin retirement (v1) flipped ukhealthform PASS→FAIL. JP
                // keeps the unconditional model (de6e/29dc6e load-bearing
                // per the R7.47/R7.48 derivation).
                && (self.doc_body_has_real_cjk
                    || page_bottom - (cursor.cursor_y + row_height) < 28.0
                    || std::env::var("OXI_S904_DISABLE").is_ok());
            if std::env::var("OXI_DUMP_ROW_LRPB").is_ok() && row_has_lrpb_at_cell_start {
                let preview: String = row
                    .cells
                    .iter()
                    .filter_map(|c| {
                        c.blocks.first().and_then(|b| match b {
                            Block::Paragraph(p) => Some(
                                p.runs
                                    .iter()
                                    .flat_map(|r| r.text.chars())
                                    .take(20)
                                    .collect::<String>(),
                            ),
                            _ => None,
                        })
                    })
                    .next()
                    .unwrap_or_default();
                eprintln!("[ROW_LRPB] row_idx={} cursor_y={:.2} row_h={:.2} page_bot={:.2} consumed_frac={:.3} row_overflows={} fire={} text={:?}",
                    row_idx, cursor.cursor_y, row_height, page_bottom,
                    consumed_row/content_height, row_overflows, lrpb_row_should_break, preview);
            }
            // Row splitting: when cantSplit=false (default) and the row overflows,
            // split it across pages rather than moving the entire row to the next page.
            // Word splits rows at the page boundary, keeping partial content on each page.
            //
            // R7.58 (Day 35 session 58, 2026-05-13): mid-row LRPB positive-evidence
            // gate. Word's `<w:lastRenderedPageBreak/>` placement tells us how Word
            // broke this row in its last saved render:
            //   - LRPB at cell=0, first paragraph, run=0: row was PUSHED whole
            //   - LRPB at any other position in row: row was SPLIT (break mid-row)
            //   - No LRPB in row: row was NOT broken by Word
            // Only enable multi-cell split when we have POSITIVE evidence Word split
            // (mid-row LRPB). Otherwise retain push-whole behavior (gate (1) un-gated
            // attempt caused 4 PASS→FAIL regressions: 29dc6e/31420af/6514/de6e, all
            // had LRPB-at-start or no LRPB; gate (2) inverted check failed because
            // it allowed split for no-LRPB rows that Word didn't break).
            // Single-cell 1x1 box tables retain prior unconditional split.
            let is_single_cell_row = row.cells.len() == 1 && num_rows == 1;
            // S993 (R2/R3, 2026-07-23, opt-out OXI_S993_DISABLE): a fixed-layout,
            // 3-cell, auto-height row in a Latin doc reads the EXACT mid-run LRPB
            // fragment anchor (the `lrpb_before` tuple flag) instead of the R7.73
            // next-paragraph approximation. This is the Word-truth-measured
            // geometry for reference__0052ba53 (3 markers, run 12 / run 2 / run 2)
            // and technical__002c1ffa; census = 2 EN docs, 0 golden / 0 JP, so
            // every other doc keeps the approximation → byte-identical.
            let s993_exact = !self.doc_body_has_real_cjk
                && table.style.layout.as_deref() == Some("fixed")
                && row.cells.len() == 3
                && row.height.is_none()
                && std::env::var("OXI_S993_DISABLE").is_err();
            // Scan row content for an LRPB that is NOT at the row-start position
            // (cell=0, first paragraph block in cell, run=0).
            let has_lrpb_mid_row = {
                let mut found = false;
                for (ci, cell) in row.cells.iter().enumerate() {
                    let mut first_para_in_cell_seen = false;
                    for block in cell.blocks.iter() {
                        if let Block::Paragraph(p) = block {
                            let is_first_para_in_cell = !first_para_in_cell_seen;
                            first_para_in_cell_seen = true;
                            for (ri, run) in p.runs.iter().enumerate() {
                                if cell_lrpb_enabled && run.has_last_rendered_page_break
                                    && !(ci == 0 && is_first_para_in_cell && ri == 0)
                                {
                                    found = true;
                                    break;
                                }
                            }
                            if found {
                                break;
                            }
                        }
                    }
                    if found {
                        break;
                    }
                }
                found
            };
            // R7.74 (Day 37, 2026-05-15): Word's implicit "table-start widow protection"
            // for HEADING-style tables (single-row + single-cell, content longer
            // than the row's available space). When such a table starts near the
            // page bottom, Word pushes it entirely to the next page even without
            // explicit keepNext/cantSplit. COM-confirmed on d4d126 T5 (25 paragraphs
            // in 1 cell of 1 row); 04b88e's multi-row form tables do NOT have this
            // implicit widow → restrict to single-row single-cell.
            let is_single_row_single_cell =
                table.rows.len() == 1 && table.rows.get(0).map_or(false, |r| r.cells.len() == 1);
            // S694 (2026-06-29, default ON, opt-out OXI_S694_DISABLE): lower the
            // R7.74 single-row-single-cell table-start widow threshold from pitch*4
            // to pitch*2.2 (GRID docs only; no-grid kept at 58.0). A long
            // regulation-box row whose start leaves ~3+ lines free at the page
            // bottom SPLITS in Word; the pitch*4 threshold pushed it WHOLE. This is
            // the tokyoshugyo 賃金-chapter "over-tallness" the S693 給月給 fix exposes:
            // under S693 the 精勤手当/賞与 box starts at free≈60pt (3.3 lines) and the
            // old pitch*4=72pt threshold widow-pushed it → a 137pt blank on p48 →
            // the whole lower chapter shifts +1. ★DERIVED window (not tuned): 3a4f's
            // LARGEST Word-PUSHES single-row box (母性健康管理, free 29.9) and
            // tokyoshugyo's Word-SPLITS box (精勤手当, free 59.9) bound the threshold;
            // the tokyoshugyo-improving window is pitch*[2.0, 2.4], and pitch*2.2
            // (~39.6pt) is its center, 9.7pt above 3a4f's 29.9 (canary 3a4f/model/
            // d4d126 byte-identical — their push-boxes all free <30 < 39.6). OXI_WIDOW_K
            // overrides for sweeping; OXI_S694_DISABLE restores the pre-S694 pitch*4.
            let widow_k_default: f32 = if std::env::var("OXI_S694_DISABLE").is_ok() {
                4.0
            } else {
                2.2
            };
            let widow_k = std::env::var("OXI_WIDOW_K")
                .ok()
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(widow_k_default);
            // Widow protection applies to a fragment of an overflowing row;
            // a complete row that fits does not leave a widow.
            let widow_break_needed = row_overflows && row_idx == 0 && has_content && is_single_row_single_cell && {
                let free_space = page_bottom - cursor.cursor_y;
                let widow_threshold = if let Some(pitch) = table_grid_pitch {
                    pitch * widow_k
                } else {
                    0.0
                };
                let fire = free_space > 0.0 && free_space < widow_threshold;
                if std::env::var("OXI_DBG_WIDOW").is_ok() && row_overflows {
                    let prev: String = row
                        .cells
                        .get(0)
                        .and_then(|c| {
                            c.blocks.iter().find_map(|b| match b {
                                Block::Paragraph(p) => Some(
                                    p.runs
                                        .iter()
                                        .flat_map(|r| r.text.chars())
                                        .take(14)
                                        .collect::<String>(),
                                ),
                                _ => None,
                            })
                        })
                        .unwrap_or_default();
                    eprintln!(
                        "[WIDOW] cur_y={:.1} free={:.1} thr={:.1} fire={} single={} txt={:?}",
                        cursor.cursor_y,
                        free_space,
                        widow_threshold,
                        fire,
                        is_single_row_single_cell,
                        prev
                    );
                }
                fire
            };

            // S533 (2026-06-10): a row carrying an inline IMAGE block is pushed
            // WHOLE to the next page instead of splitting (when it fits a fresh
            // page). Word treats the image's line as atomic — 3a4f's calendar
            // row (466pt cell: 321.75pt EMF + paragraphs, trHeight 7910 atLeast)
            // starts fresh on Word's p34; Oxi's element-level split stranded the
            // image across the boundary and the post-split cursor under-advanced,
            // overlapping the following content. Rows taller than a full page
            // still split (unavoidable).
            let row_has_image_block = row
                .cells
                .iter()
                .any(|c| c.blocks.iter().any(|b| matches!(b, Block::Image(_))));
            // S998 (2026-07-25, default ON, opt-out OXI_S998_DISABLE): the S533
            // veto above ("a row carrying an image is pushed WHOLE") is TOO
            // BROAD — it fires on any row that contains an image anywhere, but
            // Word only keeps the IMAGE LINE atomic, not the text around it. In
            // technical__0061c884 table 12 row 2 the image is the 10th of 11
            // cell blocks (9 text paragraphs before it, 1 after); Word splits
            // the row — paragraphs 0..8 on p10, image + the trailing paragraph
            // on p11 — while Oxi whole-pushed all 11 blocks to p11, cascading
            // +1 from p10 to the last page (15 vs Word 14). The DISCRIMINATOR
            // is an INTERIOR image: real (non-whitespace) content BOTH before
            // the first image AND after the last image in a cell. A TERMINAL
            // image (educational__00161422's 5 image rows, image = last block)
            // keeps the whole-push (its 14 pages match Word, at least one push
            // is load-bearing) — so the after-content half of the predicate is
            // load-bearing. The image element itself is atomic (a single
            // LayoutElement), so the existing element-level splitter moves it
            // whole to the next page when it overflows; verified the image does
            // not straddle the boundary and the post-split cursor advances (the
            // S533-era stranding bug does not recur for the interior case).
            // SCOPE: static candidates = 3 rows / 3 docs (kyotei36spec /
            // reference__0035761e / target); only the target's row actually
            // reaches ROWPUSH (row_overflows) — golden/real_en/JP are
            // byte-identical by construction.
            let s998_interior_image = std::env::var("OXI_S998_DISABLE").is_err()
                && row.cells.iter().any(|cell| {
                    let first_image = cell
                        .blocks
                        .iter()
                        .position(|b| matches!(b, Block::Image(_)));
                    let last_image = cell
                        .blocks
                        .iter()
                        .rposition(|b| matches!(b, Block::Image(_)));
                    match (first_image, last_image) {
                        (Some(first), Some(last)) => {
                            let real_para = |b: &Block| {
                                matches!(b, Block::Paragraph(p)
                                if p.runs.iter().any(|r| !r.text.trim().is_empty()))
                            };
                            // S1129 (2026-08-15, opt-in OXI_S1129=1): the BEFORE
                            // half of the predicate is not Word's rule. Probe
                            // _pb_cellimgtail_gen.py (2-cell row, cell = image +
                            // 9 tail lines, nothing before the image, 12 filler
                            // sweeps) shows Word filling the page with the tail:
                            // 7 / 5 / 3 / 1 / 0 lines kept as the row walks down,
                            // one line lost per 13.8pt — keep-all-that-fit, no
                            // whole-push. Oxi keeps NONE in every arm because
                            // this predicate needs content before the image too.
                            // What the recorded evidence actually needs is the
                            // AFTER half alone: educational__00161422's terminal
                            // images (nothing after) keep the whole-push, and
                            // the S998 target (paras before AND after) splits.
                            // educational__002a301d p5 is this probe's shape and
                            // loses 2 pages to it.
                            let after = cell.blocks[last + 1..].iter().any(real_para);
                            if std::env::var("OXI_S1129").is_ok() {
                                after
                            } else {
                                cell.blocks[..first].iter().any(real_para) && after
                            }
                        }
                        _ => false,
                    }
                });
            // S1168 (2026-08-19, default ON, opt-out OXI_S1168_DISABLE) retires
            // this veto.
            // Pushing every image-bearing row whole was S533's stand-in for
            // "an image is atomic", adopted because the element splitter
            // stranded images across the boundary -- but the atomicity is
            // GEOMETRIC, and the split site enforces it directly by refusing
            // to let the break line cross an image.
            // The probe settles that the veto is not Word: in
            // `_pb_cellimgtail_gen.py` the cell is [image, tail lines] with
            // NOTHING above the image, and Word still splits it, keeping the
            // image plus 3/2/1/0 tail lines as the image walks down the page.
            // Oxi kept ZERO in all seven arms under the veto; opt-in it
            // reproduces Word's 3/2/1/0 with image bottoms inside 0.5pt.
            // It also corrects two things the ledger had backwards: S998's
            // "terminal image" reading (002a301d's image IS terminal and Word
            // splits it anyway) and the before/after-content predicate (the
            // probe cell has nothing before its image). What makes
            // educational__00161422 push whole is not what surrounds the image
            // -- it is that the image does NOT FIT, so the line must move above
            // it and lands at the row top.
            //
            // ★GATE: shipped WITH S1169, and only together. Alone, S1168 cost
            // educational__00161422 its PASS (1.0000 -> 0.9824) -- not because
            // the rule is wrong but because retiring the veto stopped hiding a
            // dropped trailing-<w:br/> line in cells, which S1169 fixes. As a
            // bundle the EN Phase-1 A/B is n_pass 221 -> 221 with NO doc going
            // PASS -> FAIL, and two improve: 002a301d 0.6846 -> 0.9692 (pcd
            // +2 -> +1) and legal__0010437a7f75f636 0.9452 -> 0.9952 (pcd
            // -2 -> -1). Do not ship one without the other.
            // ★The residual is NOT a gap in this rule -- it is a pre-existing
            // drift the veto used to hide. Measured, not assumed:
            // `_pb_imgrowfit_gen.py` (k lines above a NON-fitting image, image
            // last, row walked down a line at a time, k = 1/2/3/6/9 x 4 offsets)
            // shows Word keeping ALL k lines in ALL 20 arms -- even a single
            // line with only 136pt free. "Keep everything above that fits" is
            // unconditional, which is what S1168 already does.
            // The Word PDF of 00161422 then says its p12 is essentially FULL
            // (last line y=535.40 against a 559.32 bottom) and p13 opens with
            // "individually on the front of the sheet" -- Word SPLITS that row
            // too. The engines differ on the CELL CONTENT, not the break: Word
            // puts a 27.6pt gap before the numbered paragraph "5. Allow about
            // ..." (476.70 -> 504.30) where Oxi runs a uniform 14.7pt pitch with
            // no paragraph gap, so Oxi fits 29 lines where Word fits 26 and the
            // boundary lands one line late.
            // (An earlier reading of "Word leaves 230pt blank" was a COM-dump
            // artefact: `paragraphs` enumerates per CELL, so the last entry for
            // a page is not the lowest thing on it. Use the PDF for page-bottom
            // questions -- the same trap as the Info6 caveat.)
            // So default-ON waits on that spacing bug, not on this rule.
            let image_atomic_push = std::env::var("OXI_S1168_DISABLE").is_ok()
                && row_has_image_block
                && row_height <= content_height
                && !s998_interior_image;

            // needs_row_split: only when overflow + table allows split.
            // widow_break_needed overrides split — we want the whole table on next page.
            // S754 (2026-07-06, default ON, opt-out OXI_S754_DISABLE): Word's default
            // "allow row to break across pages" SPLITS a multi-cell row at the
            // page boundary whenever at least one line of it fits above
            // (probethdr/probevmerge/probezgridspan/probeztbrlv Word truth —
            // fresh docx carry NO LRPB so the mid-row-LRPB evidence gate left
            // them whole-moving, +1 cascades). Guards: (a) respect explicit
            // whole-push evidence (row-start LRPB = Word pushed this row
            // whole); (b) require one grid line of room above the boundary
            // (Word moves the row whole when not even one line fits).
            // ★DISCRIMINATOR (Word PDF render truth, 7 specimens): Word splits
            // only AUTO-height rows (no explicit trHeight). All 4 probe row
            // families (auto) split; 29dc6e ③ row (trH 63.3 atLeast, fits 21.1),
            // de6e32 （３） row (trH 290.65 exact, fits 24.2) AND tokyoshugyo's
            // （参考） box row (trH 474.4 atLeast, fits 295.9!) are all pushed
            // WHOLE — Word starts a trHeight row on a fresh page even when
            // 296pt of the current page is free (p19 left ~330pt empty).
            // Fit threshold, TWO-TIER: a SINGLE-COLUMN row (cells==1, a prose
            // box — Word applies a widow-like keep) pushes whole below the S694
            // window pitch×2.2 (tokyoshugyo 遅刻早退 box fits=27.4=1.52 lines
            // PUSHED; 3a4f box 29.9 PUSHED; 精勤手当 59.9 SPLIT); a MULTI-CELL
            // data row splits with just one line of room (all 4 probe families).
            let s754_min_fit = if row.cells.len() == 1 {
                match table_grid_pitch {
                    Some(p) => p * 2.2,
                    // S999 (probe _pb_s754, default ON, opt-out OXI_S999_DISABLE): Word SPLITS a
                    // no-TYPE-docGrid / no-grid single-column row line-by-line
                    // (keep-all-that-fit, NO widow-keep) — measured sz24 keeps
                    // 5/4/3/2 lines as R=78/64/51/37, sz28 5..1 as R=87..23,
                    // one line dropped per pitch of R, no whole-push floor. The
                    // 58.0 fallback was a TYPED-grid prose-box widow-keep
                    // heuristic (tokyoshugyo 遅刻早退/3a4f boxes push at <~2
                    // lines) that does NOT apply to a no-type row. technical__
                    // 0061c884 p5 (R=41.1) Word keeps 2 lines; Oxi's 58.0
                    // whole-pushed it (+1). Use the multi-cell ~1-line floor.
                    // 0061c884 p5 (R=41.1) Word keeps 2 lines; Oxi's 58.0
                    // whole-pushed it (+1).  The S559 exposure recorded when
                    // this was first derived (-1x3 on the same doc) is GONE —
                    // the p6/p7 under-reservation it needed was closed by the
                    // S1093-S1095 inline-object work.
                    None if std::env::var("OXI_S999_DISABLE").is_err() => 14.0,
                    None => 58.0,
                }
            } else {
                let legacy = table_grid_pitch.unwrap_or(14.0);
                // S1424 (2026-09-16, default ON, opt-out OXI_S1424_DISABLE): a
                // multi-cell row splits when ONE natural cell line fits, not one
                // grid pitch. `_pb_rowfit_gen.py` split arms at 10.5 / 12 / 14pt
                // (cell lines 13.5 / 15.75 / 18 on an 18pt grid, bottom 785.2):
                // line 1 stays at top 769.5 / 767.25 / 765.75 and goes over at
                // 770.25 / 769.5 / 767.25 -- the line box, never the pitch. The
                // 10.5pt row with 16.5pt of room split in Word and moved whole here.
                // CJK-body only, like S1423: the Latin arm has its own floor
                // (legal__001410a84d3ead5f PASS -> FAIL, -1 x4, when unscoped).
                // S1485 (2026-09-19, default ON, opt-out OXI_S1485_DISABLE): the
                // same floor for a LATIN body. The Latin exclusion was covering
                // legal__001410a8's foot-only overflow (S1482); with that fixed
                // the one-line floor is Word's rule there too (`rowsplitmin.py`,
                // TNR 8-14pt: Word splits with one line box + border of room).
                if (self.doc_body_has_real_cjk || std::env::var_os("OXI_S1485_DISABLE").is_none())
                    && std::env::var_os("OXI_S1424_DISABLE").is_none() {
                    let one_line = row
                        .cells
                        .iter()
                        .filter_map(|c| c.blocks.iter().find_map(|b| match b {
                            Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.trim().is_empty()) => Some(p),
                            _ => None,
                        }))
                        .map(|p| self.estimate_para_height(
                            p, 1.0e6, row_line_pitch, table.style.para_style.as_ref(), true,
                            grid_char_pitch, grid_char_cw_ratio,
                        ))
                        .filter(|h| *h > 0.5)
                        .fold(f32::INFINITY, f32::min);
                    if one_line.is_finite() { legacy.min(one_line) } else { legacy }
                } else {
                    legacy
                }
            };
            // A floating table starts its own layout area. If no first cell
            // line fits at its text anchor, Word keeps the first row in that
            // area. Use the measured line and top inset, not a fixed font-size
            // floor; this is independent of the ordinary row splitting policy.
            let float_first_line_fit = if first_cell_line_fit.is_finite() {
                first_cell_line_fit
            } else { s754_min_fit };
            let first_row_forced = row_idx == 0 && flow_fit_offset.is_some()
                && page_bottom - cursor.cursor_y < float_first_line_fit;
            if first_row_forced {
                row_overflows = false;
            }
            // S814 (2026-07-13, experiment OXI_S814=1): the row-start-LRPB
            // veto blocks splitting even when the LRPB is STALE (uklocal
            // Annex row 2: fresh Word SPLITS the row, leaving 2 lines on
            // p36; the saved break is from different geometry) — the wp36/
            // 50/51 +1x16. The veto's necessity for the JP corpus is being
            // measured; if no JP doc relies on it, it is removed.
            let s814_no_veto = std::env::var("OXI_S814").is_ok();
            // S814 v2: the veto holds only when the fresh geometry AGREES with
            // the whole-push (remaining space too small to place a meaningful
            // split anyway). When >= s814_k remains, the saved whole-push
            // evidence contradicts the current flow (uklocal row 2: 33.2pt
            // free, fresh Word splits 2 lines onto p36) = stale -> split.
            // Row 21 (remaining ~1.3pt) keeps its veto = whole-push = Word.
            let s814_k: f32 = std::env::var("OXI_S814_K")
                .ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(28.0);
            let s814_lrpb_veto = row_has_lrpb_at_cell_start
                && !s814_no_veto
                && (page_bottom - cursor.cursor_y) < s814_k;
            // S864: an explicit atLeast row whose height is driven
            // by a run of trailing empty cell paragraphs is still splittable.
            // Word keeps the visible cell content above the boundary and lets
            // the empty tail continue on the next page; treating every trHeight
            // row as atomic pushes the whole visible row forward.
            // ★HELD OPT-IN (OXI_S864B=1, default OFF): shipping this default-ON
            // broke the JP corpus — 29dc6e8943fe (order_01) went PASS 1.0000 →
            // FAIL 0.9910, i.e. Phase-1 stopped being 87/87. Bisected to THIS
            // part alone (A/C..F leave 29dc6e at 1.0). It IS needed by its
            // target (administrative__0001ce58: with it PASS 1.0, without it
            // FAIL), so it is a JP-PASS-for-EN-PASS trade, which the merge gate
            // forbids (0 PASS→FAIL).
            // ★NO STRUCTURAL DISCRIMINATOR EXISTS (measured): 29dc6e's 14
            // matching rows and administrative's 1 are IDENTICAL in shape —
            // hRule absent (=atLeast), 2 trailing empty paras, 2-3 cells, and
            // the trHeight ranges overlap (administrative 944tw sits inside
            // 29dc6e's 255..2400). So the row shape cannot separate "Word
            // splits" from "Word keeps whole".
            // ★The rule also does not match its own stated intent: "an explicit
            // atLeast row whose height is DRIVEN BY a run of trailing empty
            // paragraphs" — but the condition never checks that the empties
            // drive the height (only that they exist), and `!= Some("exact")`
            // also admits rows with no trHeight at all. Per the no-exception
            // rule the spec is wrong and must be re-derived (candidate: compare
            // the row's natural content height against trHeight so only rows
            // actually GROWN by the empty tail split) before it can ship.
            // ★S864B-LATIN (2026-07-17, default ON via s864_part("B"), opt-out
            // OXI_S864B_DISABLE): the "NO STRUCTURAL DISCRIMINATOR" the held
            // note lamented MISSED the document-language axis. The trade was
            // JP-29dc6e-PASS→FAIL for Latin-administrative__0001ce58-PASS; the
            // two rows ARE identical in shape, but administrative is a Latin
            // (EN) doc (CJK 0) and 29dc6e is JP. Gating the split on
            // `!doc_body_has_real_cjk` fires on administrative (→ PASS 1.0) and
            // leaves EVERY JP form (29dc6e/tokumei/…) BYTE-IDENTICAL by
            // construction — 29dc6e never enters the branch. The row-split
            // decision may still be imperfect (the intent-mismatch the note
            // flags remains, so it is Latin-scoped rather than universal), but
            // within the Latin corpus it is a clean +1 doc with 0 regressions.
            // S1029 (2026-07-28, default ON, opt-out OXI_S1029_DISABLE): the
            // empty-tail split applies to a SINGLE-cell row of a MULTI-row
            // table too. forms__00160757 Member-address row (1 cell, trHeight
            // 499tw atLeast non-binding, text para + 2 run-less empties whose
            // after=160/line=259 fall to docDefaults): Word KEEPS the
            // docDefaults spacing (the s1029_cellsp probe: every variant
            // incl. multi-empty tails keeps FULL docDefaults — V9 80.40 vs
            // model 80.26, V11 103.78/103.68, so the row IS ~59pt tall), and
            // at the page bottom (free 46 < the S941 window 58) Word SPLITS
            // after the first empty (PDF row box 703.1→739.8 = 13.8 text +
            // 14.89 line259 + 8 after EXACTLY) and COLLAPSES the empties-only
            // continuation (p2 opens directly with the City row at the
            // margin). The `cells.len() > 1` gate whole-pushed it → Member
            // address +1 and a phantom p3. 1×1 tables (is_single_cell_row)
            // keep their dedicated split machinery (harassbun class).
            let s864_empty_tail_split = s864_part("B")
                && !self.doc_body_has_real_cjk
                && (row.cells.len() > 1
                    || (row.cells.len() == 1
                        && !is_single_cell_row
                        && std::env::var("OXI_S1029_DISABLE").is_err()))
                && row.height_rule.as_deref() != Some("exact")
                && row.cells.iter().any(|cell| {
                    cell.blocks
                        .iter()
                        .rev()
                        .take_while(|b| {
                            matches!(b, Block::Paragraph(p)
                        if p.runs.iter().all(|r| r.text.is_empty()))
                        })
                        .count()
                        >= 2
                })
                && (page_bottom - cursor.cursor_y) >= table_grid_pitch.unwrap_or(14.0);
            // S941 (2026-07-19, ★BUNDLE MEMBER — fires only with OXI_S940;
            // opt-out OXI_S941_DISABLE): a NON-BINDING atLeast trHeight row
            // splits when the free space clears the single-column widow
            // threshold. Word truth (uklocal rt.pdf, landscape template):
            // rows 15/20/24 (trH 51-63.75, content-driven) PUSH with 57/40/38
            // free; row 27 (trH 39, content ~116) SPLITS with ~110 free —
            // the discriminator is the FREE SPACE (K ∈ (57, 110], = the
            // no-grid single-column 58.0 / pitch×2.2 tier), NOT trH-vs-auto
            // alone. The rule presupposes Word-correct row heights (the free
            // space must be measured in the S940 geometry — at the default
            // short-line heights Oxi reaches row 15 with 86 free and would
            // mis-split), so it ships with the S935/S936/S940 set. Binding
            // rows (row_height == trh) keep the S754 whole-push (tokyoshugyo
            // trH 474 free 296 pushes — JP is excluded by scope anyway).
            // S1025 (2026-07-27, default ON, opt-out OXI_S1025_DISABLE): the
            // local-recompute for Origin A (REPORT_legal__001410a8_forms_keepNext_
            // table §6.3) — decouple S941+S942 (the non-binding atLeast row split
            // + continuation) from the OXI_S940T bundle, WITHOUT its piece-1 cell
            // estimate (the hhea line height, mod.rs:27906, which is what regresses
            // the gen2 word_png family via device-snap). legal__001410a8 Form 21
            // row 15 (trH 1097tw non-binding atLeast, 15-para right cell) SPLITS
            // correctly at the DEFAULT cell heights (the piece-1 hhea estimate is
            // NOT needed for THIS target — verified 0.9563→0.9655 = the full A+B).
            // Full golden scan (369 docs): ONLY uklocalspending changes; every
            // word_png doc is byte-identical (piece-1 not applied). This is A, on
            // top of the shipped B (S1024) — together A+B = the forms p63-65 fix.
            let s941_nonbinding_trh = std::env::var("OXI_S941_DISABLE").is_err()
                && (std::env::var("OXI_S940T_DISABLE").is_err()
                    || std::env::var("OXI_S1025_DISABLE").is_err())
                // A CJK row may also split when the first fragment can honor
                // its declared minimum height. Exact heights remain atomic.
                // S1427 (2026-09-16, default ON, opt-out OXI_S1427_DISABLE):
                // `_pb_trhsplit_gen.py` (tests/fixtures/trhsplit, ＭＳ 明朝): an
                // atLeast row taller than its minimum SPLITS when the minimum
                // fits above the bottom (n7/n9, room 74.5 vs trH 72) and moves
                // whole when it does not (room 70.5) -- the Latin S941 rule,
                // which was opt-in for CJK bodies. policies__07543a6b p9.
                // MULTI-CELL rows only: a single-cell prose box with a trHeight
                // moves whole (S754's three specimens; golden 3a4f9fbe1a83 /
                // model row 1 trH 67.2 with 67.9 of room and tokyoshugyo's
                // （参考） box went PASS -> FAIL when the split reached them).
                && (!self.doc_body_has_real_cjk
                    || (row.cells.len() > 1 && std::env::var_os("OXI_S1427_DISABLE").is_none())
                    || std::env::var("OXI_CJK_ROW_MINIMUM_SPLIT").is_ok())
                && row.height_rule.as_deref() != Some("exact")
                && row.height.map_or(false, |trh| {
                    if table_grid_pitch.is_none()
                        || (row.cells.len() > 1
                            && std::env::var("OXI_TYPED_ROW_SINGLE_FRAGMENT_DISABLE").is_err()) {
                        // A continuation may start here when its fragment can
                        // honor the declared minimum, even with only one line.
                        page_bottom - cursor.cursor_y + 0.5
                            >= trh + self.rowbox2_trh_bw(table, row)
                    } else {
                        row_height > trh + table_grid_pitch.unwrap_or(14.0)
                            && page_bottom - cursor.cursor_y
                                >= table_grid_pitch.map(|p| p * 2.2).unwrap_or(58.0)
                    }
                });
            let s754_split = (std::env::var("OXI_S754_DISABLE").is_err()
                && (row.height.is_none() || s941_nonbinding_trh)
                && !s814_lrpb_veto
                && (page_bottom - cursor.cursor_y) >= s754_min_fit)
                || s864_empty_tail_split;
            if std::env::var("OXI_DBG754").is_ok()
                && s754_split
                && row_overflows
                && !row.cant_split
                && has_content
                && !(is_single_cell_row || has_lrpb_mid_row)
                && !widow_break_needed
                && !image_atomic_push
            {
                let txt: String = row
                    .cells
                    .iter()
                    .flat_map(|c| c.blocks.iter())
                    .find_map(|b| match b {
                        Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.is_empty()) => Some(
                            p.runs
                                .iter()
                                .flat_map(|r| r.text.chars())
                                .take(16)
                                .collect::<String>(),
                        ),
                        _ => None,
                    })
                    .unwrap_or_default();
                eprintln!("[DBG754] FIRE row={} cells={} trH={:?}/{:?} row_h={:.1} cur={:.1} pbot={:.1} fits={:.1} float={} txt={:?}",
                    row_idx, row.cells.len(), row.height, row.height_rule,
                    row_height, cursor.cursor_y, page_bottom, page_bottom - cursor.cursor_y,
                    table.style.position.is_some(), txt);
            }
            // S1247 (2026-08-28, default ON, opt-out OXI_S1247_DISABLE): a row
            // that itself carries keepNext is a link in a chain, and Word does
            // not split a link. DERIVED, `_pb_kntbl2` KN arm (Word PDF truth, 21
            // filler counts): the probe's rows 0..2 all carry keepNext on their
            // leftmost cell's first paragraph, and Word SPLITS the tall row at no
            // filler count at all -- it goes from wholly on p1 (fill<=53) to
            // wholly on p2 (fill>=54) with no intermediate state, while the same
            // table without the row keepNext (the NOKN/BARE arms) splits it 3/2,
            // 2/3 across fills 55..57. So it is the row's own keepNext that
            // forbids the split, not the caption's.
            // `has_lrpb_mid_row` stays exempt: a mid-row lastRenderedPageBreak is
            // Word's own record that it DID split this row, and a recording
            // outranks a derived rule (the same precedence S1246 gives LRPB).
            let s1247_chain_row = std::env::var("OXI_S1247_DISABLE").is_err()
                && !has_lrpb_mid_row
                && s1083_kn(row_idx);
            // Even a splittable row needs room for its minimum first fragment.
            let minimum_requires_page = minimum_row_height.map_or(false, |minimum| {
                // The first fragment must fit its closing rule as well as the
                // declared minimum. Splitting off a border cannot make a
                // binding minimum shorter. Oversized minima still fill one
                // available page before continuing.
                let minimum_fragment = (minimum + s1191_foot).min(content_height);
                cursor.cursor_y + minimum_fragment > page_bottom
            });
            // Word starts a row on a new page when the first paragraph of its
            // first cell requests it. Other cells and later paragraphs do not
            // impose a row break. The existing content guard avoids blank pages.
            let explicit_row_page_break = row.cells.first()
                .and_then(|cell| cell.blocks.first())
                .map_or(false, |block| matches!(block,
                    Block::Paragraph(para) if para.style.page_break_before));
            if std::env::var("OXI_DBG754").is_ok() && row_overflows {
                eprintln!("[DBG754] gates row={} s754_split={} nonbind_trh={} cant={} content={} single={} lrpb_mid={} widow={} chain={} img={} minreq={} min={:?} explicit_pb={} hdr={} cur={:.2} pbot={:.2} rh={:.2} trH={:?}",
                    row_idx, s754_split, s941_nonbinding_trh, row.cant_split, has_content, is_single_cell_row,
                    has_lrpb_mid_row, widow_break_needed, s1247_chain_row, image_atomic_push,
                    minimum_requires_page, minimum_row_height, explicit_row_page_break, row.header,
                    cursor.cursor_y, page_bottom, row_height, row.height);
            }
            // A kept first paragraph must fit as a whole in the first fragment.
            // Paragraphs taller than a page still need a way to make progress.
            let kept_first_paragraph_requires_page = kept_first_paragraph_height > 0.0
                && kept_first_paragraph_height <= content_height
                && cursor.cursor_y + kept_first_paragraph_height > page_bottom + 0.5;
            // Fit and splitting use the same closing rule extent. A border-only
            // overflow may move the final line(s), rather than the entire row.
            // A single text line cannot be split to make room for its closing
            // rule. Use the ordinary whole-row push so keepNext predecessors
            // move with it. Multi-line cells use the common fragment boundary.
            let s1482_foot_only = std::env::var_os("OXI_S1482_DISABLE").is_none()
                && !separate_outer_edges && !row_has_multiple_text_lines
                && row_overflows && s1191_foot > 0.0
                && cursor.cursor_y + row_fit_height - s1191_foot <= page_bottom;
            let keep_float_whole = self.keep_floating_tables_together && table.style.position.is_some();
            let needs_row_split = !keep_float_whole && row_overflows
                && !s1482_foot_only
                && !kept_first_paragraph_requires_page
                && !(row.header && std::env::var("OXI_HEADER_ROW_ATOMIC_DISABLE").is_err())
                && !explicit_row_page_break
                && !row.cant_split
                && has_content
                && (is_single_cell_row || has_lrpb_mid_row || s754_split)
                && !widow_break_needed
                && !s1247_chain_row
                && !image_atomic_push
                && !minimum_requires_page;

            // S1058 (2026-08-02, default ON, opt-out OXI_S1058_DISABLE): the
            // "row fits but a saved LRPB says break" whole-push
            // (`lrpb_row_should_break`, R7.47/R7.48) must not fire when the row's
            // LRPB evidence is MID-ROW. R7.58's own semantics: an LRPB at cell 0 /
            // first paragraph / run 0 means Word PUSHED the row whole; an LRPB
            // anywhere else means Word SPLIT it. `row_has_lrpb_at_cell_start`
            // accepts ANY cell's first paragraph, so a marker on cell 1 satisfies
            // BOTH predicates and the push branch wins — replaying a whole-row
            // push on evidence of a split. Shipped with S1057: the extra footer
            // room made technical__008ae1fa's row fit by 0.5pt, which turned off
            // the overflow-split path and let this stale whole-push through.
            let s1058_midrow_lrpb_no_push =
                has_lrpb_mid_row && std::env::var("OXI_S1058_DISABLE").is_err();
            if (explicit_row_page_break || row_overflows
                || (lrpb_row_should_break && !s1058_midrow_lrpb_no_push)
                || widow_break_needed)
                && has_content
                && !needs_row_split
                && !keep_float_whole
            {
                if std::env::var("OXI_DBG_ROWPUSH").is_ok() {
                    let txt: String = row
                        .cells
                        .iter()
                        .flat_map(|c| c.blocks.iter())
                        .find_map(|b| match b {
                            Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.is_empty()) => {
                                Some(
                                    p.runs
                                        .iter()
                                        .flat_map(|r| r.text.chars())
                                        .take(14)
                                        .collect::<String>(),
                                )
                            }
                            _ => None,
                        })
                        .unwrap_or_default();
                    // `veto` is the LRPB veto; the image veto and cantSplit are
                    // separate and were the two that mattered when this was read
                    // as a single flag -- print all three.
                    eprintln!("[ROWPUSH] row={} cur={:.1} pbot={:.1} row_h={:.1} ovf={} lrpb_brk={} widow={} single={} midlrpb={} s754={} veto={} imgveto={} interiorimg={} cantsplit={} trH={:?} txt={:?}",
                        row_idx, cursor.cursor_y, page_bottom, row_height,
                        row_overflows, lrpb_row_should_break, widow_break_needed,
                        is_single_cell_row, has_lrpb_mid_row, s754_split, s814_lrpb_veto,
                        image_atomic_push, s998_interior_image, row.cant_split,
                        row.height, txt);
                }
                // S1083 (2026-08-06, default ON, opt-out OXI_S1083_DISABLE):
                // Word will not split a table INSIDE
                // a keepNext row-chain. DERIVED on technical__00501ca3 (Word
                // PDF + COM): its 24-row table 2 has keepNext on the leftmost
                // cell's first paragraph of EVERY row but row 12; Word breaks
                // p7/p8 exactly after that row-12 terminator and moves the whole
                // rows-13..23 chain (caption + 2 header rows + data) to p8 even
                // though ~135pt were free. Oxi packed rows 13-15 onto p7 (the
                // 12 mismatched paragraphs). This is the mid-table split
                // counterpart of S1024's whole-table row-chain rule.
                // Back-pull with ACTUAL geometry (the S963b/S970 contract, never
                // an estimate): the chain rows are already laid out, so their
                // extent is (cursor - chain_start_y).
                let mut s1083_moved: Vec<LayoutElement> = Vec::new();
                let mut s1083_extent = 0.0f32;
                let mut s1083_moves_header = false;
                if s1083_on && !explicit_row_page_break {
                    let mut c = row_idx;
                    while c > 0
                        && s1083_row_start.iter().any(|(ri, _)| *ri == c - 1)
                        && s1083_kn(c - 1)
                    {
                        c -= 1;
                    }
                    // Repeating leading headers stay with the first data chain.
                    if header_chain && c > 0
                        && table.rows[..c].iter().all(|r| r.header)
                        && s1083_row_start.iter().any(|(ri, _)| *ri == 0)
                    {
                        c = 0;
                    }
                    let first_on_page = s1083_row_start.first().map(|(ri, _)| *ri);
                    // never blank the page: something must remain above the chain
                    let keeps_content = Some(c) != first_on_page || !current_elements.is_empty()
                        || (header_chain && s1083_row_start.iter()
                            .find(|(ri, _)| *ri == c)
                            .map_or(false, |(_, y)| *y > page_top + 0.5));
                    if c < row_idx && keeps_content {
                        if let Some((_, cy)) =
                            s1083_row_start.iter().find(|(ri, _)| *ri == c).copied()
                        {
                            let (keep, moved): (Vec<LayoutElement>, Vec<LayoutElement>) =
                                std::mem::take(&mut elements)
                                    .into_iter()
                                    .partition(|e: &LayoutElement| e.y < cy - 0.1);
                            elements = keep;
                            s1083_moves_header = header_chain && c == 0
                                && table.rows.first().map_or(false, |r| r.header);
                            s1083_moved = moved;
                            s1083_extent = (cursor.cursor_y - cy).max(0.0);
                        }
                    }
                }
                if s1083_moved.is_empty() && row_idx > 0 {
                    if let Some(edge) = inherited_bottom_rule.as_ref()
                        .filter(|e| e.style != "none" && e.width > 0.0) {
                        let previous = &table.rows[row_idx - 1];
                        let mut grid = previous.grid_before as usize;
                        let mut x = table_x + col_widths.iter().take(grid).sum::<f32>();
                        for cell in &previous.cells {
                            let end = (grid + cell.grid_span.max(1) as usize).min(col_widths.len());
                            let width = col_widths.get(grid..end).unwrap_or(&[]).iter().sum::<f32>();
                            if width > 0.0 && cell.borders.as_ref().and_then(|b| b.bottom.as_ref()).is_none() {
                                elements.push(LayoutElement::new(x, cursor.visual_y, width, 0.0,
                                    LayoutContent::TableBorder {
                                        x1: x, y1: cursor.visual_y, x2: x + width, y2: cursor.visual_y,
                                        color: Some(format!("#{}", edge.color.as_deref().unwrap_or("000000").trim_start_matches('#'))),
                                        width: edge.width, style: Some(edge.style.clone()),
                                    }));
                            }
                            x += width;
                            grid = end;
                        }
                    }
                }
                let vmerge_cut_page = pages.len();
                let vmerge_cut_bottom = cursor.visual_y;
                // Push all accumulated elements (including previous rows) to current page
                current_elements.extend(std::mem::take(&mut elements));
                dbg_page_push(pages.len(), 0);
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });
                page_bottom += std::mem::take(&mut first_page_fit_offset);
                // S1527: the finished page's note reserve belongs to that page;
                // the continuation page starts clean and S1527 subtracts only the
                // notes whose referencing lines land on it (S740 v1 carried the
                // entry page's reserve across every continuation page).
                if row_footnotes.is_none() || std::env::var_os("OXI_S1527_DISABLE").is_none() {
                    page_bottom += std::mem::take(&mut s740_reserve);
                }
                page_bottom += advance_table_page_geometry(
                    page_geometry, pages.len() + 1, &mut page_top,
                    &mut content_height, &mut [],
                );
                let replay_header_on_restart = s728_on
                    && !s1083_moves_header
                    && s728_capture_done
                    && !s728_hdr_elems.is_empty()
                    && !row.header
                    && !widow_break_needed
                    && (s1083_moved.is_empty()
                        || (std::env::var("OXI_KEEP_REPEAT_HEADER_DISABLE").is_err()
                            && (!self.doc_body_has_real_cjk || cjk_header_chain)));
                let table_restart_top = if std::env::var("OXI_TABLE_RESTART_TOP_DISABLE").is_err()
                    && row_idx > 0
                    && !self.doc_body_has_real_cjk
                    && bug_a_enabled && s1083_moved.is_empty()
                    && (table.style.border || std::env::var("OXI_TABLE_RESTART_OUTER_EDGE_DISABLE").is_err())
                {
                    let outer = if std::env::var("OXI_TABLE_RESTART_OUTER_EDGE_DISABLE").is_err() {
                        table.style.top_border.as_ref().map_or_else(
                            || if table.style.border { table.style.border_width.unwrap_or(0.4) } else { 0.0 },
                            |edge| self.s1188_drawn(&edge.style, edge.width),
                        )
                    } else { table.style.border_width.unwrap_or(0.4) };
                    // The row (or repeated header) already includes its normal
                    // top rule. Replace that rule at a page boundary; do not
                    // add the outer rule to it a second time.
                    let outer = if separate_outer_edges {
                        let first_row = if replay_header_on_restart { 0 } else { row_idx };
                        self.table_fragment_top_width(table, &table.rows[first_row])
                            - self.s1188_edge_bw(table, first_row)
                    } else if std::env::var_os("OXI_S1622_DISABLE").is_none() {
                        // S1622 (2026-10-01, default ON, opt-out OXI_S1622_DISABLE):
                        // the same replacement for a table without separate outer
                        // edges. Word charges ONE rule at the continuation top --
                        // the row's own cell top where declared, else the table's
                        // top -- and never the row above's bottom; the row model
                        // already carries its collapsed top (s1188_edge_bw), so only
                        // the difference is added. `_pb_tblrestart_gen.py`, 6 arms
                        // (cell border none / bottom / top x table top style / none):
                        // Word's first continued row sits 0.5 above Oxi's in the
                        // three arms where the row's collapsed top was ADDED to the
                        // outer rule (bottom+style, top+style, bottom+none) and
                        // agrees in the other three; reports__0013bcb8 p3 (no cell
                        // borders) is the none+style arm.
                        let first_row = if replay_header_on_restart { 0 } else { row_idx };
                        self.table_fragment_top_width(table, &table.rows[first_row])
                            - self.s1188_edge_bw(table, first_row)
                    } else { outer };
                    page_top + outer
                } else { page_top };
                // S1621 (2026-10-01, default ON, opt-out OXI_S1621_DISABLE): the
                // continued table's top rule is drawn AT the page top; only the
                // row content sits `outer` lower. EN reports__0013bcb8 p3 (a
                // TabloKlavuzu table continued from p2, top rule from the style):
                // Word rule 71.04 against Oxi's 71.60 (drawn at page_top + 0.5)
                // and 71.10 at page_top; the cell text and the next rule (91.46 /
                // 91.48) already agree with the content at page_top + 0.5.
                if std::env::var_os("OXI_S1621_DISABLE").is_none() && !separate_outer_edges {
                    s1621_lift = table_restart_top - page_top;
                }
                cursor.set(table_restart_top);
                s1083_row_start.clear();
                // S728: replay the captured tblHeader row(s) at the new page
                // top (mid-table continuation break; NOT a widow whole-table
                // move, NOT above a header row itself). Clones shift by
                // (page_top − captured min y); TableBorder carries its own
                // content-level y1/y2 (S648: renderers draw borders from
                // those, elem.y is secondary) so shift both.
                if std::env::var("OXI_DBG728").is_ok() {
                    eprintln!("[S728] PUSH row_idx={} cap_done={} n_hdr={} hdr_h={:.1} row.header={} widow={}",
                        row_idx, s728_capture_done, s728_hdr_elems.len(), s728_hdr_h, row.header, widow_break_needed);
                }
                if replay_header_on_restart {
                    let y0 = s728_hdr_elems
                        .iter()
                        .map(|e| e.y)
                        .fold(f32::INFINITY, f32::min);
                    if y0.is_finite() {
                        let dy = table_restart_top - y0;
                        for el in &s728_hdr_elems {
                            let mut c = el.clone();
                            c.y += dy;
                            if let LayoutContent::TableBorder { y1, y2, .. } = &mut c.content {
                                *y1 += dy;
                                *y2 += dy;
                            }
                            elements.push(c);
                        }
                        cursor.set(table_restart_top + s728_hdr_h + self.repeated_header_border_delta(table, row_idx));
                    }
                }
                if !s1083_moved.is_empty() {
                    let y0 = s1083_moved
                        .iter()
                        .map(|e| e.y)
                        .fold(f32::INFINITY, f32::min);
                    let dy = cursor.cursor_y - y0;
                    for el in &s1083_moved {
                        let mut cl = el.clone();
                        cl.y += dy;
                        if let LayoutContent::TableBorder { y1, y2, .. } = &mut cl.content {
                            *y1 += dy;
                            *y2 += dy;
                        }
                        elements.push(cl);
                    }
                    cursor.set(cursor.cursor_y + s1083_extent);
                }
                // Reflow spanning text at the actual row boundary, including
                // paragraph spacing that cannot fit in the completed fragment.
                for flow in &mut vmerge_text_flows {
                    let bygrid = std::env::var_os("OXI_S1192G_DISABLE").is_none();
                    let continues = LayoutEngine::s1192_cell_at(row, flow.key, bygrid).map_or(false,
                        |c| matches!(c.v_merge.as_deref(), Some("continue") | Some("")));
                    if flow.start_row >= row_idx || !continues { continue; }
                    // A later restart in the same column owns a different span.
                    if (flow.start_row + 1..row_idx).any(|ri| {
                        LayoutEngine::s1192_cell_at(&table.rows[ri], flow.key, bygrid).map_or(true,
                            |c| !matches!(c.v_merge.as_deref(), Some("continue") | Some("")))
                    }) { continue; }
                    flow.cuts.insert(vmerge_cut_page,
                        (vmerge_cut_bottom - flow.pad_bottom, cursor.visual_y + flow.pad_top));
                    let (positions, end) = flow.paginate();
                    vmerge_absolute_ends.insert(flow.key, end);
                    let mut moved = Vec::new();
                    for pg in pages.iter_mut().skip(flow.source_page) {
                        pg.elements.retain(|e| {
                            if e.vmerge_flow_element.as_ref().map_or(false,
                                |(id, _)| std::sync::Arc::ptr_eq(id, &flow.identity)) {
                                moved.push(e.clone()); false
                            } else { true }
                        });
                    }
                    elements.retain(|e| {
                        if e.vmerge_flow_element.as_ref().map_or(false,
                            |(id, _)| std::sync::Arc::ptr_eq(id, &flow.identity)) {
                            moved.push(e.clone()); false
                        } else { true }
                    });
                    for mut e in moved {
                        let ordinal = e.vmerge_flow_element.as_ref().unwrap().1;
                        if let Some((destination, y)) = positions[ordinal] {
                            e.y = y;
                            e.vmerge_restart_overflow_to_next_page = false;
                            e.vmerge_destination_page = None;
                            if destination < pages.len() {
                                pages[destination].elements.push(e);
                            } else {
                                let offset = destination - pages.len();
                                e.y += offset as f32 * content_height;
                                e.vmerge_destination_page = (offset > 0).then_some(destination);
                                elements.push(e);
                            }
                        }
                    }
                }
                // The current row now starts on this page. Retain its actual
                // start so later keepNext checks see content above a heading.
                if s1083_on && std::env::var("OXI_TABLE_ROW_HISTORY_DISABLE").is_err() {
                    s1083_row_start.push((row_idx, cursor.cursor_y));
                }
            }

            // S1067 (2026-08-04, default ON, opt-out OXI_S1067_DISABLE): a
            // table row that whole-moves to become the FIRST content of a fresh
            // continuation page COLLAPSES its leading empty cell paragraphs to
            // zero height. Word render-truth: educational__002a301d's ROW2
            // (1-col, 8-row table → not a single-cell row, whole-moved) opens
            // p6 with "b." as its FIRST line (baseline 105.39/105.50) — the 6
            // leading 16pt empties render ~0. The faithful minimal repro
            // (s1065_reproA: 30 filler + a 2-row fixed table whose row2 = 6
            // empties + "b.") whole-moves row2 to p2 where "b." is the first
            // content at baseline 84.98 with the 6 empties collapsed. A
            // whole-TABLE move (row NOT at the page top) does NOT collapse; a
            // mid-page empty renders full (reproB control). Mirrors
            // S570/S719b's re-anchor-to-first-non-empty, which only existed in
            // the R7.56 single-cell overflow loop. The gate `!pages.is_empty()
            // && cursor==page_top` identifies a just-whole-moved row (the
            // whole-move push above does cursor.set(page_top)); a row that
            // naturally starts a fresh page has pages empty or cursor ≠
            // page_top (a following row sits at page_top + prior row_height).
            // Latin-scoped (JP form cells = the calibrated cell-height
            // tombstone area).
            // ★HELD OPT-IN (2026-08-05, `OXI_S1067=1`, default OFF = byte-identical).
            // The rule above is Word-measured, but an isolation run on its own
            // target (educational__002a301d, all three flags swept) shows it is
            // INERT there: ALL-ON 0.5984 pcd+2 == -S1067 0.5984 pcd+2, while
            // removing S1HDR or S1066 drops back to 0.2623 pcd+3. A rule with
            // no measured gate benefit buys only blast radius, so it ships when
            // a doc actually needs it (it was implemented but never verified —
            // the "12 -> 11 pages" claim was never measured).
            let s1067_row_at_page_top = std::env::var("OXI_S1067").is_ok()
                && !self.doc_body_has_real_cjk
                && !pages.is_empty()
                && (cursor.cursor_y - page_top).abs() < 0.5;

            // Second pass: render cells
            // Track actual content height per cell for row_height correction
            let is_exact_row = row.height_rule.as_deref() == Some("exact");
            let mut max_actual_cell_h: f32 = row_height;
            let elements_before_row = elements.len();
            // Keep resolved terminal spacing with its own cell paragraph. A
            // sibling's tail must not be added after the tallest continuation.
            let mut cell_terminal_spacing: std::collections::HashMap<(usize, usize), f32> =
                std::collections::HashMap::new();
            // S1431 (2026-09-16, default ON, opt-out OXI_S1431_DISABLE): (cell,
            // paragraph) -> effective space_before, so a paragraph that OPENS a
            // split row's continuation keeps its spacing at the page top.
            // `_pb_trhsplit_gen.py` SB=162 SBPARA=6 (n7, X=34/42): the paragraph
            // starting the continuation sits at 64.80 = 56.7 + 8.1, not 56.7.
            // forms__00830ac0 p4 「＿＿年度」(beforeLines 50): Word 54.0, Oxi 42.55.
            let mut s1431_cell_para_sb: std::collections::HashMap<(usize, usize), f32> =
                std::collections::HashMap::new();
            let mut split_valign_offsets: Vec<(usize, f32)> = Vec::new();
            // S1407 (2026-09-15, default ON, opt-out OXI_CELL_FRAGMENT_VALIGN_DISABLE):
            // the checkpoint's opt-in promoted -- a row without vMerge or cell
            // text boxes aligns its split fragments by the cell's vAlign.
            // technical__b243782b73e69ad1 (row 10 lrpb_mid split) and
            // creative__32ceca2e1f3ae9a8 go to PASS. Gates under the env flag:
            // golden 183/187 (same fail set), ja 179 -> 181, en 291/298 (same).
            let fragment_valign = std::env::var("OXI_CELL_FRAGMENT_VALIGN_DISABLE").is_err()
                && row.cells.iter().any(|c| c.v_merge.is_none() && c.cell_text_boxes.is_empty());
            // Element ranges exclude the outer cell borders. Keep the physical
            // content origin so alignment can use the final fragment boundary.
            let mut fragment_valign_cells: Vec<(std::ops::Range<usize>, f32, f32, f32, f32)> = Vec::new();
            let mut fragment_paragraph_after: std::collections::HashMap<(usize, usize), f32> =
                std::collections::HashMap::new();
            // S500 (L1) FALSIFIED (2026-06-06): re-centering vAlign center/bottom cells against
            // the FINAL row height (fixing early cells centered before later/taller cells set
            // max_actual_cell_h) FIXED the synthetic repro vc_2cell_auto (-1.65->+0.10) but was
            // a NO-OP on the real corpus (net -0.0005; every bottom-N page +/-0.0004, 2ea81a
            // -0.0004) — the stale-height ordering doesn't manifest in real docs (centered cells
            // are single-cell rows or similar-height) and it does NOT fix d4d126's +3.3 (the
            // over-estimate direction). Reverted. See cellY_perdoc_scoped_design.md.
            // Apply gridBefore: skip leading grid columns
            let mut grid_idx: usize = row.grid_before as usize;
            let mut cell_x = table_x
                + col_widths[..grid_idx.min(col_widths.len())]
                    .iter()
                    .sum::<f32>();
            let _num_cells = row.cells.len();
            for (cell_idx, cell) in row.cells.iter().enumerate() {
                let span = cell.grid_span.max(1) as usize;
                // vMerge="continue" cells: skip content but still draw borders
                let is_vmerge_continue = cell.v_merge.as_deref() == Some("continue")
                    || cell.v_merge.as_deref() == Some("");
                // S163 (2026-05-21): track grid_idx directly instead of recovering it
                // via cumulative-offset find with 0.5pt tolerance. The find was brittle
                // when consecutive grid columns included sub-pt spacer widths (ed025
                // Tables(7) row 4 has grid[8]=0.5pt spacer between gridSpan=2 cells;
                // cell 7's cell_x matched grid[8] under float precision instead of
                // grid[9]=33.75pt, causing 'トン' to wrap to 2 lines → +18pt content_h
                // → row growth via max_actual_cell_h → +16.5pt drift propagating
                // through pages 5-8). The first-pass row-height calc already tracks
                // grid_idx directly (line 6354); this aligns the second pass.
                // S237 (2026-05-23): removed OXI_LEGACY_GRIDIDX_FIND legacy
                // env-var fallback (was the pre-fix `col_widths.iter().find()`
                // float-precision lookup); the index-aligned path is canonical.
                let cell_start_grid = grid_idx.min(col_widths.len().saturating_sub(1));
                let cell_end_grid = (cell_start_grid + span).min(col_widths.len());
                let mut cell_w: f32 = col_widths[cell_start_grid..cell_end_grid].iter().sum();

                // S1215 (2026-08-25, opt-out `OXI_S1215_DISABLE`): a row's `w:tblPrEx`
                // may override the table's cell margins, and the parser already
                // carries it (`TableRow::cell_margins_override`) -- nothing ever
                // read it, so such a row got the table default instead.
                // 29dc6e8943fe's ③ row is the case: its tblPrEx says
                // `tblCellMar left/right = 12` (0.60pt) where the table default
                // is 108 (5.40pt), and every paragraph in that cell renders
                // 4.45pt right of Word's. Sliced out and measured against Word:
                // with no indent at all the text sits 0.84pt inside the cell's
                // own rule, not 5.4pt.
                let s1215_row_mar = if std::env::var("OXI_S1215_DISABLE").is_err() {
                    row.cell_margins_override.as_ref()
                } else {
                    None
                };
                let pad_l = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.left)
                    .or_else(|| s1215_row_mar.and_then(|m| m.left))
                    .unwrap_or(default_pad_l);
                let pad_r = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.right)
                    .or_else(|| s1215_row_mar.and_then(|m| m.right))
                    .unwrap_or(default_pad_r);
                let mut pad_t = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.top)
                    .unwrap_or(row_default_pad_t);
                let pad_b = cell
                    .margins
                    .as_ref()
                    .and_then(|m| m.bottom)
                    .unwrap_or(row_default_pad_b);
                // S1575 (2026-09-26, default ON, opt-out OXI_S1575_DISABLE): a cell's
                // top/bottom margin is ROW-wide -- every cell of the row takes the
                // row's largest. reference__009644b1: only the label cells carry tcMar
                // 100/100; Word starts the content cell's first line on the label's
                // baseline (PDF 256.13 both) and the 3-line row is 5 + 43.92 + 5,
                // Oxi's content cell had no margin (43.9) and page 2 ran ~50pt short.
                let (pad_t, pad_b) = if std::env::var_os("OXI_S1575_DISABLE").is_none() {
                    let mt = row.cells.iter().map(|c| c.margins.as_ref().and_then(|m| m.top).unwrap_or(row_default_pad_t)).fold(pad_t, f32::max);
                    let mb = row.cells.iter().map(|c| c.margins.as_ref().and_then(|m| m.bottom).unwrap_or(row_default_pad_b)).fold(pad_b, f32::max);
                    (mt, mb)
                } else { (pad_t, pad_b) };
                #[allow(unused_mut)]
                let mut pad_t = pad_t;

                // S494b/S496 tblInd PER-CELL absorption (default-ON, opt-out OXI_S496_TBLIND_DISABLE): the
                // leading-edge column cell (grid col 0) of a NON-nested tblInd table absorbs
                // its left margin — Word renders its content at margin + tblInd and its left
                // border at margin + tblInd - cellMargin. Shift only THIS cell's border+content
                // left by its left margin and widen it so the right edge / column advance are
                // unchanged (other cells stay put). This replaces the table_x translate, which
                // moved every column's border and regressed border-visible docs (15076df). Only
                // when the legacy table_x path did NOT already absorb (border_offset ~0, i.e.
                // not the Some(0)+style-border case), tblInd present, and at the true leading
                // column (cell_start_grid==0 — gridBefore rows whose first cell is offset are
                // skipped, which is why 15076df's content does not move).
                // S496 GATE FOUND (2026-06-05): the absorb-vs-literal split is the document
                // compatibilityMode, NOT any table-structure feature (S494b ruled all of those
                // out). Word 2013+ (compatibilityMode 15) changed table layout so the leading
                // cell does NOT absorb its left margin; Word 2010 (mode <= 14) DOES. Verified
                // 100% across S494b's set: all 3 absorbers (e3c545/04b88e/34140b) are mode 14,
                // all 15 regressors (tokumei/kyodokenkyu/order forms a1d6e4/d4d126/15076df/...)
                // are mode 15. The FULL corpus affected set is exactly those 3 mode-14 docs with
                // positive tblInd (every other mode-14 doc has no positive tblInd, every mode-15
                // doc is excluded => byte-identical). Render-truth (e3c545 p4): Word puts the
                // leading code-block cell text at margin+tblInd, Oxi was at margin+tblInd+pad_l
                // (+5.4pt over for the default 108tw cell margin). Gate on ANY positive tblInd
                // (not > pad_l) so the tblInd~=cellMargin tables (e3c545 108tw) also absorb.
                // opt-out OXI_S496_TBLIND_DISABLE. spec_tblind_cellmargin_absorption memory.
                let lead_absorb =
                    self.compat_mode <= 14 && table.style.indent.map_or(false, |v| v > 0.1);
                if cell_start_grid == 0
                    && !is_nested
                    && lead_absorb
                    && std::env::var("OXI_S496_TBLIND_DISABLE").is_err()
                {
                    cell_x -= pad_l;
                    // S766 = the S585c 本体 (root) fix (2026-07-08, default ON,
                    // opt-out OXI_S766_DISABLE). The S496 lead_absorb outsets the
                    // leading cell's LEFT border by cellMar (cell_x -= pad_l). For a
                    // MULTI-cell row the leading cell's RIGHT edge is a column
                    // boundary Word keeps, so also widen (cell_w += pad_l = keep the
                    // right edge). But for a SINGLE-cell row the leading cell IS the
                    // whole table and there is no column boundary to the right:
                    //   • tblW=dxa/pct → Word shifts the WHOLE table left by cellMar
                    //     (BOTH borders −cellMar, declared width KEPT, overflowing the
                    //     page). Widening the sole cell instead put its RIGHT border
                    //     +1 cellMar too far (tokyoshugyo dxa note Oxi 527.05/W522.2)
                    //     AND the S713 wrap that subtracts pads from cell_w over-wide
                    //     (435−9.9=425.1 vs Word 430.05−9.9=420.15). Not widening
                    //     makes cell_w = declared → border and wrap both = Word.
                    //   • tblW=auto → Word CLAMPS the box to the page (S591); s585c's
                    //     eff_cell_w clamp handles it, and its over-wide TRIGGER
                    //     (cell_w > content_width) RELIES on the widen (the tokyoshugyo
                    //     条文 boxes: declared 422.90 < content 425.2, only the widen
                    //     427.85 crosses the threshold). So KEEP the widen for auto —
                    //     removing it drops the clamp+compression → over-wrap (0.56).
                    // See [[tokyoshugyo_wrap_not_cellheight]] (S713 "no-widen remains
                    // a follow-up") / [[tokumei_form_family_ssim]].
                    // S1163 (2026-08-17, default ON, opt-out OXI_S1163_DISABLE):
                    // the `auto` carve-out above is about the OVER-WIDE box that
                    // S585c then clamps. A single-cell auto box that FITS is not
                    // that case, and widening it puts its right border a cellMar
                    // past Word: _pb_tblanchor plain_ind567 (tblW auto, grid
                    // 400pt inside a 425pt column) measures Word 108.02..508.18
                    // = the declared 400.16, Oxi 108.00..513.40 = 405.40. Keep the
                    // widen only where the clamp needs it.
                    // Gate: probe 10/10 on left border, text and RIGHT border;
                    // Phase 1 95/96 with zero per-doc change; SSIM sentinel one
                    // document changed -- e3c545_LOD_Handbook, one of the three
                    // mode-14 leading-cell absorbers S496 was derived on, +0.0033
                    // over 12 pages -- and nothing worse. tokyoshugyo, whose
                    // 条文 boxes are the reason the widen exists, is untouched:
                    // they overflow, so this carve-out does not reach them
                    // (0.8607 -> 0.8607, all 90 pages identical).
                    // S1200 (2026-08-23, opt-in `OXI_S1200`): the box "fits" up to the
                    // S1196 cap, not up to the bare content area. Word lets an auto
                    // table hang its cell margins outside the text area (S1196:
                    // cap = content - tblInd + inset_L + inset_R when a tblInd is
                    // declared), and the single-column path never reaches that rule
                    // because the shrink branch requires more than one column.
                    //
                    // WORD RENDER-TRUTH (tokyoshugyo p20 条文 box, rules read out of
                    // Word's own PDF): grid 8458tw = 422.90, drawn 92.18..515.14 =
                    // 422.96 -- the DECLARED width, shifted left by the cell margin,
                    // its right border 4.87 past the right text margin and inside the
                    // 5.40 tolerance. Word does NOT clamp it. Oxi widens it by a cell
                    // margin (428.30) purely so the S585c clamp trigger fires, and the
                    // clamp then approximately undoes the widen -- two errors that
                    // leave 0.66pt, which is exactly what makes its 「…手待時間」）
                    // line hold one character Word pushes out.
                    let s1200_cap = if std::env::var("OXI_S1200").is_ok() {
                        content_width - table.style.indent.unwrap_or(0.0) + pad_l + pad_r
                    } else {
                        content_width - pad_l
                    };
                    let s1163_fits = std::env::var("OXI_S1163_DISABLE").is_err()
                        && table.style.width_type.as_deref() == Some("auto")
                        && cell_w + pad_l <= s1200_cap + pad_l;
                    let s766_shift_whole = row.cells.len() == 1
                        && (matches!(table.style.width_type.as_deref(), Some("dxa") | Some("pct"))
                            || s1163_fits)
                        && std::env::var("OXI_S766_DISABLE").is_err();
                    // S1240 (2026-08-27, default ON, opt-out OXI_S1240_DISABLE): a
                    // MULTI-cell row does NOT widen its leading cell. The widen
                    // above rests on S766's stated assumption that "the leading
                    // cell's RIGHT edge is a column boundary Word keeps" — that
                    // half was never measured (S1163 measured only the SINGLE-cell
                    // case). MEASURED (_pb_tblindwide_gen.py, 11 arms, boundaries
                    // read off Word's own PDF rules): the absorption moves the
                    // WHOLE grid left, so the boundary moves with it and every
                    // column keeps its declared gridCol.
                    //   left_border = margin + tblInd − cellMar_left
                    //   boundary_i  = left_border + Σ gridCol_1..i   (no widen)
                    // The default-cellMar arms cannot see this (a doc with no
                    // table style has cellMar ≈ 0, so absorb and no-absorb
                    // coincide); the arms that declare a margin separate them:
                    //   ind567 3col cellMar 200tw: Word left 90.38 = 72+28.35−10,
                    //     boundary1 190.37 = left+100.0 EXACT (widen predicts
                    //     200.35); Oxi gave col1 110.0 vs the grid's 100.0.
                    //   ind113 5col cellMar 108tw (= technical__002c1ffa's own
                    //     table, reproduced): Word left 72.26 vs absorb 72.25,
                    //     boundary1 164.18 vs left+91.90 = 164.15 EXACT (widen
                    //     predicts 169.55); all five columns land on the grid
                    //     within 0.05 and each cell's text at its border + 5.4.
                    //   cm15 + cellMar 200tw: left 100.58 = margin+tblInd, text
                    //     at +10 — compat 15 does not absorb, so the S496 gate
                    //     itself is re-confirmed, not widened.
                    //   compat 11 reproduces compat 14 EXACTLY on both margin-
                    //     bearing arms (190.37 / 164.18), so the law covers the
                    //     whole legacy regime, not just mode 14 — that matters
                    //     because tokyoshugyo (compat 11) is one of the three
                    //     Phase-1 baseline docs this reaches.
                    // WITNESS technical__002c1ffa65f3a566 (compat 14, tblInd 113,
                    // 5-col legislation-history table, tcW == gridCol): Oxi drew
                    // col1 at 97.30 = grid 91.90 + one cellMar, so its cell text
                    // wrapped ~one word later than Word's per line, every row came
                    // out shorter, and the Endnote-3 table carried a constant
                    // sub-page offset from Word p331 onward — the doc's pcd −1
                    // (367 pages vs 368). The single-cell paths are untouched:
                    // S585c's over-wide clamp still depends on the widen, and it
                    // only ever reaches a 1-cell row.
                    let s1240_no_widen = std::env::var("OXI_S1240_DISABLE").is_err()
                        && row.cells.len() > 1;
                    if !s766_shift_whole && !s1240_no_widen {
                        cell_w += pad_l;
                    }
                }

                // S585c (2026-07-01, default ON, opt-out OXI_S585C_DISABLE): the
                // ONE consistent clamp for the over-wide single-cell AUTO box (the
                // tokyoshugyo 条文/解説 regulation boxes). Word clamps such a box so
                // its content-RIGHT aligns with the page text margin and its right
                // border sits at content_right + cellMar; the declared gridCol
                // overflow is discarded. Oxi kept cell_w = declared gridCol, so the
                // right border landed +1 cellMar too far (measured tokyoshugyo:
                // right border Oxi 520.15 vs Word 514.9, LEFT border 92.6 vs 92.18 =
                // aligned) AND the wrap was ~1 cellMar too wide. S585b/S594/S585N/
                // PGCAP/PROPCELL/LEGACYCELL each patched only the RENDER wrap, leaving
                // the border and the estimate/render wraps inconsistent. This computes
                // a single `eff_cell_w` = the width whose right border matches Word,
                // and threads it through BOTH the border/shading rendering AND the
                // render wrap so every border-x/wrap path is derived from one geometry.
                // SCOPE = the validated s585_cellmar envelope (single-cell, inherited
                // cellMar, tblW=auto → Word fits to page; dxa/pct keep wide) + a
                // geometric overshoot trigger (the border actually exceeds Word's
                // position). is_nested excluded (nested boxes reference a parent cell,
                // not the page margin — the deferred S585N case). See
                // [[tokyoshugyo_wrap_not_cellheight]] (S585c).
                let content_right_edge = start_x + content_width;
                let s585c_over = cell_w - content_width;
                let s585c_clamp_base = std::env::var("OXI_S585C_DISABLE").is_err()
                    && !is_nested
                    && row.cells.len() == 1
                    && !table.style.has_explicit_cellmar
                    && cell_w > content_width
                    && (s585c_over < 5.0
                        || (table.style.width_type.as_deref() == Some("auto")
                            && s585c_over < 11.0))
                    && cell_x + cell_w > content_right_edge + pad_r + 0.5;
                // S767b = the S585c 本体 detection completion (2026-07-08, default ON,
                // opt-out OXI_S767_DISABLE). s585c's over-wide TRIGGER is
                // `cell_w > content_width` — POSITION-AGNOSTIC: it measures the cell's
                // width against the full page content, ignoring that a large tblInd
                // pushes the cell RIGHT. A single-cell AUTO box whose declared width
                // FITS the page (cell_w ≤ content_width) but is shifted over the
                // page-right by a big tblInd (cell_x + cell_w > page content-right +
                // cellMar) is MISSED — its border overflows the page and Word clamps
                // it, but s585c never fires. e3c545's tblInd=18 RDF code boxes:
                // border 549.4 / wrap 469.30 vs Word 544.0 / 463.90 (rt.pdf-confirmed
                // right border 543.8). ADDITIVE (only fires where s585c_clamp_base is
                // FALSE) so it cannot change any currently-clamped cell — it only adds
                // the tblInd-overshoot case s585c missed. Border-overshoot amount gate
                // (< 11 = the auto-fit envelope) mirrors the eff_cell_w clamp geometry.
                let s585c_border_over = (cell_x + cell_w) - (content_right_edge + pad_r);
                let s767b_tblind = std::env::var("OXI_S767_DISABLE").is_err()
                    && !s585c_clamp_base
                    && !is_nested
                    && row.cells.len() == 1
                    && !table.style.has_explicit_cellmar
                    && table.style.width_type.as_deref() == Some("auto")
                    && s585c_border_over > 0.5
                    && s585c_border_over < 11.0;
                let s585c_clamp = s585c_clamp_base || s767b_tblind;
                let eff_cell_w = if s585c_clamp {
                    (content_right_edge + pad_r - cell_x).min(cell_w).max(0.0)
                } else {
                    cell_w
                };
                // S585c wrap+compress (2026-07-02, coupled with the border clamp,
                // opt-out OXI_S585C_WRAP_DISABLE): narrowing the WRAP to the clamped
                // content (eff_cell_w − 2×cellMar = Word's true content width) is
                // geometrically the matching half. Narrow ALONE over-wraps (Oxi had no
                // cell 約物 compression → tokyoshugyo 90→91 pages), but Word fits its
                // content within the clamped width via per-line 約物 compression — the
                // DERIVED cell model (~0.235em kanji-fit break cap + line-end ぶら下げ,
                // = legacy_cell_break + cell_bura below, formerly OXI_LEGACYCELL). Narrow
                // COUPLED with that compression recovers Word's line counts EXACTLY
                // (tokyoshugyo 90pg = the compensating-wide baseline, now geometrically
                // correct: border, wrap, AND 約物 positions all at Word's geometry).
                // SCOPE = LEGACY (compat<15) + compressPunctuation — Word's legacy cell
                // 約物 oikomi/ぶら下げ (S568/S572 for the body; this is the cell analog).
                // e3c545 (the only OTHER s585c word_png doc) is Latin/non-compressPunctuation
                // → s585c_narrow=false → border-only (unchanged, +0.0328). Full-corpus
                // Phase-1: 0 real flips (only tokyoshugyo, unchanged score). See
                // [[tokyoshugyo_wrap_not_cellheight]] / [[char_budget_wall]].
                let s585c_narrow = s585c_clamp
                    && std::env::var("OXI_S585C_WRAP_DISABLE").is_err()
                    && self.compat_mode < 15
                    && self.compress_punctuation;

                // Round 30 (2026-04-09): When cell top/bottom padding is 0 and
                // the table has borders, add the border width as implicit padding.
                // Word positions text below the top border line, not at the border.
                // COM-confirmed minimal repro: Table Grid with tcMar=0 all sides,
                // MS Mincho 12pt → text_y = topMargin + 0.5pt (= border width).
                // S359 (2026-05-27): test confirmed Round 30 is load-bearing
                // (OXI_S359_NO_ROUND30=1 caused -0.0186 corpus regression).
                // S386 (2026-05-27): double-border-count hypothesis FALSIFIED
                // (see height-calc site above; -0.0082 corpus regression).
                if self.rowbox2_pad_on() {
                    // ROWBOX2: generalized Round30 (see the first-pass site).
                    // S870: Latin docs use the ROW's rule (see the helper).
                    pad_t += self.rowbox2_border_pad_row(table, row_idx, cell);
                } else if pad_t == 0.0 && table.style.border {
                    let bw = table.style.border_width.unwrap_or(0.4);
                    pad_t = bw;
                }

                // Emit cell shading (background fill) before cell content
                if let Some(ref shading_color) = cell.shading {
                    if !shading_color.is_empty() && shading_color != "auto" {
                        let color_hex = if shading_color.starts_with('#') {
                            shading_color.clone()
                        } else {
                            format!("#{}", shading_color)
                        };
                        elements.push(LayoutElement::new(
                            cell_x,
                            cursor.visual_y,
                            eff_cell_w,
                            row_height,
                            LayoutContent::CellShading { color: color_hex },
                        ));
                    }
                }

                // 2026-04-19: Use content area (cell_w - padding) for wrap width.
                // Previous comment claimed "Word uses cell_w" but b35 組織的管理措置
                // cell wraps at 4 chars (= 4×10.5=42pt fits in 49.05pt inner-pad area)
                // not 6 chars (which would require 59.85pt cell_w with overflow).
                let _inner_w = (cell_w - pad_l - pad_r).max(0.0);
                let mut cell_elements: Vec<LayoutElement> = Vec::new();
                // Session 131: vertical writing anchor — Word reports
                // Information(6) for ALL paragraphs in a vert-text cell at
                // the cell top y (= row top). Snapshot the cell-entry content_h
                // so all vert paragraphs emit at that relative_y. This matches
                // the 2ea81a COM-confirmed pattern where 予納する理由 / （い
                // ずれかを選択） / empty all report y=478 (row top).
                let vert_cell_anchor_h: f32 = 0.0;
                let mut content_h: f32 = 0.0;
                // S488 (CLASS E step 3): record each cell block's content_h-relative
                // top so in-cell floating text boxes with relV="paragraph" can be
                // anchored to their SPECIFIC paragraph (not the cell top). Indexed
                // by block_pos; absolute para top = cell_block_tops[idx] + dy (the
                // dy applied to cell_elements below). Declared at cell-loop scope so
                // it survives past the `if !is_vmerge_continue` block to the text-box
                // emit site. Only consumed under OXI_S487_ENABLE.
                let mut cell_block_tops: Vec<f32> = Vec::new();
                let cell_float_flow = LayoutEngine::cell_float_enabled(cell) && !self.is_vert_writing_active(cell);
                let mut float_tops = vec![0.0; cell.blocks.len()];
                let fixed_float_positions = row_float_positions.get(&cell_start_grid);

                // Layout blocks in document order (paragraphs and nested tables interleaved)
                let is_exact = row.height_rule.as_deref() == Some("exact");
                // R7.32: count Paragraph blocks within this cell so each cell
                // paragraph can be distinguished in the dump output.
                let mut cell_para_counter: usize = 0;
                // R7.73: track whether the immediately-previous cell paragraph
                // carried a `<w:lastRenderedPageBreak/>` on a non-run-0 run.
                // Reset to false at each cell start.
                let mut prev_cell_para_had_mid_lrpb: bool = false;
                // S753: lazily-computed column layout for a tbRlV cell
                // ((cols, max_col_len); see s753_vert_cell_columns).
                let mut s753_vert_cols: Option<(Vec<(usize, String, f32, f32, f32)>, f32)> = None;
                // S427 (2026-05-29): track previous cell paragraph's space_after
                // for adjacent-paragraph spacing collapse (see pre-pass comment).
                let s427_collapse = std::env::var("OXI_S427_DISABLE").is_err();
                let mut prev_cell_sa: Option<f32> = None;
                // S939: prev paragraph's (contextual_spacing, style_id).
                let mut s939_prev_r: Option<(bool, Option<&str>)> = None;
                // S1075: (previous cell paragraph's after_autospacing, its numId)
                let mut s1075_prev_r: Option<(bool, Option<&str>)> = None;
                if !is_vmerge_continue {
                    // S428 (2026-05-29): index of the last cell block that carries
                    // real content (a non-empty paragraph or a nested table). Used to
                    // gate the empty-paragraph zero-glyph element emission below to
                    // only INTERIOR empty paragraphs (those followed by content). A
                    // trailing empty paragraph must NOT get an element, else it would
                    // overflow a mid-cell page split alone and spawn a near-blank
                    // continuation page (e3c545: a lone trailing empty cell paragraph
                    // created a blank page 5, cascading every later page +1).
                    let last_content_block_pos: Option<usize> = cell
                        .blocks
                        .iter()
                        .enumerate()
                        .filter(|(_, b)| match b {
                            Block::Paragraph(p) => p.runs.iter().any(|r| !r.text.is_empty()),
                            _ => true,
                        })
                        .map(|(i, _)| i)
                        .last();
                    // Cell-autospace (OXI_CELLAS): first/last Paragraph block positions for
                    // container-edge suppression of before/afterAutospacing. See
                    // cell_effective_spacing.
                    let first_para_pos = cell
                        .blocks
                        .iter()
                        .position(LayoutEngine::is_cell_spacing_paragraph);
                    let last_para_pos = cell
                        .blocks
                        .iter()
                        .rposition(LayoutEngine::is_cell_spacing_paragraph);
                    // S716: the post-nested-table stub paragraph is skipped entirely
                    // (no element, no content_h advance) — Word collapses it to ~0.
                    let s716_stub_render = self.nested_table_stub_pos(cell);
                    // S751: empty hideMark cell renders nothing (matches the
                    // pre-pass zero-height; without this the placement pass grew
                    // max_actual_cell_h back and the S648 correction re-inflated
                    // the row).
                    let hidden_tail_para = cell.blocks.last().and_then(|block| match block {
                        Block::Paragraph(para) if LayoutEngine::hidden_cell_final_line(cell, cell.blocks.len() - 1, para) =>
                            Some(LayoutEngine::without_hidden_cell_after(para)),
                        _ => None,
                    });
                    let s1311_tail_render = self.s1311_hidemark_tail_pos(cell);
                    let s751_hide_render = cell.hide_mark
                        && std::env::var("OXI_S751_DISABLE").is_err()
                        && s1311_tail_render.is_none()
                        && cell.blocks.iter().all(|b| {
                            matches!(b, Block::Paragraph(p)
                        if p.runs.iter().all(|r| r.text.is_empty()))
                        });
                    // S1067: collapse this cell's leading empty paragraphs when the
                    // row was just whole-moved to a fresh page top (see the row-level
                    // flag above). Reset to false once a content-bearing paragraph is
                    // placed so subsequent empties in the same cell render normally.
                    let mut s1067_skip_empties = s1067_row_at_page_top;
                    // An exact row clips an ordinary cell, but a restart cell
                    // owns the full vertical merge. Use its complete fixed
                    // height; an automatic continuation can grow for content.
                    let mut exact_cell_limit = is_exact.then_some(row_height);
                    if cell.v_merge.as_deref() == Some("restart") && is_exact {
                        for next_row in &table.rows[row_idx + 1..] {
                            let mut next_grid = next_row.grid_before as usize;
                            let continuation = next_row.cells.iter().find(|next_cell| {
                                let at_start = next_grid == cell_start_grid;
                                next_grid += next_cell.grid_span.max(1) as usize;
                                at_start
                            });
                            let continues = continuation.is_some_and(|next_cell|
                                next_cell.grid_span.max(1) == cell.grid_span.max(1)
                                    && matches!(next_cell.v_merge.as_deref(), Some("continue") | Some("")));
                            if !continues { break; }
                            exact_cell_limit = match (exact_cell_limit, next_row.height, next_row.height_rule.as_deref()) {
                                (Some(total), Some(height), Some("exact")) => Some(total + height),
                                _ => None,
                            };
                        }
                    }
                    for (block_pos, block) in cell.blocks.iter().enumerate() {
                        // S488: snapshot this block's content_h-relative top (aligns with
                        // block_pos via enumerate; pushed before the exact-clip break so
                        // blocks that fit are all recorded).
                        debug_assert_eq!(cell_block_tops.len(), block_pos);
                        cell_block_tops.push(content_h);
                        // Clip content that overflows exact row height
                        if exact_cell_limit.is_some_and(|limit| content_h + pad_t >= limit) {
                            break;
                        }
                        if Some(block_pos) == s716_stub_render || Some(block_pos) == s1311_tail_render {
                            continue;
                        }
                        if s751_hide_render {
                            continue; // S751
                        }
                        match block {
                            Block::Math(math_block)
                                if std::env::var("OXI_S1244_DISABLE").is_err() =>
                            {
                                // S1244: emit cell equations at the cell
                                // content origin and advance by the S652
                                // model — the estimate arm above mirrors this.
                                let mfs: f32 = 10.5;
                                let (me, mbb) = crate::layout::math::emit_math_block(
                                    math_block,
                                    cell_x + pad_l,
                                    content_h,
                                    mfs,
                                );
                                if !me.is_empty() {
                                    let adv = LayoutEngine::s1244_math_advance(&me, &mbb, mfs);
                                    for mut e in me {
                                        e.cell_row_index = Some(row_idx);
                                        e.cell_col_index = Some(cell_idx);
                                        cell_elements.push(e);
                                    }
                                    content_h += adv;
                                }
                            }
                            Block::Table(nested) => {
                                // COM-confirmed: nested table width = outer cell width - 2 × padding
                                let nested_x = cell_x + pad_l;
                                let nested_content_w = (cell_w - pad_l - pad_r).max(0.0);
                                let mut nested_y = LayoutCursor::new(content_h);
                                let mut dummy_pages = Vec::new();
                                let mut dummy_elems = Vec::new();
                                let nested_elements = self.layout_table(
                                    nested,
                                    nested_x,
                                    &mut nested_y,
                                    nested_content_w,
                                    table_grid_pitch,
                                    grid_char_pitch,
                                    grid_char_cw_ratio,
                                    0.0,
                                    99999.0,
                                    0.0,
                                    99999.0,
                                    &mut dummy_pages,
                                    &mut dummy_elems,
                                    block_idx,
                                    page,
                                    true,
                                    None,
                                    None,
                                    0.0,
                                    0.0,
                                    false, // S740: nested notes counted at the outer row
                                    None,
                                );
                                if std::env::var("OXI_DBG_NEST").is_ok() {
                                    let emax = nested_elements
                                        .iter()
                                        .map(|e| match &e.content {
                                            LayoutContent::TableBorder { y2, .. } => *y2,
                                            _ => e.y + e.height,
                                        })
                                        .fold(f32::NEG_INFINITY, f32::max);
                                    eprintln!("[NEST_DBG] cursor_y={:.2} visual_y={:.2} elem_max={:.2} n_elems={}",
                            nested_y.cursor_y, nested_y.visual_y, emax, nested_elements.len());
                                }
                                for mut elem in nested_elements {
                                    elem.cell_ancestor_path.insert(0, (row_idx, cell_idx, block_pos));
                                    cell_elements.push(elem);
                                }
                                // S756 (2026-07-06, default ON, opt-out OXI_S756_DISABLE):
                                // anchor the cell content FOLLOWING a nested table at the
                                // VISUAL track, not the cursor track. The render-only
                                // per-row overheads (S200/S661/S666 advance_split +0.5)
                                // accumulate on visual_y only — probenest's 20-row inner
                                // table painted to 843.5 while cursor_y said 834.0, so the
                                // following （続き） paragraph was placed 9.5pt INSIDE the
                                // nested table's last row (row-split continuation showed
                                // it overlapping at y=752.5 vs the row line at 748.5).
                                content_h = if std::env::var("OXI_S756_DISABLE").is_err() {
                                    nested_y.visual_y
                                } else {
                                    nested_y.cursor_y
                                };
                                prev_cell_sa = None; // S427: nested table breaks paragraph adjacency
                                s939_prev_r = None;
                                s1075_prev_r = None;
                            }
                            Block::Paragraph(para) => {
                                let hidden_final_mark = !self.is_vert_writing_active(cell)
                                    && LayoutEngine::hidden_cell_final_line(cell, block_pos, para);
                                                                let para = if hidden_final_mark { hidden_tail_para.as_ref().unwrap() } else { para };
                                // Session 131 (2026-05-20): vertical writing early-exit.
                                // For tbRlV cells, emit one Text element per paragraph at
                                // relative_y=0 (Word's COM Information(6) on a vert-cell
                                // paragraph returns the row-top y for all paragraphs in
                                // that cell). Cell content_h grows by vert_para_height so
                                // row-height calc reflects the vertical-text extent.
                                // The renderer (S132 GDI, S133 DWrite) is responsible for
                                // actual 90° CW rotation when emitting glyphs; this layout
                                // step only ensures positional correctness for pagination.
                                if self.is_vert_writing_active(cell)
                                    && std::env::var("OXI_S753_DISABLE").is_err()
                                {
                                    // S753 (2026-07-05): multi-column wrapped vert-cell emit —
                                    // the text wraps into the FINAL row height (chars per
                                    // column = floor((row_h−1)/fs)), columns advance
                                    // RIGHT→LEFT, each paragraph starts a new column, the
                                    // block is horizontally centred in the cell (overflow
                                    // past the cell border allowed = Word). content_h = the
                                    // real flow extent so the existing vAlign v_offset
                                    // centres the block vertically like Word (2ea81a top gap
                                    // 22.25 ≈ (row 120.6 − maxcol 72.1)/2). Derivation at
                                    // s753_vert_cell_columns.
                                    if s753_vert_cols.is_none() {
                                        // vMerge=restart: the flow height = the MERGED span,
                                        // not the single row (7ead52b 連絡担当窓口 spans 3
                                        // rows — single-row capacity wrongly wrapped its 6
                                        // chars into 2 columns where Word paints 1). Walk the
                                        // continue-rows (same grid-position logic as the
                                        // vAlign span walk at ~16715) summing declared
                                        // trHeights (row_height as fallback per row).
                                        let mut s753_avail = row_height;
                                        if cell.v_merge.as_deref() == Some("restart") {
                                            let target_grid = cell_start_grid;
                                            for next_ri in (row_idx + 1)..table.rows.len() {
                                                let next_row = &table.rows[next_ri];
                                                let mut next_grid = next_row.grid_before as usize;
                                                let mut continues = false;
                                                for next_cell in &next_row.cells {
                                                    let next_span =
                                                        next_cell.grid_span.max(1) as usize;
                                                    if next_grid == target_grid {
                                                        if matches!(
                                                            next_cell.v_merge.as_deref(),
                                                            Some("continue") | Some("")
                                                        ) {
                                                            continues = true;
                                                        }
                                                        break;
                                                    }
                                                    next_grid += next_span;
                                                }
                                                if !continues {
                                                    break;
                                                }
                                                s753_avail += next_row.height.unwrap_or(row_height);
                                            }
                                        }
                                        let (cols, max_len) = self.s753_vert_cell_columns(
                                            cell,
                                            cell_w,
                                            s753_avail,
                                            table_grid_pitch,
                                            table.style.para_style.as_ref(),
                                            grid_char_pitch,
                                            grid_char_cw_ratio,
                                        );
                                        // Cap at the single-row content box: a merged-span
                                        // flow (max_len > row) must not re-inflate THIS row
                                        // via max_actual_cell_h/S648 (the S751 lesson).
                                        content_h +=
                                            max_len.min((row_height - pad_t - pad_b).max(0.0));
                                        s753_vert_cols = Some((cols, max_len));
                                    }
                                    let first_run_style = para
                                        .runs
                                        .first()
                                        .map(|r| r.style.clone())
                                        .unwrap_or_default();
                                    let first_run_fs =
                                        self.resolve_font_size(&first_run_style, &para.style);
                                    let para_text: String =
                                        para.runs.iter().flat_map(|r| r.text.chars()).collect();
                                    let font_family = self
                                        .resolve_font_family_for_text(
                                            &para_text,
                                            &first_run_style,
                                            &para.style,
                                        )
                                        .map(|s| s.to_string());
                                    if let Some((cols, _)) = s753_vert_cols.as_ref() {
                                        for (bp, col_text, x_rel, col_w, _len) in cols.iter() {
                                            if *bp != block_pos || col_text.is_empty() {
                                                continue;
                                            }
                                            let mut elem = LayoutElement::new(
                                                cell_x + x_rel,
                                                0.0,
                                                *col_w,
                                                first_run_fs,
                                                LayoutContent::Text {
                                                    text: col_text.clone(),
                                                    font_size: first_run_fs,
                                                    font_family: font_family.clone(),
                                                    bold: self.resolve_bold(
                                                        &first_run_style,
                                                        &para.style,
                                                    ),
                                                    italic: first_run_style.italic,
                                                    underline: first_run_style.underline,
                                                    underline_style: first_run_style
                                                        .underline_style
                                                        .clone(),
                                                    strikethrough: first_run_style.strikethrough,
                                                    double_strikethrough: first_run_style
                                                        .double_strikethrough,
                                                    color: first_run_style.color.clone(),
                                                    highlight: first_run_style.highlight.clone(),
                                                    character_spacing: 0.0,
                                                    field_type: None,
                                                    text_scale: first_run_style
                                                        .text_scale
                                                        .unwrap_or(100.0),
                                                    is_vertical: true,
                                                    effects: TextEffects {
                                                        shadow: first_run_style.shadow,
                                                        emboss: first_run_style.emboss,
                                                        imprint: first_run_style.imprint,
                                                        outline: first_run_style.outline,
                                                        no_fill: first_run_style.no_fill,
                                                    },
                                                },
                                            );
                                            elem.paragraph_index = block_idx;
                                            elem.cell_paragraph_index = Some(cell_para_counter);
                                            elem.cell_row_index = Some(row_idx);
                                            elem.cell_col_index = Some(cell_idx);
                                            cell_elements.push(elem);
                                        }
                                    }
                                    cell_para_counter += 1;
                                    continue;
                                }
                                if self.is_vert_writing_active(cell) {
                                    let vert_h = self.vert_para_height(para);
                                    let first_run_style = para
                                        .runs
                                        .first()
                                        .map(|r| r.style.clone())
                                        .unwrap_or_default();
                                    let first_run_fs =
                                        self.resolve_font_size(&first_run_style, &para.style);
                                    let para_text: String =
                                        para.runs.iter().flat_map(|r| r.text.chars()).collect();
                                    let font_family = self
                                        .resolve_font_family_for_text(
                                            &para_text,
                                            &first_run_style,
                                            &para.style,
                                        )
                                        .map(|s| s.to_string());
                                    // Word's COM Information(6) returns the cell-top y for
                                    // ALL vert-cell paragraphs (verified on 2ea81a tbl=1
                                    // row=8: 予納する理由, （いずれかを選択）, empty para
                                    // all report y=478). Anchor all vert paragraphs at the
                                    // cell-entry content_h, not the running content_h.
                                    let mut elem = LayoutElement::new(
                                        cell_x + pad_l,
                                        vert_cell_anchor_h,
                                        (cell_w - pad_l - pad_r).max(0.0),
                                        first_run_fs,
                                        LayoutContent::Text {
                                            text: para_text,
                                            font_size: first_run_fs,
                                            font_family,
                                            bold: self.resolve_bold(&first_run_style, &para.style),
                                            italic: first_run_style.italic,
                                            underline: first_run_style.underline,
                                            underline_style: first_run_style
                                                .underline_style
                                                .clone(),
                                            strikethrough: first_run_style.strikethrough,
                                            double_strikethrough: first_run_style
                                                .double_strikethrough,
                                            color: first_run_style.color.clone(),
                                            highlight: first_run_style.highlight.clone(),
                                            character_spacing: 0.0,
                                            field_type: None,
                                            text_scale: first_run_style.text_scale.unwrap_or(100.0),
                                            // Session 132: flag for renderer rotation.
                                            is_vertical: true,
                                            effects: TextEffects {
                                                shadow: first_run_style.shadow,
                                                emboss: first_run_style.emboss,
                                                imprint: first_run_style.imprint,
                                                outline: first_run_style.outline,
                                                no_fill: first_run_style.no_fill,
                                            },
                                        },
                                    );
                                    elem.paragraph_index = block_idx;
                                    elem.cell_paragraph_index = Some(cell_para_counter);
                                    elem.cell_row_index = Some(row_idx);
                                    elem.cell_col_index = Some(cell_idx);
                                    cell_elements.push(elem);
                                    content_h += vert_h;
                                    cell_para_counter += 1;
                                    continue;
                                }
                                // S1067: collapse a just-whole-moved row's leading empty
                                // paragraphs (before ANY spacing math — Word's p6 opens with
                                // "b." as the first line; the 6 empties render ~0, so no
                                // space_before survives either). The skip predicate is the
                                // codebase's standard "no visible text" test. Reset the flag
                                // so a later empty in the same cell renders full height.
                                if s1067_skip_empties
                                    && para.runs.iter().all(|r| r.text.trim().is_empty())
                                {
                                    continue;
                                }
                                s1067_skip_empties = false;
                                // Resolve cell line spacing in the same order as the
                                // height estimate: direct/paragraph style, table style,
                                // then document defaults.
                                let cell_default_line_spacing = table.style.para_style.as_ref()
                                    .filter(|ps| ps.line_spacing.is_some());
                                // A table overrides document defaults only when it
                                // declares line spacing, as in the height estimate.
                                let preserve_cell_defaults =
                                    std::env::var("OXI_CELL_DEFAULT_LINE_DISABLE").is_err();
                                let effective_line_spacing =
                                    if para.style.line_spacing_from_doc_defaults {
                                        if preserve_cell_defaults {
                                            cell_default_line_spacing.and_then(|ps| ps.line_spacing)
                                                .or(para.style.line_spacing)
                                        } else { None }
                                    } else {
                                        para.style.line_spacing.or_else(|| {
                                            table
                                                .style
                                                .para_style
                                                .as_ref()
                                                .and_then(|ps| ps.line_spacing)
                                        })
                                    };
                                let effective_line_rule =
                                    if para.style.line_spacing_from_doc_defaults {
                                        if preserve_cell_defaults {
                                            cell_default_line_spacing.map_or(
                                                para.style.line_spacing_rule.as_deref(),
                                                |ps| ps.line_spacing_rule.as_deref())
                                        } else { None }
                                    } else {
                                        para.style.line_spacing_rule.as_deref().or_else(|| {
                                            table
                                                .style
                                                .para_style
                                                .as_ref()
                                                .and_then(|ps| ps.line_spacing_rule.as_deref())
                                        })
                                    };
                                let style_has_explicit_rule = effective_line_rule == Some("exact")
                                    || effective_line_rule == Some("atLeast");
                                // S855 (2026-07-15): render cell before/after reset keys on
                                // whether the DIRECT pPr set before/after (has_direct_before_after),
                                // not has_direct_spacing — a direct line-only spacing must not
                                // preserve the docDefaults-inherited before/after in a cell.
                                // Kept in sync with estimate_para_height_inner / cell_para_spacing.
                                let (reset_before, reset_after) = self.cell_spacing_reset_sides(
                                    &para.style,
                                    style_has_explicit_rule,
                                    true,
                                );
                                let tbl_has_ls = table
                                    .style
                                    .para_style
                                    .as_ref()
                                    .and_then(|ps| ps.line_spacing)
                                    .is_some();
                                // S699 (2026-06-30): the table-style line-spacing override must NOT fire
                                // when the paragraph STYLE itself sets explicit line spacing. ECMA-376
                                // precedence (see comment above) is table style pPr < paragraph style <
                                // direct, so a paragraph-style value (line_spacing present and NOT inherited
                                // from docDefaults) outranks the table style. Without this guard, a cell
                                // whose Normal style sets line=276 (1.15x) was wrongly reset to the
                                // TableGrid style's line=240 (Single) → rows ~1.92pt too short, cumulative
                                // (test_table_grid: Oxi 13.44pt/row vs Word 15.36, drift ~9.6pt over 5
                                // rows). Corpus blast radius = 1 (test_table_grid is the only doc with a
                                // Normal-style line!=240 + a styled table); gen2 (docDefaults line=276,
                                // from_doc_defaults=true) and the form family are unchanged. Opt-out
                                // OXI_S699_DISABLE.
                                let para_style_explicit_ls = para.style.line_spacing.is_some()
                                    && !para.style.line_spacing_from_doc_defaults
                                    && std::env::var("OXI_S699_DISABLE").is_err();
                                let (effective_line_spacing, effective_line_rule) = if tbl_has_ls
                                    && !para.style.has_direct_spacing
                                    && !para_style_explicit_ls
                                {
                                    let tbl_ls = table
                                        .style
                                        .para_style
                                        .as_ref()
                                        .and_then(|ps| ps.line_spacing);
                                    let tbl_lr = table
                                        .style
                                        .para_style
                                        .as_ref()
                                        .and_then(|ps| ps.line_spacing_rule.as_deref());
                                    (tbl_ls, tbl_lr)
                                } else {
                                    (effective_line_spacing, effective_line_rule)
                                };
                                // S136 (2026-05-20): OXI_SB_NO_SUPPRESS=1 disables the first-cell-para
                                // sb suppression (Day 33 part 17). TR_V200-V203 + R1A re-measurement
                                // show Word DOES apply sb. Default off; env var enables revert behavior.
                                // S239 (2026-05-23): removed OXI_LEGACY_SB_SUPPRESS and
                                // OXI_SB_NO_SUPPRESS legacy env-var fallbacks (LEGACY var
                                // default false → suppression branch was dead). S151
                                // default ON since 2026-05-21.
                                // S936: a docDefaults-sourced side takes the table style's
                                // declared value (cell_para_spacing / estimate mirror).
                                let (s936_sb, s936_sa) = self.s936_tbl_style_dd_override(
                                    &para.style,
                                    table.style.para_style.as_ref(),
                                );
                                let effective_space_before = if let Some(v) = s936_sb {
                                    v
                                } else if reset_before {
                                    // Day 33 part 17 (2026-05-10): Word suppresses spacing.before
                                    // for the first paragraph in a cell. COM-confirmed via 8 repros
                                    // (row1_attr_isolation): R1A_spacing_lineRule has spacing.before=4.35pt
                                    // + lineRule=exact 12pt → Word renders row at 12.5pt (no spacing
                                    // applied), Oxi was rendering at 16.85pt (+4.35pt over-pump).
                                    // Same for R1A_all4. Affects 备考 cluster docs (d4d126/de6e/etc)
                                    // where row 1 cell has style "ac"+spacing.before+lineRule=exact.
                                    0.0
                                } else if let (Some(bl), Some(pitch)) =
                                    (para.style.before_lines, table_grid_pitch)
                                {
                                    bl / 100.0 * pitch
                                } else {
                                    para.style
                                        .space_before
                                        .or_else(|| {
                                            table
                                                .style
                                                .para_style
                                                .as_ref()
                                                .and_then(|ps| ps.space_before)
                                        })
                                        .unwrap_or(0.0)
                                };
                                let effective_space_after = if let Some(v) = s936_sa {
                                    Some(v)
                                } else if reset_after {
                                    None
                                } else if let (Some(al), Some(pitch)) =
                                    (para.style.after_lines, table_grid_pitch)
                                {
                                    // Session 94 (2026-05-18) fix: afterLines was parsed into
                                    // IR but not applied in cell rendering path. Body path at
                                    // mod.rs:4554 already had this. Symmetric with before_lines
                                    // handling at mod.rs:6497. TR33 (afterLines only, no after
                                    // twip) measured Word pitch 13.50pt vs Oxi 12.00pt = +1.5pt
                                    // gap closed by reading afterLines.
                                    Some(al / 100.0 * pitch)
                                } else {
                                    para.style.space_after.or_else(|| {
                                        table
                                            .style
                                            .para_style
                                            .as_ref()
                                            .and_then(|ps| ps.space_after)
                                    })
                                };
                                // Cell autospace override (before/afterAutospacing → 13.75,
                                // overriding explicit, edge-suppressed). See cell_effective_spacing.
                                let s952_tbl =
                                    table.style.para_style.as_ref().map_or(false, |ts| {
                                        ts.before_autospacing || ts.after_autospacing
                                    });
                                let (effective_space_before, effective_space_after) =
                                    if para.style.before_autospacing
                                        || para.style.after_autospacing
                                        || para.style.contextual_spacing
                                        || s952_tbl
                                    {
                                        let (sb, sa) = self.cell_effective_spacing(
                                            para,
                                            table.style.para_style.as_ref(),
                                            Some(block_pos) == first_para_pos,
                                            Some(block_pos) == last_para_pos,
                                            effective_space_before,
                                            effective_space_after.unwrap_or(0.0),
                                        );
                                        (sb, Some(sa))
                                    } else {
                                        (effective_space_before, effective_space_after)
                                    };
                                if fragment_valign {
                                    fragment_paragraph_after.insert((cell_idx, cell_para_counter),
                                        effective_space_after.unwrap_or(0.0));
                                }
                                if std::env::var_os("OXI_CELL_EMPTY_LINES").is_some()
                                    && Some(block_pos) == last_para_pos
                                {
                                    cell_terminal_spacing.insert(
                                        (cell_idx, cell_para_counter),
                                        effective_space_after.unwrap_or(0.0) + pad_b,
                                    );
                                }
                                // S427: collapse this paragraph's space_before against the
                                // previous cell paragraph's space_after (max(sa,sb), not sum).
                                if s427_collapse {
                                    if let Some(psa) = prev_cell_sa {
                                        content_h -= psa.min(effective_space_before);
                                    }
                                }
                                // A split continuation carries the paragraph's resolved
                                // before spacing, after same-style contextual suppression.
                                // Keep adjacency before updating the previous-style state.
                                let carry_space_before = if !self.preserve_same_style_cell_spacing
                                    && !self.doc_body_has_real_cjk
                                    && std::env::var("OXI_S939_DISABLE").is_err()
                                    && prev_cell_sa.is_some()
                                    && para.style.contextual_spacing
                                    && s939_prev_r.is_some_and(|p| para.style.style_id.as_deref() == p.1)
                                { 0.0 } else {
                                    // The previous after spacing already owns its part
                                    // of the collapsed boundary. Only the remaining
                                    // before spacing travels with a fresh paragraph.
                                    let previous_credit = if s427_collapse {
                                        prev_cell_sa.map_or(0.0, |after| after.min(effective_space_before))
                                    } else { 0.0 };
                                    (effective_space_before - previous_credit).max(0.0)
                                };
                                // S939: layered contextualSpacing collapse inside the cell.
                                content_h -= self.s939_cell_ctx_credit(
                                    prev_cell_sa,
                                    s939_prev_r.map_or(false, |p| p.0),
                                    s939_prev_r.and_then(|p| p.1),
                                    &para.style,
                                    effective_space_before,
                                );
                                // S1075: same-list adjacency inside the cell.
                                content_h -= self.s1075_cell_list_credit(
                                    prev_cell_sa,
                                    s1075_prev_r.map_or(false, |p| p.0),
                                    s1075_prev_r.and_then(|p| p.1),
                                    &para.style,
                                    effective_space_before,
                                );
                                s939_prev_r = Some((
                                    para.style.contextual_spacing,
                                    para.style.style_id.as_deref(),
                                ));
                                s1075_prev_r = Some((
                                    para.style.after_autospacing,
                                    para.style.num_id.as_deref(),
                                ));
                                content_h += effective_space_before;
                                s1431_cell_para_sb.insert((cell_idx, cell_para_counter), carry_space_before);
                                if cell_float_flow {
                                    let floor = float_replay.and_then(|r| r.origins.get(&(row_idx, cell_idx, block_pos)))
                                        .copied().unwrap_or(0.0);
                                    content_h = content_h.max(floor);
                                }
                                let para_content_start_h = content_h;
                                if cell_float_flow { float_tops[block_pos] = content_h; }
                                {
                                    // Paragraph indentation within cell (relative to cell content area)
                                    // COM-confirmed: *Chars multiplier = 10.5pt always
                                    let p_indent_left = para
                                        .style
                                        .indent_left
                                        .or_else(|| {
                                            self.s1349_left_pt(para, grid_char_pitch, grid_char_cw_ratio)
                                        })
                                        .unwrap_or(0.0);
                                    let p_indent_right = para
                                        .style
                                        .indent_right
                                        .or_else(|| {
                                            para.style.indent_right_chars.map(|c| self.s1349_default_chars_pt(c, para, grid_char_pitch, grid_char_cw_ratio))
                                        })
                                        .unwrap_or(0.0);
                                    // When both firstLine (twip) and firstLineChars exist,
                                    // twip value is authoritative (pre-computed by Word).
                                    let p_first_line_indent_raw = para
                                        .style
                                        .indent_first_line
                                        .or_else(|| {
                                            para.style
                                                .indent_first_line_chars
                                                .map(|c| LayoutEngine::s1214_chars_pt(c, para, true, grid_char_pitch, grid_char_cw_ratio))
                                        })
                                        .unwrap_or(0.0);
                                    // COM-confirmed (2026-04-25): numbered list + hanging + suff=tab/default
                                    // => marker consumes hanging, text starts at `left`. See body path.
                                    let p_list_consumes_hanging = para.style.list_marker.is_some()
                                        && p_first_line_indent_raw < 0.0
                                        && matches!(
                                            para.style.list_suff.as_deref(),
                                            None | Some("tab")
                                        );
                                    let p_first_line_indent = if p_list_consumes_hanging {
                                        0.0
                                    } else {
                                        p_first_line_indent_raw
                                    };
                                    // Day 33 part 57 (2026-05-12): use cell_w (not inner_w with padding
                                    // subtracted) for wrap width. Matches estimate path comment at
                                    // mod.rs:5677: "Word allows text to extend into cell margins for
                                    // wrapping purposes". 191cb row 3 cell 0 (16 CJK chars, cell_w=104pt,
                                    // inner_w=94.1pt): Oxi was wrapping at 8 chars (94.1pt limit) but
                                    // Word wraps at 9 chars (94.5pt fits in 104pt). The estimate-vs-
                                    // render inconsistency was the source of the over-pump.
                                    //
                                    // Session 126 (2026-05-20) — A/B tested OXI_CELL_INNER_WRAP=1
                                    // (= switch wrap_base to inner_w). Phase 1: 53/55 → 49/55. 3a4f
                                    // went 11 paras delta=-1 → 1314 paras delta=+1 (catastrophic).
                                    // Confirms Word's rule is doc-dependent: 191cb uses cell_w extension,
                                    // b35 uses sub-inner_w fill-justify. No simple toggle works.
                                    // Pre-S125 conclusion "accept b35 limit" re-confirmed.
                                    // S172 (2026-05-22): conditional inner_w wrap for d77a-class cells.
                                    // Discriminator: hanging-indent paragraph + single-cell row + cell
                                    // within body width. This matches d77a/29dc6e/b35/31420af's
                                    // structure (single-column body-width-sized tables with hanging
                                    // paragraphs) while excluding 1636d (multi-cell), a47e (cell > body),
                                    // and 191cb (multi-cell narrow).
                                    // S237 (2026-05-23): removed OXI_LEGACY_NO_CELL_HANG_INNER
                                    // legacy env-var fallback during hardening pass.
                                    // S301 (2026-05-26): subtract cell padding from wrap budget for
                                    // 2-cell-row hanging-indent paragraphs in tblLayout="fixed" tables
                                    // when the paragraph (or its style chain) sets `<w:wordWrap w:val="0"/>`.
                                    // COM-confirmed discriminator vs 191cb (regressed with broader gate):
                                    //   29dc6e/d4d126: pStyle="ac" → wordWrap=0 inherited → Word subtracts
                                    //   191cb: paragraph has no pStyle, default wordWrap=true → Word doesn't
                                    // The wordWrap=0 paragraphs are CJK-aware (line-break anywhere in CJK,
                                    // including mid-word for Latin). Word treats their wrap budget more
                                    // conservatively (subtracts cellMar). Standard wordWrap=true paragraphs
                                    // wrap on word boundaries and use the full cell width.
                                    // Env-gated default ON now that the discriminator is tight enough:
                                    //   OXI_S301_DISABLE=1 reverts to pre-S301 (S172-only) behavior.
                                    let cell_hang_inner = p_first_line_indent_raw < 0.0
                                        && row.cells.len() == 1
                                        && cell_w <= content_width;
                                    let s301_layout_fixed = std::env::var("OXI_S301_DISABLE")
                                        .is_err()
                                        && table.style.layout.as_deref() == Some("fixed")
                                        && (pad_l + pad_r) > 0.0
                                        && row.cells.len() == 2
                                        && cell_w <= content_width
                                        && !para.style.word_wrap; // tight discriminator: wordWrap=0 only
                                                                  // S413 (2026-05-29) — gate v4 implementation behind
                                                                  // OXI_S412_ENABLE (default OFF / opt-in). Default
                                                                  // behavior unchanged — gate only fires when the env
                                                                  // var is set, allowing local A/B validation against
                                                                  // ed025 + 1ec1 without baseline risk. See full
                                                                  // discriminator rationale at S411/S412 comment
                                                                  // block below.
                                                                  //
                                                                  // S413 A/B VALIDATION RESULT (full renderer rebuild,
                                                                  // ed025+1ec1 caches cleared, OFF vs ON):
                                                                  //   Phase 1 (pagination): 53/55 UNCHANGED. No page-break
                                                                  //     movement; ed025 per-page para counts identical
                                                                  //     (kinsoku force-fit blocks the 2-line rewrap per S409).
                                                                  //   Phase 2 (element IoU): UNCHANGED. ed025 0.9179,
                                                                  //     1ec1 0.9853 — ZERO per-element IoU delta on both.
                                                                  //     Element IoU measures cell/line bbox, NOT text-start
                                                                  //     x, so the intra-cell text shift is invisible to it.
                                                                  //   Gate firing confirmed: text-start x shifts exactly
                                                                  //     -9.9pt (= cellMar 99+99 dxa) on every fire cell.
                                                                  //   Direct Word comparison (text-matched cells):
                                                                  //     1ec1 i=37 "　　　　○": Word x=316.0,
                                                                  //       Oxi OFF=356.45 (Δ40.5), ON=346.55 (Δ30.6)
                                                                  //     ed025 × col: Word x=364.0,
                                                                  //       Oxi OFF=401.75 (Δ37.8), ON=391.85 (Δ27.9)
                                                                  //     → ON moves Oxi +9.9pt TOWARD Word on BOTH docs
                                                                  //       (cellMar subtraction is DIRECTIONALLY CORRECT),
                                                                  //       but a ~28-31pt residual cell-x offset remains
                                                                  //       (pre-existing, larger than cellMar, NOT addressed
                                                                  //       by this gate — likely cell column x-origin).
                                                                  // DECISION: KEEP default OFF. Gate is directionally
                                                                  // validated but yields no Phase 1/Phase 2 gain (does not
                                                                  // meet "IoU strictly increases" merge gate). Scaffold
                                                                  // retained for combined future work: (a) Phase 3 SSIM
                                                                  // gate where intra-cell text-x becomes visible,
                                                                  // (b) the ~28pt residual cell-x fix, (c) S409 kinsoku
                                                                  // rebalance to actually rewrap ed025.
                                                                  //
                                                                  // S414 (2026-05-29) CAVEAT — the model below is on
                                                                  // SHAKY GROUND. The fire cells are predominantly
                                                                  // jc=RIGHT-ALIGNED (ed025 226/262, 1ec1 i=37), not
                                                                  // left-edge-wrapped. For right-aligned text Word anchors
                                                                  // at content_right and position is set by TEXT WIDTH, not
                                                                  // a left-edge wrap budget. The S413 -9.9pt shift toward
                                                                  // Word was coincidence of magnitude (~cellMar), not the
                                                                  // correct mechanism. The real ~40pt residual (1ec1 col3
                                                                  // "　　　　税": Word x=316.0 vs Oxi 356.45; neighbors
                                                                  // col0/col2/col4 all match Word) is specific to
                                                                  // right-aligned + firstLine-indent cells and needs COM
                                                                  // glyph measurement before any fix. This gate may be
                                                                  // RETIRED in S415+ rather than promoted. Do NOT enable
                                                                  // by default without re-deriving from right-aligned
                                                                  // positioning data.
                                                                  // S418: discriminator now uses has_explicit_cellmar
                                                                  // (author-declared <w:tblCellMar> in this table's
                                                                  // tblPr) instead of the default_cell_margins.is_some()
                                                                  // PROXY. S417e caught the proxy over-firing on 04b88e
                                                                  // (which has default margins but no explicit tblCellMar)
                                                                  // and regressing its x-fidelity 0.7309 -> 0.7000. The
                                                                  // explicit-only condition matches the S412 v4 analysis
                                                                  // (262 ed025 + 1 1ec1, 0 in 04b88e/3a4f/51 others).
                                                                  // S419 SHIP (2026-05-29): default ON (opt-out
                                                                  // OXI_S412_DISABLE, S301 pattern). COM-validated
                                                                  // correctness fix — matches Word TRUE rendering
                                                                  // (S416 GetPoint): ed025/1ec1 right-aligned firstLine
                                                                  // cellMar cells move to Word's rendered x (1ec1 col4
                                                                  // x-IoU 0.84->0.998). Ships on its own merit like S408:
                                                                  // the phase gates can't see horizontal fixes (Phase 1
                                                                  // x-independent, Phase 2 Y-only) but it regresses none
                                                                  // (Phase 1 53/55, Phase 2 0.9647, lib 142/0/6) and the
                                                                  // x_fidelity_diff diagnostic confirms the improvement.
                                    let s412_cellmar_subtract = std::env::var("OXI_S412_DISABLE")
                                        .is_err()
                                        && p_first_line_indent_raw > 0.0
                                        && para.style.indent_first_line_chars.is_some()
                                        && row.cells.len() >= 3
                                        && table.style.layout.as_deref() != Some("fixed")
                                        && table.style.has_explicit_cellmar
                                        && cell_w <= content_width;
                                    // S405-S411 ed025 chain (2026-05-28):
                                    // S408 shipped × U+00D7 fullwidth correctness fix (safe).
                                    // S409 isolated S405 padding-subtract impact:
                                    //   - Only 2 docs regress: 3a4f (-0.6415), 04b88e (-0.3905)
                                    //   - 53 other docs unchanged
                                    //   - ed025 score UNCHANGED even with S405 because Oxi's
                                    //     kinsoku force-fit puts `）` on same line (line-start
                                    //     prohibited → forced onto current line, no actual
                                    //     wrap to 2 lines)
                                    // → S405 alone doesn't even fix ed025. Need BOTH:
                                    //   1. narrower S405 gate (avoid 3a4f/04b88e regression)
                                    //   2. kinsoku REBALANCE algorithm (look BACKWARD when
                                    //      prohibited char would be alone on next line —
                                    //      pull preceding char too so prohibited char has
                                    //      companion). Current Oxi force-fits in this case.
                                    // ed025 needs (1) AND (2) together to render 2 lines
                                    // matching Word's 5-char + 2-char split.
                                    //
                                    // S411 (2026-05-28) — narrower S405 gate hypothesis v3 from
                                    // XML attribute comparison across ed025 / 3a4f / 04b88e:
                                    // Candidate gate: `has_tblCellMar AND cells_in_row >= 3
                                    //                  AND tblLayout != "fixed"`.
                                    // Per-doc fire counts on positive-firstLine table cells:
                                    //   ed025  : 262/381 (fires on T16 target + similar tables)
                                    //   3a4f   :   2/177 (down from 192 unrestricted)
                                    //   04b88e :   0/47  (FULLY protected)
                                    // The 2 residual 3a4f cells are in a nested table with
                                    // firstLine=5twip (0.25pt — negligible) and NO
                                    // firstLineChars attribute (raw twip indent, not
                                    // char-based).
                                    //
                                    // S412 (2026-05-28) — STRENGTHENED gate v4: add
                                    // `firstLineChars is not None` constraint. Discriminator
                                    // interpretation: Word subtracts cellMar ONLY when
                                    // (a) author explicitly declared tblCellMar in tblPr,
                                    // (b) row is multi-column, (c) layout is auto, AND
                                    // (d) indent is CHAR-BASED (firstLineChars set,
                                    //     signalling CJK-aware authoring intent).
                                    // Raw-twip firstLine without firstLineChars uses cell_w
                                    // as-is (3a4f nested table pattern).
                                    //
                                    // Corpus-wide v4 fire counts (54/55 baseline docs walked):
                                    //   ed025 : 262 cells (target tables; T16 + similar)
                                    //   1ec1  :   1 cell  (NEW; 6-cell row + cellmar=99/99
                                    //                      + firstLineChars=200 — structurally
                                    //                      identical to ed025 rule)
                                    //   3a4f  :   0 cells (FULLY PROTECTED — both v3 residuals
                                    //                     lacked firstLineChars)
                                    //   04b88e:   0 cells (FULLY PROTECTED — no tblCellMar)
                                    //   51 other docs: 0 fires
                                    // Total: 263 cells across 2 docs only.
                                    //
                                    // Status: HYPOTHESIS (not implemented). 1ec1 is currently
                                    // Phase 1 PASS (score 1.0) IoU 0.9853 — applying the gate
                                    // may improve or regress it. Pre-ship validation needed:
                                    // (1) COM-measure 1ec1's tbl[0] tr[2] tc[3] p[0] cell to
                                    //     verify Word's wrap budget there. Same for one ed025
                                    //     T16 cell. Both should show cell_w - cellMar usage.
                                    // (2) Implement gate behind OXI_S412_DISABLE env var,
                                    //     A/B test on baseline. ed025 corpus score will
                                    //     likely stay flat without kinsoku rebalance
                                    //     (S409 blocker); 1ec1 is the leading validator.
                                    // (3) ed025 full improvement requires BOTH S412 gate
                                    //     AND kinsoku rebalance per S409.
                                    //
                                    // Both ed025 and 3a4f's with-tblCellMar tables use
                                    // identical cellmar=99/99 dxa — value alone is not the
                                    // discriminator (rejected). The PRESENCE of
                                    // firstLineChars + tblCellMar + multi-column +
                                    // auto-layout is the discriminator.
                                    // S531 (2026-06-09): a SINGLE-cell table reserves its cellMar as
                                    // padding, so the wrap budget is cell_w - pad_l - pad_r (like Word).
                                    // 683f's `af`-styled 解説 cell (style cellMar 108/108, single cell,
                                    // body-width, left/justified flowing text) fit 45 chars/line vs Word's
                                    // 44 because wrap_base used the full cell_w. Gated to:
                                    //   - single-cell rows (row.cells.len()==1): excludes the S412/S417e
                                    //     right-aligned MULTI-column tabular cells (04b88e x-anchor
                                    //     regression came from those; this never touches them).
                                    //   - default_cell_margins.is_some(): a REAL declared/inherited cellMar
                                    //     (the 4.95pt hardcoded fallback is None -> never fires).
                                    //   - non right/center alignment: cellMar-as-wrap-budget applies to
                                    //     left-to-right flowing/justified text, not right-anchored.
                                    //   - cell_w <= content_width: a body-width single-cell block.
                                    // cell_hang_inner already covers single-cell HANGING-indent paras; this
                                    // adds the non-hanging case. opt-out OXI_S531_DISABLE.
                                    // !has_explicit_cellmar: only when the cellMar is INHERITED (from the
                                    // table style / default table style), NOT author-declared in this
                                    // table's tblPr. 6295e189's form cells set tblCellMar=52tw directly in
                                    // tblPr (has_explicit_cellmar=true) and Word does NOT reduce their wrap
                                    // budget there (subtracting regressed it -0.0036); 683f's `af`-style
                                    // cellMar is inherited (has_explicit_cellmar=false) and Word DOES reduce
                                    // it. This also keeps s531 out of S412's author-declared territory.
                                    let s531_singlecell_cellmar = std::env::var("OXI_S531_DISABLE")
                                        .is_err()
                                        && row.cells.len() == 1
                                        && table.style.default_cell_margins.is_some()
                                        && !table.style.has_explicit_cellmar
                                        && !matches!(
                                            para.alignment,
                                            Alignment::Right | Alignment::Center
                                        )
                                        && cell_w <= content_width;
                                    // S559 SHIP (2026-06-13, default ON, opt-out OXI_S559_DISABLE): a
                                    // JUSTIFIED single-cell AUTOFIT-SQUEEZED table reserves Word's DEFAULT
                                    // 108tw cellMar even though default_cell_margins.is_none() (so the s531
                                    // gate above — which requires is_some() — never fires on it). This is
                                    // the 3a4f para-2234 ⑦ over-pack: Oxi packed ⑦ on 1 line (39 chars),
                                    // Word wraps to 2 (L1=37; COM: gridCol 8244 − cellMar 216 − firstLine
                                    // 210 = 7818tw). The −18pt loss pulled para 2260 to p80 (Word p81) =
                                    // 3a4f's sole Phase-1 FAIL.
                                    // DISCRIMINATOR (why ⑦ reserves but the 86 same-signature cells don't):
                                    //   - row.cells.len()==1, !has_explicit_cellmar, non-right (= s531 scope)
                                    //   - tcW − gridCol >= 8pt: the cell is AUTOFIT-SQUEEZED (its preferred
                                    //     width exceeds the laid column by ~one cellMar; ⑦ tcW 8458 >
                                    //     gridCol 8244 = 214tw). diff==0 cells (got their preferred width)
                                    //     are excluded.
                                    //   - JUSTIFIED (jc=both): ⑦ is jc=both via style a7; the structurally
                                    //     IDENTICAL p19 cell is explicit jc=left and Word does NOT reserve
                                    //     cellMar there (Oxi-OFF matched it at 2 lines). Firing on p19 (jc=
                                    //     left) over-wrapped it 2→3, and the +1 line at page 19 cascaded the
                                    //     {1:1323} pagination regression. Restricting to Justify excludes p19.
                                    // VALIDATION: full corpus Phase-1 pagination 54/55 → 55/55, 0 PASS→FAIL,
                                    // mean_score 1.0000. Only 2 paras change line count corpus-wide under the
                                    // rule (⑦ 1→2 = the fix; one p94 justified cell 119→121, post-2260, no
                                    // new delta). NOTE: jc-vs-left is the empirical discriminator here but ⑦
                                    // also has left=0 while p19 has left=459, so the two are confounded —
                                    // "justl0" (Justify AND left≈0) gives identical corpus results. The pad
                                    // subtracted is Oxi's 4.95pt fallback (Word's true default is 5.4pt/108tw;
                                    // the 0.9pt gap is within ⑦'s wrap slack so the rule still fires correctly).
                                    // Env OXI_S559_CELLMAR overrides the rule for A/B testing: all / tcwgt /
                                    // just (= default) / justl0.
                                    let s559_disabled = std::env::var("OXI_S559_DISABLE").is_ok();
                                    let s559_mode = std::env::var("OXI_S559_CELLMAR").ok();
                                    let s559_tcw_gt =
                                        cell.width.map_or(false, |tcw| tcw - cell_w >= 8.0);
                                    let s559_base = !s559_disabled
                                        && row.cells.len() == 1
                                        && !table.style.has_explicit_cellmar
                                        && !matches!(
                                            para.alignment,
                                            Alignment::Right | Alignment::Center
                                        )
                                        && cell_w <= content_width;
                                    let s559_justified = matches!(
                                        para.alignment,
                                        Alignment::Justify | Alignment::Distribute
                                    );
                                    let s559_cellmar = s559_base
                                        && match s559_mode.as_deref() {
                                            Some("all") => true,
                                            Some("tcwgt") => s559_tcw_gt,
                                            Some("justl0") => {
                                                s559_tcw_gt
                                                    && s559_justified
                                                    && p_indent_left.abs() < 1.0
                                            }
                                            // default (None) or "just" = the shipped rule
                                            _ => s559_tcw_gt && s559_justified,
                                        };
                                    // S562 (2026-06-14): a hanging+span>1 cellMar-subtract gate
                                    // (OXI_S562) was PROTOTYPED here for roudoujoken's r7 (5)裁量
                                    // cell and CONFIRMED to render r7 correctly (17→18 lines = Word)
                                    // — but the roudoujoken −1 pagination was UNCHANGED. So the r7
                                    // cell over-fit is a REAL but NON-OPERATIVE bug: its +1 line
                                    // (~14pt) on the pages-1-2 form table does NOT tip ８.「休暇」 on
                                    // page 3. The operative −1 cause is a page-3 cascade still
                                    // unidentified (r7, s475/pi16, and the 記載要領 paras all ruled
                                    // out). Gate removed (no merge-gate benefit + cell-wrap risk).
                                    // See memory session560 for the full ruled-out chain.
                                    // S585 (2026-06-16, default ON, opt-out OXI_S585_DISABLE):
                                    // a full-page-width single-cell table whose declared gridCol
                                    // SLIGHTLY exceeds the page content area subtracts its cellMar
                                    // from the wrap budget (Word fits the content to the page).
                                    // tokyoshugyo's regulation-box tables: cell_w=427.85 (fixed
                                    // layout keeps the declared gridCol 8557tw) > content_width
                                    // 8504tw=425.2 — Word's wrap ≈ gridCol − 2×cellMar (the ④para
                                    // line is 412.4pt wide, NOT the full 427.85); Oxi used
                                    // wrap_base=cell_w → fit ~1-2 more chars/line → 16 paras
                                    // under-wrap by 1 line → the doc-wide −1 page drift.
                                    // The +5pt cell_w over (Oxi's right border at content+2×cellMar
                                    // vs Word's +1×) DISQUALIFIES s531/s559 (both require
                                    // cell_w ≤ content_width), so this gate handles cell_w > content.
                                    // DISCRIMINATOR `over < 5pt`: a TRUE full-page table exceeds the
                                    // content area by < one cellMar (tokyoshugyo +2.65pt = ~half the
                                    // 99tw cellMar). A genuinely-WIDE table (harassbun +19.65,
                                    // 1636 +7.1) overflows the page deliberately and Word keeps its
                                    // full cell_w — firing on those over-wrapped them PASS→FAIL
                                    // (the cell-wrap tombstone). Single-cell, non-right, inherited
                                    // cellMar (!has_explicit_cellmar — author-declared cellMar is
                                    // S412/S418 territory). Corpus-validated: Phase-1 65/69 unchanged
                                    // (0 PASS→FAIL; 1636/harassbun stay PASS), tokyoshugyo
                                    // 0.7107→0.8071 (page count 89→90=Word). RESIDUAL +1×282
                                    // oscillation = per-cell wrap-narrowing variance (Word narrows
                                    // some cells by < 2×cellMar) — the page COUNT is right but the
                                    // per-cell line distribution isn't exact (deferred). See
                                    // [[tokyoshugyo_wrap_not_cellheight]].
                                    // S591 (2026-06-16): the S585b single-cell cellMar-subtract
                                    // DISCRIMINATOR. The over-amount alone is FALSIFIED (canary:
                                    // OXI_S585_OVER=11 → 1636 PASS→FAIL — 1636's over∈[5,11) cell
                                    // is in a tblW=dxa table Word keeps wide). The TRUE rule is
                                    // tblW TYPE (docx tblGrid + 3-doc analysis): a tblW=auto table
                                    // is AUTO-SIZED → Word fits it to the page content (CLAMP);
                                    // tblW=dxa declares an explicit width → Word HONORS it
                                    // (KEEP-WIDE, overflow). tokyoshugyo's regulation boxes
                                    // (T15/T41/T57/T101, over +4.5..+9.9) are ALL tblW=auto → clamp;
                                    // harassbun (+19.6) and 1636 (+14.3) are tblW=dxa → keep.
                                    // RULE: clamp if over<5.0 (preserve the S585b ship value, all
                                    // canary-validated) OR (tblW=auto AND over < 11 = up to ~2×cellMar,
                                    // the auto-fit full-page envelope). The old <5.0 MISSED T15/T57/
                                    // T101 (over +9.9, tblW=auto) → the 賃金 chapter stayed 1pg short.
                                    // ★NOTE the body↔cell COUPLING (--pagedelta): clamping cells alone
                                    // over-fills (the ×0.6667-short body compensates over-wide cells);
                                    // tokyoshugyo PASS needs body(S590)+cells+S586 jointly. This fixes
                                    // the CELL piece correctly (canary-clean by tblW=dxa exclusion).
                                    // OXI_S585_OVER tunes the auto bound (default 11).
                                    let s585_auto_over: f32 = std::env::var("OXI_S585_OVER")
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(11.0);
                                    let s585_tblw_auto =
                                        table.style.width_type.as_deref() == Some("auto");
                                    let s585_over = cell_w - content_width;
                                    let s585_cellmar = std::env::var("OXI_S585_DISABLE").is_err()
                                        && row.cells.len() == 1
                                        && !table.style.has_explicit_cellmar
                                        && !matches!(
                                            para.alignment,
                                            Alignment::Right | Alignment::Center
                                        )
                                        && cell_w > content_width
                                        && (s585_over < 5.0
                                            || (s585_tblw_auto && s585_over < s585_auto_over));
                                    // S594 (2026-06-17, opt-IN OXI_S594=1): narrow the S585b cell
                                    // wrap_base by ONE EXTRA cellMar. S585c: Oxi's cell right border is
                                    // at content+2×cellMar vs Word's +1× → Oxi's cell_w over-computes by
                                    // ~1 cellMar, so wrap_base=cell_w−2×pad=417.95 is still ~5pt wider
                                    // than Word's true content ≈413 (=cell_w−3×cellMar). The 66 residual
                                    // CELL over-fit roots (S7m) are this width over. Subtract one more
                                    // cellMar to reach Word's content (single-cell S585b tables only;
                                    // 3a4f/model 第N条 are in BODY → unaffected).
                                    let s594_extra = if let Ok(k) = std::env::var("OXI_S594_K") {
                                        // tunable extra narrowing (pt) for the cell wrap_base sweep
                                        k.parse::<f32>().unwrap_or(0.0)
                                    } else if std::env::var("OXI_S594").ok().as_deref() == Some("1")
                                    {
                                        pad_l
                                    } else {
                                        0.0
                                    };
                                    // S585N (SCAFFOLD, OXI_S585N=1, default OFF=byte-identical): a
                                    // NESTED-table cell (is_nested) with an explicit tcW subtracts its OWN
                                    // cellMar from wrap_base. The 賃金 参考 box's innermost cell (tcW=423.55)
                                    // OVERFLOWS its nested container (≈416.85), so the s585b/s531/s559 gates
                                    // (all require cell_w≤content_width or autofit-squeeze) miss it → Oxi
                                    // wraps at full cell_w, ~10.6pt wider than Word (content x84.55 vs x90.7).
                                    // S585N → wrap_base 413.65 ≈ Word 412.9 (CELLX-verified). ★HELD: this is
                                    // a RENDER-correct fix but PAGINATION-NEUTRAL on tokyoshugyo (gate
                                    // byte-identical with/without, S586 on or off) — the 参考-box ~1-char/line
                                    // over-fit is render-real but NON-OPERATIVE (the S562/r7-cell pattern).
                                    // Shipping S585N default-ON needs corpus SSIM/IoU A/B over nested-table
                                    // docs (deferred). Gated to is_nested (3a4f p19 is top-level, untouched).
                                    // ★OPERATIVE #2 cause (CORRECTED 2026-06-22 s2 — NOT a row-split): a
                                    // DISTRIBUTED BODY sub-pt over-fit in the 賃金 chapter. Under S586 the −1
                                    // onsets at Oxi p47, but p47's content tracks Word at a CONSTANT −39pt
                                    // offset (no internal drift, landmark-Y verified) ⇒ the ~2 over-fit lines
                                    // are inherited from p46/earlier, where the over-fit lines are BODY 解説
                                    // paragraphs (no CELLX). Components: (a) body page-bottom leniency (Oxi
                                    // fits «給月給»/«し引く» at p46 bottom, last-line bottom ~761 > content
                                    // 756.85, the S603/S576 typed-grid mid-para leniency wall); (b) per-line
                                    // body wrapping; (c) small ~1.3pt/解説-box spacing (Word 37.3 vs Oxi 36.0
                                    // ×~20 boxes ≈ 26pt). No single dominant lever. Per-line localization is
                                    // LIMITED by the GDI dump being per-RUN for body (not per-glyph) — needs
                                    // glyph-level Oxi instrumentation. The fix is a chapter-wide vertical-
                                    // fidelity pass + #1(S586), convergent with the sub-pt spacing/leniency
                                    // walls. See [[tokyoshugyo_wrap_not_cellheight]].
                                    let s585n_nested = std::env::var("OXI_S585N").ok().as_deref()
                                        == Some("1")
                                        && is_nested
                                        && cell.width.is_some()
                                        && !matches!(
                                            para.alignment,
                                            Alignment::Right | Alignment::Center
                                        );
                                    // S713 render mirror (see the estimate-side s713_cellmar comment):
                                    // legacy (compat<=14) single-cell explicit-tblCellMar row wraps
                                    // at cell_w - pads. Word render-truth tokyoshugyo (注) cell.
                                    let s713_cellmar_render = std::env::var("OXI_S713_DISABLE")
                                        .is_err()
                                        && row.cells.len() == 1
                                        && table.style.has_explicit_cellmar
                                        && self.compat_mode <= 14;
                                    // S767 = the S585c 本体 wrap-consistency (2026-07-08, default ON,
                                    // opt-out OXI_S767_DISABLE). S585c clamps the BORDER of an over-wide
                                    // AUTO cell to the page (eff_cell_w) but the WRAP kept subtracting pads
                                    // from the PRE-CLAMP cell_w — so a clamped cell that is NOT the
                                    // justified-compress narrow case (which already uses eff_cell_w above)
                                    // wrapped ~1 cellMar too wide. e3c545's LEFT-aligned Latin RDF code
                                    // blocks (compressPunctuation off → s585c_narrow=false): border clamped
                                    // to 544.0 but wrap = cell_w-pads = 469.30 vs Word's eff_cell_w-pads =
                                    // 463.90 → Oxi packed ~1 char/line more, mis-breaking the long @prefix
                                    // lines. Use eff_cell_w (== cell_w when NOT clamped, so non-clamped
                                    // cells are byte-identical) in the pad-subtracting branches; the
                                    // `else { cell_w }` text-into-margins case (S562) is left untouched.
                                    // This completes S585c's stated "one eff_cell_w → border AND wrap"
                                    // for the non-compress path. See [[tokumei_form_family_ssim]].
                                    let wrap_cell_w = if std::env::var("OXI_S767_DISABLE").is_ok() {
                                        cell_w
                                    } else {
                                        eff_cell_w
                                    };
                                    // S1121 (2026-08-14): thread WHETHER this base already
                                    // subtracted the cell padding, so the S493J alignment
                                    // adjust below cannot subtract it a SECOND time. S493J's
                                    // own condition list (3 flags) predates S768/S531/S559/
                                    // S585n/CELLPAIR/S713 — on a pure-Latin document S768
                                    // makes EVERY cell take the pad-subtracted base, so the
                                    // alignment area lost pad_l+pad_r twice and a right-
                                    // aligned number cell (hmrc's 1-8 boxes: avail 2.40 <
                                    // line 3.55) collapsed to LEFT-aligned.
                                    let mut s1121_pad_in_base = true;
                                    let wrap_base = if self.celllaw() {
                                        // S1173 render side: the measured law, in
                                        // front of the allowlist it supersedes.
                                        LayoutEngine::celllaw_twips(
                                            (wrap_cell_w
                                                - self.celllaw_inset(table, cell, pad_l, false)
                                                - self.celllaw_inset(table, cell, pad_r, true))
                                            .max(0.0),
                                        )
                                    } else if s585c_narrow
                                        && matches!(
                                            para.alignment,
                                            Alignment::Justify | Alignment::Distribute
                                        ) {
                                        // S585c: the clamped box wraps within its Word content area
                                        // (eff_cell_w - 2×cellMar). Supersedes s585_cellmar/s594 (which
                                        // subtracted from the UNCLAMPED cell_w, leaving the wrap ~1 cellMar
                                        // wider than the clamped border). Derived from the SAME eff_cell_w
                                        // the border/shading use → border-x and wrap are consistent. Paired
                                        // with legacy_cell_break + cell_bura (below) so the narrower wrap
                                        // fits Word's line counts via cell 約物 compression + ぶら下げ.
                                        (eff_cell_w - pad_l - pad_r).max(0.0)
                                    } else if s585_cellmar {
                                        (wrap_cell_w - pad_l - pad_r - s594_extra).max(0.0)
                                    } else if cell_hang_inner
                                        || s301_layout_fixed
                                        || s412_cellmar_subtract
                                        || s531_singlecell_cellmar
                                        || s559_cellmar
                                        || s585n_nested
                                        || self.cellpair_active()
                                        || s713_cellmar_render
                                        || (std::env::var("OXI_S768_DISABLE").is_err()
                                            && !self.doc_body_has_real_cjk)
                                    {
                                        // S768 (opt-in): pure-Latin document → wrap at cell_w − cellMar
                                        // (Word's content area). See the estimate-side s768_latin_wrap comment.
                                        (wrap_cell_w - pad_l - pad_r).max(0.0)
                                    } else if let Some(f) = std::env::var("OXI_CELLPAD")
                                        .ok()
                                        .map(|v| v.parse::<f32>().unwrap_or(1.0))
                                    {
                                        // OXI_CELLPAD (2026-08-18, opt-IN, default OFF =
                                        // byte-identical): subtract the padding for EVERY cell
                                        // rather than for the eight cases above. kojin's page-17
                                        // table is the case the allowlist misses: a CJK jc=left
                                        // cell, where `wrap_base = cell_w` leaves the budget
                                        // pad_l+pad_r = 10.8pt too wide and Oxi draws «雇用保険被保»
                                        // out to 210.25 past the cell's own right rule at 209.05,
                                        // where Word wraps after five characters and stretches
                                        // them to fill (measured 2026-08-18: Word's column
                                        // positions match Oxi's within 0.1pt, so the geometry is
                                        // right and only the budget is wrong).
                                        // ★Held opt-in because the same narrowing is what the
                                        // S559/S585c balance documents as over-correcting
                                        // tokyoshugyo by +1x645 -- at Word's wrap width Oxi
                                        // produces MORE lines, since it does not compress 約物
                                        // inside cells the way Word does. Whether kojin behaves
                                        // like tokyoshugyo is a measurement, not a deduction.
                                        // ★The value is the FRACTION of the padding to take, so
                                        // the sweep can ask whether a partial subtraction keeps
                                        // kojin's gain without costing 04b88e/34140b the line
                                        // that pushes a paragraph onto the next page.
                                        (wrap_cell_w - (pad_l + pad_r) * f).max(0.0)
                                    } else {
                                        s1121_pad_in_base = false;
                                        cell_w
                                    };
                                    let mut wrap_w =
                                        (wrap_base - p_indent_left - p_indent_right).max(0.0);
                                    let mut first_line_wrap_w = if p_first_line_indent < 0.0 {
                                        (wrap_base
                                            - (p_indent_left + p_first_line_indent).max(0.0)
                                            - p_indent_right)
                                            .max(0.0)
                                    } else {
                                        (wrap_w - p_first_line_indent).max(0.0)
                                    };
                                    // LEGACYCELL wrap narrow (2026-06-23, OXI_LEGACYCELL=1): a legacy (compat≤14)
                                    // compressPunctuation justified single-cell box has its wrap ~1 cellMar TOO WIDE
                                    // (S585c: Oxi wrap_base = cell_w − 2×pad vs Word's content = cell_w − 3×cellMar).
                                    // Subtract 1 more cellMar so the wrap = Word's content-right, enabling the
                                    // derived small-cap break (the 約物 then overflows enough to wrap a kanji like
                                    // Word). Scoped to LEGACY (3a4f/model = compat15, EXCLUDED). See [[char_budget_wall]].
                                    if std::env::var("OXI_LEGACYCELL").ok().as_deref() == Some("1")
                                        && self.compat_mode < 15
                                        && self.compress_punctuation
                                        && row.cells.len() == 1
                                        && matches!(
                                            para.alignment,
                                            Alignment::Justify | Alignment::Distribute
                                        )
                                    {
                                        // cap the wrap at the page text-margin (= Word's content-right for a
                                        // full-page-width 条文 box); only narrows (never widens), so a cell
                                        // already inside the margin is untouched.
                                        let pg_right = start_x + content_width;
                                        let cont_left = cell_x + pad_l + p_indent_left;
                                        wrap_w = wrap_w.min((pg_right - cont_left).max(0.0));
                                        let fl_left = cell_x
                                            + pad_l
                                            + (p_indent_left + p_first_line_indent).max(0.0);
                                        first_line_wrap_w =
                                            first_line_wrap_w.min((pg_right - fl_left).max(0.0));
                                    }
                                    // PROPCELL WRAP CAP (tokyoshugyo/d77a commentary boxes, 2026-06-23, default
                                    // ON, opt-out OXI_PROPCELL_DISABLE): a jc=LEFT single-cell
                                    // box in a PROPORTIONAL CJK font (MS PMincho/PGothic — the 参考/ガイドライン
                                    // 抜粋 commentary boxes) has its wrap ~1 cellMar TOO WIDE (S585c: Oxi cell
                                    // right border = content+2×cellMar vs Word +1×). Oxi wraps at x515.5 vs Word
                                    // x510 → fits ~1 char/line more → content shifts up → page-bottom over-fit →
                                    // the discrete −1 page flips (page-20 趣旨 box: «…ことか|ら、» Oxi 2 / Word 3
                                    // lines). CAP the wrap at the page text-margin (start_x+content_width = Word's
                                    // content-right ≈ x510), paired with the bounded oikomi below (the trailing
                                    // 約物 then overflows enough to oikomi, matching Word). UNLIKE the S585c
                                    // 条文-box PGCAP (which over-corrects because Oxi doesn't compress 約物 in
                                    // JUSTIFIED cells), the jc=LEFT commentary boxes break at NATURAL proportional
                                    // width + oikomi the line-end 約物 (no 約物 compression involved). Scoped to
                                    // proportional + jc=left + single-cell. See [[tokyoshugyo_wrap_not_cellheight]].
                                    let propcell_wrap_cap =
                                        std::env::var("OXI_PROPCELL_DISABLE").is_err()
                                            && row.cells.len() == 1
                                            && !matches!(
                                                para.alignment,
                                                Alignment::Right | Alignment::Center
                                            )
                                            && para.runs.iter().any(|r| {
                                                self.resolve_font_family_for_text(
                                                    &r.text,
                                                    &r.style,
                                                    &para.style,
                                                )
                                                .map_or(false, |f| {
                                                    f.contains("Ｐ明朝")
                                                        || f.contains("Ｐゴシック")
                                                        || f.contains("PMincho")
                                                        || f.contains("PGothic")
                                                })
                                            });
                                    if propcell_wrap_cap {
                                        let pg_right = start_x + content_width;
                                        let cont_left = cell_x + pad_l + p_indent_left;
                                        wrap_w = wrap_w.min((pg_right - cont_left).max(0.0));
                                        let fl_left = cell_x
                                            + pad_l
                                            + (p_indent_left + p_first_line_indent).max(0.0);
                                        first_line_wrap_w =
                                            first_line_wrap_w.min((pg_right - fl_left).max(0.0));
                                    }
                                    // OXI_PGCAP (tokyoshugyo #2c, half of the coupled fix): cap the cell
                                    // content wrap at the PAGE TEXT-MARGIN right (start_x+content_width), not
                                    // the over-wide cell border (S585c +1-cellMar over). Must be paired with
                                    // OXI_CELLCOMP (cell 約物 compression) — alone it over-corrects (exposes
                                    // the missing compression). See [[tokyoshugyo_wrap_not_cellheight]].
                                    // S585c-FIX: gate PGCAP to FULL-WIDTH SINGLE-cell boxes (the 条文/解説 cells).
                                    // BUG found 2026-06-22: PGCAP fired on NARROW multi-column cells positioned
                                    // near the right edge (their content_right > margin because they're far-right,
                                    // even though narrow) → spuriously wrapped them (e.g. schedule «41:33»→«41:3»+«3»).
                                    // row.cells.len()==1 excludes the multi-column schedule/calc cells.
                                    if std::env::var("OXI_PGCAP").ok().as_deref() == Some("1")
                                        && row.cells.len() == 1
                                    {
                                        let pg_off: f32 = std::env::var("OXI_PGCAP_OFF")
                                            .ok()
                                            .and_then(|v| v.parse().ok())
                                            .unwrap_or(0.0);
                                        let pg_right = start_x + content_width + pg_off;
                                        let cont_left = cell_x + pad_l + p_indent_left;
                                        wrap_w = wrap_w.min((pg_right - cont_left).max(0.0));
                                        let fl_left = cell_x
                                            + pad_l
                                            + (p_indent_left + p_first_line_indent).max(0.0);
                                        first_line_wrap_w =
                                            first_line_wrap_w.min((pg_right - fl_left).max(0.0));
                                    }
                                    // ★tokyoshugyo #2c COUPLING PROVEN (2026-06-22, OXI_PGCAP experiment,
                                    // reverted): capping the cell wrap at the page text-margin (start_x+
                                    // content_width = x510, = Word's 第３２条 fill) narrows 条文 cells to Word's
                                    // wrap-right BUT over-corrects massively (oxi 91>90, +1×645) — at the SAME
                                    // wrap-width as Word, Oxi produces MORE lines (packs FEWER chars/line)
                                    // because Oxi does NOT compress 約物 in cells where Word does. ⇒ #2c is
                                    // wrap-width ⊕ cell-約物-compression, MUTUALLY COMPENSATING (the 90pg
                                    // baseline is the S559/S585c balance: over-wide cells compensate the
                                    // missing compression). The fix needs BOTH together — narrow the wrap
                                    // (PGCAP/S594) AND add Word's cell 約物 compression (the 6de22246 char-
                                    // budget cell-wrapper wall). See [[tokyoshugyo_wrap_not_cellheight]].
                                    if std::env::var("OXI_DUMP_CELLX").is_ok() {
                                        let preview: String = para
                                            .runs
                                            .iter()
                                            .flat_map(|r| r.text.chars())
                                            .take(8)
                                            .collect();
                                        eprintln!(
                                "[CELLX] cell_x={:.2} cell_w={:.2} pad_l={:.2} pad_r={:.2} \
                                 ind_l={:.2} ind_r={:.2} first_ind={:.2} wrap_base={:.2} \
                                 wrap_w={:.2} first_line_wrap_w={:.2} hang_inner={} s301={} text={:?}",
                                cell_x, cell_w, pad_l, pad_r, p_indent_left, p_indent_right,
                                p_first_line_indent, wrap_base, wrap_w, first_line_wrap_w,
                                cell_hang_inner, s301_layout_fixed, preview
                            );
                                    }

                                    // 2026-04-19: Render list marker (numPr) for cells too.
                                    // Body renders at mod.rs:1939; cells previously skipped it.
                                    // b35 p1 "事務処理体制を整備" row: numId=5 ilvl=0 → □ marker.
                                    let list_marker_info: Option<(String, f32, f32)> =
                                        para.style.list_marker.as_ref().map(|marker| {
                                            let marker_style = s1037_marker_style(para)
                                                .cloned()
                                                .unwrap_or_else(|| {
                                                    para.runs
                                                        .first()
                                                        .map(|r| r.style.clone())
                                                        .unwrap_or_default()
                                                });
                                            let marker_fs =
                                                self.resolve_font_size(&marker_style, &para.style);
                                            let marker_metrics =
                                                &*self.metrics_for(&marker_style, &para.style);
                                            // S692 (2026-06-29, SHIPPED default ON, opt-out OXI_MARKERCJK_DISABLE): a CELL numbered-list
                                            // label that is CJK/full-width (「第３４条」, the tokyoshugyo 賃金
                                            // regulation article markers) had its WIDTH computed with metrics_for
                                            // (the ASCII Century font) → the full-width digits 「３４」 fell to the
                                            // proportional ~7.5pt (marker w=36, vs the GDI RENDER's 10.5/char = 42
                                            // = Word). The under-counted marker reserve (s592_marker_reserve = w +
                                            // fs*0.25) starts the body ~6pt too LEFT → the body line over-fits
                                            // (fits 「当」 where Word wraps). Use the render width (font_size) for
                                            // full-width marker chars. See [[tokyoshugyo_wrap_not_cellheight]].
                                            let marker_cjk =
                                                std::env::var("OXI_MARKERCJK_DISABLE").is_err();
                                            let marker_width: f32 = marker
                                                .chars()
                                                .map(|c| {
                                                    if marker_cjk && crate::font::is_fullwidth(c) {
                                                        marker_fs
                                                    } else {
                                                        self.registry.char_width_pt_with_fallback(
                                                            c,
                                                            marker_fs,
                                                            marker_metrics,
                                                        )
                                                    }
                                                })
                                                .sum();
                                            (marker.clone(), marker_fs, marker_width)
                                        });
                                    // S592→S641 (2026-06-17 found, flipped DEFAULT ON 2026-06-22, opt-out
                                    // OXI_S641_DISABLE): a SPACE-suffix numbered CELL paragraph (the 賃金
                                    // regulation 第N条 boxes) renders its number INLINE at the cell-left
                                    // indent (NOT outdented), body flowing AFTER it. Word does not outdent a
                                    // space-suffix number (no tab). Oxi outdented the marker (marker_x
                                    // −list_indent) AND did not reserve marker width on line-1 → body over-fit
                                    // + overlapped the marker. FIX: reserve marker width on line-1 (wrap),
                                    // place the marker at cell-left, start the body after it. Cross-doc PDF:
                                    // tokyoshugyo 第４条 x97.6, 3a4f 第２条 x95.7 = cell-left (validated on BOTH).
                                    // GATE: tokyoshugyo 0.9638→0.9746 (the 第３２条 cell 1→2 lines = Word), full
                                    // Phase-1 81/84 0 PASS→FAIL (only tokyoshugyo's pagination changes; 3a4f's
                                    // marker render shifts to cell-left = Word but pagination byte-identical).
                                    // See [[tokyoshugyo_wrap_not_cellheight]].
                                    let s641_cell_space = std::env::var("OXI_S641_DISABLE")
                                        .is_err()
                                        && matches!(para.style.list_suff.as_deref(), Some("space"))
                                        && para.style.list_indent.unwrap_or(0.0) > 0.5;
                                    let s592_cell_space = s641_cell_space;
                                    let s592_marker_reserve = if s592_cell_space {
                                        list_marker_info
                                            .as_ref()
                                            .map(|(_, fs, w)| w + fs * 0.25)
                                            .unwrap_or(0.0)
                                    } else {
                                        0.0
                                    };
                                    if s592_cell_space {
                                        first_line_wrap_w =
                                            (first_line_wrap_w - s592_marker_reserve).max(0.0);
                                    }
                                    // S718: TAB-suffix list body starts at the next defaultTabStop
                                    // when it precedes ind_l (see s718_list_tab_pull). Widen line 0's
                                    // budget by the pull; the emit shifts line 0's text left to match.
                                    let s718_pull = if !s592_cell_space {
                                        list_marker_info
                                            .as_ref()
                                            .map(|(_, _, w)| {
                                                self.s718_list_tab_pull(
                                                    p_indent_left,
                                                    para.style.list_indent.unwrap_or(18.0),
                                                    *w,
                                                    (-p_first_line_indent_raw).max(0.0),
                                                    self.s1590_exact_marker_w(para),
                                                    &para.style.tab_stops,
                                                )
                                            })
                                            .unwrap_or(0.0)
                                    } else {
                                        0.0
                                    };
                                    first_line_wrap_w += s718_pull;
                                    let cell_float_wrap = if cell_float_flow {
                                        let obstacles = if let Some(positions) = fixed_float_positions {
                                            LayoutEngine::cell_float_obstacles_at(cell, positions)
                                        } else {
                                            LayoutEngine::cell_float_obstacles(cell, &float_tops, block_pos, wrap_base)
                                        };
                                        Some(self.measure_cell_float_para(para, wrap_base, row_line_pitch,
                                            table.style.para_style.as_ref(), grid_char_pitch, grid_char_cw_ratio,
                                            true, true, content_h, &obstacles).1)
                                    } else { None };

                                    // Collect runs into lines with greedy wrapping
                                    // Tuple: (text, font_size, width, bold, italic, underline, underline_style, strikethrough, font_family, color, highlight, character_spacing, text_scale)
                                    // S993 (R2/R3, 2026-07-23): the trailing `bool` is
                                    // `lrpb_before` — this fragment's originating run carried a
                                    // mid-run `<w:lastRenderedPageBreak/>`. Consumed only when
                                    // s993_exact (fixed 3-cell auto Latin row); replaces the
                                    // R7.73 next-paragraph approximation for those rows.
                                    let mut lines: Vec<
                                        Vec<(
                                            String,
                                            f32,
                                            f32,
                                            bool,
                                            bool,
                                            bool,
                                            Option<String>,
                                            bool,
                                            Option<String>,
                                            Option<String>,
                                            Option<String>,
                                            f32,
                                            f32,
                                            bool,
                                            Option<String>,
                                            bool, // S1312: fragment's run carries ruby
                                            RunStyle, // Original run settings for cell glyph advances
                                        )>,
                                    > = Vec::new();
                                    // S1299 (2026-09-04): the LAST field is the run's own
                                    // `w:eastAsia` family. Without it the cell path rebuilt a
                                    // RunStyle from this tuple carrying only size/bold/italic/
                                    // ascii-family, so `metrics_for_text` resolved the East
                                    // Asian face from the PARAGRAPH instead — i.e. through
                                    // docDefaults' `eastAsiaTheme` to the theme — even for a run
                                    // that names ＭＳ 明朝 outright. b5f706e9 is the witness: its
                                    // cells say `w:eastAsia="ＭＳ 明朝"` on every run, yet 391 of
                                    // its glyphs moved 1.5pt when the theme's EA font changed.
                                    let mut current_line: Vec<(
                                        String,
                                        f32,
                                        f32,
                                        bool,
                                        bool,
                                        bool,
                                        Option<String>,
                                        bool,
                                        Option<String>,
                                        Option<String>,
                                        Option<String>,
                                        f32,
                                        f32,
                                        bool,
                                        Option<String>,
                                        bool, // S1312: fragment's run carries ruby
                                        RunStyle, // Original run settings for cell glyph advances
                                    )> = Vec::new();
                                    // Task P (2026-07-22, default ON, opt-out OXI_S982_DISABLE): a cell-inline OLE object
                                    // (Equation.DSMT4, step 3 routed it to run.style.inline_object_*)
                                    // is carried as a paragraph-local `&Image` registry + a U+F8FE
                                    // sentinel fragment (the S703c F8FF pattern; the 13-element cell
                                    // tuple carries no style fields, so the image ref cannot ride the
                                    // tuple). Default (s982_cell=false) never pushes — byte-identical.
                                    let s982_cell = std::env::var("OXI_S982_DISABLE").is_err();
                                    let mut cell_inline_objects: Vec<&Image> = Vec::new();
                                    // S1252: the same carrier for a STRUCTURED inline oMath, with
                                    // its own U+F8FD sentinel. `educational__002a301d` puts all six
                                    // of its `n.  f(x)=…` paragraphs in TABLE CELLS, and a cell
                                    // fragment the emit does not know DRAWS NOTHING — the probe's
                                    // `cellmath` arm went from 12 text elements to 3 when the
                                    // parser routed the maths without this.
                                    let mut cell_inline_math: Vec<&crate::ir::MathBlock> =
                                        Vec::new();
                                    let mut line_x: f32 = 0.0;
                                    // Session 118 jc=both refactor — gated by OXI_JCBOTH_REFACTOR env var.
                                    // When enabled, calls compute_compression from jc_both_compress module
                                    // for wrap-decision lookahead. Track per-char CharContext alongside
                                    // string buf so we can pass to compute_compression.
                                    // S166 (2026-05-21): default ON. Baseline: mean IoU 0.9301 → 0.9303,
                                    // 15076df 0.8799 → 0.8850 (+0.005), no other doc moved.
                                    // S238 (2026-05-23): removed OXI_LEGACY_NO_JCBOTH_REFACTOR
                                    // legacy env-var fallback during hardening pass.
                                    let jc_gate_active =
                                        true && matches!(
                                            para.alignment,
                                            Alignment::Justify | Alignment::Distribute
                                        ) && self.balance_single_byte_double_byte_width
                                            && self.compress_punctuation;
                                    // S1082: the S825 compat-15 justified space-shrink capacity,
                                    // ported to this cell breaker (see the effective_wrap site).
                                    let s1082_cell_shrink = !self.doc_body_has_real_cjk
                                        && self.compat_mode >= 15
                                        && self.compat_mode_explicit
                                        && matches!(
                                            para.alignment,
                                            Alignment::Justify | Alignment::Distribute
                                        )
                                        && std::env::var("OXI_S1082_DISABLE").is_err()
                                        && std::env::var("OXI_S825_DISABLE").is_err();
                                    // S1174 (opt-in): the derived legacy cell 約物
                                    // compression. Both settings must be declared --
                                    // compat 15 turns the mechanism off entirely.
                                    // ★The gate is the ALIGNMENT, with compatibility as the
                                    // second way in -- not compatibility alone. Swept both:
                                    //   compat 11, jc=left  -> compresses
                                    //   compat 15, jc=left  -> does not
                                    //   compat 15, jc=both  -> compresses, identically to 11
                                    // Reading the first two as "legacy only" missed the third,
                                    // and the third is a47e: compatibilityMode 15 with justified
                                    // cells, where Word squeezes 「．」 from 9.599 to 6.125 to
                                    // hold a line 1.02pt over its budget and Oxi wraps instead.
                                    let s1174_yakucomp =
                                        std::env::var("OXI_YAKUCOMP_DISABLE").is_err()
                                            && self.compress_punctuation
                                            && (self.compat_mode <= 14
                                                || matches!(
                                                    para.alignment,
                                                    Alignment::Justify | Alignment::Distribute
                                                ));
                                    // S1176 (opt-in): the CJK space-shrink. Justified only --
                                    // jc=left gave nothing at all across the sweep.
                                    let s1176_space =
                                        std::env::var("OXI_CJKSPACE").ok().as_deref() == Some("1")
                                            && matches!(
                                                para.alignment,
                                                Alignment::Justify | Alignment::Distribute
                                            );
                                    // S497b FALSIFIED (2026-06-05): extending the compute_compression wrap
                                    // lookahead to left-aligned compressPunctuation cells (to model Word's
                                    // end-of-line yakumono oikomi at wrap for non-justified paras) was a NO-OP
                                    // on the whole tokumei family + ed025c/3a4f (dwrite ΔTOTAL +0.0000). The
                                    // tokumei cells do NOT wrap early at yakumono boundaries — S497 (the
                                    // line-start-prohibited hang) already covered the one real case (15076df).
                                    // The remaining tokumei gap is cumulative sub-pixel precision / weight, not
                                    // fixable wrap-precision. Reverted; not gated.
                                    let mut current_line_chars: Vec<
                                        crate::layout::jc_both_compress::CharContext,
                                    > = Vec::new();
                                    let mut is_first_line = true;
                                    // S1169 (2026-08-19, default ON, opt-out
                                    // OXI_S1169_DISABLE): a paragraph that ENDS
                                    // with a hard break keeps the empty line that
                                    // break opened. The break handler below pushes
                                    // the finished line and leaves `current_line`
                                    // empty; the final flush then skips it because
                                    // it is empty, so the line silently vanishes --
                                    // in CELLS only, since the body path builds its
                                    // lines elsewhere and already agrees with Word.
                                    // This cell renderer also disagreed with its OWN
                                    // line-count estimate, which counts the break
                                    // (`lines += 1`, the Session-109 mirror) -- the
                                    // estimate==render invariant that path exists to
                                    // keep. Probe `_pb_trailbr_gen.py`: Word puts TWO
                                    // pitches between a paragraph ending in <w:br/>
                                    // and the next (27.60 vs 13.80 plain, 41.40 for
                                    // two breaks), in a cell exactly as in the body;
                                    // Oxi matched the body arms and lost the cell one.
                                    // Cost: educational__00161422's p12 fits 29 lines
                                    // where Word fits 26, moving the row boundary a
                                    // line late (S927 fixed this same class for
                                    // <w:cr>: "dropping it collapsed a genuine
                                    // trailing empty line").
                                    let mut s1169_trailing_break = false;
                                    let mut explicit_break_lines = std::collections::HashSet::new();
                                    // R7.51 (2026-05-13): autoSpaceDE state for CJK↔Latin transitions.
                                    // Tracks the last emitted character across runs/buffers so we can
                                    // detect transitions and add Word's 2.5pt (10.5pt font) gap. The
                                    // body renderer (break_into_lines) applies this; this cell-renderer
                                    // path historically did not, causing d77a58 w_i=47 wrap mismatch
                                    // (5 lines Oxi vs 6 lines Word).
                                    let mut cell_edge_tab_nowrap = false;
                                    let mut cell_tab_char_pack = false;
                                    let mut prev_char_emitted: Option<char> = None;
        let mut prev_char_gap: Option<f32> = None;
                                    let mut prev_char_ruby = false; // S1316: the emitted char came from a ruby field
                                    // S443: the widened oikomi must fire ONLY on tab-bearing
                                    // (list-marker) paragraphs. 3a4f has hanging-indent paras
                                    // but ZERO hanging+tab paras; gating oikomi on hanging
                                    // alone made it fire on 3a4f's tab-less hanging cells and
                                    // cratered it (909 paras +1). Restricting to para_has_tab
                                    // excludes 3a4f entirely while keeping d77a's カ/タ list items.
                                    let para_has_tab =
                                        para.runs.iter().any(|r| r.text.contains('\t'));
                                    // S586 (2026-06-16, opt-IN OXI_S586=1, default OFF = byte-identical;
                                    // SCAFFOLD held pending the coupled #2 fix): page-44 約物 OIKOMI. A LEGACY
                                    // (compat<15) compressPunctuation CELL line whose trailing char would
                                    // orphan (<=2 chars to para end) by a SMALL overflow (<= OXI_S586_CAP,
                                    // default 3.5pt) is pulled up by collapsing a 約物 immediately before an
                                    // OPENING bracket (the 、「 inter-space fully vanishes, -7.5pt). DISCRIMINATOR
                                    // derived from the 4-firing dataset + Word PDF (S585c): of 4 orphan+small
                                    // lines, ONLY page-44 «…については、「育児…» has 、 before an opener (Word
                                    // collapses=oikomi); the 3 with 、 before a KANJI are Word OIDASHI. Fires on
                                    // EXACTLY 1 corpus line and eliminates region-2 +1×282 (the SOLE region-2
                                    // root). ★HELD default-OFF: page-44 alone (OXI_S586=1) flips region-2
                                    // +283→−1×417, oxi 90→89pg (gate-verified 2026-06-22), exposing #2.
                                    // ★#2 ROOT (CORRECTED 2026-06-22, supersedes the "page-bottom over-fit /
                                    // NOT separately pinnable" framing): it is a CELL wrap_base OVER-WIDTH.
                                    // Oxi wraps the 賃金-chapter regulation / 参考 box cells at the full cell_w
                                    // where Word reserves cellMar. PINNED on the 配偶者手当 参考 box (depth-2
                                    // NESTED table, innermost cell explicit tcW=8471=423.55pt): Oxi places
                                    // content at x84.55 / wrap_base=423.55 (CELLX), Word at x90.7 / wrap≈412.9
                                    // (PDF, 3 nested borders 79.6/85.3/91.9) → Oxi over-fits ~1 char/line
                                    // («要因とな» vs Word «要因と») → packs the chapter 1pg tight. It falls
                                    // through EVERY cellMar-subtract gate: s531 needs default_cell_margins
                                    // .is_some() (this is None=4.95 fallback); s559 needs autofit-squeezed
                                    // (tcW−cell_w≥8; here tcW==cell_w) + justified; S585b needs cell_w > PAGE
                                    // content (here cell_w<page content); OXI_S559_CELLMAR=all ALSO misses it
                                    // (the nested cell_w 423.55 > its nested container ≈416.85 → cell_w≤
                                    // content_width fails). DISTRIBUTED (region-1 −17 baseline + region-2
                                    // masked); the page-44 over-wrap COMPENSATES the cumulative over-width →
                                    // 90=90 by accident. The blanket subtract-pads regressed Phase-1 53→49
                                    // (line 10874) and S559's broad form cascaded {1:1323} on 3a4f p19 → the
                                    // FIX is correct nested-table width/position resolution (the explicit tcW
                                    // exceeds its nested container; Oxi lays the cell at the outer-table x not
                                    // the nested-inset x), a focused canary-gated session. Ship #1+#2 together.
                                    // See [[tokyoshugyo_wrap_not_cellheight]], [[char_budget_wall]].
                                    let s586_para_chars: Vec<char> =
                                        para.runs.iter().flat_map(|r| r.text.chars()).collect();
                                    // S586 (2026-06-22, flipped DEFAULT ON, opt-out OXI_S586_DISABLE):
                                    // the page-44 約物→opening-bracket cell oikomi. Held opt-in for
                                    // 16+ sessions pending "#2" (the chapter under-count) — now found:
                                    // #2 = the dropped VML canvas figures (S640/VMLCANVAS), NOT the
                                    // char-budget wall. S586 (remove the compensating page-44 over-wrap)
                                    // + VMLCANVAS (add the dropped figure height) ship TOGETHER:
                                    // tokyoshugyo 0.8077→0.9638, full Phase-1 gate 0 PASS→FAIL (only
                                    // tokyoshugyo changes). The baseline 90=90 was a COMPENSATING balance
                                    // (page-44 +1 ⊕ figures −1). See [[tokyoshugyo_wrap_not_cellheight]].
                                    let s586_orphan = std::env::var("OXI_S586_DISABLE").is_err()
                                        && self.compat_mode < 15
                                        && self.compress_punctuation;
                                    let s586_cap: f32 = std::env::var("OXI_S586_CAP")
                                        .ok()
                                        .and_then(|v| v.parse().ok())
                                        .unwrap_or(3.5);
                                    let mut s586_run_offset = 0usize;

                                    // S993 (R2/R3): pending flag for the exact mid-run LRPB
                                    // fragment anchor. Set when entering a mid-run LRPB run
                                    // (p0/r0 excluded — that is a continuation/row-start marker,
                                    // R7.64), taken by the FIRST emitted fragment (mem::take).
                                    // Not reset on non-LRPB runs → an empty LRPB run carries to
                                    // the next visible run's first fragment.
                                    let mut s993_lrpb_pending = false;
                                    for (run_idx, run) in para.runs.iter().enumerate() {
                                        if run.has_last_rendered_page_break
                                            && !(cell_para_counter == 0 && run_idx == 0)
                                        {
                                            s993_lrpb_pending = true;
                                        }
                                        let font_size =
                                            self.resolve_font_size(&run.style, &para.style);
                                        // S1518 (2026-09-21, default ON, opt-out OXI_S1518_DISABLE):
                                        // a CELL sub/superscript run that states its own size is
                                        // still drawn at 0.65 of it -- the body path's S899, which
                                        // never reached cells. technical__01242a0a Table 3 header:
                                        // 'Q' 10pt + 'shoulder-' subscript sz=20 in a 38.7pt cell;
                                        // Word's PDF sets 'shoulder-' at 6.5pt (26.2pt wide) on the
                                        // Q's line, Oxi measured it at 10pt (37.7pt) and broke it
                                        // 'shoulde'/'r-', so the header row grew 2 -> 4 lines and
                                        // 25 paragraphs ran a page late.
                                        let font_size = if std::env::var_os("OXI_S1518_DISABLE").is_none()
                                            && run.style.font_size.is_some()
                                            && matches!(
                                                run.style.vertical_align,
                                                Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript)
                                            ) {
                                            LayoutEngine::vertical_align_font_size(font_size)
                                        } else {
                                            font_size
                                        };
                                        let bold = self.resolve_bold(&run.style, &para.style);
                                        let font_family = self
                                            .resolve_font_family_for_text(
                                                &run.text,
                                                &run.style,
                                                &para.style,
                                            )
                                            .map(|s| s.to_string());

                                        // Task P step 4 (2026-07-22, default ON, opt-out OXI_S982_DISABLE): a cell-inline
                                        // OLE object → a U+F8FE{index} atomic fragment (width = the
                                        // object extent). Registry holds the &Image for the emit
                                        // (step 5). Do NOT `continue` — a same-run trailing text
                                        // (target object 1) must still be processed. Gated on
                                        // s982_cell → 0 hits in default (byte-identical).
                                        // S1252: a structured inline oMath run -> a U+F8FD{index}
                                        // atomic fragment of the maths ADVANCE, drawn at step 5.
                                        if s982_cell {
                                            if let (Some((ow, _)), Some(mb)) = (
                                                run.style.inline_object_extent,
                                                run.style.inline_math.as_deref(),
                                            ) {
                                                let ew = cell_float_wrap.as_ref().map_or_else(|| if is_first_line {
                                                    first_line_wrap_w
                                                } else {
                                                    wrap_w
                                                }, |wrap| wrap.frame(lines.len(), wrap_w, first_line_wrap_w, p_indent_left, p_first_line_indent).width);
                                                if line_x + ow > ew && !current_line.is_empty() {
                                                    lines.push(std::mem::take(&mut current_line));
                                                    line_x = 0.0;
                                                    current_line_chars.clear();
                                                    is_first_line = false;
                                                }
                                                let index = cell_inline_math.len();
                                                cell_inline_math.push(mb);
                                                current_line.push((
                                                    format!("\u{F8FD}{index}"),
                                                    font_size,
                                                    ow,
                                                    bold,
                                                    run.style.italic,
                                                    false,
                                                    None,
                                                    false,
                                                    font_family.clone(),
                                                    None,
                                                    None,
                                                    0.0,
                                                    100.0,
                                                    std::mem::take(&mut s993_lrpb_pending),
                                                    run.style.font_family_east_asia.clone(),
                                                    run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                    run.style.clone(),
                                                ));
                                                line_x += ow;
                                                continue;
                                            }
                                        }
                                        if s982_cell {
                                            if let (Some((ow, _)), Some(img)) = (
                                                run.style.inline_object_extent,
                                                run.style.inline_object_image.as_deref(),
                                            ) {
                                                let ew = cell_float_wrap.as_ref().map_or_else(|| if is_first_line {
                                                    first_line_wrap_w
                                                } else {
                                                    wrap_w
                                                }, |wrap| wrap.frame(lines.len(), wrap_w, first_line_wrap_w, p_indent_left, p_first_line_indent).width);
                                                if line_x + ow > ew && !current_line.is_empty() {
                                                    lines.push(std::mem::take(&mut current_line));
                                                    line_x = 0.0;
                                                    current_line_chars.clear();
                                                    is_first_line = false;
                                                }
                                                let index = cell_inline_objects.len();
                                                cell_inline_objects.push(img);
                                                if std::env::var("OXI_DBG_CELLOLE").is_ok() {
                                                    eprintln!("[CELL-OLE] phase=builder enabled=1 index={} w={} h={}",
                                            index, ow, img.height);
                                                }
                                                current_line.push((
                                                    format!("\u{F8FE}{index}"),
                                                    font_size,
                                                    ow,
                                                    bold,
                                                    run.style.italic,
                                                    false,
                                                    None,
                                                    false,
                                                    font_family.clone(),
                                                    None,
                                                    None,
                                                    0.0,
                                                    100.0,
                                                    std::mem::take(&mut s993_lrpb_pending),
                                                    run.style.font_family_east_asia.clone(),
                                                    run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                    run.style.clone(),
                                                ));
                                                line_x += ow;
                                            }
                                        }

                                        // S703c (2026-06-30): a `combine` run (割注 / two-lines-in-one)
                                        // in a TABLE CELL renders as warichu. Push ONE compact tuple
                                        // whose text is SENTINEL-encoded ("\u{F8FF}{brackets}\u{F8FF}
                                        // {realtext}") — the cell emit decodes it and draws 2 small
                                        // rows + brackets (avoids a tuple-arity change). Gated on
                                        // combine → only kyotei's 1 run fires; other cells unchanged.
                                        if run.style.combine
                                            && !run.text.trim().is_empty()
                                            && std::env::var("OXI_S703_DISABLE").is_err()
                                        {
                                            let n = run.text.chars().count();
                                            let rows_chars = (n + 1) / 2;
                                            let brackets = run
                                                .style
                                                .combine_brackets
                                                .clone()
                                                .unwrap_or_else(|| "none".to_string());
                                            let has_br = brackets != "none";
                                            let cw = rows_chars as f32 * (font_size * 0.5)
                                                + if has_br { font_size * 0.8 } else { 0.0 };
                                            let ew = cell_float_wrap.as_ref().map_or_else(|| if is_first_line {
                                                first_line_wrap_w
                                            } else {
                                                wrap_w
                                            }, |wrap| wrap.frame(lines.len(), wrap_w, first_line_wrap_w, p_indent_left, p_first_line_indent).width);
                                            if line_x + cw > ew && !current_line.is_empty() {
                                                lines.push(std::mem::take(&mut current_line));
                                                line_x = 0.0;
                                                current_line_chars.clear();
                                                is_first_line = false;
                                            }
                                            let encoded =
                                                format!("\u{F8FF}{}\u{F8FF}{}", brackets, run.text);
                                            current_line.push((
                                                encoded,
                                                font_size,
                                                cw,
                                                bold,
                                                run.style.italic,
                                                run.style.underline,
                                                run.style.underline_style.clone(),
                                                run.style.strikethrough,
                                                font_family.clone(),
                                                run.style.color.clone(),
                                                run.style
                                                    .highlight
                                                    .clone()
                                                    .or_else(|| run.style.shading.clone()),
                                                0.0,
                                                100.0,
                                                std::mem::take(&mut s993_lrpb_pending),
                                                run.style.font_family_east_asia.clone(),
                                                run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                run.style.clone(),
                                            ));
                                            line_x += cw;
                                            continue;
                                        }

                                        // Split text character by character for wrapping
                                        let cs = if run.style.fit_text.is_some() || run.style.ruby_spread {
                                            run.style.character_spacing.unwrap_or(0.0)
                                        } else {
                                            snap_character_spacing(
                                                run.style.character_spacing.unwrap_or(0.0),
                                            )
                                        };
                                        let mut buf = String::new();
                                        let mut buf_w: f32 = 0.0;
                                        // S118: per-char context for jc_both_compress integration.
                                        let mut buf_chars: Vec<
                                            crate::layout::jc_both_compress::CharContext,
                                        > = Vec::new();
                                        // S1443: a page break inside a cell is inert (see the breaker note).
                                        let s586_run_chars: Vec<char> = run.text.chars().filter(|&c| !(c == '\x0C' && std::env::var_os("OXI_S1443_DISABLE").is_none())).collect();
                                        // Shaped cluster advances for complex-script text, the
                                        // same family choice as the body breaker (line_break.rs).
                                        let cell_deva_adv: Option<Vec<f32>> = if std::env::var_os("OXI_INDIA_CELL_DEVA_DISABLE").is_none()
                                            && s586_run_chars.iter().any(|&c| crate::font::is_complex_script(c))
                                        {
                                            let shaped: String = s586_run_chars.iter().collect();
                                            let emit_fam = self
                                                .resolve_font_family_for_text(&shaped, &run.style, &para.style)
                                                .map(|s| s.to_string());
                                            let emit_fam = if std::env::var_os("OXI_INDIA_CS_SHAPE_DISABLE").is_none() {
                                                run.style.font_family_cs.clone()
                                                    .or_else(|| para.style.default_run_style.as_ref().and_then(|s| s.font_family_cs.clone()))
                                                    .or(emit_fam)
                                            } else { emit_fam };
                                            let shape_fam = match emit_fam {
                                                Some(f) if crate::font::shape::family_covers(&f, '\u{0915}') => f,
                                                Some(f) if std::env::var_os("OXI_INDIA_CS_SHAPE_DISABLE").is_none()
                                                    && !self.registry.supports_family(&f)
                                                    && crate::font::runtime::resolve(&f, false, false).is_none()
                                                    && crate::font::shape::family_covers("Mangal", '\u{0915}') => "Mangal".to_string(),
                                                _ => "Nirmala UI".to_string(),
                                            };
                                            crate::font::shape::cluster_advances(&shape_fam, run.style.bold, run.style.italic, &shaped, font_size)
                                                .map(|(adv, _)| adv)
                                                .filter(|adv| adv.len() == s586_run_chars.len())
                                        } else {
                                            None
                                        };
                                        for (s586_ci, ch) in
                                            s586_run_chars.iter().copied().enumerate()
                                        {
                                            // Session 109 (2026-05-19): honour soft line breaks
                                            // (<w:br/>) and column/page break markers within table
                                            // cells. The OOXML parser converts <w:br/> to '\n' in
                                            // the run text (ooxml.rs:2569); the body renderer's
                                            // break_into_lines branch at line ~5109 picks them up,
                                            // but THIS cell-renderer ran them through the regular
                                            // char-width path, emitting a literal '\n' Text element
                                            // (~5pt wide) and keeping subsequent content on the
                                            // same visual line. LLA canary surfaced this as the
                                            // identical p.1 L14 mismatch across a1d6e4 / d4d126 /
                                            // de6e32 tokumei docs (S109).
                                            if matches!(ch, ' ' | '\t' | '\n' | '\x0B' | '\x0C') { cell_tab_char_pack = false; }
                                            let tab_break = ch == '\t' && p_first_line_indent < 0.0
                                                && (!current_line.is_empty() || !buf.is_empty())
                                                && self.cell_tab_word_overflows(para, s586_run_offset + s586_ci,
                                                    line_x + buf_w + if is_first_line { (p_indent_left + p_first_line_indent).max(0.0) } else { p_indent_left },
                                                    wrap_w + p_indent_left, p_indent_left, p_first_line_indent);
                                            if tab_break && std::env::var("OXI_CELL_TAB_CHAR_PACK_DISABLE").is_err() { cell_tab_char_pack = true; }
                                            if ch == '\n' || ch == '\x0B' || ch == '\x0C' || tab_break {
                                                if !buf.is_empty() {
                                                    current_line.push((
                                                        buf.clone(),
                                                        font_size,
                                                        buf_w,
                                                        bold,
                                                        run.style.italic,
                                                        run.style.underline,
                                                        run.style.underline_style.clone(),
                                                        run.style.strikethrough,
                                                        font_family.clone(),
                                                        run.style.color.clone(),
                                                        run.style
                                                            .highlight
                                                            .clone()
                                                            .or_else(|| run.style.shading.clone()),
                                                        cs,
                                                        run.style.text_scale.unwrap_or(100.0),
                                                        std::mem::take(&mut s993_lrpb_pending),
                                                        run.style.font_family_east_asia.clone(),
                                                        run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                        run.style.clone(),
                                                    ));
                                                    buf.clear();
                                                    buf_w = 0.0;
                                                    current_line_chars.extend(buf_chars.drain(..));
                                                }
                                                if current_line.is_empty() && !tab_break
                                                    && !self.doc_body_has_real_cjk
                                                    && std::env::var("OXI_CELL_BREAK_FONT_DISABLE").is_err()
                                                {
                                                    current_line.push((String::new(), font_size, 0.0, bold,
                                                        run.style.italic, run.style.underline,
                                                        run.style.underline_style.clone(), run.style.strikethrough,
                                                        font_family.clone(), run.style.color.clone(),
                                                        run.style.highlight.clone().or_else(|| run.style.shading.clone()),
                                                        cs, run.style.text_scale.unwrap_or(100.0), false,
                                                        run.style.font_family_east_asia.clone(), false, run.style.clone()));
                                                }
                                                if !tab_break { explicit_break_lines.insert(lines.len()); }
                                                lines.push(std::mem::take(&mut current_line));
                                                s1169_trailing_break = true;
                                                cell_edge_tab_nowrap = false;
                                                line_x = 0.0;
                                                current_line_chars.clear();
                                                is_first_line = false;
                                                prev_char_emitted = None;
                prev_char_gap = None;
                                                prev_char_ruby = false;
                                                if !tab_break { continue; }
                                            }
                                            // S443 (2026-05-30, SHIP, default ON, opt-out OXI_S443_DISABLE):
                                            // TAB-STOP advancement in the CELL wrap path. The body path
                                            // (mod.rs:5986-6014) advances a '\t' to the next tab stop, but the
                                            // cell path historically treated '\t' as a ~0-width char (S442 root
                                            // cause: d77a item J — marker カ + tab + body — wrapped 1 line in
                                            // Oxi vs 2 in Word because the missing ~12pt tab advance left room
                                            // to fit the trailing す。). Mirror the body formula: tab stops are
                                            // in absolute coords from the cell content-left; line_x+buf_w is
                                            // relative to the line's wrap start, so add the line-start indent
                                            // offset to convert. Gated to hanging-indent paragraphs
                                            // (first_line_indent<0 = the list-marker pattern) to bound the
                                            // blast radius; combined with the para_has_tab oikomi gate below
                                            // this is perfectly isolated (only d77a moves: +0.0306; 3a4f and
                                            // all other 54 docs EXACTLY unchanged; Phase 1 54/55).
                                            if ch == '\t'
                                                && std::env::var("OXI_S443_DISABLE").is_err()
                                                && p_first_line_indent < 0.0
                                            {
                                                let indent_off = if is_first_line {
                                                    (p_indent_left + p_first_line_indent).max(0.0)
                                                } else {
                                                    p_indent_left
                                                };
                                                let abs_pos = line_x + buf_w + indent_off;
                                                let (next_pos, _) = self.cell_next_tab_stop(
                                                    para, abs_pos, p_indent_left, p_first_line_indent);
                                                if std::env::var("OXI_CELL_EDGE_TAB_DISABLE").is_err()
                                                    && !self.doc_body_has_real_cjk
                                                    && (next_pos * 20.0).round() >= ((wrap_w + p_indent_left) * 20.0).round()
                                                {
                                                    cell_edge_tab_nowrap = true;
                                                }
                                                let tab_w = (next_pos - abs_pos).max(0.0);
                                                // A tab is a positioned control, not a glyph.
                                                // Keep its source text and measured advance in
                                                // its own fragment so every renderer starts the
                                                // next visible fragment at the tab stop. The
                                                // wrapper already used this width for capacity.
                                                if !buf.is_empty() {
                                                    current_line.push((buf.clone(), font_size, buf_w, bold,
                                                        run.style.italic, run.style.underline,
                                                        run.style.underline_style.clone(), run.style.strikethrough,
                                                        font_family.clone(), run.style.color.clone(),
                                                        run.style.highlight.clone().or_else(|| run.style.shading.clone()),
                                                        cs, run.style.text_scale.unwrap_or(100.0),
                                                        std::mem::take(&mut s993_lrpb_pending),
                                                        run.style.font_family_east_asia.clone(), run.ruby.is_some(),
                                                        run.style.clone()));
                                                    line_x += buf_w;
                                                    buf.clear();
                                                    buf_w = 0.0;
                                                    current_line_chars.extend(buf_chars.drain(..));
                                                }
                                                current_line.push(("\t".to_string(), font_size, tab_w, bold,
                                                    run.style.italic, run.style.underline,
                                                    run.style.underline_style.clone(), run.style.strikethrough,
                                                    font_family.clone(), run.style.color.clone(),
                                                    run.style.highlight.clone().or_else(|| run.style.shading.clone()),
                                                    cs, run.style.text_scale.unwrap_or(100.0),
                                                    std::mem::take(&mut s993_lrpb_pending),
                                                    run.style.font_family_east_asia.clone(), run.ruby.is_some(),
                                                    run.style.clone()));
                                                line_x += tab_w;
                                                current_line_chars.push(
                                                    crate::layout::jc_both_compress::CharContext {
                                                        ch,
                                                        natural_advance: tab_w,
                                                        font_size,
                                                    },
                                                );
                                                prev_char_emitted = Some(ch);
                prev_char_gap = Some(self.natural_autospace_after(ch, &run.style, &para.style, font_size, cs));
                                                prev_char_ruby = run.style.ruby_field;
                                                continue;
                                            }
                                            // S763: metrics keep the legacy quote class, EXCEPT for
                                            // S1052 (a curly quote glued to ASCII text = Latin).
                                            let cm = &*self.metrics_for_char_in(
                                                ch,
                                                !self.s1052_cell_latin_quote(
                                                    &s586_run_chars,
                                                    s586_ci,
                                                ),
                                                &run.style,
                                                &para.style,
                                            );
                                            // S1204 (2026-08-24, opt-in `OXI_S1204`): snap the
                                            // SIZE to the 600dpi device grid once, then take the
                                            // advance at that size -- as opposed to S1203, which
                                            // snapped every advance and is falsified. Rounding the
                                            // size happens once, so it does not conflict with the
                                            // alternating 10.444/10.560 that says the ORIGIN is what
                                            // gets snapped per glyph.
                                            //   10.5pt -> 10.5 x 600/72 = 87.5 (a tie) -> 88 -> 10.56
                                            // tokyoshugyo p30 needs exactly this: its budget is
                                            // 389.69 (cell rules 92.06..522.22 out of Word's PDF,
                                            // pad 4.95, line origin 127.58) and Word refuses the
                                            // 37th fullwidth glyph, which is 388.50 at 10.50 but
                                            // 390.72 at 10.56.
                                            //
                                            // ★FALSIFIED too, and with the SAME signature as S1203:
                                            // 3a4f9fbe 1.0000 -> 0.8334, ed025 1.0000 -> 0.9633.
                                            // So 10.56 is not a global property of the face at
                                            // 10.5pt -- neither snapping each advance nor snapping
                                            // the size once survives the corpus. Whatever charges
                                            // tokyoshugyo p30 its extra 0.74pt is LOCAL, and every
                                            // local input there has been read and matches Word: the
                                            // cell rules (92.06..522.22 vs Oxi 92.05..522.10), the
                                            // padding (4.95, from the Normal Table default -- the
                                            // table declares no tblCellMar, no tblInd, no tblStyle)
                                            // and the line origin (127.58 vs 127.60). Both flags are
                                            // kept only so the two models are not tried a third time.
                                            let s1204_fs = if std::env::var("OXI_S1204").is_ok() {
                                                (font_size * 600.0 / 72.0 + 0.5).floor() * 72.0 / 600.0
                                            } else {
                                                font_size
                                            };
                                            let mut cw = self
                                                .registry
                                                .char_width_pt_with_fallback(ch, s1204_fs, cm);

                                            // S1203 (2026-08-24, opt-in `OXI_S1203`): snap the
                                            // break-time advance to the 600dpi device grid
                                            // (1/600 inch = 0.12pt, half up).
                                            //
                                            // Word's own advances, read glyph by glyph out of its
                                            // PDF, are all multiples of 0.12: ＭＳ Ｐ明朝 「5.280
                                            // （5.164 て9.483, and ＭＳ 明朝 fullwidth 10.560 --
                                            // which is 10.5pt snapped UP (10.5 x 600/72 = 87.5, a
                                            // tie). tokyoshugyo p30 turns on exactly that 0.06:
                                            // Word's cell there measures 92.06..522.22 off the PDF
                                            // rules, so the budget is 389.24, and «…ありますす。»
                                            // is 37 glyphs -- 388.50 at 10.50 each (fits, which is
                                            // what Oxi does) but 390.72 at 10.56 (does not, which is
                                            // what Word does).
                                            //
                                            // ★FALSIFIED, parked so it is not tried again. Snapping
                                            // EVERY advance is broadly wrong: 3a4f9fbe 1.0000 ->
                                            // 0.8334, ed025 1.0000 -> 0.9690, tokyoshugyo -> 0.9931.
                                            // The tell was already in the measurement -- ＭＳ 明朝
                                            // fullwidth comes out of Word's PDF as BOTH 10.444 and
                                            // 10.560 along one line. That alternation is the
                                            // signature of snapping the cumulative ORIGIN to the
                                            // grid, not each advance: the running total then stays
                                            // within 0.06 of n x 10.5 and never drifts, so it cannot
                                            // be what pushes 37 glyphs over a 389.24 budget.
                                            if std::env::var("OXI_S1203").is_ok() {
                                                cw = (cw * 600.0 / 72.0 + 0.5).floor() * 72.0 / 600.0;
                                            }
                                            if let Some(adv) = cell_deva_adv.as_ref() {
                                                if crate::font::is_complex_script(ch)
                                                    || s586_run_chars.get(s586_ci.wrapping_sub(1)).map_or(false, |&p| crate::font::is_complex_script(p))
                                                {
                                                    cw = adv[s586_ci];
                                                }
                                            }
                                            // S869 (2026-07-16, default ON, opt-out OXI_S869_DISABLE):
                                            // LATINEM for the CELL wrapper. The cell
                                            // breaker is a SEPARATE greedy wrapper from break_into_lines,
                                            // so LATINEM (no-kern Latin breaks at the UN-ROUNDED em
                                            // advance) never reached a cell: char_width_pt_with_fallback
                                            // rounds every char to 10tw (0.5pt), which for Calibri 10pt
                                            // biases +0.144/char. forms__000ee7c0 render-truth: Word
                                            // fits "Has student previously received an Individual
                                            // Evaluation?" on ONE line at 233.45pt = EXACTLY the em
                                            // width from the real Calibri TTF (Oxi's width table is
                                            // unit-exact); the 10tw round summed 241.50 > wrap 238.05
                                            // so Oxi wrapped to 2 lines (+11.67pt) on four rows. The em
                                            // width is CORRECT (verified on policies__0009e9db too: its
                                            // "Entry Level 3 and Level 1-2" float cell renders 1 line in
                                            // Word = the S869-ON wrap; float #1 bottom 641.92 = Word
                                            // 641.26). The policies PASS→FAIL this initially caused was
                                            // an S872 exposure (an empty para placed in a float gap its
                                            // line box crosses), fixed there — the S869+S870+S871+S872
                                            // set ships together (forms 1.000 + policies 1.000).
                                            // Same rule, same scope as break_into_lines' LATINEM.
                                            if std::env::var("OXI_S869_DISABLE").is_err() {
                                                let kern_active =
                                                    std::env::var("OXI_KERNBREAK_DISABLE").is_err()
                                                        && run
                                                            .style
                                                            .kern
                                                            .or_else(|| {
                                                                para.style
                                                                    .default_run_style
                                                                    .as_ref()
                                                                    .and_then(|rs| rs.kern)
                                                            })
                                                            .map_or(false, |k| {
                                                                k > 0.0 && font_size >= k
                                                            });
                                                if !kern_active
                                                    && std::env::var("OXI_LATINEM_DISABLE").is_err()
                                                    && latinem_in_scope(self.doc_body_has_real_cjk)
                                                    && !kinsoku::is_cjk(ch)
                                                    && cell_has_em_advance(cm, ch)
                                                {
                                                    cw = cm.char_width_em(ch) * font_size;
                                                } else if kern_active
                                        // S1017 (2026-07-26, opt-out
                                        // OXI_S1017_DISABLE): KERNBREAK for the
                                        // CELL wrapper — a kern-active Latin run
                                        // (w:kern <= font size, e.g.
                                        // forms__0011786c docDefaults kern=2 →
                                        // Arial 12pt) breaks at the em advance +
                                        // legacy kern-pair adjustment (the body
                                        // break_into_lines model, mod.rs:15993),
                                        // not the 10tw-rounded fallback.
                                        // "Marks/Scars/Tattoos:" fits ONE line in
                                        // Word (112.116pt via the T,a −1.330pt
                                        // pair) but Oxi's cell wrapper (unrounded
                                        // em 113.446 > wrap 112.70) wrapped 2
                                        // lines → +13.799pt → +1. The estimate
                                        // mirror (count_cell_lines) MUST match.
                                        && std::env::var("OXI_S1017_DISABLE").is_err()
                                        && !self.doc_body_has_real_cjk
                                        && !kinsoku::is_cjk(ch)
                                        && cell_has_em_advance(cm, ch)
                                                {
                                                    cw = cm.char_width_em(ch) * font_size;
                                                    if let Some(&next) =
                                                        s586_run_chars.get(s586_ci + 1)
                                                    {
                                                        if !kinsoku::is_cjk(next) {
                                                            cw += self.registry.latin_kern_em(
                                                                &cm.family,
                                                                cm.units_per_em,
                                                                ch,
                                                                next,
                                                            ) * font_size;
                                                        }
                                                    }
                                                }
                                            }
                                            let mut deferred_fullwidth_grid = 0.0;
                                            if std::env::var_os("OXI_CELL_BALANCED_SPACE").is_some()
                                                && self.balance_single_byte_double_byte_width
                                                && balanced_cjk_space(ch,
                                                    (s586_run_offset + s586_ci).checked_sub(1)
                                                        .and_then(|i| s586_para_chars.get(i).copied()),
                                                    s586_para_chars.get(s586_run_offset + s586_ci + 1).copied())
                                            {
                                                cw = font_size * 0.5 + cm.synthetic_bold_advance * font_size;
                                            }

                                            if std::env::var("OXI_DBG1244").is_ok()
                                                && para
                                                    .runs
                                                    .iter()
                                                    .any(|r| r.text.contains("To be a member"))
                                            {
                                                eprintln!("[DBG1244] ch={:?} cw={:.3}", ch, cw);
                                            }
                                            // S691 (2026-06-29) FALSIFIED: forcing full-width digits (U+FF1x)
                                            // to font_size in the cell break (OXI_FWDIGIT) was a NO-OP — the
                                            // 第N条-marker break-width discrepancy (--dump-layout 9.5/char vs
                                            // --dump-glyphs painted 10.5/char) is NOT the full-width digit
                                            // char-width (already 10.5 here) but a charSpace/marker-path cause
                                            // (UNISOLATED). The chapter cell over-fit root is still open.
                                            // OXI_CJKADV_K=<pt>: experimental — widen the fullwidth-CJK cell
                                            // break advance (Word's typed-docGrid CJK advance ≈10.56 vs Oxi em
                                            // 10.50). Applied to chars whose natural advance ≈ font_size (full-
                                            // width). Tests whether matching Word's advance + cell width aligns
                                            // the 賃金 chapter. Affects break AND render (page-count test only).
                                            if let Ok(k) = std::env::var("OXI_CJKADV_K") {
                                                if (cw - font_size).abs() < 1.0 {
                                                    cw += k.parse::<f32>().unwrap_or(0.0);
                                                }
                                            }
                                            // 2026-04-19: Apply charSpace as ABSOLUTE delta (not fs-scaled).
                                            // COM-measured b35 fs=9 → 8.3pt, fs=10.5 → 9.8pt: both are
                                            // fs − |charSpace_pt| (0.663pt), NOT fs × ratio.
                                            // Previous formula (fs × pitch/default_fs) over-compressed
                                            // when fs<default_fs. Correct: cw = fs + charSpace_pt where
                                            // charSpace_pt = pitch − default_fs (negative for compressPunc).
                                            // S342 (2026-05-27): see effective_char_pitch at line 4073 for
                                            // OXI_S342_NO_SNAP_GATE gate-drop rationale.
                                            // S344 (2026-05-27): refine S342 to require fs < default_fs
                                            // when snap_to_grid=false. See count_cell_lines comment.
                                            // S342 SHIP (2026-05-27): default ON. Drops snap_to_grid gate from
                                            // char-grid (horizontal compression) per OOXML §17.3.1.32. Env-var
                                            // preserved as opt-OUT.
                                            let s342_no_snap_gate =
                                                std::env::var("OXI_S342_NO_SNAP_GATE")
                                                    .map(|v| v != "0" && v != "false")
                                                    .unwrap_or(true);
                                            let s344_fs_gate =
                                                std::env::var("OXI_S344_FS_LT_DEFAULT")
                                                    .map(|v| v != "0" && v != "false")
                                                    .unwrap_or(false);
                                            let snap_ok = s342_no_snap_gate
                                                || s344_fs_gate
                                                || para.style.snap_to_grid;
                                            if run.style.fit_text.is_none() && snap_ok {
                                                if let (Some(ratio), Some(pitch)) =
                                                    (grid_char_cw_ratio, grid_char_pitch)
                                                {
                                                    if ratio > 0.0
                                                        && pitch > 0.0
                                                        && cw > 0.0
                                                        && crate::font::is_fullwidth(ch)
                                                    {
                                                        let default_fs = pitch / ratio;
                                                        let char_space_pt = pitch - default_fs;
                                                        // R7.59 hybrid (see break_into_lines comment).
                                                        // S141 H6: skip expansion when font_size < default_fs
                                                        let h6_skip =
                                                            std::env::var("OXI_H6_GRID_GATE")
                                                                .is_ok()
                                                                && char_space_pt > 0.0
                                                                && font_size < default_fs;
                                                        let h7_skip =
                                                            std::env::var("OXI_H7_GRID_GATE_LE")
                                                                .is_ok()
                                                                && char_space_pt > 0.0
                                                                && font_size <= default_fs;
                                                        // S151 H8 default ON: skip positive char_grid_extra
                                                        // S239 (2026-05-23): removed OXI_LEGACY_GRID_KERN.
                                                        // S466 NOTE (2026-05-31): the CELL visible-wrap mirror of
                                                        // the break_into_lines h8 change (make cells grid-fit 36
                                                        // like Word instead of natural 37) was TRIED here and
                                                        // REVERTED — it traded the p7 partial-fix (-0.0138→-0.0065)
                                                        // for NEW p5/p6 cascade regressions (a1d6 p5 -0.020, 6514
                                                        // p5 -0.020, d4d126 p6 -0.014), netting 8 regressions vs 4
                                                        // at the same family mean (+0.0017 vs +0.0019). The
                                                        // b35123-class cell-wrap re-flow cascade (memory) fired.
                                                        // Cell visible-wrap stays h8-natural; S466 is body-only.
                                                        // S466CELL re-test (2026-06-25): the cell PAINT advance
                                                        // (grid_cs_adj, 13165) is the grid pitch (10.855), but this
                                                        // wrap decision uses natural (10.5) → the painted line over-
                                                        // packs ~1 char past the cell border. Mirror the body h8
                                                        // (expand fs>=default to grid) so wrap == paint == Word.
                                                        // Re-tested vs the post-S660/661/664/666 baseline (the 2026-
                                                        // 05-31 revert predated those vertical fixes). Default OFF.
                                                        let s466cell =
                                                            std::env::var("OXI_S466CELL").is_ok();
                                                        // S1210 (2026-08-24, opt-in `OXI_S1210=1`):
                                                        // the cell mirror of the additive pitch
                                                        // derived in break_into_lines. a1d6e4ef's
                                                        // note cell is the specimen: Word breaks
                                                        // 「者及び利用者、」 after 7 characters, and
                                                        // only the grid pitch explains it. The cell
                                                        // is 67.74pt wide (measured off the PDF's own
                                                        // rules: 53.64..152.90 less 0.6pt margins and
                                                        // the paragraph's indents), 8 x 9.3547 =
                                                        // 74.84 overflows it by 0.76em -- past any
                                                        // 約物 credit S1209 can pay -- while 8 x 9.00
                                                        // = 72.00 overflows by only 0.47em, which the
                                                        // half-em credit DOES cover. That is why the
                                                        // ※2 note sits a page early in Oxi.
                                                        let s1210 = std::env::var("OXI_S1210_DISABLE").is_err();
                                                        let h8_skip = char_space_pt > 0.0
                                                            && !s1210
                                                            && (!s466cell
                                                                || font_size < default_fs);
                                                        // S344: when snap_to_grid=false and S344 enabled,
                                                        // skip compression unless fs < default_fs.
                                                        let s344_skip = s344_fs_gate
                                                            && !para.style.snap_to_grid
                                                            && font_size >= default_fs;
                                                        // S1374 (2026-09-13, default ON, opt-out
                                                        // OXI_S1374_DISABLE): no charSpace, no
                                                        // character cell -- the body rule (S466 /
                                                        // S1315) applied to cells. MEASURED: BIZ
                                                        // UDP明朝 8pt, 22 chars, linesAndChars
                                                        // linePitch=325 without charSpace: Word
                                                        // 147.0pt in a cell and in the body; the
                                                        // cell walk here widened every character
                                                        // to 8.0 (168.0) and wrapped
                                                        // forms__00830ac053a2c57a's cells.
                                                        let s1374_no_cell = char_space_pt.abs() < 0.01
                                                            && std::env::var("OXI_S1374_DISABLE").is_err();
                                                        if !(h6_skip
                                                            || h7_skip
                                                            || h8_skip
                                                            || s344_skip
                                                            || s1374_no_cell)
                                                        {
                                                            let grid_base = if char_space_pt > 0.0 && cell_resolved_em(cm, &self.registry, ch, font_size) < 0.99 { cw } else { font_size };
                                cw = if char_space_pt >= 0.0
                                                                && !s1210
                                                            {
                                                                grid_base * pitch / default_fs
                                                            } else {
                                                                // S1210: additive for both signs.
                                                                grid_base + cell_cjk_grid_increment(ch, cm, &self.registry, font_size, char_space_pt)
                                                            };
                                                            if (std::env::var_os("OXI_CELL_GRID_SINGLE_BYTE").is_some()
                                            || std::env::var_os("OXI_S1449_DISABLE").is_none()) {
                                                                deferred_fullwidth_grid = cw - grid_base;
                                                                cw = grid_base;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // Adjacent punctuation shares an advance across run boundaries.
                                            // Run segmentation does not change Word's cell wrapping.
                                            let pair_next = s586_para_chars
                                                .get(s586_run_offset + s586_ci + 1).copied();
                                            if pair_next.map_or(false, |next| {
                                                let legacy = self.compat_mode < 15 && self.compress_punctuation;
                                                let kern = run.style.kern.or_else(|| para.style.default_run_style.as_ref().and_then(|r| r.kern))
                                                    .map_or(false, |k| k > 0.0 && font_size >= k);
                                                (legacy && kinsoku::is_yakumono_closing(ch)
                                                    && kinsoku::is_yakumono_trigger(next))
                                                    || ((legacy || kern) && kinsoku::is_yakumono_opening(ch)
                                                        && kinsoku::is_yakumono_opening(next))
                                            }) {
                                                cw = cw.min(font_size * 0.5);
                                                deferred_fullwidth_grid *= 0.5;
                                            }
                                            if let Some(scale) = run.style.text_scale {
                                                if (scale - 100.0).abs() > 0.01 {
                                                    cw *= scale / 100.0;
                                                }
                                            }
                                            cw += deferred_fullwidth_grid;
                if run.style.fit_text.is_none() {
                    if let Some((base, extra)) = cell_proportional_space_grid(ch, cm, font_size,
                        grid_char_pitch, grid_char_cw_ratio, self.balance_single_byte_double_byte_width) {
                        cw = base * run.style.text_scale.unwrap_or(100.0) / 100.0 + extra;
                    }
                }
                                            if (std::env::var_os("OXI_CELL_GRID_SINGLE_BYTE").is_some()
                                            || std::env::var_os("OXI_S1449_DISABLE").is_none()) && run.style.fit_text.is_none() {
                                                cw += cell_single_byte_grid_increment(ch, grid_char_pitch,
                                                    grid_char_cw_ratio, self.balance_single_byte_double_byte_width);
                                            }
                                            // Session 56 Finding 3: balanceSingleByteDoubleByteWidth
                                            // doubles cs for CJK fullwidth chars.
                                            // Day 37 (2026-05-14): EXCLUDE fitText runs from balance
                                            // doubling. resolve_fit_text_runs (mod.rs:1408) computes the
                                            // per_em_cs as (target − natural) / denom_em — this value is
                                            // the FINAL effective cs that should be applied at render to
                                            // hit the target width. Applying balance doubling on top of
                                            // this doubled value over-expands by 2×, causing fitText
                                            // paragraphs to wrap at the cell boundary (ed025c "(2) ○○
                                            // 奨励金" wraps at "○○" because cw=10.5+21+21=52.5pt × 3 chars
                                            // = 157.5pt > first_line_wrap_w=170pt approximately. The
                                            // correct cs=21pt gives cw=31.5pt × 3 = 94.5pt which fits).
                                            let balance_extra_cs = if self
                                                .balance_single_byte_double_byte_width
                                                && crate::font::is_fullwidth(ch)
                                                && run.style.fit_text.is_none() && !run.style.ruby_spread
                                            {
                                                cs
                                            } else {
                                                0.0
                                            };
                                            let cw = cw + cs + balance_extra_cs;
                                            // R7.51 (2026-05-13): autoSpaceDE for CJK↔Latin transitions
                                            // in cell renderer. The body renderer (break_into_lines) already
                                            // applies this 2.5pt gap (at 10.5pt) but the cell-renderer loop
                                            // here historically did not. d77a58 w_i=47 wrap mismatch
                                            // (5 lines Oxi vs 6 lines Word) traced to missing auto-space
                                            // around "URL" / "1.0" / "CC BY" Latin runs within CJK text.
                                            // Formula matches break_into_lines: ((fs/2)+0.5).floor()*0.5.
                                            // Session 95 (2026-05-18) split DE (alpha) vs DN (digit).
                                            let auto_space_extra = {
                                                let prev_cjk_ideo = prev_char_emitted.map_or(
                                                    false,
                                                    kinsoku::is_cjk_ideograph_or_kana,
                                                );
                                                let prev_alpha = prev_char_emitted
                                                    .map_or(false, |c| c.is_ascii_alphabetic());
                                                let prev_digit = prev_char_emitted
                                                    .map_or(false, |c| c.is_ascii_digit());
                                                let cur_cjk_ideo =
                                                    kinsoku::is_cjk_ideograph_or_kana(ch);
                                                let cur_alpha = ch.is_ascii_alphabetic();
                                                let cur_digit = ch.is_ascii_digit();
                                                let de_boundary = (prev_cjk_ideo && cur_alpha)
                                                    || (prev_alpha && cur_cjk_ideo);
                                                let dn_boundary = (prev_cjk_ideo && cur_digit)
                                                    || (prev_digit && cur_cjk_ideo);
                                                // S1316: no auto-space against a ruby field.
                                                let s1316_adj = std::env::var("OXI_S1316_DISABLE").is_err()
                                                    && (prev_char_ruby || run.style.ruby_field);
                                                if !s1316_adj
                                                    && ((de_boundary && para.style.auto_space_de)
                                                        || (dn_boundary && para.style.auto_space_dn))
                                                {
                                                    if let Some(gap) = prev_char_gap { gap } else if std::env::var("OXI_AUTOSPACE2_DISABLE").is_err()
                                                    {
                                                        s1175_autospace(
                                                            font_size,
                                                            cs,
                                                            self.balance_single_byte_double_byte_width
                                                                && run.style.fit_text.is_none() && !run.style.ruby_spread,
                                                        )
                                                    } else {
                                                        s546_autospace_extra(font_size)
                                                    }
                                                } else {
                                                    0.0
                                                }
                                            };
                                            let cw = cw + auto_space_extra;
                                            // OXI_DUMP_GLYPHW=1: the cell breaker's per-CHARACTER
                                            // budget, broken into the terms that build it. The
                                            // archive's 04b88e note stalled at "Oxi has no
                                            // per-glyph instrument, so this cannot be split
                                            // further" -- a run-level width cannot say whether a
                                            // 2pt difference is one wide glyph, a doubled letter
                                            // space, or a CJK/Latin gap on the wrong side of a
                                            // joint. Set OXI_DUMP_GLYPHW_TEXT to a substring to
                                            // restrict the dump to the paragraph that contains it.
                                            if std::env::var("OXI_DUMP_GLYPHW").is_ok() {
                                                let want = std::env::var("OXI_DUMP_GLYPHW_TEXT")
                                                    .unwrap_or_default();
                                                let ptext: String = para
                                                    .runs
                                                    .iter()
                                                    .flat_map(|r| r.text.chars())
                                                    .collect();
                                                if want.is_empty() || ptext.contains(&want) {
                                                    eprintln!(
                                        "[GLYPHW] ch={:?} fw={} base={:.3} cs={:.3} bal={:.3} aki={:.3} cw={:.3} line_x={:.3} buf_w={:.3} fs={:.2} face={:?} upem={}",
                                        ch, crate::font::is_fullwidth(ch),
                                        cw - cs - balance_extra_cs - auto_space_extra,
                                        cs, balance_extra_cs, auto_space_extra, cw,
                                        line_x, buf_w, font_size,
                                        cm.family, cm.units_per_em
                                    );
                                                }
                                            }
                                            let effective_wrap = if cell_edge_tab_nowrap {
                                                f32::MAX / 4.0
                                            } else { cell_float_wrap.as_ref().map_or_else(|| if is_first_line {
                                                first_line_wrap_w
                                            } else {
                                                wrap_w
                                            }, |wrap| wrap.frame(lines.len(), wrap_w, first_line_wrap_w, p_indent_left, p_first_line_indent).width) };
                                            // S1082 (2026-08-06, opt-out OXI_S1082_DISABLE): the
                                            // CELL wrapper gets the S825 compat-15 justified
                                            // SPACE-SHRINK capacity that break_into_lines has.
                                            // The cell wrapper is a separate greedy breaker (the
                                            // S869 LATINEM / S1017 KERNBREAK family) and it had no
                                            // shrink credit at all, so a justified cell line was
                                            // fitted at its NATURAL width where Word compresses
                                            // the inter-word spaces.
                                            // MEASURED (technical__00501ca3, Word PDF per-char
                                            // origins, Times New Roman 9.96 in a 259.35pt cell):
                                            // the line «change, and the new value of the
                                            // parameter.  A printed copy of the » is 265.77 at
                                            // hmtx-natural, and Word paints it with its 14 spaces
                                            // at 2.281-2.490 (natural 2.490) for a VISIBLE extent
                                            // of 258.61 — inside Oxi's own 259.35 budget. Oxi
                                            // dropped «the», the right cell took a 10th line, the
                                            // last row grew 8.4pt and «(Table Added 2018)» spilled
                                            // to the next page.
                                            // The capacity is S825's, per space: 0.25 x em (the cs
                                            // term is folded in — a cell run's cs is 0 in every
                                            // measured case and 0.25 vs 0.24 on it is < 0.01pt).
                                            // Word can shrink inter-word spaces on the final line too.
                                            // The terminal-word exclusion conflated a trailing-space
                                            // endpoint with the usable cell width. Keep the same
                                            // quarter-space capacity for interior and terminal words.
                                            let effective_wrap = effective_wrap
                                                + if s1082_cell_shrink {
                                                    (current_line_chars
                                                        .iter()
                                                        .chain(buf_chars.iter())
                                                        .filter(|c| c.ch == ' ')
                                                        .map(|c| c.natural_advance)
                                                        .sum::<f32>())
                                                        * 0.25
                                                } else {
                                                    0.0
                                                };
                                            // S1174 (opt-in `OXI_YAKUCOMP=1`): a LEGACY
                                            // compressPunctuation cell line does not break the
                                            // moment it runs out of room -- it takes the shortfall
                                            // out of its 約物, half an em each, and breaks only
                                            // when that pool cannot cover it. See
                                            // `kinsoku::cell_yaku_capacity` for the sweep this
                                            // comes from. Adding the pool to the budget is the
                                            // same shape S1082 uses for justified Latin spaces.
                                            // ★This is the second half of S1173: at Word's true
                                            // cell budget Oxi produces MORE lines than Word on
                                            // exactly the twelve legacy compressPunctuation
                                            // documents, because it wraps where Word squeezes.
                                            // Trailing spaces don't trigger line wrapping (Word behavior)
                                            let is_space = ch == ' ' || ch == '\u{3000}';
                                            // S118 wrap-decision lookahead: when env var ON + gate active,
                                            // call compute_compression on (current_line + buf + ch) and only
                                            // wrap if line CANNOT fit even with priority compression applied.
                                            // S119 tuned kanji_max_savings 0.6% → 0.1% to reduce over-fit.
                                            // S121 fix: require run.cs < 0 (matches original S112 trigger).
                                            // S122 refinement: require run.cs ≤ -0.1pt (= ≤ -2tw). The
                                            // d1e8ac8 doc has a custom style "一太郎" with `cs=-1` (= -0.05pt,
                                            // 1 twip negative — effectively no compression). My gate `cs<0`
                                            // fired on those paragraphs and shifted them slightly, causing
                                            // -0.03 SSIM regression on d1e8 p.1. S113 grid showed Word
                                            // actually compresses at cs∈{-5,-9,-15,-20}tw = {-0.25..-1.0pt};
                                            // cs=-1tw=-0.05pt is below Word's compression threshold.
                                            if std::env::var_os("OXI_DEBUG_FIT_CELL").is_some()
                                                && run.style.fit_text.is_some() {
                                                eprintln!("[FIT_CELL] id={:?} ch={:?} cs={:.9} cw={:.9} x={:.9} buf={:.9} cap={:.9} sum={:.9}",
                                                    run.style.fit_text_id, ch, cs, cw, line_x, buf_w, effective_wrap, line_x + buf_w + cw);
                                            }
                                            let would_overflow_natural =
                                                line_x + buf_w + cw > effective_wrap;
                                            let run_has_neg_cs = cs <= -0.1;
                                            // S586 orphan + small-overflow + 約物→opener oikomi (scaffold).
                                            let s586_overflow_fixed = if s586_orphan
                                                && would_overflow_natural
                                                && !kinsoku::is_cjk_compressible(ch)
                                            {
                                                let gpos = s586_run_offset + s586_ci;
                                                let chars_to_end =
                                                    s586_para_chars.len().saturating_sub(gpos);
                                                let overflow =
                                                    (line_x + buf_w + cw) - effective_wrap;
                                                // OXI_S586_FIT=1: experimental — fit unconditionally when the
                                                // page-44 signature (約物→opener cluster present) fires, with a
                                                // generous overflow cap. Tests whether fitting the page-44 box
                                                // at the S594-narrowed width aligns the 賃金 chapter (Δ0).
                                                let s586_fit_mode =
                                                    std::env::var("OXI_S586_FIT").ok().as_deref()
                                                        == Some("1");
                                                let fit_cap = if s586_fit_mode {
                                                    std::env::var("OXI_S586_FITCAP")
                                                        .ok()
                                                        .and_then(|v| v.parse().ok())
                                                        .unwrap_or(15.0)
                                                } else {
                                                    s586_cap
                                                };
                                                if chars_to_end <= 2 && overflow <= fit_cap {
                                                    let seq: Vec<char> = current_line_chars
                                                        .iter()
                                                        .chain(buf_chars.iter())
                                                        .map(|c| c.ch)
                                                        .chain(std::iter::once(ch))
                                                        .collect();
                                                    // collapse a 約物 immediately before an opening bracket
                                                    // to ~3.0pt (full 、「 inter-space vanish), natural − 3.0.
                                                    let collapse: f32 = seq.iter().enumerate()
                                            .filter(|(i, &c)| matches!(c, '、' | '。' | '，' | '．')
                                                && seq.get(i + 1).map_or(false, |&n| kinsoku::is_yakumono_opening(n)))
                                            .map(|_| (font_size - 3.0).max(0.0))
                                            .sum();
                                                    if s586_fit_mode {
                                                        // FULL-CLUSTER capacity: Word fits the orphan char by
                                                        // compressing the whole line's 約物 cluster (page-44:
                                                        // 、→3.0 [-7.5], ・→9.4 [-1.1]×2, 」→8.3 [-2.2] ≈ -11.9pt).
                                                        // Gate stays on the 約物→opener signature (`collapse>0`);
                                                        // available = opener-adjacent 約物 (full -7.5) + every
                                                        // other compressible 約物 (up to half-em). Fit iff the
                                                        // line (incl. ch) minus available compression ≤ wrap.
                                                        let mut available = collapse; // opener-adjacent 約物
                                                        for (i, &c) in seq.iter().enumerate() {
                                                            let opener_adj = matches!(
                                                                c,
                                                                '、' | '。' | '，' | '．'
                                                            ) && seq
                                                                .get(i + 1)
                                                                .map_or(false, |&n| {
                                                                    kinsoku::is_yakumono_opening(n)
                                                                });
                                                            if opener_adj {
                                                                continue;
                                                            } // already counted in `collapse`
                                                            if matches!(
                                                                c,
                                                                '、' | '。'
                                                                    | '，'
                                                                    | '．'
                                                                    | '」'
                                                                    | '』'
                                                                    | '）'
                                                                    | '】'
                                                                    | '〕'
                                                                    | '・'
                                                            ) {
                                                                available += font_size / 2.0;
                                                            }
                                                        }
                                                        collapse > 0.0
                                                            && (line_x + buf_w + cw - available)
                                                                <= effective_wrap
                                                    } else {
                                                        collapse > 0.0
                                                            && (line_x + buf_w + cw - collapse)
                                                                <= effective_wrap
                                                    }
                                                } else {
                                                    false
                                                }
                                            } else {
                                                false
                                            };
                                            // OXI_CELLCOMP (tokyoshugyo #2c, the other half): enable
                                            // compute_compression for cs=0 JUSTIFIED cells (the 条文 boxes are
                                            // align=Justify cs=0 → no 約物 compression by default → over-wrap).
                                            // Paired with OXI_PGCAP (wrap-to-margin): narrow wrap + 約物
                                            // compression together = Word's cell rendering. See
                                            // [[tokyoshugyo_wrap_not_cellheight]].
                                            let cell_comp_active =
                                                std::env::var("OXI_CELLCOMP").ok().as_deref()
                                                    == Some("1")
                                                    && matches!(
                                                        para.alignment,
                                                        Alignment::Justify | Alignment::Distribute
                                                    );
                                            // S466CELL 約物-oikomi (2026-06-25): when S466CELL grid-expansion is
                                            // on, a compressPunctuation cell over-wraps (grid-pitch kanji, no 約物
                                            // compression) where Word fits one more char by compressing close 約物
                                            // (、。」） up to half-em — DERIVED from the tokumei_08_01 family Word PDF
                                            // (a1d6e4 （２）cell: Word 35 chars w/ 約物→8.64, Oxi-grid 34). S497b
                                            // (this lookahead for jc=left cells) was a NO-OP on the DEFAULT em-wrap
                                            // (no overflow → never triggers); WITH grid-expansion it now fires.
                                            // jc-independent (the family cells are jc=left). Default OFF (no env).
                                            // Scoped to POSITIVE charSpace (ratio>1) = where grid-expansion fires.
                                            let s466_oikomi = std::env::var("OXI_S466CELL").is_ok()
                                                && self.compress_punctuation
                                                && grid_char_cw_ratio.map_or(false, |r| r > 1.0);
                                            // OXI_CELLBURA (tokyoshugyo #2c): cell line-end 約物 ぶら下げ
                                            // (mirrors the body S601). A hangable line-end 約物 (。、，．・ +
                                            // closing brackets) hangs PAST the wrap when the preceding content
                                            // fits (line_x+buf_w ≤ wrap) — its width is not counted, so the
                                            // line is not broken before it. Required to pair with OXI_PGCAP:
                                            // PGCAP caps content at the margin, and Word HANGS the trailing
                                            // 約物 past it (de-ぶら下げ content wrap is uniformly the margin).
                                            // Without it PGCAP wraps the trailing 約物 → +1 line/sentence-cell.
                                            let cell_bura_active = (std::env::var("OXI_CELLBURA")
                                                .ok()
                                                .as_deref()
                                                == Some("1")
                                                || std::env::var("OXI_LEGACYCELL").ok().as_deref()
                                                    == Some("1")
                                                || s585c_narrow)
                                                && matches!(
                                                    para.alignment,
                                                    Alignment::Justify | Alignment::Distribute
                                                );
                                            // OXI_LEGACYCELL (2026-06-23): the REFINED cell break = the derived
                                            // TWO-BUDGET model. LibreOffice (guess.cxx compress-to-fit + SwHangingPortion)
                                            // + the synthetic cell dataset (_cb_derive_cell.py) showed Word's BREAK fits a
                                            // would-wrap KANJI with only ~0.235em (2.46pt@10.5) 約物 compression — NOT
                                            // half-em (the JUSTIFY budget) — + ぶら下げ for line-end 約物. compute_compression
                                            // (CELLCOMP) uses half-em for the BREAK → over-fits (fits 又 where Word wraps).
                                            // This uses the small kanji-fit cap for the break; cell_bura handles line-end
                                            // ぶら下げ; compute_compression/justify keeps half-em for the render. Scoped to
                                            // LEGACY (compat≤14) compressPunctuation justified cells (3a4f/model = compat15,
                                            // EXCLUDED → their compensating baseline preserved). See [[char_budget_wall]].
                                            let legacy_cell_break = ((std::env::var("OXI_LEGACYCELL").ok().as_deref() == Some("1")
                                    || s585c_narrow)
                                    && self.compat_mode < 15 && self.compress_punctuation
                                    && matches!(para.alignment, Alignment::Justify | Alignment::Distribute))
                                    // OXI_CELLPAIR: the small-cap 約物 credit for ANY justified
                                    // compressPunctuation cell (compat-independent), paired with
                                    // the universal subtract boundary. 191cb なお: needs 1.45pt
                                    // over 3 、 (~0.48/、, well under the 2.5 cap).
                                    || (self.cellpair_active()
                                        && self.compress_punctuation
                                        && matches!(para.alignment, Alignment::Justify | Alignment::Distribute));
                                            let legacy_cell_cap: f32 =
                                                std::env::var("OXI_LEGACYCELL_CAP")
                                                    .ok()
                                                    .and_then(|v| v.parse().ok())
                                                    .unwrap_or(2.5);
                                            // S720: ぶら下げ (hanging punctuation) applies to PERIODS/COMMAS
                                            // only (JIS X 4051) — NOT closing brackets. Word render-truth
                                            // tokyoshugyo p20 (イ): the para-final «）» at overflow 2.6pt is
                                            // OIDASHI'd (間」）to the next line), not hung; the old list let
                                            // cell_bura hang any closing bracket → the box lost a line.
                                            let is_cell_hangable =
                                                if std::env::var("OXI_S720_DISABLE").is_err() {
                                                    matches!(ch, '。' | '、' | '，' | '．' | '・')
                                                } else {
                                                    matches!(
                                                        ch,
                                                        '。' | '、'
                                                            | '，'
                                                            | '．'
                                                            | '・'
                                                            | '）'
                                                            | '」'
                                                            | '』'
                                                            | '】'
                                                            | '〕'
                                                            | '］'
                                                            | '｝'
                                                    )
                                                };
                                            // Uncompressed CJK cells keep the terminal mark
                                            // inside the available width, including legacy modes.
                                            // S1429 (2026-09-16, default ON, opt-out OXI_S1429_DISABLE):
                                            // promoted from the OXI_CJK_CELL_NATURAL_LINE_END opt-in.
                                            // `_pb_hang_gen.py` (tests/fixtures/hang): a doNotCompress
                                            // cell never hangs its line-final mark (compat 14 and 15,
                                            // jc left and both, 11.8 and 11.2 char widths all push
                                            // 「二、」 down); a compressPunctuation cell keeps the S1174
                                            // half-em. policies__07543a6b p28 「服用した / 後、横紋筋」.
                                            let cell_natural_line_end = self.doc_body_has_real_cjk
                                                && !self.compress_punctuation
                                                && (std::env::var("OXI_CJK_CELL_NATURAL_LINE_END").is_ok()
                                                    || std::env::var_os("OXI_S1429_DISABLE").is_none());
                                            let would_overflow = if cell_bura_active
                                                && !cell_natural_line_end
                                                && is_cell_hangable
                                                && (line_x + buf_w) <= effective_wrap
                                                && !(current_line.is_empty() && buf.is_empty())
                                            {
                                                false
                                            } else if s586_overflow_fixed {
                                                false
                                            } else if (s1174_yakucomp || s1176_space || std::env::var_os("OXI_CELL_AKI").is_some())
                                                && would_overflow_natural
                                            {
                                                // S1174: the pool is HALF AN EM FOR THE LINE, not
                                                // half an em per 約物. Sweeping the count
                                                // (`_cw_yaku_n.py`, 12-character lines carrying 0
                                                // to 8 closing brackets) the line breaks at the
                                                // same 0.495-0.500em of shortfall whether it holds
                                                // one 約物 or eight. A line-final 約物 adds its own
                                                // half em on top, by hanging -- which is why the
                                                // five-character arms that ended in one held out to
                                                // a full em and read as "additive".
                                                // S1175: each CJK/Latin joint on the line is aki
                                                // too, and gives up about half of itself before
                                                // the line breaks -- 501 widths took the gap from
                                                // 1.560 down to 0.840 in device steps and only
                                                // then wrapped. Unlike the 約物 pool these ARE
                                                // additive: the probe's two joints both gave way
                                                // together.
                                                let joints = if std::env::var("OXI_AUTOSPACE2_DISABLE").is_err()
                                                {
                                                    let mut n = 0.0f32;
                                                    let mut prev: Option<char> = None;
                                                    for c in current_line_chars
                                                        .iter()
                                                        .chain(buf_chars.iter())
                                                        .map(|c| c.ch)
                                                        .chain(std::iter::once(ch))
                                                    {
                                                        if let Some(p) = prev {
                                                            if if std::env::var_os("OXI_CELL_AKI").is_some() { cell_aki_joint(p, c, para.style.auto_space_de, para.style.auto_space_dn) } else { is_cjk_latin_joint(p, c) } {
                                                                n += 1.0;
                                                            }
                                                        }
                                                        prev = Some(c);
                                                    }
                                                    n * 0.5
                                                        * s1175_autospace(
                                                            font_size,
                                                            cs,
                                                            self.balance_single_byte_double_byte_width
                                                                && run.style.fit_text.is_none() && !run.style.ruby_spread,
                                                        )
                                                } else {
                                                    0.0
                                                };
                                                // S1176: an ideographic space on a JUSTIFIED line
                                                // squeezes to a quarter of itself -- 0.75em of give.
                                                // It happens at compatibilityMode 15, where 約物 do
                                                // not compress at all, so this is the space-shrink
                                                // (S825/S1082's mechanism) and not the 約物 one.
                                                // Measured both ways: jc=both holds out to 0.748em
                                                // of shortfall, jc=left wraps the moment it is short.
                                                // Capped at one em: the sweep can only ever ask about ONE extra
                                                // character, so a shortfall never exceeds an em and "two spaces still
                                                // hold at 1.0em" is the most the measurement can say. Granting 0.75em
                                                // per space uncapped extrapolates far past that -- a form line with six
                                                // of them would get four and a half ems -- and that is what put the
                                                // tokumei documents 0.048 down on SSIM.
                                                let spaces = if s1176_space {
                                                    (0.75
                                                        * current_line_chars
                                                            .iter()
                                                            .chain(buf_chars.iter())
                                                            .filter(|c| c.ch == '\u{3000}')
                                                            .map(|c| c.natural_advance)
                                                            .sum::<f32>())
                                                    .min(font_size)
                                                } else {
                                                    0.0
                                                };
                                                // S1174 refined: one 0.5em for the line from
                                                // the closing marks however many there are, plus
                                                // 0.5em from EACH opening bracket. A line-final
                                                // closing mark is taken whatever the shortfall --
                                                // it hangs -- so it short-circuits the pool.
                                                // S1198 (2026-08-23, opt-out
                                                // `OXI_S1198_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): the COMBINATION rule for
                                                // a line that carries both classes of 約物.
                                                //
                                                // `_cw_yaku_class.py` swept one class at a time and
                                                // left the mixture undetermined; the implementation
                                                // guessed "0.5em if any type A, plus 0.5em per
                                                // opening bracket, uncapped". `_cw_yaku_mix.py`
                                                // (new) alternates the classes down a 30-character
                                                // line so one width sweep reports every (nA, nB)
                                                // combination, in BOTH orders. 8 arms x 3 counts,
                                                // all 24 land on:
                                                //
                                                //   last 約物 is type A -> 0.50em, whatever nA/nB
                                                //   last 約物 is type B -> min(1.00em,
                                                //                             0.5em*nB + 0.5em*[nA>0])
                                                //
                                                // i.e. the pool is CAPPED at one em, and a trailing
                                                // type-A mark collapses it to half an em. The
                                                // uncapped additive form was over by up to 0.5em on
                                                // six of the eight arms.
                                                //
                                                // ★ENVELOPE: measured at jc=both, compat 11, with
                                                // compressPunctuation -- which is where the corpus
                                                // specimens live. The cached artifacts behind the
                                                // ORIGINAL per-class table were built WITHOUT
                                                // compressPunctuation, and re-running the tool at
                                                // its own defaults (compat 15) silently measures a
                                                // different engine, so the arms are rebuilt here.
                                                let s1198 =
                                                    std::env::var("OXI_S1198_DISABLE").is_err();
                                                // S1208 (2026-08-24, opt-out `OXI_S1208_DISABLE`;
                                                // reachable only under the opt-in `OXI_YAKUCOMP`): the
                                                // flat half em is the pool a line gets when the character
                                                // it is squeezing in follows an ORDINARY one. When the
                                                // character immediately before it is a FULLWIDTH SPACE,
                                                // every mark on the line lends its own half em instead, up
                                                // to one and a half. And a fullwidth space lends NOTHING on
                                                // its own, so it stops counting as a mark at all.
                                                // DERIVED with `_pb_bodyyaku{,2,3,4,6}_gen.py` and confirmed
                                                // cell-side, arm for arm, with `_pb_cellyaku_gen.py` (Word
                                                // COM, the paragraph's right indent swept in 0.25pt steps,
                                                // MS Mincho 10.5pt, jc=both, compat 11; 19 arms, each one
                                                // monotone with a one-step flip bracket, body and cell
                                                // readings IDENTICAL):
                                                //     marks only, n = 1..4                 0.50em
                                                //     fullwidth spaces only, 1..4          0.00em
                                                //     n marks + a fullwidth space
                                                //       immediately before the last char   0.50em*min(n,3)
                                                //     the same space 1..6 characters back  0.50em
                                                //     a comma / period / closing bracket
                                                //       in that same position              0.50em
                                                // The render says the same thing glyph by glyph: at the
                                                // released widths EVERY mark carries its full half em; at
                                                // the flat width two of them share one half em between them.
                                                // tokyoshugyo p76's marked item is exactly this line, and
                                                // Word compresses BOTH of its commas by 3.9pt to keep the
                                                // last character on it, which a flat half em cannot pay for.
                                                // MEASURED BUT NOT IMPLEMENTED: an OPENING bracket in that
                                                // position releases the pool too AND adds its own half em
                                                // (1.50em for two marks). No corpus line reaches it, so it
                                                // is recorded here rather than coded.
                                                let s1208 =
                                                    std::env::var("OXI_S1208_DISABLE").is_err();
                                                let s1209 =
                                                    std::env::var("OXI_S1209_DISABLE").is_err();
                                                let mut s1216_n_a = 0.0f32;
                                                let mut s1216_sum_a = 0.0f32;
                                                let punctuation_unit = self.cell_punctuation_unit(para, font_size, grid_char_pitch, grid_char_cw_ratio);
                    let yaku = if s1174_yakucomp {
                                                    let mut has_a = false;
                                                    let mut b = 0.0f32;
                                                    let mut last_a = false;
                                                    let mut saw_mark = false;
                                                    let mut n_a = 0.0f32;
                                                    let mut n_marks = 0.0f32;
                                                    let mut last_ch: Option<char> = None;
                                                    for (i, c) in current_line_chars
                                                        .iter()
                                                        .chain(buf_chars.iter())
                                                        .enumerate()
                                                    {
                                                        last_ch = Some(c.ch);
                                                        if s1208 && c.ch == '\u{3000}' {
                                                            continue;
                                                        }
                                                        if kinsoku::cell_yaku_type_b(c.ch) {
                                                            saw_mark = true;
                                                            last_a = false;
                                                            // S1197 (opt-in `OXI_S1197`): a
                                                            // LINE-INITIAL opening bracket has
                                                            // nothing left to give -- JIS X 4051
                                                            // line-start processing already removed
                                                            // its left aki (the S757 principle).
                                                            // Held opt-in: S1198 subsumes the cases
                                                            // it was written for.
                                                            if i == 0
                                                                && std::env::var("OXI_S1197").is_ok()
                                                            {
                                                                continue;
                                                            }
                                                            b += 0.5 * c.natural_advance;
                                                            n_marks += 1.0;
                                                        } else if kinsoku::cell_yaku_type_a(c.ch) {
                                                            has_a = true;
                                                            saw_mark = true;
                                                            last_a = true;
                                                            n_a += 1.0;
                                                            n_marks += 1.0;
                                                            s1216_n_a += 1.0;
                                                            s1216_sum_a += c.natural_advance;
                                                        }
                                                    }
                                                    // S1209 (2026-08-24, opt-out `OXI_S1209_DISABLE`; reachable
                                                    // only under the opt-in `OXI_YAKUCOMP`): ONE rule for the
                                                    // whole pool, replacing S1174's additive form and S1198's
                                                    // last-mark form.
                                                    //
                                                    //   a line's mark credit is HALF AN EM, flat -- however many
                                                    //   marks it carries and whatever their class -- EXCEPT when
                                                    //   the character immediately before the one being squeezed
                                                    //   in is a FULLWIDTH SPACE or an OPENING BRACKET, and then
                                                    //   every mark on the line lends its own half em, capped at
                                                    //   one and a half. A fullwidth space is not itself a mark;
                                                    //   an opening bracket is.
                                                    //
                                                    // S1198 read 'the last mark is type B' where the truth is 'a
                                                    // releasing character stands next to the squeeze': its arms
                                                    // put a mark every third character, so the last mark was
                                                    // always within three of the line end and the two readings
                                                    // could not be told apart. 3a4f9fbe p80's 第５９条 cell is
                                                    // where they differ -- one comma and one opening bracket
                                                    // fourteen characters from the end -- and Word wraps the line
                                                    // that S1198's full em let Oxi pack one more character into.
                                                    // DERIVED with _pb_bodyyaku4/6/8_gen.py (35 arms, body AND
                                                    // cell, identical readings), then STATED AS A PREDICTION and
                                                    // tested against fresh arms in _pb_bodyyaku9_gen.py: nA 0..3
                                                    // x nB 0..2 x the releasing character in {ordinary, fullwidth
                                                    // space, opening bracket, period}, 40 arms, 40 of 40
                                                    // predicted exactly, each monotone with a one-step flip
                                                    // bracket. Word COM throughout, the paragraph's right indent
                                                    // swept in 0.25pt steps, MS Mincho 10.5pt, jc=both, compat 11.
                                                    let released = last_ch == Some('\u{3000}')
                                                        || last_ch.map_or(false, kinsoku::cell_yaku_type_b);
                                                    if !s1209 {
                                                        if s1208 && last_ch == Some('\u{3000}') {
                                                            (0.5 * punctuation_unit * n_a).min(1.5 * punctuation_unit) + b
                                                        } else if s1198 {
                                                            if !saw_mark {
                                                                0.0
                                                            } else if last_a {
                                                                0.5 * punctuation_unit
                                                            } else {
                                                                (b + if has_a { 0.5 * punctuation_unit } else { 0.0 })
                                                                    .min(font_size)
                                                            }
                                                        } else {
                                                            (if has_a { 0.5 * punctuation_unit } else { 0.0 }) + b
                                                        }
                                                    } else if released {
                                                        (0.5 * punctuation_unit * n_marks).min(1.5 * punctuation_unit)
                                                    } else if n_marks > 0.0 {
                                                        0.5 * punctuation_unit
                                                    } else {
                                                        0.0
                                                    }
                                                } else {
                                                    0.0
                                                };
                                                // ★Word does not spend the whole pool. Squeezed lines across d77a58 use
                                                // 8-72% of what is available; the lines it declines (34140b/04b88e's
                                                // 「（年度）」 cells) would have needed 85-98%. So the pool is a ceiling
                                                // with a fraction under it -- swept via OXI_YAKUK.
                                                // S1206 (2026-08-24, opt-out
                                                // `OXI_S1206_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): the CJK/Latin joint pool
                                                // does NOT add on top of the 約物 pool.
                                                //
                                                // `_cw_yaku_joint.py` puts both on one line and
                                                // reads the transition points. A line ending in an
                                                // ordinary character lends 0.502 / 0.508 / 0.503 em
                                                // with 約物 alone, and 0.502 / 0.511 / 0.385 / 0.504
                                                // with 約物 AND one or two joints -- the same half
                                                // em. Adding the terms is what let tokyoshugyo p30
                                                // take a 7.725pt overflow on 5.25 (約物) + 3.94
                                                // (three joints) where Word takes none.
                                                // S1207 (2026-08-24, opt-out
                                                // `OXI_S1207_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): a PROPORTIONAL Japanese
                                                // face gets NO 約物 credit at all.
                                                //
                                                // `_cw_yaku_joint.py` run twice over the same arm,
                                                // once per face (the sweep font had been hardcoded
                                                // to ＭＳ 明朝, so every 約物 measurement ever taken
                                                // here was a monospaced one). Lines ending in an
                                                // ordinary character:
                                                //   ＭＳ 明朝   0.502 / 0.508 / 0.503 em
                                                //   ＭＳ Ｐ明朝 0.008 / -0.001 / 0.004 / -0.005 em
                                                // i.e. half an em against nothing. That is what
                                                // 3a4f9fbe p20 needs -- a proportional cell where
                                                // Word refuses an overflow of only 0.212em.
                                                let s1207_proportional = std::env::var(
                                                    "OXI_S1207_DISABLE",
                                                )
                                                .is_err()
                                                    && (matches!(
                                                        cm.family.as_str(),
                                                        "MS PMincho" | "MS PGothic" | "HGPGothicM"
                                                    ) || ['\u{3001}', '\u{3002}'].iter().all(|&mark| {
                                                        let width = cm.char_width_em(mark);
                                                        width > 0.0 && width < 0.99
                                                    }));
                                                // S1216 (2026-08-25, opt-out
                                                // `OXI_S1216_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): a docGrid that COMPRESSES
                                                // (negative charSpace) leaves no 約物 credit -- the
                                                // aki is already spent by the grid.
                                                //
                                                // `_pb_yakudist_gen.py`, 576 arms: one 、 placed 1,
                                                // 2, 3, 5, 8, 10, 12 and 16 characters before the
                                                // squeeze, right indent swept in 0.25pt steps, in a
                                                // cell at ＭＳ 明朝 10.5pt.
                                                //     no character grid  0.500em at EVERY distance
                                                //     charSpace = +1453  0.524em at EVERY distance
                                                //     charSpace = -2714  **0.000em**, every distance
                                                // (The sweep also kills the "the pool decays with
                                                // distance" reading that b35123 first suggested --
                                                // distance does not enter it at all.)
                                                //
                                                // b35123fe8efc p2 is the case: charSpace=-2714, its
                                                // 、 ten characters back, Oxi granting 5.25pt and
                                                // packing one more character than Word.
                                                // The first cut of this (no snapToGrid test) made
                                                // b35123fe8efc WORSE -- p2 -0.0126 -> -0.0387 --
                                                // because it took the credit from the non-snapping
                                                // paragraphs too. Scoped, it gives +0.0017 there
                                                // and touches no other document.
                                                // REFINED (2026-08-25): the grid only spends the
                                                // aki of a paragraph that SNAPS to it. b35123fe8efc
                                                // carries both kinds on one page --
                                                // 「…接続している場合、不正ア」 snaps (Word gives no
                                                // credit) while 「…キャビネット等に保」 carries
                                                // `w:snapToGrid w:val="0"` (Word gives the half em).
                                                // The 576 arms all snap, which is why they read a
                                                // flat zero and the corpus refused it.
                                                let s1216_compressing_grid =
                                                    std::env::var("OXI_S1216_DISABLE").is_err()
                                                        && para.style.snap_to_grid
                                                        && match (grid_char_pitch, grid_char_cw_ratio)
                                                        {
                                                            (Some(p), Some(r)) if r > 0.0 => {
                                                                (p - p / r) < -0.01
                                                            }
                                                            _ => false,
                                                        };
                                                let pool_terms = if s1216_compressing_grid {
                                                    // S1216 v2 (2026-08-26): not zero -- the mark's
                                                    // credit is its CURRENT advance minus half an
                                                    // un-gridded em. The one formula reproduces
                                                    // every pool reading taken so far:
                                                    //   no grid   10.5  - 5.25 = 5.25 (S1209's 0.5em)
                                                    //   cs +1453  10.85 - 5.25 = 5.60 (poolorder 0.524em)
                                                    //   cs -2714  9.837 - 5.25 = 4.59 -- measured
                                                    //     T in (4.37, 4.62] on b35123's own row
                                                    //     (_pb_b35_ablate.py), and the SAME credit
                                                    //     splits its twin lines correctly: p1
                                                    //     overflow 0.96 <= 4.59 packs 39, p2
                                                    //     overflow 4.66 > 4.59 wraps to 38 -- the
                                                    //     v1 zero got p2 right and p1 wrong.
                                                    // (v1's zero came from the synthetic probe's
                                                    // fixed-layout cell; fixed genuinely lowers the
                                                    // credit (T ~2.75) and is rare in the corpus --
                                                    // not modeled yet.)
                                                    (s1216_sum_a - s1216_n_a * 0.5 * font_size).max(0.0)
                                                } else if s1207_proportional {
                                                    0.0
                                                } else if std::env::var("OXI_S1206_DISABLE").is_ok()
                                                {
                                                    yaku + joints + spaces
                                                } else {
                                                    yaku.max(joints + spaces)
                                                };
                                                let pool = pool_terms
                                                    * std::env::var("OXI_YAKUK")
                                                        .ok()
                                                        .and_then(|v| v.parse::<f32>().ok())
                                                        .unwrap_or(1.0);
                                                let pool = pool.max(self.modern_cell_punctuation_capacity(
                                                    para, &run.style,
                                                    current_line_chars.iter().chain(buf_chars.iter()).map(|c| c.ch),
                                                    font_size, ch, grid_char_pitch,
                                                ));
                                                let pool = LayoutEngine::cell_english_pair_capacity(para, &run.style, current_line_chars.iter().chain(buf_chars.iter()).map(|c| c.ch), font_size, pool);
                                                if std::env::var("OXI_DUMP_GLYPHW").is_ok() {
                                                    let want = std::env::var("OXI_DUMP_GLYPHW_TEXT")
                                                        .unwrap_or_default();
                                                    let ptext: String = para
                                                        .runs
                                                        .iter()
                                                        .flat_map(|r| r.text.chars())
                                                        .collect();
                                                    if want.is_empty() || ptext.contains(&want) {
                                                        let head: String = current_line_chars
                                                            .iter()
                                                            .chain(buf_chars.iter())
                                                            .map(|c| c.ch)
                                                            .take(6)
                                                            .collect();
                                                        let bs: String = current_line_chars
                                                            .iter()
                                                            .chain(buf_chars.iter())
                                                            .filter(|c| kinsoku::cell_yaku_type_b(c.ch))
                                                            .map(|c| format!("{}:{:.2}", c.ch, c.natural_advance))
                                                            .collect::<Vec<_>>()
                                                            .join(",");
                                                        eprintln!(
                                            "[POOL] ch={:?} yaku={:.3} joints={:.3} spaces={:.3} pool={:.3} head={:?} typeB=[{}]",
                                            ch, yaku, joints, spaces, pool, head, bs
                                        );
                                                    }
                                                }
                                                // ★Only where it was measured: the
                                                // 250/250 hang was swept at jc=both. Do not
                                                // extend an unconditional rule past its
                                                // envelope.
                                                // S1199 (2026-08-23, opt-out
                                                // `OXI_S1199_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): a RUN of closing marks
                                                // does not hang. The unconditional short-circuit
                                                // below came from a sweep that only ever put ONE
                                                // mark at the boundary.
                                                //
                                                // `_pb_hang2.py` layers the boundary by how many
                                                // closers form the group there (801-1301 widths,
                                                // 3 arms, ＭＳ 明朝 10.5pt, the specimen's compat):
                                                //
                                                //   one closer overflows   -> hangs   441/441
                                                //   two or more overflow   -> hangs     1/154,
                                                //                                       2/180,
                                                //                                       0/158
                                                //
                                                // tokyoshugyo's 「…手待時間」） is the second case:
                                                // Word pushes 間」）to the next line as a group
                                                // where Oxi fits 間 and hangs both marks.
                                                let s1199_run_of_closers =
                                                    std::env::var("OXI_S1199_DISABLE").is_err()
                                                        && s586_run_chars
                                                            .get(s586_ci + 1)
                                                            .map_or(false, |n| {
                                                                kinsoku::cell_yaku_type_a(*n)
                                                            });
                                                // S1205 (2026-08-24, opt-out
                                                // `OXI_S1205_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): the hang is CAPPED at half
                                                // the mark's own advance -- it lends its right-side
                                                // aki, not itself.
                                                //
                                                // The unconditional "take a line-final closing mark
                                                // whatever the shortfall" came from a reading of
                                                // 441/441 whose demand was computed at one em per
                                                // character, which inflates it. `_pb_tailhang.py`
                                                // measures the boundary directly, with the advances
                                                // read off Word's own PDF: the mark hangs up to
                                                // 0.517em of overflow and is pushed from 0.521
                                                // (0.527 / 0.532 with two characters after it), so
                                                // the threshold is half an em and whether the mark
                                                // ends the PARAGRAPH makes no difference.
                                                //
                                                // tokyoshugyo p30 is the case: its 。 overflows the
                                                // 389.69 budget by 9.31pt = 0.887em, so Word declines
                                                // to hang it and drops す。 to its own line.
                                                let s1205_hang_ok = std::env::var("OXI_S1205_DISABLE")
                                                    .is_ok()
                                                    || ((line_x + buf_w + cw) - effective_wrap)
                                                        <= 0.5 * cw
                                                            + if std::env::var_os("OXI_CJK_CELL_COMPRESSION_HANG").is_some() {
                                                                // Interior punctuation compression and the
                                                                // final mark's hanging allowance can coexist.
                                                                // The pool excludes the candidate mark.
                                                                pool
                                                            } else {
                                                                0.0
                                                            }
                                                            + 0.01;
                                                // S1213 (2026-08-24, opt-out
                                                // `OXI_S1213_DISABLE`; reachable only under the
                                                // opt-in `OXI_YAKUCOMP`): a line that carries a TAB
                                                // does not hang its last mark at all.
                                                //
                                                // `_pb_poolorder_gen.py` runs the same three arms
                                                // (36 characters, right indent swept in 0.25pt
                                                // steps) through five switches. The line-final mark
                                                // buys, over the mark-free control:
                                                //     body                        +1.00em (the
                                                //                                 whole advance)
                                                //     cell                        +0.50em
                                                //     cell, ＭＳ Ｐゴシック        +0.50 x 6.98pt
                                                //     cell + hanging indent       the same 3.5pt
                                                //     cell + marker AND TAB       +0.00 -- nothing
                                                // compat 11 and 15 read identically, and the pool a
                                                // DIFFERENT mark on the line lends (0.5em) survives
                                                // the tab -- only the hang dies.
                                                //
                                                // d77a58 p9 is the case: 「カ<tab>本利用ルールは…
                                                // あります。」 overflows its budget by 0.78pt, Oxi
                                                // hangs the 。 for 3.49pt of allowance and keeps
                                                // す。 on the line, Word drops them both. The budget
                                                // itself is already right on both sides -- the
                                                // column sweep in `_pb_d77a_budget.py` puts Word's
                                                // flip at +1.00pt and Oxi's own [CELLX] trace reads
                                                // first_line_wrap_w=415.60, the same number.
                                                //
                                                // ENVELOPE: measured with the tab LEADING the line
                                                // (a marker, then a tab, then the text), which is
                                                // the corpus shape. A tab in the middle of a line is
                                                // not covered by any arm.
                                                let s1213_tabbed_line = std::env::var(
                                                    "OXI_S1213_DISABLE",
                                                )
                                                .is_err()
                                                    && current_line_chars
                                                        .iter()
                                                        .chain(buf_chars.iter())
                                                        .any(|c| c.ch == char::from(9u8));
                                                if s1174_yakucomp
                                                    && kinsoku::cell_yaku_can_hang(ch)
                                                    && !s1199_run_of_closers
                                                    && !s1213_tabbed_line
                                                    && s1205_hang_ok
                                                    && matches!(
                                                        para.alignment,
                                                        Alignment::Justify | Alignment::Distribute
                                                    )
                                                {
                                                    false
                                                } else
                                                {
                                                    ((line_x + buf_w + cw) - effective_wrap) > pool
                                                }
                                            } else if legacy_cell_break
                                                && !s1174_yakucomp
                                                && would_overflow_natural
                                            {
                                                // ★S1174 SUPERSEDES this branch when it is on.
                                                // Both are the same shape -- a per-約物 break budget
                                                // -- so leaving both in place counts the pool twice:
                                                // the S1174 credit already widened `effective_wrap`,
                                                // and this then allows another `n_yak * cap` on top.
                                                // That is what put tokyoshugyo 78 paragraphs AHEAD
                                                // of Word (delta -1) under OXI_YAKUCOMP=1 while
                                                // CELLLAW alone left it at 1.0.
                                                // The caps differ because they were found different
                                                // ways: this one was tuned to ~2.5pt at 10.5 (0.238em)
                                                // against the OLD over-wide budget, the S1174 one is
                                                // 0.5em read off a controlled sweep. They cannot both
                                                // apply.
                                                // small-cap capacity: fit a (non-hangable) kanji iff its overflow ≤
                                                // Σ(small cap per compressible mid-line 約物). cap ~2.5pt@10.5 = the
                                                // derived kanji-fit budget, scaled by fs. Line-end 約物 hang via cell_bura.
                                                let n_yak = current_line_chars
                                                    .iter()
                                                    .chain(buf_chars.iter())
                                                    .filter(|c| {
                                                        matches!(
                                                            c.ch,
                                                            '、' | '。'
                                                                | '，'
                                                                | '．'
                                                                | '」'
                                                                | '』'
                                                                | '）'
                                                                | '】'
                                                                | '〕'
                                                                | '・'
                                                        )
                                                    })
                                                    .count()
                                                    as f32;
                                                // S721 (2026-07-03, default ON, opt-out OXI_S721_DISABLE):
                                                // PARAGRAPH-TAIL ORPHAN ELIMINATION — Word compresses mid-line
                                                // 約物 up to ~3.9pt each (vs the derived ~2.5 normal break cap)
                                                // when fitting the char removes a short (≤2-glyph) final line.
                                                // Word render-truth p76 ⑧火災等: «…をとり、␣␣␣␣に» — に is
                                                // the para's LAST glyph; Word compresses 、×2 to 6.60/6.63
                                                // (−3.90 each) to fit it on one line (avoiding a 1-char orphan
                                                // continuation); mid-para overflows keep the small cap (the
                                                // 変形-line flips at a FLAT cap ≥3.68 stay byte-identical —
                                                // the orphan gate IS the discriminator the flat sweep lacked).
                                                // Consistent with ikujidetail る。 (S595: demand ~11 > 3.9 →
                                                // Word wraps even at the tail) and the S590 derivation (normal
                                                // break compression rare/small).
                                                let s721_gpos = s586_run_offset + s586_ci;
                                                let s721_k: usize = std::env::var("OXI_S721_K")
                                                    .ok()
                                                    .and_then(|v| v.parse().ok())
                                                    .unwrap_or(1);
                                                let s721_tail =
                                                    s586_para_chars.len().saturating_sub(s721_gpos)
                                                        <= s721_k;
                                                let s721_cap = if s721_tail
                                                    && std::env::var("OXI_S721_DISABLE").is_err()
                                                {
                                                    std::env::var("OXI_S721_CAP")
                                                        .ok()
                                                        .and_then(|v| v.parse().ok())
                                                        .unwrap_or(3.9)
                                                } else {
                                                    legacy_cell_cap
                                                };
                                                let budget = n_yak * s721_cap * (font_size / 10.5);
                                                ((line_x + buf_w + cw) - effective_wrap) > budget
                                            } else if ((jc_gate_active
                                                && (run_has_neg_cs || cell_comp_active))
                                                || s466_oikomi)
                                                && would_overflow_natural
                                            {
                                                let ch_ctx =
                                                    crate::layout::jc_both_compress::CharContext {
                                                        ch,
                                                        natural_advance: cw,
                                                        font_size,
                                                    };
                                                let mut trial: Vec<
                                                    crate::layout::jc_both_compress::CharContext,
                                                > = Vec::with_capacity(
                                                    current_line_chars.len() + buf_chars.len() + 1,
                                                );
                                                trial.extend(current_line_chars.iter().cloned());
                                                trial.extend(buf_chars.iter().cloned());
                                                trial.push(ch_ctx);
                                                let r = crate::layout::jc_both_compress::compute_compression(
                                        &trial, effective_wrap, true,
                                    );
                                                !r.fits
                                            } else {
                                                // CELL half-em best-fit oikomi for cs=0 justified cells
                                                // ATTEMPTED + REVERTED (2026-06-22, OXI_CELLHE): tokyoshugyo's
                                                // 条文 boxes are align=Justify cs=0 → no 約物 compression →
                                                // over-wrap (+283 root). Half-em best-fit (oikomi iff overflow ≤
                                                // half-em + compute_compression.fits) gained tokyoshugyo
                                                // 0.8077→0.8141 (fixed only 14/283 — the rest are page-44-style
                                                // HEAVY 約物→opener over-wraps = S586 domain, not half-em) AND
                                                // preserved b837/b35/1636/ed025/1ec1/29dc6e but REGRESSED a1d6
                                                // 1.0→0.9930 (PASS→FAIL). Net negative + no n_pass gain (tks
                                                // stays FAIL). The cell +283 needs page-44 heavy compression +
                                                // per-line best-fit + a1d6-safe gate = multi-session. Reverted.
                                                would_overflow_natural
                                            };
                                            // S1201 (2026-08-23, default ON 2026-08-26 (opt-out `OXI_S1201_DISABLE`)): a character that
                                            // would DRAG a run of closing marks with it must fit
                                            // TOGETHER WITH THE WHOLE RUN. 」）cannot begin a line
                                            // (行頭禁則), so taking the character before them commits
                                            // the line to all three.
                                            //
                                            // `_pb_hang2.py`: when the character at the boundary is
                                            // followed by closing marks, Word takes it 3 times out of
                                            // 634 -- and S1199 measured that a run of two or more
                                            // never hangs, so there is no relief at the end either.
                                            // The specimen shows the same thing where the character
                                            // fits NATURALLY, which the squeeze-side rule does not
                                            // cover: tokyoshugyo p20's line 2 has 17.6pt of slack and
                                            // Word still pushes 間」）down, because 間+」+） is
                                            // 371.15 against a 367.6 budget.
                                            //
                                            // Scoped to a run of TWO OR MORE: a single trailing mark
                                            // hangs (441/441), so the character before it is free.
                                            //
                                            // ★The first cut keyed the run on `cell_yaku_type_a` and
                                            // collapsed tokyoshugyo to 0.5607. That class is a
                                            // COMPRESSION class and contains U+3000, which may begin a
                                            // line -- so «１年　　　　» and «…第１０４　　　» were torn
                                            // apart at the digit. Keying on what actually cannot start
                                            // a line makes it inert at default (10 documents unchanged)
                                            // and worth +0.0056 inside the cell bundle
                                            // (OXI_CELLLAW + OXI_YAKUCOMP + OXI_AUTOSPACE2:
                                            // tokyoshugyo 0.9906 -> 0.9962, its p20 box breaking
                                            // «…「手待時» / «間」）» exactly as Word does).
                                            let would_overflow = if would_overflow
                                                || std::env::var("OXI_S1201_DISABLE").is_ok()
                                                || (std::env::var_os("OXI_CJK_CELL_COMPRESSED_CLOSER_DISABLE").is_none() /* S1408 */
                                                    && self.compress_punctuation
                                                    && matches!(para.alignment, Alignment::Justify | Alignment::Distribute))
                                                || kinsoku::is_line_start_prohibited(ch)
                                            {
                                                would_overflow
                                            } else {
                                                let mut group = 0.0f32;
                                                let mut n = 0usize;
                                                let mut k = s586_ci + 1;
                                                while let Some(c) = s586_run_chars.get(k) {
                                                    // ★the run is what CANNOT START A LINE, which is
                                                    // the whole justification for dragging the
                                                    // preceding character down. `cell_yaku_type_a`
                                                    // is a COMPRESSION class and includes U+3000,
                                                    // which may begin a line: keying on it broke
                                                    // «１年　　　　» and «…第１０４　　　» apart at
                                                    // the digit. The probe only ever put closing
                                                    // brackets and punctuation there.
                                                    if !kinsoku::is_line_start_prohibited(*c) {
                                                        break;
                                                    }
                                                    group += self
                                                        .registry
                                                        .char_width_pt_with_fallback(
                                                            *c, font_size, cm,
                                                        );
                                                    n += 1;
                                                    k += 1;
                                                }
                                                // Three cases, layered the way `_pb_hang2.py`
                                                // layers the boundary:
                                                //   ch already overflows, any closers follow
                                                //       -> never squeezed (3/634)
                                                //   ch fits, TWO OR MORE closers follow
                                                //       -> the group must fit (they cannot hang: 1/154)
                                                //   ch fits, ONE closer follows
                                                //       -> that mark hangs (441/441), ch is free
                                                if n == 0 {
                                                    false
                                                } else if would_overflow_natural {
                                                    true
                                                } else if n >= 2 {
                                                    (line_x + buf_w + cw + group) > effective_wrap
                                                } else {
                                                    false
                                                }
                                            };
                                            // Permanent instrument (env-gated): the cell breaker's fit test for
                                            // a character near the wrap edge. OXI_DBG_CELLWRAP=<text prefix of
                                            // the paragraph> limits it to one paragraph.
                                            if let Some(pre) = std::env::var_os("OXI_DBG_CELLWRAP") {
                                                let head: String = para.runs.iter().flat_map(|r| r.text.chars()).take(40).collect();
                                                if head.starts_with(pre.to_string_lossy().as_ref())
                                                    && (line_x + buf_w + cw) > effective_wrap - 12.0
                                                {
                                                    eprintln!("[CELLWRAP] ch={:?} fs={:.2} line_x={:.2} buf={:?} buf_w={:.2} cw={:.3} wrap={:.2} nat={} ovf={}",
                                                        ch, font_size, line_x, buf, buf_w, cw, effective_wrap, would_overflow_natural, would_overflow);
                                                }
                                            }
                                            // PROPCELL OIKOMI (tokyoshugyo/d77a commentary boxes, 2026-06-23,
                                            // default ON, opt-out OXI_PROPCELL_DISABLE):
                                            // a jc=LEFT cell in a PROPORTIONAL CJK font (MS PMincho/PGothic/
                                            // HGPGothicM — the 参考/ガイドライン抜粋 commentary boxes) FORCE-FITS
                                            // an overflowing line-end 約物 (S421 oikomi is gated to cellmar/tab
                                            // cells, neither fires here) → it hangs the trailing 、/。 ~6.3pt past
                                            // the margin where Word OIKOMI's the cluster (pushes 約物+companion to
                                            // the next line). Word PDF (page-20 趣旨 box): «…ことか|ら、» — Word
                                            // breaks at か, Oxi packs «から、» (、 to x516.3) → 2 lines vs Word 3 →
                                            // content shifts up ~1 line → page-bottom over-fit → −1 page flips.
                                            // Enable the S421 oikomi when the 約物 overflows past OXI_PROPCELL_BOUND
                                            // (default 5.0pt ≈ Word's measured ~5.9pt max line-end 約物 hang).
                                            let propcell_font =
                                                font_family.as_deref().map_or(false, |f| {
                                                    f.contains("Ｐ明朝")
                                                        || f.contains("Ｐゴシック")
                                                        || f.contains("PMincho")
                                                        || f.contains("PGothic")
                                                });
                                            let propcell_oikomi =
                                                std::env::var("OXI_PROPCELL_DISABLE").is_err()
                                                    && !matches!(
                                                        para.alignment,
                                                        Alignment::Justify | Alignment::Distribute
                                                    )
                                                    && propcell_font;
                                            let propcell_over =
                                                (line_x + buf_w + cw) - effective_wrap;
                                            let propcell_bound: f32 =
                                                std::env::var("OXI_PROPCELL_BOUND")
                                                    .ok()
                                                    .and_then(|v| v.parse().ok())
                                                    .unwrap_or(5.0);
                                            // S720 (2026-07-03, default ON, opt-out OXI_S720_DISABLE): extend
                                            // the S643 propcell oikomi to JUSTIFIED proportional-font cell
                                            // paragraphs. S643 excluded Justify to protect the MONOSPACE
                                            // justified 条文 boxes — but the FONT gate already excludes those;
                                            // the justified+proportional combination (the（参考）ガイドライン
                                            // box's (ア)(イ)(ウ) items, ＭＳ Ｐ明朝, jc=both inherited) was
                                            // left in force-fit land. Word render-truth p20 (イ): the para-final
                                            // «間」）» — Word OIDASHI's the trailing «）» at overflow 2.6pt
                                            // (kinsoku cascade pulls 間」 down → L3 = 間」）), while Oxi
                                            // force-fit all three → the box lost a line → 特に crept onto p20
                                            // (wi=434). Word's natural Ｐ明朝 advances = Oxi's (±0.06, measured
                                            // unstretched last lines); the 2.6 overflow is real. Bound default
                                            // 1.0 (< the measured 2.6 oidashi; OXI_S720_BOUND to sweep) — the
                                            // jc=LEFT arm keeps its 5.0.
                                            let s720_just_prop = std::env::var("OXI_S720_DISABLE")
                                                .is_err()
                                                && propcell_font
                                                && matches!(
                                                    para.alignment,
                                                    Alignment::Justify | Alignment::Distribute
                                                )
                                                && propcell_over
                                                    > std::env::var("OXI_S720_BOUND")
                                                        .ok()
                                                        .and_then(|v| v.parse::<f32>().ok())
                                                        .unwrap_or(1.0);
                                            let would_overflow = LayoutEngine::fit_text_cell_overflow(
                                                &para.runs, run_idx, s586_ci,
                                                line_x + buf_w, effective_wrap,
                                            ).unwrap_or(would_overflow);
                                            if !is_space
                                                && would_overflow
                                                && !(current_line.is_empty() && buf.is_empty())
                                            {
                                                // Kinsoku: line-start-prohibited chars (）。、etc.) stay on current line
                                                if kinsoku::is_line_start_prohibited(ch)
                        && !(ch.is_ascii() && !self.doc_body_has_real_cjk
                            && std::env::var("OXI_ASCII_CELL_CLOSER_DISABLE").is_err())
                    {
                                                    // S421 (2026-05-29): kinsoku OIKOMI (押し下げ).
                                                    // The old behavior force-fit the prohibited char
                                                    // onto the current line (S409 bug) → ed025's cell
                                                    // （× × ×） rendered 1 line instead of Word's 2-line
                                                    // 5+2. Word pulls the preceding char down so the
                                                    // prohibited char is not alone at line start
                                                    // (COM-confirmed S420 on ）。、」). Mirrors the body
                                                    // path oikomi at mod.rs:6231-6254. Only pops from
                                                    // `buf` (current run's pending chars); falls back to
                                                    // the old force-fit when buf cannot supply the
                                                    // companion (empty / would empty the line).
                                                    // S421 SHIP: default ON (opt-out OXI_S421_DISABLE).
                                                    // Phase 1 53/55→54/55 (ed025 PASS 709/709),
                                                    // Phase 2 0.9647→0.9651, 3a4f unchanged 0.9757.
                                                    // S421b: restrict oikomi to S412 cells. Blanket
                                                    // oikomi catastrophically regressed 3a4f (score
                                                    // 0.79→0.20, 909 paras +1) and 34140b — those
                                                    // docs rely on the legacy force-fit / margin
                                                    // extension. Tying oikomi to the same
                                                    // discriminator as the S412 budget narrowing
                                                    // fires it ONLY on the 263 ed025+1ec1 cells where
                                                    // Word's narrowed budget forces the wrap.
                                                    // S713: the narrowed budget (cell_w - pads) forces
                                                    // wraps Word resolves by OIKOMI/OIDASHI, so the
                                                    // S421 oikomi ties to the same discriminator (the
                                                    // S421b precedent: oikomi follows the budget gate).
                                                    // tokyoshugyo (注) L2: 。 overflows 9.45pt > the
                                                    // compression budget -> Word pulls す down (L3 =
                                                    // す。); the old force-fit kept 。 on L2.
                                                    // CJK documents only (the probe's scope): reports__003862302b
                                                    // (Latin, compat 15) lost its fit when this fired on its cells.
                                                    let s1583_left = std::env::var_os("OXI_S1583_DISABLE").is_none()
                                                        && self.doc_body_has_real_cjk
                                                        && self.compat_mode >= 15
                                                        && self.compat_mode_explicit
                                                        && !matches!(para.alignment, Alignment::Justify | Alignment::Distribute);
                                                    if std::env::var("OXI_S421_DISABLE").is_err()
                                                        && (s412_cellmar_subtract
                                                            || (std::env::var("OXI_S443_DISABLE")
                                                                .is_err()
                                                                && p_first_line_indent < 0.0
                                                                && para_has_tab)
                                                            || (propcell_oikomi
                                                                && propcell_over > propcell_bound)
                                                            || s720_just_prop
                                                            || s713_cellmar_render
                                                            || cell_natural_line_end
                                                            // S1583 (2026-09-27, default ON, opt-out
                                                            // OXI_S1583_DISABLE): a compat-15 NON-justified
                                                            // cell paragraph pushes the character before a
                                                            // line-start-prohibited mark down with it
                                                            // (追い出し) instead of force-fitting the mark.
                                                            // `_pb_hang_bracket_gen.py` cell arms (3686tw,
                                                            // linesAndChars 350/1382, jc=left): 16 あ + 、/。
                                                            // /）/」 = 2 lines in Word, like 17 あ; jc=both
                                                            // keeps the mark. legal__0adfa250 p4
                                                            // 「□内職　□その他（　　　　　　）」.
                                                            || s1583_left)
                                                    {
                                                        let ch_ctx = crate::layout::jc_both_compress::CharContext { ch, natural_advance: cw, font_size };
                                                        let mut carry: Vec<crate::layout::jc_both_compress::CharContext> = vec![ch_ctx];
                                                        loop {
                                                            let head = carry[0].ch;
                                                            let tail = buf.chars().last();
                                                            let need =
                                                                kinsoku::is_line_start_prohibited(
                                                                    head,
                                                                ) || tail.map_or(
                                                                    false,
                                                                    kinsoku::is_line_end_prohibited,
                                                                // S1583: an ideographic space is no break point
                                                                // inside a bracket group -- Word moves the whole
                                                                // 「（　　　）」 (0adfa250 p2 「自宅・その他」 /
                                                                // 「（　　　）」).
                                                                ) || (s1583_left && head == '\u{3000}');
                                                            let remaining_on_line = buf
                                                                .chars()
                                                                .count()
                                                                + current_line
                                                                    .iter()
                                                                    .map(|f| f.0.chars().count())
                                                                    .sum::<usize>();
                                                            if !need || remaining_on_line <= 1 {
                                                                break;
                                                            }
                                                            if buf.pop().is_some() {
                                                                if let Some(pc) = buf_chars.pop() {
                                                                    buf_w -= pc.natural_advance;
                                                                    carry.insert(0, pc);
                                                                }
                                                            } else if std::env::var(
                                                                "OXI_S443_DISABLE",
                                                            )
                                                            .is_err()
                                                            {
                                                                // S443: when buf is empty, pop the companion
                                                                // from current_line (already-flushed chars).
                                                                // d77a J's overflowing 「。」 has its companion
                                                                // 「す」 in current_line, not buf — the S421
                                                                // buf-only oikomi could not reach it and
                                                                // force-fit instead. Pop from current_line_chars
                                                                // (per-char ctx) + trim the matching glyph off
                                                                // the last fragment's text/width.
                                                                if let Some(pc) =
                                                                    current_line_chars.pop()
                                                                {
                                                                    if let Some(frag) =
                                                                        current_line.last_mut()
                                                                    {
                                                                        if frag.0.pop().is_some() {
                                                                            frag.2 -=
                                                                                pc.natural_advance;
                                                                            if frag.0.is_empty() {
                                                                                current_line.pop();
                                                                            }
                                                                        }
                                                                    }
                                                                    carry.insert(0, pc);
                                                                } else {
                                                                    break;
                                                                }
                                                            } else {
                                                                break; // can't pop across run boundary here
                                                            }
                                                        }
                                                        if carry.len() >= 2 {
                                                            // Oikomi succeeded: flush remaining buf as
                                                            // line1 tail, push line1, seed line2 with carry.
                                                            if !buf.is_empty() {
                                                                current_line.push((
                                                                    buf.clone(),
                                                                    font_size,
                                                                    buf_w,
                                                                    bold,
                                                                    run.style.italic,
                                                                    run.style.underline,
                                                                    run.style
                                                                        .underline_style
                                                                        .clone(),
                                                                    run.style.strikethrough,
                                                                    font_family.clone(),
                                                                    run.style.color.clone(),
                                                                    run.style
                                                                        .highlight
                                                                        .clone()
                                                                        .or_else(|| {
                                                                            run.style
                                                                                .shading
                                                                                .clone()
                                                                        }),
                                                                    cs,
                                                                    run.style
                                                                        .text_scale
                                                                        .unwrap_or(100.0),
                                                                    std::mem::take(
                                                                        &mut s993_lrpb_pending,
                                                                    ),
                                                                    run.style.font_family_east_asia.clone(),
                                                                    run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                                    run.style.clone(),
                                                                ));
                                                                buf.clear();
                                                                buf_w = 0.0;
                                                                current_line_chars
                                                                    .extend(buf_chars.drain(..));
                                                            }
                                                            lines.push(std::mem::take(
                                                                &mut current_line,
                                                            ));
                                                            line_x = 0.0;
                                                            current_line_chars.clear();
                                                            is_first_line = false;
                                                            for c in &carry {
                                                                buf.push(c.ch);
                                                                buf_w += c.natural_advance;
                                                            }
                                                            buf_chars.extend(carry);
                                                            continue;
                                                        }
                                                        // else: fall through to force-fit below
                                                    }
                                                    // Force-fit (old behavior; also oikomi fallback):
                                                    // add to buffer and break AFTER this char.
                                                    buf.push(ch);
                                                    buf_w += cw;
                                                    buf_chars.push(crate::layout::jc_both_compress::CharContext {
                                            ch, natural_advance: cw, font_size,
                                        });
                                                    if !buf.is_empty() {
                                                        current_line.push((
                                                            buf.clone(),
                                                            font_size,
                                                            buf_w,
                                                            bold,
                                                            run.style.italic,
                                                            run.style.underline,
                                                            run.style.underline_style.clone(),
                                                            run.style.strikethrough,
                                                            font_family.clone(),
                                                            run.style.color.clone(),
                                                            run.style.highlight.clone().or_else(
                                                                || run.style.shading.clone(),
                                                            ),
                                                            cs,
                                                            run.style.text_scale.unwrap_or(100.0),
                                                            std::mem::take(&mut s993_lrpb_pending),
                                                            run.style.font_family_east_asia.clone(),
                                                            run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                            run.style.clone(),
                                                        ));
                                                        buf.clear();
                                                        buf_w = 0.0;
                                                        current_line_chars
                                                            .extend(buf_chars.drain(..));
                                                    }
                                                    lines.push(std::mem::take(&mut current_line));
                                                    line_x = 0.0;
                                                    current_line_chars.clear();
                                                    is_first_line = false;
                                                    continue;
                                                }
                                                // CELLWORD (2026-07-08, default ON, opt-out
                                                // OXI_CELLWORD_DISABLE): Latin WORD wrap in the cell
                                                // wrapper. The greedy per-char fill breaks English
                                                // mid-word («Activi|ty», «h|armed», «ris|k» —
                                                // uk_risk_assessment headers; Word/LibreOffice both
                                                // wrap at the word boundary; column geometry is
                                                // IDENTICAL across all three so the per-char break is
                                                // the sole divergence). When the overflowing char is
                                                // a Latin word char and the line tail is a Latin word
                                                // run, pull the WHOLE partial word down (S421-oikomi
                                                // carry mechanics, crossing run boundaries via
                                                // current_line like S443). Fall through to the plain
                                                // char-break when the line would empty (word longer
                                                // than the line = Word char-breaks too).
                                                // PURE-LATIN-paragraph scope (v1): the mixed CJK+Latin
                                                // cell case regressed the calibrated tokumei form
                                                // family (6514 −0.1096 / d4d126 −0.0965, the S559
                                                // balanced-compensation wall) — their embedded Latin
                                                // tokens' char-break behavior is part of the tuned
                                                // row-height balance. A paragraph with ANY CJK char
                                                // keeps the legacy per-char fill; pure-Latin
                                                // paragraphs (English forms) get the word wrap.
                                                // v2 (2026-07-12): in a LATIN document the gate uses
                                                // REAL-CJK (ideograph/kana) — kinsoku::is_cjk counts
                                                // General Punctuation (U+2010-2044: – ' ' " "), so an
                                                // English cell para with an EN DASH or curly quote
                                                // fell back to the per-char fill and split mid-word
                                                // («cha|rge», «ex|cept», «fr|aud» — uklocalspending
                                                // p9 exclusions table; Word wraps at the word boundary
                                                // and needs 1 more line per affected row = part of the
                                                // wp36+ −1 cascade). JP docs keep the v1 gate
                                                // byte-identical (the second disjunct requires
                                                // !doc_body_has_real_cjk).
                                                // S818: in a Latin document the token wrap engages for ANY
                                                // non-whitespace overflow char (an URL overflowing at '/',
                                                // a parenthesized token at '('), not just alphanumerics.
                                                let s818_cell = !self.doc_body_has_real_cjk
                                                    && std::env::var("OXI_S818_DISABLE").is_err();
                                                let cellword_ok = std::env::var(
                                                    "OXI_CELLWORD_DISABLE",
                                                )
                                                .is_err()
                                                    && (if s818_cell {
                                                        !is_break_space(ch) // S1241
                                                    } else {
                                                        ch.is_ascii_alphanumeric()
                                                    })
                                                    && (para.runs.iter().all(|r| {
                                                        !r.text.chars().any(kinsoku::is_cjk)
                                                    }) || (!self.doc_body_has_real_cjk
                                                        && para.runs.iter().all(|r| {
                                                            !r.text.chars().any(
                                                                kinsoku::is_cjk_ideograph_or_kana,
                                                            )
                                                        })));
                                                if cellword_ok && !cell_tab_char_pack
                                                    && !self.hinted_alphabet_break(ch, &run.style, &para.style) {
                                                    // Move a trailing token to a fresh line before splitting it.
                                                    // Within an overlong token, prefer hyphens; otherwise
                                                    // fill the line character by character. Slashes do not
                                                    // trigger backtracking. Keep the estimator in sync.
                                                    let p_alnum =
                                                        |c: char| c.is_ascii_alphanumeric();
                                                    // S1241: NBSP is not a break
                                                    // opportunity -> not a token boundary.
                                                    let p_hyphen = |c: char| !is_break_space(c) && !s1600_cell_break_after_dash(c);
                                                    let p_token = |c: char| !is_break_space(c);
                                                    let preds: &[&dyn Fn(char) -> bool] =
                                                        if s818_cell {
                                                            &[&p_hyphen, &p_token]
                                                        } else {
                                                            &[&p_alnum]
                                                        };
                                                    let mut consumed = false;
                                                    for pred in preds {
                                                        let ch_ctx = crate::layout::jc_both_compress::CharContext { ch, natural_advance: cw, font_size };
                                                        let mut carry: Vec<crate::layout::jc_both_compress::CharContext> = vec![ch_ctx];
                                                        let mut hit_boundary = false;
                                                        loop {
                                                            let tail =
                                                                buf.chars().last().or_else(|| {
                                                                    current_line.last().and_then(
                                                                        |f| f.0.chars().last(),
                                                                    )
                                                                });
                                                            let remaining_on_line = buf
                                                                .chars()
                                                                .count()
                                                                + current_line
                                                                    .iter()
                                                                    .map(|f| f.0.chars().count())
                                                                    .sum::<usize>();
                                                            let joined_space = s818_cell && tail == Some(' ')
                                                                && cell_nbsp_cluster_blocks_break(
                                                                    current_line_chars.iter().chain(buf_chars.iter()).map(|c| c.ch),
                                                                    carry.iter().map(|c| c.ch).chain(std::iter::once(ch)),
                                                                );
                                                            if !tail.map_or(false, |c| pred(c)) && !joined_space {
                                                                hit_boundary =
                                                                    remaining_on_line >= 1;
                                                                break;
                                                            }
                                                            if remaining_on_line <= 1 {
                                                                break;
                                                            }
                                                            if buf.pop().is_some() {
                                                                if let Some(pc) = buf_chars.pop() {
                                                                    buf_w -= pc.natural_advance;
                                                                    carry.insert(0, pc);
                                                                }
                                                            } else if let Some(pc) =
                                                                current_line_chars.pop()
                                                            {
                                                                if let Some(frag) =
                                                                    current_line.last_mut()
                                                                {
                                                                    if frag.0.pop().is_some() {
                                                                        frag.2 -=
                                                                            pc.natural_advance;
                                                                        if frag.0.is_empty() {
                                                                            current_line.pop();
                                                                        }
                                                                    }
                                                                }
                                                                carry.insert(0, pc);
                                                            } else {
                                                                break;
                                                            }
                                                        }
                                                        // A legal hyphen can already be the last character
                                                        // on the line, so only the overflowing character moves.
                                                        let after_hyphen = buf.chars().last().or_else(||
                                                            current_line.last().and_then(|f| f.0.chars().last())).map_or(false, s1600_cell_break_after_dash);
                                                        if hit_boundary && (carry.len() >= 2 || (s818_cell && after_hyphen)) {
                                                            if !buf.is_empty() {
                                                                current_line.push((
                                                                    buf.clone(),
                                                                    font_size,
                                                                    buf_w,
                                                                    bold,
                                                                    run.style.italic,
                                                                    run.style.underline,
                                                                    run.style
                                                                        .underline_style
                                                                        .clone(),
                                                                    run.style.strikethrough,
                                                                    font_family.clone(),
                                                                    run.style.color.clone(),
                                                                    run.style
                                                                        .highlight
                                                                        .clone()
                                                                        .or_else(|| {
                                                                            run.style
                                                                                .shading
                                                                                .clone()
                                                                        }),
                                                                    cs,
                                                                    run.style
                                                                        .text_scale
                                                                        .unwrap_or(100.0),
                                                                    std::mem::take(
                                                                        &mut s993_lrpb_pending,
                                                                    ),
                                                                    run.style.font_family_east_asia.clone(),
                                                                    run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                                    run.style.clone(),
                                                                ));
                                                                buf.clear();
                                                                buf_w = 0.0;
                                                                current_line_chars
                                                                    .extend(buf_chars.drain(..));
                                                            }
                                                            lines.push(std::mem::take(
                                                                &mut current_line,
                                                            ));
                                                            line_x = 0.0;
                                                            current_line_chars.clear();
                                                            is_first_line = false;
                                                            for c in &carry {
                                                                buf.push(c.ch);
                                                                buf_w += c.natural_advance;
                                                            }
                                                            buf_chars.extend(carry);
                                                            consumed = true;
                                                            break;
                                                        }
                                                        // un-carry: restore anything popped (carry[..len-1],
                                                        // in reading order — popping drained buf's TAIL then
                                                        // current_line's tail, so appending keeps order) back
                                                        // onto buf so the next stage / plain wrap sees the
                                                        // original text (reachable when the word fills the
                                                        // whole line).
                                                        if carry.len() >= 2 {
                                                            let n = carry.len() - 1;
                                                            for c in carry.drain(..n) {
                                                                buf.push(c.ch);
                                                                buf_w += c.natural_advance;
                                                                buf_chars.push(c);
                                                            }
                                                        }
                                                    }
                                                    if consumed {
                                                        prev_char_emitted = Some(ch);
                prev_char_gap = Some(self.natural_autospace_after(ch, &run.style, &para.style, font_size, cs));
                                                        prev_char_ruby = run.style.ruby_field;
                                                        continue;
                                                    }
                                                }
                                                // Flush buffer to current line, then wrap
                                                if !buf.is_empty() {
                                                    current_line.push((
                                                        buf.clone(),
                                                        font_size,
                                                        buf_w,
                                                        bold,
                                                        run.style.italic,
                                                        run.style.underline,
                                                        run.style.underline_style.clone(),
                                                        run.style.strikethrough,
                                                        font_family.clone(),
                                                        run.style.color.clone(),
                                                        run.style
                                                            .highlight
                                                            .clone()
                                                            .or_else(|| run.style.shading.clone()),
                                                        cs,
                                                        run.style.text_scale.unwrap_or(100.0),
                                                        std::mem::take(&mut s993_lrpb_pending),
                                                        run.style.font_family_east_asia.clone(),
                                                        run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                        run.style.clone(),
                                                    ));
                                                    buf.clear();
                                                    buf_w = 0.0;
                                                    current_line_chars.extend(buf_chars.drain(..));
                                                }
                                                lines.push(std::mem::take(&mut current_line));
                                                line_x = 0.0;
                                                current_line_chars.clear();
                                                is_first_line = false;
                                            }
                                            // S497 (2026-06-05, SHIP default-on, opt-out OXI_S497_DISABLE):
                                            // a line-start-prohibited char (）。、」 etc.) must NEVER be the
                                            // first char of a cell line (kinsoku 行頭禁則). When the preceding
                                            // char force-fit + broke the line, the next prohibited char lands
                                            // alone at the head of a new line (15076df y408: the closing ） on
                                            // its own line). Word keeps it on the previous line, hanging past
                                            // the wrap limit (burasagari). Pull it back onto the previous
                                            // flushed line. GATE: Phase-1 54/55 with ZERO pagination change
                                            // (no PASS<->FAIL, no score change on any of 55 docs); SSIM
                                            // +0.0027 on 15076df, byte-identical on the rest of the corpus
                                            // (the prohibited-char-starts-cell-line case is rare — only
                                            // 15076df in a 120-doc scan + the 12-doc tokumei/control sample,
                                            // 3a4f/ed025c tombstones unchanged). lib 142/0/6.
                                            if std::env::var("OXI_S497_DISABLE").is_err()
                                                && buf.is_empty()
                                                && current_line.is_empty()
                                                && !lines.is_empty()
                                                && kinsoku::is_line_start_prohibited(ch)
                                            {
                                                if let Some(last) = lines.last_mut() {
                                                    // S993: a kinsoku-prohibited char joins the
                                                    // PREVIOUS line — never the LRPB anchor. Pass
                                                    // false so s993_lrpb_pending survives to the
                                                    // next real fragment.
                                                    last.push((
                                                        char_to_string(ch),
                                                        font_size,
                                                        cw,
                                                        bold,
                                                        run.style.italic,
                                                        run.style.underline,
                                                        run.style.underline_style.clone(),
                                                        run.style.strikethrough,
                                                        font_family.clone(),
                                                        run.style.color.clone(),
                                                        run.style
                                                            .highlight
                                                            .clone()
                                                            .or_else(|| run.style.shading.clone()),
                                                        cs,
                                                        run.style.text_scale.unwrap_or(100.0),
                                                        false,
                                                        run.style.font_family_east_asia.clone(),
                                                        run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                        run.style.clone(),
                                                    ));
                                                    prev_char_emitted = Some(ch);
                prev_char_gap = Some(self.natural_autospace_after(ch, &run.style, &para.style, font_size, cs));
                                                    prev_char_ruby = run.style.ruby_field;
                                                    continue;
                                                }
                                            }
                                            buf.push(ch);
                                            buf_w += cw;
                                            buf_chars.push(
                                                crate::layout::jc_both_compress::CharContext {
                                                    ch,
                                                    natural_advance: cw,
                                                    font_size,
                                                },
                                            );
                                            prev_char_emitted = Some(ch);
                prev_char_gap = Some(self.natural_autospace_after(ch, &run.style, &para.style, font_size, cs));
                                            prev_char_ruby = run.style.ruby_field;
                                        }
                                        if !buf.is_empty() {
                                            current_line.push((
                                                buf,
                                                font_size,
                                                buf_w,
                                                bold,
                                                run.style.italic,
                                                run.style.underline,
                                                run.style.underline_style.clone(),
                                                run.style.strikethrough,
                                                font_family,
                                                run.style.color.clone(),
                                                run.style
                                                    .highlight
                                                    .clone()
                                                    .or_else(|| run.style.shading.clone()),
                                                cs,
                                                run.style.text_scale.unwrap_or(100.0),
                                                std::mem::take(&mut s993_lrpb_pending),
                                                run.style.font_family_east_asia.clone(),
                                                run.ruby.is_some(),  // S1312: this fragment's run carries ruby
                                                run.style.clone(),
                                            ));
                                            line_x += buf_w;
                                            current_line_chars.extend(buf_chars.drain(..));
                                        }
                                        s586_run_offset += s586_run_chars.len();
                                    }
                                    if !current_line.is_empty()
                                        || (s1169_trailing_break
                                            && std::env::var("OXI_S1169_DISABLE").is_err())
                                    {
                                        lines.push(current_line);
                                    }

                                    if lines.is_empty() {
                                        // Day 33 part 9: when pPr/rPr explicitly sets font size on an empty
                                        // paragraph in a cell, use it. Pre-fix code always used
                                        // self.default_font_size (10.5pt), inflating cell content height
                                        // for fs=8pt cells (bd90b00 table 0 rows 1, 4: +2.1pt per empty
                                        // cell → row +2.25pt → table exit +4pt drift → 備考 overflow
                                        // = Class A FAIL root cause). Narrow fix: only override when
                                        // ppr_rpr.font_size is Some, leaving the default-fallback path
                                        // (664c38001b40 form-cells) unchanged.
                                        //
                                        // S403 (2026-05-28) verified empty-cell-paragraph height
                                        // is NOT the source of ed025 Phase 1 -1 delta. Diagnostic
                                        // (OXI_S403_DUMP_CELL) confirmed Oxi uses lh=18.0 for
                                        // every cell paragraph (matches Word's per-gap median
                                        // 18.0pt across 33 gaps).
                                        //
                                        // S404 (2026-05-28) verified PARSER is also CORRECT.
                                        // Oxi IR for T16 row1 cell 2 has exactly 98 paragraphs
                                        // (60 TEXT + 38 EMPTY) — matches XML count and matches
                                        // Word's 98 per-cell distribution (Word p13: 18 TEXT/
                                        // 16 EMPTY = 34; p14: 34/5 = 39; p15: 8/17 = 25).
                                        // Per-page TOTAL paragraph count also matches (34/39/25
                                        // in both). Only the TEXT placement differs by ONE:
                                        //   Word: p13 18T/16E, p14 34T/5E
                                        //   Oxi:  p13 19T/15E, p14 33T/6E
                                        // → 1 TEXT that Word puts on p14 is on Oxi p13, swapped
                                        //   with 1 EMPTY going the other direction.
                                        //
                                        // The actual root cause is sub-pt height-accumulation
                                        // drift across 34 cell paragraphs that lands one
                                        // boundary paragraph on the wrong side of page_bottom.
                                        // Both per-paragraph height (18.0pt) and total paragraph
                                        // count (98) match Word; the layout sums must differ by
                                        // <0.5pt per paragraph and accumulate to >18pt at the
                                        // boundary. Needs per-paragraph y-trace COM measurement
                                        // vs Oxi to locate the drift source.
                                        let (pprrpr_fs, empty_lh) = self.cell_mark_line_height(
                                            para, effective_line_spacing, effective_line_rule, row_line_pitch);
                                        if std::env::var("OXI_DBG_EMPTYLH").is_ok() {
                                            eprintln!("[EMPTYLH] row={} cell={} lh={:.2} pprrpr_fs={:?} default_fs={} ls={:?} rule={:?} snap={}",
                                    row_idx, cell_idx, empty_lh, pprrpr_fs, self.default_font_size,
                                    effective_line_spacing, effective_line_rule, para.style.snap_to_grid);
                                        }
                                        // S428 (2026-05-29): emit a zero-glyph Text element for the
                                        // empty cell paragraph so the row-split / re-anchor logic
                                        // (mod.rs ~9215) treats it as a real line box. Without an
                                        // element, an empty paragraph that falls at a mid-cell page
                                        // boundary is invisible to the split: the re-anchor snaps the
                                        // first VISIBLE overflow text to page_top, dropping the empty
                                        // line's height and shifting the whole continuation up ~1 line
                                        // (e3c545 page 4: cell_para 10 empty between cpi 9/11 → all of
                                        // page 4 rendered ~12pt too high). Both renderers skip empty
                                        // text (GDI TextOutW of "" draws nothing; DWrite early-returns),
                                        // and both phase gates exclude empty paragraphs (pagination_diff
                                        // MIN_MATCH_LEN, dml_diff `if not text`), so this only affects
                                        // the split's positioning of NON-empty content. Opt-out:
                                        // OXI_S428_DISABLE.
                                        // Empty paragraph marks need an 18pt usable lane beside
                                        // a cell float; use the same obstacle frame as estimation.
                                        let mut empty_x = cell_x + pad_l;
                                        if std::env::var_os("OXI_CELL_EMPTY_FLOAT_DISABLE").is_none() {
                                            if let Some(wrap) = cell_float_wrap.as_ref() {
                                                let mut empty_wrap = wrap.clone();
                                                empty_wrap.heights = vec![empty_lh];
                                                // A zero-glyph mark uses the logical padded lane,
                                                // not the visible text clip inset by half a border.
                                                let logical_lane = std::env::var_os("OXI_CELL_EMPTY_LOGICAL_LANE_DISABLE").is_none();
                                                let empty_base = (cell_w - pad_l - pad_r).max(0.0);
                                                let empty_width = (empty_base - p_indent_left - p_indent_right).max(0.0);
                                                let empty_first_width = if p_first_line_indent < 0.0 {
                                                    (empty_base - (p_indent_left + p_first_line_indent).max(0.0)
                                                        - p_indent_right).max(0.0)
                                                } else { (empty_width - p_first_line_indent).max(0.0) };
                                                let frame = empty_wrap.frame_with_minimum(0,
                                                    if logical_lane { empty_width } else { wrap_w },
                                                    if logical_lane { empty_first_width } else { first_line_wrap_w },
                                                    p_indent_left, p_first_line_indent, 18.0);
                                                content_h += frame.gap;
                                                empty_x += frame.left;
                                            }
                                        }
                                        let is_interior_empty = last_content_block_pos
                                            .map_or(false, |last| block_pos < last);
                                        if (is_interior_empty || std::env::var_os("OXI_CELL_EMPTY_LINES").is_some())
                                            && std::env::var("OXI_S428_DISABLE").is_err()
                                        {
                                            let mut empty_el = LayoutElement::new(
                                                empty_x,
                                                content_h,
                                                0.0,
                                                empty_lh,
                                                LayoutContent::Text {
                                                    text: String::new(),
                                                    font_size: pprrpr_fs
                                                        .unwrap_or(self.default_font_size),
                                                    font_family: None,
                                                    bold: false,
                                                    italic: false,
                                                    underline: false,
                                                    underline_style: None,
                                                    strikethrough: false,
                                                    double_strikethrough: false,
                                                    color: None,
                                                    highlight: None,
                                                    character_spacing: 0.0,
                                                    field_type: None,
                                                    text_scale: 100.0,
                                                    is_vertical: false,
                                                    effects: TextEffects::default(),
                                                },
                                            );
                                            empty_el.paragraph_index = block_idx;
                                            empty_el.cell_paragraph_index = Some(cell_para_counter);
                                            empty_el.cell_row_index = Some(row_idx);
                                            empty_el.cell_col_index = Some(cell_idx);
                                            cell_elements.push(empty_el);
                                        }
                                        content_h += empty_lh;
                                    }

                                    // S1242 (2026-08-27, default ON, opt-out OXI_S1242_DISABLE):
                                    // a JUSTIFIED cell's Latin text arrives as ONE fragment per
                                    // line, so the word-space slack distribution below has no
                                    // space fragments to widen — jc=both table cells rendered
                                    // ragged-left (administrative__00018048's bullet cells; the
                                    // body path splits words and justifies fine). Split each
                                    // multi-word fragment at spaces into word/space pieces whose
                                    // widths are NORMALIZED to the fragment's measured width
                                    // (interior positions approximate, line edges exact).
                                    // Sentinel fragments (F8FE/F8FF) and space-free text pass
                                    // through untouched; non-justified paragraphs are skipped.
                                    if (para.alignment == Alignment::Justify
                                        || para.alignment == Alignment::Distribute)
                                        && std::env::var("OXI_S1242_DISABLE").is_err()
                                    {
                                        let fm = self.registry.default_metrics();
                                        for line in lines.iter_mut() {
                                            if !s1082_cell_shrink && line.iter().any(|t| {
                                                !t.0.is_empty() && t.0.trim().is_empty()
                                            }) {
                                                continue; // already has space fragments
                                            }
                                            if !line.iter().any(|t| {
                                                !t.0.starts_with('\u{F8FD}')
                                                    && !t.0.starts_with('\u{F8FE}')
                                                    && !t.0.starts_with('\u{F8FF}')
                                                    && (if s1082_cell_shrink { t.0.as_str() } else { t.0.trim() }).contains(' ')
                                            }) {
                                                continue;
                                            }
                                            let mut out = Vec::with_capacity(line.len() * 2);
                                            for frag in line.drain(..) {
                                                let (text, fs, tw) =
                                                    (frag.0.clone(), frag.1, frag.2);
                                                if text.starts_with('\u{F8FD}')
                                                    || text.starts_with('\u{F8FE}')
                                                    || text.starts_with('\u{F8FF}')
                                                    || !(if s1082_cell_shrink { text.as_str() } else { text.trim() }).contains(' ')
                                                {
                                                    out.push(frag);
                                                    continue;
                                                }
                                                // Alternating word/space pieces.
                                                let mut pieces: Vec<String> = Vec::new();
                                                let mut cur = String::new();
                                                let mut cur_is_space: Option<bool> = None;
                                                for ch in text.chars() {
                                                    let sp = ch == ' ';
                                                    if cur_is_space != Some(sp)
                                                        && !cur.is_empty()
                                                    {
                                                        pieces.push(std::mem::take(&mut cur));
                                                    }
                                                    cur_is_space = Some(sp);
                                                    cur.push(ch);
                                                }
                                                if !cur.is_empty() {
                                                    pieces.push(cur);
                                                }
                                                let raw: Vec<f32> = pieces
                                                    .iter()
                                                    .map(|p| {
                                                        p.chars()
                                                            .map(|c| {
                                                                self.registry
                                                                    .char_width_pt_with_fallback(
                                                                        c, fs, &fm,
                                                                    )
                                                            })
                                                            .sum::<f32>()
                                                    })
                                                    .collect();
                                                let raw_sum: f32 = raw.iter().sum();
                                                let scale = if raw_sum > 0.01 {
                                                    tw / raw_sum
                                                } else {
                                                    1.0
                                                };
                                                for (p, rw) in
                                                    pieces.into_iter().zip(raw.into_iter())
                                                {
                                                    let mut nf = frag.clone();
                                                    nf.0 = p;
                                                    nf.2 = rw * scale;
                                                    out.push(nf);
                                                }
                                            }
                                            *line = out;
                                        }
                                    }
                                    let total_lines = lines.len();
                                    let mut float_row_height = 0.0_f32;
                                    // S1626: the paragraph's ruby GROUPS, consumed in order as
                                    // their base fragments are laid out (index, base chars seen).
                                    // A ruby whose base runs differ in formatting (forms__01c5a769:
                                    // spacing 150 on the first base char, 30 on the second) parses
                                    // into consecutive runs that each carry the same ruby; they are
                                    // ONE annotation.
                                    let mut s1626_rubies: Vec<Vec<&Run>> = Vec::new();
                                    for r in para.runs.iter() {
                                        let Some(rb) = r.ruby.as_ref() else { continue };
                                        let same = s1626_rubies.last().and_then(|g| g.last()).and_then(|l| l.ruby.as_ref())
                                            .map_or(false, |prev| prev.text == rb.text && prev.base == rb.base);
                                        let adjacent = s1626_rubies.last().and_then(|g| g.last())
                                            .map_or(false, |l| std::ptr::eq(*l, &para.runs[para.runs.iter().position(|x| std::ptr::eq(x, r)).unwrap_or(0).saturating_sub(1)]));
                                        if same && adjacent {
                                            s1626_rubies.last_mut().unwrap().push(r);
                                        } else {
                                            s1626_rubies.push(vec![r]);
                                        }
                                    }
                                    let mut s1626_ri = 0usize;
                                    let mut s1626_seen = 0usize;
                                    for (line_idx, line) in lines.iter().enumerate() {
                                        if hidden_final_mark && line_idx + 1 == lines.len() && line.is_empty() {
                                            continue;
                                        }
                                        let float_frame = cell_float_wrap.as_ref().map(|wrap|
                                            wrap.frame(line_idx, wrap_w, first_line_wrap_w, p_indent_left, p_first_line_indent));
                                        if let Some(frame) = float_frame { content_h += frame.gap; }
                                        // CELLLINE instrument: reliable per-LINE cell text + char count.
                                        // OXI_DUMP_CELLLINE. Localizes cell wrap over-fit (Oxi chars/line
                                        // vs Word PDF) — the GDI dump is per-RUN for these cells.
                                        if std::env::var("OXI_DUMP_CELLLINE").is_ok() {
                                            let lt: String =
                                                line.iter().map(|t| t.0.as_str()).collect();
                                            if !lt.trim().is_empty() {
                                                eprintln!("[CELLLINE] li={} cy={:.1} nc={} wrapw={:.1} «{}»",
                                        line_idx, content_h, lt.chars().count(), wrap_w,
                                        lt.chars().take(44).collect::<String>());
                                            }
                                        }
                                        // Clip content that overflows exact row height
                                        if exact_cell_limit.is_some_and(|limit| content_h + pad_t >= limit) {
                                            break;
                                        }
                                        // Line height = max of all runs in line (in_table_cell=true: no default font minimum)
                                        // OXI_CELLPAIR: whitespace-only fragments do NOT contribute to the
                                        // line height (b35123 row2: the note's leading sz-21 「　」 run must
                                        // not lift the fs-9 line from 11.7 to 13.5 — Word row arithmetic
                                        // 17.5+11.7+11.7+3.5 = 44.4 ≈ measured 44.6; with 13.5 it would be
                                        // 46.4 = Oxi's wrong row). Word sizes the line by its INK runs.
                                        let cellpair_ws = self.cellpair_active();
                                        let ignore_ascii_spaces = !self.doc_body_has_real_cjk
                                            && line.iter().any(|(t, ..)| !t.trim().is_empty());
                                        let tab_mark_line = !self.doc_body_has_real_cjk
                                            && !explicit_break_lines.contains(&line_idx)
                                            && line.iter().any(|f| f.0.contains('\t'))
                                            && line.iter().all(|f| f.0.chars().all(|c| c == '\t'));
                                        let tab_mark = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                        let tab_mark_fs = self.resolve_font_size(&tab_mark, &para.style);
                                        let tab_mark_metrics = &*self.metrics_for_para_mark_g(&tab_mark, &para.style, true);
                                        let mut lh: f32 = line
                                            .iter()
                                            .filter(|(t, ..)| !cellpair_ws || !t.trim().is_empty())
                                            .filter(|(t, ..)| !ignore_ascii_spaces || t.is_empty()
                                                || !t.chars().all(|c| c == ' ' || c == '\t'))
                                            .map(
                                                |(
                                                    _text,
                                                    fs,
                                                    _,
                                                    s1629_bold,
                                                    s1629_italic,
                                                    _,
                                                    _,
                                                    _,
                                                    font_family,
                                                    _,
                                                    _,
                                                    _,
                                                    _,
                                                    _,
                                                    _,
                                                    _, // S1312 ruby flag
                                                    _source_style,
                                                )| {
                                                    // Cell lines use the actual requested face, as
                                                    // body lines and the cell height pre-pass do.
                                                    // A compensating excess in surrounding flow
                                                    // cannot justify sizing bold text as regular.
                                                    // Complex-script text takes the same line box as in
                                                    // the body (cs face if installed, else Nirmala UI);
                                                    // the resolved family is the Latin slot when the cs
                                                    // face is missing (igrsup_md_v4: Word 16.2, TNR 13.8).
                                                    let deva_cell = std::env::var_os("OXI_INDIA_CELL_DEVA_DISABLE").is_none()
                                                        && _text.chars().any(crate::font::is_complex_script);
                                                    let metrics = &*if deva_cell {
                                                        self.metrics_for_text(_text, _source_style, &para.style)
                                                    } else {
                                                        match font_family.as_deref() {
                                                            Some(ff) => self.registry.get_with_style(ff, *s1629_bold, *s1629_italic),
                                                            None => self.registry.default_metrics(),
                                                        }
                                                    };
                                                    // S1119 cells (measured 2026-08-14,
                                                    // `_pb_symline_gen.py ... cell`): Word applies
                                                    // the SAME fallback inside a table cell. The
                                                    // deltas against each font's control arm are
                                                    // identical to the body arms to 0.001pt —
                                                    // Arial ballot +3.281 (body +3.281), Calibri
                                                    // black-square −1.594 (body −1.593), Calibri
                                                    // diamond −0.938 (body −0.937). Only the BASE
                                                    // differs (a cell line carries no external
                                                    // leading, ~1.5pt lower). This site was left
                                                    // unwired at first ship precisely because that
                                                    // was unmeasured; the cell variant settles it.
                                                    // ★Apply the face as a DELTA, not a
                                                    // substitution. Oxi's cell model already
                                                    // matches Word for the run font (Calibri 18pt
                                                    // control: Oxi 20.447 vs Word 20.438), but the
                                                    // substituted face has no GDI cell entry and
                                                    // falls to a different branch, so swapping the
                                                    // metrics wholesale moved Calibri black-square
                                                    // only −0.223 where Word moves −1.594.
                                                    // Word's per-face drop IS the plain natural
                                                    // difference, verified against the Calibri
                                                    // control at 18pt (±0.06, the rounding class):
                                                    //   Cambria Math (typo)  −0.928 vs −0.938
                                                    //   Courier New  (win)   −1.641 vs −1.594
                                                    //   Segoe UI Sym (win)   +1.911 vs +1.968
                                                    let base = self.line_height_inner(
                                                        *fs,
                                                        effective_line_spacing,
                                                        effective_line_rule,
                                                        metrics,
                                                        para.style.snap_to_grid,
                                                        row_line_pitch,
                                                        true,
                                                    );
                                                    let base = if matches!(effective_line_rule, None | Some("auto"))
                                                        && matches!(metrics.family.as_str(), "Symbol" | "Wingdings")
                                                        && _text.chars().any(|c| matches!(c as u32, 0xF000..=0xF0FF))
                                                    {
                                                        metrics.natural_line_height_hhea(*fs) * effective_line_spacing.unwrap_or(1.0)
                                                    } else { base };
                                                    match self.s1119_run_face(_text, metrics) {
                                                        Some(fb) => {
                                                            base + fb.natural_line_height_hhea(*fs)
                                                                - metrics.natural_line_height_hhea(*fs)
                                                        }
                                                        None => base,
                                                    }
                                                },
                                            )
                                            .fold(0.0_f32, f32::max);
                                        if tab_mark_line {
                                            lh = self.line_height_inner(tab_mark_fs, effective_line_spacing,
                                                effective_line_rule, tab_mark_metrics, para.style.snap_to_grid,
                                                row_line_pitch, true);
                                        }
                                        if effective_line_rule != Some("exact")
                                            && line.iter().any(|f| f.0.chars().any(|c| matches!(c as u32, 0xF000..=0xF0FF))
                                                && matches!(f.8.as_deref(), Some("Symbol") | Some("Wingdings")))
                                        {
                                            let mut ascent = 0.0f32;
                                            let mut descent = 0.0f32;
                                            let mut natural_max = 0.0f32;
                                            for f in line.iter().filter(|f| !f.0.trim().is_empty()) {
                                                let m = &*self.registry.get(f.8.as_deref().unwrap_or("Calibri"));
                                                let leading = (m.ascent + m.descent + m.line_gap - m.win_ascent - m.win_descent).max(0.0);
                                                natural_max = natural_max.max(m.natural_line_height_hhea(f.1));
                                                ascent = ascent.max((m.win_ascent + leading) * f.1);
                                                descent = descent.max(m.win_descent * f.1);
                                            }
                                            // Baseline overflow is added once, outside the line-spacing multiple.
                                            if effective_line_rule == Some("atLeast") {
                                                lh = lh.max(ascent + descent);
                                            } else {
                                                lh += (ascent + descent - natural_max).max(0.0);
                                            }
                                        }
                                        // The final blank line belongs to the paragraph mark.
                                        // Earlier blank lines ending at explicit breaks keep
                                        // the break run's font box (Word's 32 controls).
                                        if LayoutEngine::paragraph_mark_only(para)
                                            || (line_idx + 1 == lines.len()
                                                && line.iter().all(|f| LayoutEngine::mark_spacing_only_text(&f.0))) {
                                            lh = self.cell_mark_line_height(para, effective_line_spacing,
                                                effective_line_rule, row_line_pitch).1;
                                        }
                                        if std::env::var("OXI_DBG_CELLLH").is_ok() {
                                            let fams: Vec<String> = line.iter().map(|f| format!("{}@{}", f.8.as_deref().unwrap_or("-"), f.1)).collect();
                                            let head: String = line.iter().flat_map(|f| f.0.chars()).take(10).collect();
                                            eprintln!("[CELLLH] lh={:.3} ls={:?} lr={:?} pitch={:?} snap={} n={} {:?} «{}»",
                                                lh, effective_line_spacing, effective_line_rule, row_line_pitch, para.style.snap_to_grid, line.len(), fams, head);
                                        }
                                        if lh == 0.0 {
                                            // whitespace-only line: fall back to all fragments
                                            lh = line.iter().map(|(_text, fs, _, _, _, _, _, _, font_family, _, _, _, _, _, _, _, _)| {
                                    let metrics = &*match font_family.as_deref() {
                                        Some(ff) => self.registry.get(ff),
                                        None => self.registry.default_metrics(),
                                    };
                                    self.line_height_inner(*fs, effective_line_spacing, effective_line_rule, metrics, para.style.snap_to_grid, row_line_pitch, true)
                                }).fold(0.0_f32, f32::max);
                                        }
                                        // S973 (2026-07-21, opt-out OXI_S973_DISABLE): a soft
                                        // break inside a CELL paragraph opens a line whether or
                                        // not anything lands on it, and Word gives that line the
                                        // paragraph's full height. With no fragments the two
                                        // folds above both yield 0, so an empty subline occupied
                                        // nothing. MEASURED (tools/metrics/_pb_cellbr_gen.py, one
                                        // document, marker rows above and below each case against
                                        // a break-free control, Times New Roman 12): a leading
                                        // break adds 13.800, a trailing one 13.830, two
                                        // consecutive 27.600, three 41.420 and two mid-paragraph
                                        // 27.600 — exactly one 13.8pt line per break, wherever it
                                        // sits. This is the sole axis separating
                                        // policies__0028d1be's one 41.67pt auto-spacing boundary
                                        // from the sixteen 27.8pt ones: `Physical Demands` opens
                                        // with a <w:br/>.
                                        if lh == 0.0 && std::env::var("OXI_S973_DISABLE").is_err() {
                                            // Same resolution as the empty-PARAGRAPH branch
                                            // above (22794): the paragraph mark's own rPr when
                                            // it names a size, else the document default.
                                            // Word gives the empty line the SAME height as a text
                                            // line of this paragraph, so take it from the
                                            // paragraph's own already-resolved fragments rather
                                            // than the paragraph mark (whose eastAsia chain would
                                            // price a Latin line at the CJK 83/64 box — measured
                                            // 15.56 against Word's 13.8 on this very boundary).
                                            lh = lines.iter().flat_map(|l| l.iter())
                                    .filter(|(text, ..)| !text.trim().is_empty())
                                    .map(|(_text, fs, _, _, _, _, _, _, font_family, _, _, _, _, _, _, _, _)| {
                                        let metrics = &*match font_family.as_deref() {
                                            Some(ff) => self.registry.get(ff),
                                            None => self.registry.default_metrics(),
                                        };
                                        self.line_height_inner(*fs, effective_line_spacing,
                                            effective_line_rule, metrics,
                                            para.style.snap_to_grid, row_line_pitch, true)
                                    })
                                    .fold(0.0_f32, f32::max);
                                        }
                                        // S1312 (2026-09-05, default ON, opt-out OXI_S1312_DISABLE):
                                        // every LINE that carries a ruby run grows by the ruby
                                        // expansion -- in a cell as in the body. DERIVED
                                        // (`_pb_cellruby_gen.py`, HG丸 10pt, hps 5 / raise 9, span
                                        // over a no-ruby control): body 1 line +3.75, 2 lines both
                                        // with ruby +7.50, ruby on the first line only +3.75; cell
                                        // 1 line +3.75, 3 lines all with ruby +11.25, first only
                                        // +3.75. Oxi gave the cell +2.84 / +0.15 (the estimate's
                                        // once-per-paragraph term, never the render) and the body
                                        // +4.18 once. Witness: correspondence__04a3e3e1's content
                                        // rows, two ruby lines short each (-8/row).
                                        if std::env::var("OXI_S1312_DISABLE").is_err()
                                            && !(effective_line_rule == Some("exact")
                                                && std::env::var("OXI_EXACT_CELL_RUBY").is_ok())
                                            && line.iter().any(|t| t.15)
                                        {
                                            let s1312_fs = self.resolve_font_size(&RunStyle::default(), &para.style);
                                            let exp = self.s1396_ruby_expansion(para, s1312_fs);
                                            // S1624 (2026-10-01, default ON, opt-out OXI_S1624_DISABLE):
                                            // on a typed grid the ruby goes into the line's grid cells
                                            // first -- the row is the natural line plus the expansion,
                                            // rounded UP to whole cells, not the snapped line plus it.
                                            // `_pb_gridruby_gen.py` (forms__01c5a769 host, lines 360,
                                            // hps 8, base 10.5 / 14 x hpsRaise 5..25pt, 12 arms): Word's
                                            // rows are 36 / 36 / 54 / 54 / 54 / 72 and 36 / 36 / 54 /
                                            // 54 / 72 / 72 (+0.48 rule), = ceil((nat + exp) / 18) cells
                                            // on all 12; Oxi gave 36 + exp (40.5 .. 80.5). Witness: the
                                            // form's «ふりがな／氏名» row, Word 36.48 against Oxi 48.5.
                                            // An exact line is never grid-snapped and carries no
                                            // expansion (S1396): legal__03512306's «ふりがな／氏名»
                                            // cells are line=400 exact and stay 20pt in Word.
                                            match row_line_pitch.filter(|p| *p > 0.0 && para.style.snap_to_grid) {
                                                Some(pitch) if std::env::var_os("OXI_S1624_DISABLE").is_none()
                                                    && exp > 0.0
                                                    && effective_line_rule != Some("exact") => {
                                                    let nat = line.iter()
                                                        .filter(|t| !t.0.trim().is_empty())
                                                        .map(|t| {
                                                            let metrics = &*match t.8.as_deref() {
                                                                Some(ff) => self.registry.get(ff),
                                                                None => self.registry.default_metrics(),
                                                            };
                                                            self.line_height_inner(t.1, effective_line_spacing,
                                                                effective_line_rule, metrics,
                                                                para.style.snap_to_grid, None, true)
                                                        })
                                                        .fold(0.0_f32, f32::max);
                                                    let cells = ((nat + exp) / pitch - 1e-3).ceil().max(1.0);
                                                    lh = lh.max(cells * pitch);
                                                }
                                                _ => lh += exp,
                                            }
                                        }
                                        // S1517 cell site (2026-09-21, default ON, opt-out
                                        // OXI_S1517_DISABLE): a raised/lowered run shifts its
                                        // own box by w:position and the CELL line is the
                                        // union of the shifted boxes -- the body path's S655,
                                        // which never reached cells. position_probe2.py
                                        // (TNR, line 240, Info6 pitch): same-size run lowered
                                        // 2pt +1.5; 10pt run beside 14pt text lowered 2pt
                                        // +0.75, raised 5pt +1.5, raised 10pt +6, lowered 7pt
                                        // +6; a 14pt run beside 10pt text lowered 2pt +0,
                                        // lowered 5pt +0.75 (the top shrinks). A flat
                                        // |position| broke 002a301d/0016b30b/0019967c
                                        // (lowered inline OLE equations); objects stay out.
                                        // technical__01242a0a 'Density (kg/m3)': the 3 at
                                        // position 10 -> Word row 25.5, Oxi 19.5.
                                        if std::env::var_os("OXI_S1517_DISABLE").is_none()
                                            && (!para.style.snap_to_grid || row_line_pitch.is_none())
                                            && line.iter().any(|t| t.16.position.is_some() && LayoutEngine::s1517_is_text(&t.0, &t.16))
                                        {
                                            let (mut a0, mut d0, mut a1, mut d1) = (0.0_f32, 0.0_f32, 0.0_f32, 0.0_f32);
                                            for t in line.iter().filter(|t| LayoutEngine::s1517_is_text(&t.0, &t.16)) {
                                                let m = &*match t.8.as_deref() {
                                                    Some(ff) => self.registry.get(ff),
                                                    None => self.registry.default_metrics(),
                                                };
                                                // A sub/superscript run's box is its DECLARED
                                                // (unscaled) size shifted by the raw position:
                                                // position_probe3.py -- sup/sub alone grow the
                                                // line 0, sup pos+10 grows it like a plain run
                                                // pos+10 (+4.5/+5.25), sub pos-8 +3.75, and a
                                                // 10pt-declared superscript beside 14pt text 0.
                                                // Word's PDF: 01242a0a's Density row is 25pt.
                                                let fs = if matches!(
                                                    t.16.vertical_align,
                                                    Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript)
                                                ) {
                                                    t.16.font_size.unwrap_or_else(|| self.resolve_font_size(&RunStyle::default(), &para.style))
                                                } else {
                                                    t.1
                                                };
                                                let (asc, des) = (m.word_ascent_pt(fs), m.word_descent_pt(fs));
                                                let pos = t.16.position.unwrap_or(0.0);
                                                a0 = a0.max(asc);
                                                d0 = d0.max(des);
                                                a1 = a1.max(asc + pos);
                                                d1 = d1.max(des - pos);
                                            }
                                            lh = (lh + (a1 + d1) - (a0 + d0)).max(0.0);
                                        }

                                        // S1125 (2026-08-15, opt-out OXI_S1125_DISABLE): a CELL
                                        // bullet's line 0 grows for its SYMBOL numbering marker
                                        // exactly as a body bullet line does — the S820b/S1112
                                        // ascent-overflow model, unmultiplied: overflow =
                                        // (marker_asc − text_asc − text_ext)⁺, advanced at the
                                        // paragraph ENTRY like S821 (Word truth: the exit pitch
                                        // b2→next-row is IDENTICAL 13.82 with and without the
                                        // Symbol marker — only the pitch INTO the bullet grows).
                                        // Word truth (_pb_bulletpitch_gen.py, NDIS-faithful
                                        // table-style cell arms): Symbol marker 11.64/11.76 vs
                                        // Arial marker / no marker 11.16; the model's 9.2+0.54 =
                                        // 9.74 sits centred in Word's 9.64/9.76 px alternation.
                                        // This is the NDIS technical__0043bfe0 −1 root: every
                                        // price row carries 1-2 Symbol bullets, Oxi's cell path
                                        // priced them flat (the S795/S1112 body sites are
                                        // layout_paragraph-only) → −0.4..−0.56/bullet → rows 13%
                                        // short from p78 on. Symbol-only (a non-Symbol marker
                                        // measured NO growth — B/J arms; the S1037 tall-marker
                                        // cell case stays unmeasured/unwired); exact/atLeast
                                        // exact clips the marker; minimum absorbs only the
                                        // overflow covered by its floor. Latin docs only (JP corpus
                                        // byte-identical).
                                        if line_idx == 0
                                            && !self.doc_body_has_real_cjk
                                            && !matches!(
                                                effective_line_rule,
                                                Some("exact")
                                            )
                                            && (list_marker_info
                                                .as_ref()
                                                .map_or(false, |(m, _, _)| m.contains('\u{F0B7}'))
                                                || self.s1614_aum_marker_asc(para).is_some())
                                            && std::env::var("OXI_S1125_DISABLE").is_err()
                                        {
                                            let mut s1125_asc: f32 = 0.0;
                                            let mut s1125_desc: f32 = 0.0;
                                            let mut s1125_best_sum: f32 = 0.0;
                                            let mut s1125_ext: f32 = 0.0;
                                            for (t, fs, _, _, _, _, _, _, font_family, _, _, _, _, _, _, _, _) in
                                                line.iter()
                                            {
                                                if t.trim().is_empty() {
                                                    continue;
                                                }
                                                let m = &*match font_family.as_deref() {
                                                    Some(ff) => self.registry.get(ff),
                                                    None => self.registry.default_metrics(),
                                                };
                                                s1125_asc = s1125_asc.max(m.win_ascent * *fs);
                                                s1125_desc = s1125_desc.max(m.win_descent * *fs);
                                                let win_sum = (m.win_ascent + m.win_descent) * *fs;
                                                let fext = (m.natural_line_height_hhea(*fs)
                                                    - win_sum)
                                                    .max(0.0);
                                                if win_sum > s1125_best_sum {
                                                    s1125_best_sum = win_sum;
                                                    s1125_ext = fext;
                                                }
                                            }
                                            if s1125_asc > 0.0 {
                                                // Microsoft Symbol winAsc 2059/2048 (the S795 const)
                                                let mfs = list_marker_info
                                                    .as_ref()
                                                    .map(|(_, f, _)| *f)
                                                    .unwrap_or(0.0);
                                                // S1614: an AUM numbering marker uses its own ascent.
                                                let marker_asc = self.s1614_aum_marker_asc(para)
                                                    .unwrap_or(2059.0 / 2048.0 * mfs);
                                                content_h += LayoutEngine::minimum_marker_entry_overflow(
                                                    (marker_asc - s1125_asc - s1125_ext).max(0.0),
                                                    s1125_asc + s1125_desc + s1125_ext,
                                                    lh, effective_line_rule);
                                            }
                                        }
                                        // Task P step 6 (2026-07-22, default ON, opt-out OXI_S982_DISABLE): grow the cell
                                        // line to fit an inline OLE object. The F8FE fragment carries
                                        // the object EXTENT as width but line_height_inner priced it
                                        // at the run font_size (~13.8), so a 17-21pt object overflowed
                                        // (negative y in step 5). ★S1066b correction (2026-08-05): the
                                        // FAITHFUL repro (real Navigator/Compatible cell paragraphs in
                                        // minimal tables, Word COM) proves Word's cell line =
                                        // max(paragraph_line, image_height) PER LINE — a raised
                                        // (w:position=-6) inline image sits within the line box, so it
                                        // does NOT add the text win-descent on top. nav_on row grows
                                        // +3.84pt over nav_off (16 atLeast → 19.8 = max(16, 19.79));
                                        // comp 3 images across 2 lines grow +7.20 (≈2×3.8). The old
                                        // `obj_h + descent + extra` over-counted by descent (21.92 vs
                                        // Word 19.8). target = max(lh, obj_h) (+ S875 auto-extra for a
                                        // multiple>1 auto rule, which is 0 for atLeast). Gated on
                                        // s982_cell → a default line has no F8FE fragment so obj_h=0
                                        // and this is inert (byte-identical).
                                        if s982_cell {
                                            // S1252: a maths box grows the cell line the same way
                                            // an object does - max(line, box), the S1066b model.
                                            let math_h: f32 = line
                                                .iter()
                                                .filter_map(|(text, fs, ..)| {
                                                    text.strip_prefix('\u{F8FD}')
                                                        .and_then(|t| t.parse::<usize>().ok())
                                                        .map(|i| (i, *fs))
                                                })
                                                .filter_map(|(i, fs)| {
                                                    cell_inline_math.get(i).map(|mb| {
                                                        let (_, a, d) =
                                                            crate::layout::math::inline_math_ink(
                                                                mb, fs,
                                                            );
                                                        a + d
                                                    })
                                                })
                                                .fold(0.0_f32, f32::max);
                                            let obj_h: f32 = line
                                                .iter()
                                                .filter_map(|(text, ..)| {
                                                    text.strip_prefix('\u{F8FE}')
                                                        .and_then(|s| s.parse::<usize>().ok())
                                                })
                                                .filter_map(|i| {
                                                    cell_inline_objects.get(i).map(|im| im.height)
                                                })
                                                .fold(0.0_f32, f32::max)
                                                .max(math_h);
                                            if obj_h > 0.0 {
                                                // S875 auto-rule extra leading (multiple > 1 only).
                                                let factor = if matches!(
                                                    para.style.line_spacing_rule.as_deref(),
                                                    None | Some("auto")
                                                ) {
                                                    para.style
                                                        .line_spacing
                                                        .map(|l| l.max(1.0))
                                                        .unwrap_or(1.0)
                                                } else {
                                                    1.0
                                                };
                                                let extra = if factor > 1.0 {
                                                    lh * (1.0 - 1.0 / factor)
                                                } else {
                                                    0.0
                                                };
                                                let target = obj_h.max(lh) + extra;
                                                if std::env::var("OXI_DBG_CELLOLE").is_ok() {
                                                    eprintln!("[CELL-OLE] phase=height obj_h={:.2} extra={:.2} text_lh={:.2} chosen={:.2}",
                                            obj_h, extra, lh, target);
                                                }
                                                if target > lh {
                                                    lh = target;
                                                }
                                            }
                                        }

                                        // Paragraph indentation: first line uses indent_left + first_line_indent
                                        let line_indent = p_indent_left
                                            + if line_idx == 0 {
                                                p_first_line_indent
                                            } else {
                                                0.0
                                            };
                                        let line_indent = line_indent + float_frame.map_or(0.0, |frame|
                                            frame.left - if line_idx == 0 { (p_indent_left + p_first_line_indent).max(0.0) } else { p_indent_left });

                                        // Calculate line total width for alignment
                                        let line_total_w: f32 = line
                                            .iter()
                                            .map(|(_, _, tw, _, _, _, _, _, _, _, _, _, _, _, _, _, _)| tw)
                                            .sum();
                                        // S502 (2026-06-08, FALSIFIED as a clean win — NOT shipped):
                                        // hypothesized that docGrid linesAndChars cells must center/right-align
                                        // on the GRID-EXPANDED width (natural tw sum + per-fullwidth-char
                                        // charSpace delta), not the natural width, because the render injects
                                        // that delta as character_spacing (~9964) so the rendered line is wider.
                                        // An idealized repro (long pure-fullwidth center line, charSpace=+1453)
                                        // confirmed +3.87pt: Oxi centered on natural 276 vs Word's grid 284.3.
                                        // SIGN: positive charSpace→expand correct, negative (b35 −2714)→natural
                                        // (clamp ≥0). BUT on REAL docs the effect is SUB-PIXEL and net-negative:
                                        // the only affected set is 5 mode-15 tokumei/order docs (linesAndChars
                                        // + charSpace>0 + jc=center-in-cell); their center lines are SHORT, and
                                        // a per-glyph position A/B (vs Word PDF) showed losses (~3.4pt total,
                                        // the longer p6 "匿名データの利用に当たって" line ON 1.45/OFF 0.85 ×4
                                        // docs) outweighing wins (29dc6e (名称) 0.26/0.44, d4d126 0.27/1.15;
                                        // ~1.1pt). The idealized repro did not generalize — Word's real centering
                                        // does not match the simple grid-expand model at this scale. SSIM-
                                        // invisible either way. jc=RIGHT was separately confounded by a real
                                        // merged/gridSpan cell-width error on 29dc6e ※ cells (Oxi ~4.6pt too
                                        // narrow; natural-width right-align was compensating it). Reverted to
                                        // natural-width alignment; left this note so the lever is not retried.
                                        let effective_wrap = if line_idx == 0 {
                                            first_line_wrap_w
                                        } else {
                                            wrap_w
                                        };
                                        let effective_wrap = float_frame.map_or(effective_wrap, |frame| frame.width);

                                        // Justify: non-last lines for jc=both, all lines for distribute
                                        let is_last_line = line_idx == total_lines - 1;
                                        let should_justify = (para.alignment == Alignment::Justify
                                            && !is_last_line)
                                            || para.alignment == Alignment::Distribute;

                                        // Alignment within the cell CONTENT area (cell_w - pad_l - pad_r).
                                        // S493j (2026-06-04): the common-case wrap_base = cell_w (NOT minus
                                        // padding — see ~8982), so right/center alignment within effective_wrap
                                        // overflowed by ~pad_l+pad_r. Right-aligned cell numbers then collided
                                        // with the next cell (2ea81a 相続税 row: right-aligned "2,000,000"
                                        // ended at the cell border, overlapping "被相続人" in the 備考 cell;
                                        // Word leaves the ~5.4pt cell right-margin). Subtract the padding for
                                        // ALIGNMENT only when wrap_base didn't already (wrapping unchanged →
                                        // Phase-1 safe). Opt-out OXI_S493J_DISABLE.
                                        // S1121: if the wrap base ALREADY subtracted the
                                        // padding (any of the 9 branches above), do not
                                        // subtract it again here — S493J only exists for the
                                        // legacy `cell_w` base. Opt-out OXI_S1121_DISABLE
                                        // restores the double subtraction.
                                        let pad_adjust = if (s1121_pad_in_base
                                            && std::env::var("OXI_S1121_DISABLE").is_err())
                                            || cell_hang_inner
                                            || s301_layout_fixed
                                            || s412_cellmar_subtract
                                            || std::env::var("OXI_S493J_DISABLE").is_ok()
                                        {
                                            0.0
                                        } else {
                                            pad_l + pad_r
                                        };
                                        let align_avail = (effective_wrap - pad_adjust).max(0.0);
                                        let align_offset = if should_justify {
                                            0.0
                                        } else {
                                            match para.alignment {
                                                Alignment::Center => {
                                                    (align_avail - line_total_w).max(0.0) / 2.0
                                                }
                                                Alignment::Right => {
                                                    (align_avail - line_total_w).max(0.0)
                                                }
                                                _ => 0.0,
                                            }
                                        };

                                        // Justify: CJK punctuation compression + space/gap distribution
                                        let aki_plan = if std::env::var_os("OXI_CELL_AKI").is_some() {
                                            let fragments: Vec<(&str, f32)> = line.iter().map(|f| (
                                                f.0.as_str(), s1175_autospace(f.1, f.11,
                                                    self.balance_single_byte_double_byte_width && !f.15),
                                            )).collect();
                                            plan_cell_aki(&fragments, line_total_w - effective_wrap,
                                                |a, b| cell_aki_joint(a, b, para.style.auto_space_de, para.style.auto_space_dn))
                                        } else { Vec::new() };
                                        let mut frag_width_adj: Vec<f32> = (0..line.len())
                                            .map(|i| -aki_plan.get(i).map_or(0.0, |p| p.width_reduction)).collect();
                                        // Preserve the same per-character reductions used to
                                        // shorten a fragment when emitting grid glyph positions.
                                        let mut frag_char_width_adj: Vec<Vec<f32>> = line.iter()
                                            .map(|f| vec![0.0; f.0.chars().count()]).collect();
                                        let mut frag_spacing: Vec<f32> = vec![0.0; line.len()];
                                        let mut justify_char_spacing: f32 = 0.0;
                                        // A justified terminal line may contract its word spaces,
                                        // but must never stretch. Trailing blank fragments are not
                                        // part of the visible extent or the compression capacity.
                                        if s1082_cell_shrink && is_last_line && para.alignment == Alignment::Justify {
                                            if let Some(last_visible) = line.iter().rposition(|f| !f.0.trim().is_empty()) {
                                                let visible_width: f32 = line[..=last_visible].iter().map(|f| f.2).sum();
                                                let deficit = (visible_width - effective_wrap).max(0.0);
                                                let spaces: Vec<(usize, f32)> = line[..last_visible].iter().enumerate()
                                                    .filter(|(_, f)| !f.0.is_empty() && f.0.chars().all(|c| c == ' '))
                                                    .map(|(i, f)| (i, f.2.max(0.0) * 0.25)).collect();
                                                let capacity: f32 = spaces.iter().map(|(_, w)| *w).sum();
                                                if deficit > 0.0 && capacity > 0.0 {
                                                    let shrink = deficit.min(capacity);
                                                    for (i, credit) in spaces {
                                                        frag_spacing[i] -= shrink * credit / capacity;
                                                    }
                                                }
                                            }
                                        }

                                        // 2026-04-19: allow single-fragment justify for CJK content.
                                        // Word distributes chars within a single CJK run for jc=both
                                        // non-last lines (b35 "組織的管" row: 4 chars spread across cell).
                                        if should_justify && !line.is_empty() {
                                            let mut slack = effective_wrap - line_total_w - frag_width_adj.iter().sum::<f32>();

                                            // Phase 1: CJK punctuation compression (only when overflowing)
                                            let limit_cell_compression = std::env::var_os("OXI_CELL_PUNCT_COMPRESSION_LIMIT").is_some();
                                            if slack < 0.0 {
                                                for (
                                                    fi,
                                                    (text, fs, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _),
                                                ) in line.iter().enumerate()
                                                {
                                                    for (char_index, ch) in text.chars().enumerate() {
                                                        if kinsoku::is_cjk_compressible(ch) {
                                                            let fm =
                                                                self.registry.default_metrics();
                                                            let char_w = self
                                                                .registry
                                                                .char_width_pt_with_fallback(
                                                                    ch, *fs, &fm,
                                                                );
                                                            // Do not compress beyond the actual deficit and
                                                            // then distribute that artificial excess as spaces.
                                                            let savings = if limit_cell_compression {
                                                                (char_w * 0.5).min((-slack).max(0.0))
                                                            } else { char_w * 0.5 };
                                                            frag_width_adj[fi] -= savings;
                                                            frag_char_width_adj[fi][char_index] -= savings;
                                                            slack += savings;
                                                        }
                                                    }
                                                }
                                            }

                                            // Phase 2: Distribute slack at word spaces, then CJK gaps.
                                            // A positioned tab has a fixed stop advance; it is
                                            // never a stretchable word space. Keep its advance
                                            // independent of the justified line's remaining slack.
                                            if slack > 0.0 {
                                                let space_count = line.iter()
                                        .enumerate()
                                        .filter(|(i, (text, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _))| *i < line.len() - 1 && text.trim().is_empty() && !text.contains('\t'))
                                        .count();
                                                if space_count > 0 {
                                                    let per_space = slack / space_count as f32;
                                                    for (
                                                        fi,
                                                        (
                                                            text,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _,
                                                            _, // S1312 ruby flag
                                                            _source_style,
                                                        ),
                                                    ) in line.iter().enumerate()
                                                    {
                                                        if fi < line.len() - 1
                                                            && text.trim().is_empty() && !text.contains('\t')
                                                        {
                                                            frag_spacing[fi] += per_space;
                                                        }
                                                    }
                                                } else if std::env::var("OXI_S1242_DISABLE")
                                                    .is_err()
                                                    && line.len() > 1
                                                    && (0..line.len() - 1).any(|i| {
                                                        line[i].0.ends_with(' ')
                                                            || line[i + 1].0.starts_with(' ')
                                                    })
                                                {
                                                    // S1242b: the doc's runs arrive pre-split
                                                    // at word boundaries with the spaces glued
                                                    // to the words ("member " "of ") — no pure-
                                                    // space fragment exists, so the branch above
                                                    // finds nothing. Distribute the slack at
                                                    // fragment boundaries that carry a space on
                                                    // either side (administrative__00018048's
                                                    // justified bullet cells).
                                                    let bounds: Vec<usize> = (0..line.len() - 1)
                                                        .filter(|&i| {
                                                            line[i].0.ends_with(' ')
                                                                || line[i + 1].0.starts_with(' ')
                                                        })
                                                        .collect();
                                                    let per_space = slack / bounds.len() as f32;
                                                    for i in bounds {
                                                        frag_spacing[i] += per_space;
                                                    }
                                                } else {
                                                    // No word spaces: distribute between ALL CJK character gaps
                                                    // Only activate when line is noticeably short (>10% slack);
                                                    // COM-confirmed 2026-04-19: for b35 "法令の理解" row with
                                                    // 4% slack Word does NOT distribute, showing natural widths.
                                                    let has_cjk = line.iter().any(|(text, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _)| text.chars().any(|c| kinsoku::is_cjk(c)));
                                                    let slack_ratio = if effective_wrap > 0.0 {
                                                        slack / effective_wrap
                                                    } else {
                                                        0.0
                                                    };
                                                    if has_cjk && slack_ratio > 0.10 {
                                                        let total_chars: usize = line.iter()
                                                .map(|(text, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _)| text.chars().count())
                                                .sum();
                                                        if total_chars > 1 {
                                                            let per_char_gap =
                                                                slack / (total_chars - 1) as f32;
                                                            for (
                                                                fi,
                                                                (
                                                                    text,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _,
                                                                    _, // S1312 ruby flag
                                                                    _source_style,
                                                                ),
                                                            ) in line.iter().enumerate()
                                                            {
                                                                let n = text.chars().count();
                                                                if n > 1 {
                                                                    frag_width_adj[fi] +=
                                                                        per_char_gap
                                                                            * (n - 1) as f32;
                                                                }
                                                                if fi < line.len() - 1 && n > 0 {
                                                                    frag_spacing[fi] +=
                                                                        per_char_gap;
                                                                }
                                                            }
                                                            // Pass per-char gap to renderer for visual spread.
                                                            justify_char_spacing = per_char_gap;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        // Apply text_y_offset to center/bottom-align text within line_height
                                        // per spec §13.4 note "GDI TextOutW character cell = fontSize".
                                        // For grid-snapped cell lines (line_height = n*pitch), this centers
                                        // the character cell; for exact spacing, bottom-aligns.
                                        let cell_max_fs: f32 = line
                                            .iter()
                                            .filter(|(t, ..)| !ignore_ascii_spaces || t.is_empty() || !t.chars().all(|c| c == ' ' || c == '\t'))
                                            .map(|(_, fs, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _)| *fs)
                                            .fold(0.0_f32, f32::max);
                                        // S175 (2026-05-22): match body's S166 fix — use word_line_height_table_cell
                                        // (font's natural height incl. ascent+descent) as centering height,
                                        // not raw font_size. The +2pt table-cell drift cluster (15f9/338c92/
                                        // 8efcd/cb8be/04b88e/b5f706/29dc6e and others, 62 docs with the
                                        // adjustLineHeightInTable XML tag) is caused by treating font_size
                                        // as natural height in the centering formula. Body was fixed in
                                        // S166; cell was missed.
                                        // S237 (2026-05-23): removed OXI_LEGACY_CELL_FONT_CENTERING
                                        // legacy env-var fallback during hardening pass.
                                        let cell_centering_height: f32 = if line.is_empty() {
                                            cell_max_fs
                                        } else {
                                            line.iter().filter(|(t, ..)| !ignore_ascii_spaces || t.is_empty() || !t.chars().all(|c| c == ' ' || c == '\t'))
                                    .map(|(text, fs, _, bold, italic, _underline, _us, _strikethrough, font_family, _color, _hl, _cs, _ts, _, ea, _ruby, _)| {
                                        let mut rs = RunStyle::default();
                                        rs.font_size = Some(*fs);
                                        rs.bold = *bold;
                                        rs.italic = *italic;
                                        if let Some(ff) = font_family { rs.font_family = Some(ff.clone()); }
                                        // S1299: give back the run's OWN eastAsia face. Without
                                        // it `metrics_for_text` resolves the East Asian family
                                        // from the paragraph — through docDefaults' eastAsiaTheme
                                        // to the theme — for a run that names one outright.
                                        if let (Some(e), true) =
                                            (ea, std::env::var("OXI_S1299_DISABLE").is_err())
                                        {
                                            rs.font_family_east_asia = Some(e.clone());
                                            rs.has_explicit_east_asia = true;
                                        }
                                        let m = &*self.metrics_for_text(text, &rs, &para.style);
                                        m.word_line_height_table_cell(*fs)
                                    })
                                    .fold(0.0_f32, f32::max)
                                        };
                                        let cell_max_fs = if tab_mark_line { tab_mark_fs } else { cell_max_fs };
                                        let cell_centering_height = if tab_mark_line {
                                            tab_mark_metrics.word_line_height_table_cell(tab_mark_fs)
                                        } else { cell_centering_height };
                                        let cell_text_y_off =
                                            match (effective_line_rule, effective_line_spacing) {
                                                (Some("exact"), Some(_))
                                                | (Some("atLeast"), Some(_)) => {
                                                    // Session 76 Mech A fix (2026-05-17): cells are
                                                    // body context — top-align text within line box
                                                    // for exact/atLeast (matches Word). Shape/textbox
                                                    // context bottom-aligns but cells are never shape.
                                                    // Session 78 Mech A v2 refinement: cell offset =
                                                    // 0.25pt (5 twips) per Session 70 B5/B6 repros.
                                                    //
                                                    // S462 (2026-05-31) ★ SHIP — BOTTOM-align cell text
                                                    // for exact/atLeast line spacing. The flat 0.25 top-
                                                    // align is correct only when the exact line value ≈
                                                    // natural glyph height (slack≈0, the B5/B6 repros).
                                                    // When the exact line is MUCH larger than the glyph
                                                    // (large slack), Word places the glyph at the BOTTOM
                                                    // of the line box — its documented exact-spacing rule
                                                    // "extra space goes ABOVE the glyph". Pixel-measured
                                                    // de6e32 p7 (12pt CJK list in line=480=24pt exact in a
                                                    // 1-cell table): Word offset ≈10pt = (lh − natural) vs
                                                    // Oxi's old ~0 → the whole list rendered ~8pt too HIGH.
                                                    // (Misdiagnosed as charGrid under-wrap S461; the wrap
                                                    // is IDENTICAL.) (lh − cell_centering_height) self-
                                                    // adjusts: ≈0 at slack≈0 (B5/B6 unaffected), ~10pt at
                                                    // large slack. GATE (full 410-pg): mean 0.9079→0.9098
                                                    // (+0.0020), bottom-3 +0.0115 / bottom-5 +0.0547 /
                                                    // bottom-10 +0.1732, <0.70 7→3, ≥0.99 47→64; 29 up
                                                    // (tokumei -1/-2/-3/-4 p7 all +0.112, p4 +0.03-0.04;
                                                    // 459f +0.039, 34140b +0.020), only 3 tiny regress
                                                    // (b35/a47e/1ec1 p1 ≈−0.003, opposite-direction
                                                    // charGrid/textbox). Render-only (cell text_y_off) →
                                                    // element.y / pagination / Phase-1 54/55 / Phase-2
                                                    // 0.9692 preserved. Override OXI_S462_CELL_EXACT
                                                    // (center / top).
                                                    let mode = std::env::var("OXI_S462_CELL_EXACT")
                                                        .unwrap_or_else(|_| "bottom".to_string());
                                                    match mode.as_str() {
                                                        "center" => {
                                                            ((lh - cell_centering_height).max(0.0)
                                                                / 2.0
                                                                * 2.0
                                                                + 0.5)
                                                                .floor()
                                                                / 2.0
                                                        }
                                                        "top" => 0.25,
                                                        _ => (lh - cell_centering_height).max(0.0),
                                                    }
                                                }
                                                _ => {
                                                    // Single/auto grid-snapped: center within lh using natural lh.
                                                    // S383 (2026-05-27, FALSIFIED): hypothesized the +1.0pt cluster
                                                    // was over-centering (center in row_height not grid-snapped lh).
                                                    // Env-gated test: net -0.0013, 0 improve / 9 regress, and
                                                    // b5f706 (the target) did NOT improve — so the +1.0 is NOT in
                                                    // the cell centering window. Re-localized to the body→table
                                                    // transition cursor advance (the +0.5 extra is in the
                                                    // body→table GAP: Word 19.0pt vs Oxi 19.5pt), not cell internals.
                                                    let raw =
                                                        (lh - cell_centering_height).max(0.0) / 2.0;
                                                    // S360 (2026-05-27): cell centering uses same CEIL-half-up
                                                    // rounding as body (S328). The +1.0pt table dy cluster
                                                    // (S357: 400 paragraphs / 16 docs) may stem from this CEIL
                                                    // over-application. Env-gated FLOOR / ROUND variants to test.
                                                    let use_floor =
                                                        std::env::var("OXI_S360_CELL_FLOOR")
                                                            .map(|v| v != "0" && v != "false")
                                                            .unwrap_or(false);
                                                    let use_round =
                                                        std::env::var("OXI_S360_CELL_ROUND")
                                                            .map(|v| v != "0" && v != "false")
                                                            .unwrap_or(false);
                                                    if use_floor {
                                                        (raw * 2.0).floor() / 2.0
                                                    } else if use_round {
                                                        (raw * 2.0).round() / 2.0
                                                    } else {
                                                        (raw * 2.0 + 0.5).floor() / 2.0
                                                    }
                                                }
                                            };
                                        // S453 (2026-05-30, Phase 3) ★ SSIM-validated cell-glyph vertical
                                        // correction. Oxi's table-cell first-line glyph renders ~1.5pt too
                                        // HIGH vs Word (Word reserves more leading above the first line than
                                        // Oxi's centering gives). Resolves the S451b glyph-vs-box-top
                                        // question via pixels: of the v2 box-top offset (~−3pt), ~1.5pt is a
                                        // REAL glyph error and ~1.5pt is box-top measurement convention.
                                        // EVIDENCE (DWrite SSIM): b5f706 δ-sweep peaks at +1.5 (0.7949→0.8024);
                                        // full 51-doc corpus d0 0.8456→d15 0.8495 (+0.0039), bottom-5 sum
                                        // +0.0396, 33 up / 3 down (b35/e8caed regress — opposite-direction
                                        // charGrid-compression family, S430; not in bottom-N). text_y_off is a
                                        // RENDER-time glyph offset, so element.y / layout / pagination are
                                        // UNCHANGED → Phase-1 (54/55) & Phase-2 (IoU 0.9692) sentinels exactly
                                        // preserved. Override/disable via OXI_S453_CELL_GLYPH_DY (set 0 to off).
                                        // TODO refine: magnitude is doc-dependent (d77a/04b88e want ~2pt) —
                                        // a leading-proportional δ would recover b35 and over-correct d77a.
                                        let mut cell_glyph_dy =
                                            std::env::var("OXI_S453_CELL_GLYPH_DY")
                                                .ok()
                                                .and_then(|v| v.parse::<f32>().ok())
                                                .unwrap_or(1.5);
                                        // S660 (2026-06-24, default ON, opt-out OXI_S660_DISABLE, override
                                        // OXI_S660_EXTRA): a table cell whose exact line box COMPRESSES below the
                                        // natural CJK 83/64 cell renders ~1pt too HIGH even after the flat S453
                                        // +1.5. Gold-standard per-line (Word PDF vs Oxi --dump-glyphs): the
                                        // tokumei form family has exact line 12.7 < natural 83/64 13.62 (10.5pt)
                                        // → +0.96-1.45 still too high; 459f05's exact 15.0 > natural → aligned at
                                        // 1.5. Add an extra +1.0 DY ONLY for these compressed CJK cells. The
                                        // reference is the 83/64 natural (word_line_height_no_grid), NOT
                                        // word_line_height_table_cell (=12.625 for 10.5pt, BELOW tokumei's 12.7 lh
                                        // → would never fire). SCOPE: expanded cells (459f05 lh>natural) AND the
                                        // gen2 family (auto/single, lh≈natural or larger fonts) are byte-identical
                                        // (NOT compressed) → fully protected. Full corpus SSIM A/B (extra=1.0):
                                        // only 30 docs change, net +0.1161, 18 improve (the form bottom-N: 6514
                                        // +0.0333, a1d6 +0.0241, d4d126 +0.0178, de6e32 +0.0164, tokumei_08_09
                                        // +0.0139, 2ea81a +0.0081, b35 +0.0013), 6 regress. RESIDUAL: 31420
                                        // (−0.0136) + order_09 (−0.0076) have CUMULATIVE per-line drift (offset
                                        // grows top→bottom), not a uniform per-cell offset, so a constant DY
                                        // over-corrects them — the deeper cumulative line-pitch wall, deferred.
                                        // Render-only (text_y_off) → element.y / pagination / Phase-1 / IoU
                                        // preserved by construction (same as S453).
                                        let s660_extra =
                                            if std::env::var("OXI_S660_DISABLE").is_ok() {
                                                0.0
                                            } else {
                                                std::env::var("OXI_S660_EXTRA")
                                                    .ok()
                                                    .and_then(|v| v.parse::<f32>().ok())
                                                    .unwrap_or(1.0)
                                            };
                                        if s660_extra != 0.0 && !line.is_empty() {
                                            let mut natural_cjk = 0.0_f32;
                                            for (
                                                text,
                                                fs,
                                                _,
                                                bold,
                                                italic,
                                                _u,
                                                _us,
                                                _st,
                                                ff,
                                                _c,
                                                _h,
                                                _cs,
                                                _ts,
                                                _,
                                                ea,
                                                _, // S1312 ruby flag
                                                _source_style,
                                            ) in line.iter()
                                            {
                                                let mut rs = RunStyle::default();
                                                rs.font_size = Some(*fs);
                                                rs.bold = *bold;
                                                rs.italic = *italic;
                                                if let Some(f) = ff {
                                                    rs.font_family = Some(f.clone());
                                                }
                                                // S1299: the run's own eastAsia face.
                                                if let (Some(e), true) =
                                                    (ea, std::env::var("OXI_S1299_DISABLE").is_err())
                                                {
                                                    rs.font_family_east_asia = Some(e.clone());
                                                    rs.has_explicit_east_asia = true;
                                                }
                                                let m =
                                                    &*self.metrics_for_text(text, &rs, &para.style);
                                                if m.is_cjk_83_64_font() {
                                                    let n = m.word_line_height_no_grid(*fs);
                                                    if n > natural_cjk {
                                                        natural_cjk = n;
                                                    }
                                                }
                                            }
                                            if natural_cjk > 0.0 && lh + 0.05 < natural_cjk {
                                                cell_glyph_dy += s660_extra;
                                            }
                                        }
                                        // S664 (2026-06-25, default ON, opt-out OXI_S664_DISABLE, override
                                        // OXI_S664_DY): a NEGATIVE-charSpace docGrid (文字詰め horizontal
                                        // compression, w:charSpace<0) renders its cell text ~0.5pt too HIGH
                                        // (render-anchor: the painted cell glyph baseline for these
                                        // compressed-grid forms lands ~0.5pt above Word). RENDER-ONLY cell
                                        // glyph push-down (added to cell_text_y_off, like S660 cell_glyph_dy);
                                        // content_h / row height / pagination UNCHANGED. SCOPE = page-level
                                        // grid_char_space_raw < 0 (doc-derivable): the 3 corpus word_png
                                        // neg-charSpace docs (b35123 +0.0105 [the worst tokumei doc], 191cb5
                                        // tokumei_08_10 +0.0026, albalunaSS neutral) all improve/neutral, 0
                                        // regress; positive-charSpace (d4d126 etc.) and type=lines docs are
                                        // EXCLUDED (mixed/regress under a uniform DY). DY=0.5 = the joint SSIM
                                        // peak (b35123 + 191cb5). See [[tokumei_form_family_ssim]].
                                        let s664_dy: f32 =
                                            if page.grid_char_space_raw.map_or(false, |cs| cs < 0)
                                                && std::env::var("OXI_S664_DISABLE").is_err()
                                            {
                                                std::env::var("OXI_S664_DY")
                                                    .ok()
                                                    .and_then(|v| v.parse().ok())
                                                    .unwrap_or(0.5)
                                            } else {
                                                0.0
                                            };
                                        // S698 (2026-06-30, default ON, opt-out OXI_S698_DISABLE, override
                                        // OXI_S698_DY): a PURE-LATIN (non-CJK) table cell renders ~0.76pt too
                                        // LOW. The S453 flat +1.5 cell_glyph_dy was calibrated on CJK cells
                                        // (which render too HIGH without it — Word reserves more leading above
                                        // a CJK first line); Latin glyphs need LESS leading, so +1.5 overshoots.
                                        // MEASURED uniform +0.76 within-cell across 4 gen2 Latin tables
                                        // (gen2_050/051/044/069 — Word PNG vs Oxi DWrite cell-text centroid),
                                        // with LibreOffice matching Word exactly (dY ~−0.01). Found via the
                                        // LibreOffice bug-finder (gen2 p2 table bands 0.65 vs Libra 0.98). The
                                        // root is grid-independent (S453 over-corrects Latin leading) so the
                                        // gate is purely "no CJK-83/64 fragment in the cell line" → the tuned
                                        // CJK form family (tokumei/1ec1/b35123, MS Mincho cells) is BYTE-
                                        // IDENTICAL by construction (s660 also keys off the CJK fragment, never
                                        // co-fires here). RENDER-ONLY (text_y_off) → element.y / row height /
                                        // pagination / Phase-1 / IoU preserved (same as S453/S660/S664).
                                        let s698_reduce: f32 = if std::env::var("OXI_S698_DISABLE")
                                            .is_err()
                                            && !line.is_empty()
                                        {
                                            let mut has_cjk = false;
                                            for (
                                                text,
                                                fs,
                                                _,
                                                bold,
                                                italic,
                                                _u,
                                                _us,
                                                _st,
                                                ff,
                                                _c,
                                                _h,
                                                _cs,
                                                _ts,
                                                _,
                                                ea,
                                                _, // S1312 ruby flag
                                                _source_style,
                                            ) in line.iter()
                                            {
                                                let mut rs = RunStyle::default();
                                                rs.font_size = Some(*fs);
                                                rs.bold = *bold;
                                                rs.italic = *italic;
                                                if let Some(f) = ff {
                                                    rs.font_family = Some(f.clone());
                                                }
                                                // S1299: the run's own eastAsia face.
                                                if let (Some(e), true) =
                                                    (ea, std::env::var("OXI_S1299_DISABLE").is_err())
                                                {
                                                    rs.font_family_east_asia = Some(e.clone());
                                                    rs.has_explicit_east_asia = true;
                                                }
                                                if self
                                                    .metrics_for_text(text, &rs, &para.style)
                                                    .is_cjk_83_64_font()
                                                {
                                                    has_cjk = true;
                                                    break;
                                                }
                                            }
                                            if has_cjk {
                                                0.0
                                            } else {
                                                std::env::var("OXI_S698_DY")
                                                    .ok()
                                                    .and_then(|v| v.parse::<f32>().ok())
                                                    .unwrap_or(0.76)
                                            }
                                        } else {
                                            0.0
                                        };
                                        let cell_text_y_off =
                                            cell_text_y_off + cell_glyph_dy + s664_dy - s698_reduce;
                                        // S1625 (2026-10-01, default ON, opt-out OXI_S1625_DISABLE):
                                        // a RUBY line on a typed grid centres the ruby-plus-base BLOCK
                                        // in its grid cells, not the base em box. The block runs from
                                        // the annotation box's top (raise + its upper box part) to the
                                        // base box's bottom (the lower part of the 83/64 box), so the
                                        // base baseline sits at lh/2 + (raise + up(ruby) - down(base))/2.
                                        // `_pb_gridruby_gen.py` (12 arms, base 10.5 / 14, hps 8, raise
                                        // 5..25pt over 2..4 cells): Word's baseline-from-rule fits that
                                        // to +0.29..+0.32 on every arm (the half rule above the cell);
                                        // forms__002abc3e (base 12, hps 5, raise 11) +0.44. Oxi kept the
                                        // em centring, 3-5pt high and blind to the raise.
                                        let cell_text_y_off = match row_line_pitch
                                            .filter(|p| *p > 0.0 && para.style.snap_to_grid)
                                        {
                                            Some(_) if std::env::var_os("OXI_S1625_DISABLE").is_none()
                                                && effective_line_rule != Some("exact")
                                                && line.iter().any(|t| t.15) =>
                                            {
                                                let base_t = line.iter().filter(|t| t.15)
                                                    .max_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(std::cmp::Ordering::Equal));
                                                let ruby = para.runs.iter().filter_map(|r| r.ruby.as_ref())
                                                    .max_by_key(|r| r.hps_raise_halfpt.unwrap_or(0));
                                                match (base_t, ruby) {
                                                    (Some(bt), Some(rb)) => {
                                                        let bfs = bt.1;
                                                        let bm = match bt.8.as_deref() {
                                                            Some(ff) => self.registry.get(ff),
                                                            None => self.registry.default_metrics(),
                                                        };
                                                        let hps = rb.hps_halfpt.map(|h| h as f32 / 2.0).unwrap_or(bfs / 2.0);
                                                        let raise = rb.hps_raise_halfpt.map(|h| h as f32 / 2.0)
                                                            .unwrap_or_else(|| ruby::default_hps_raise_pt(bfs, hps));
                                                        let rm = rb.annotation_fonts.first()
                                                            .map(|n| self.registry.get(n))
                                                            .unwrap_or_else(|| bm.clone());
                                                        let parts = |m: &FontMetrics, fs: f32| {
                                                            let extra = (LayoutEngine::s1367_cjk_box(m, fs)
                                                                - (m.win_ascent + m.win_descent) * fs).max(0.0) / 2.0;
                                                            (m.win_ascent * fs + extra, m.win_descent * fs + extra)
                                                        };
                                                        let (up_r, _) = parts(&rm, hps);
                                                        let (_, down_b) = parts(&bm, bfs);
                                                        let baseline = lh / 2.0 + (raise + up_r - down_b) / 2.0;
                                                        // renderer: baseline = y + text_y_off - 1.0 + ascent
                                                        baseline - bm.win_ascent * bfs + 1.0
                                                    }
                                                    _ => cell_text_y_off,
                                                }
                                            }
                                            _ => cell_text_y_off,
                                        };
                                        // Ordinary horizontal text on one cell line shares a
                                        // baseline even when its faces or sizes differ. Preserve
                                        // the logical line box and express the paint baseline
                                        // explicitly; using each face's own ascent lowers the
                                        // shorter face beside a taller symbol or formatting run.
                                        // Grid, ruby, positioned scripts and embedded objects
                                        // retain their existing composed geometry.
                                        let shared_cell_baseline = if row_line_pitch.is_none()
                                            && effective_line_rule != Some("exact")
                                            && para.style.list_marker.is_none()
                                            && line.iter().all(|f| !f.15 && !f.16.combine
                                                && f.16.vertical_align.is_none()
                                                && f.16.position.map_or(true, |p| p.abs() < 0.001)
                                                && f.16.inline_object_extent.is_none()
                                                && !f.0.starts_with('\u{F8FD}'))
                                        {
                                            let mut smallest = f32::INFINITY;
                                            let mut largest = 0.0f32;
                                            for f in line.iter().filter(|f| !f.0.trim().is_empty()) {
                                                let m = &*self.registry.get_with_style(
                                                    f.8.as_deref().unwrap_or("Calibri"), f.3, f.4);
                                                let ascent = m.baseline_ascent() * f.1;
                                                smallest = smallest.min(ascent);
                                                largest = largest.max(ascent);
                                            }
                                            if largest > smallest + 0.001 {
                                                Some(cell_text_y_off - 1.0 + largest)
                                            } else { None }
                                        } else { None };
                                        // S592: line-1 body starts AFTER the inline marker (no overlap).
                                        // S718: line-1 body pulled left to the defaultTabStop position.
                                        let mut rx = if s592_cell_space && line_idx == 0 {
                                            s592_marker_reserve
                                        } else if line_idx == 0 {
                                            -s718_pull
                                        } else {
                                            0.0_f32
                                        };
                                        // Emit list marker on the first line of the paragraph.
                                        if line_idx == 0 {
                                            if let Some((ref mk_text, mk_fs, mk_w)) =
                                                list_marker_info
                                            {
                                                let list_indent =
                                                    para.style.list_indent.unwrap_or(18.0);
                                                let marker_style = s1037_marker_style(para)
                                                    .cloned()
                                                    .unwrap_or_else(|| {
                                                        para.runs
                                                            .first()
                                                            .map(|r| r.style.clone())
                                                            .unwrap_or_default()
                                                    });
                                                // Session 75 Phase D: y is LINE BOX TOP; renderer adds cell_text_y_off.
                                                // S592: place a space-suffix number at cell-left (no outdent).
                                                let marker_x = if s592_cell_space {
                                                    cell_x + pad_l + line_indent
                                                } else {
                                                    cell_x + pad_l + line_indent - list_indent
                                                };
                                                let mut marker_el = LayoutElement::new(
                                                    marker_x,
                                                    content_h,
                                                    mk_w,
                                                    lh,
                                                    LayoutContent::Text {
                                                        text: mk_text.clone(),
                                                        font_size: mk_fs,
                                                        font_family: self
                                                            .resolve_font_family_for_text(
                                                                mk_text,
                                                                &marker_style,
                                                                &para.style,
                                                            )
                                                            .map(|s| s.to_string()),
                                                        bold: self.resolve_bold(
                                                            &marker_style,
                                                            &para.style,
                                                        ),
                                                        italic: marker_style.italic,
                                                        underline: marker_style.underline,
                                                        underline_style: marker_style
                                                            .underline_style
                                                            .clone(),
                                                        strikethrough: marker_style.strikethrough,
                                                        double_strikethrough: marker_style
                                                            .double_strikethrough,
                                                        color: self
                                                            .resolve_color(
                                                                &marker_style,
                                                                &para.style,
                                                            )
                                                            .map(|s| s.to_string()),
                                                        highlight: marker_style.highlight.clone(),
                                                        character_spacing: 0.0,
                                                        field_type: None,
                                                        text_scale: 100.0,
                                                        is_vertical: false,
                                                        effects: TextEffects::default(),
                                                    },
                                                );
                                                // Session 72 Phase A: populate text_y_off.
                                                marker_el.text_y_off = cell_text_y_off;
                                            self.set_cell_exact_baseline(&mut marker_el, effective_line_rule, effective_line_spacing.unwrap_or(lh));
                                                // A marker travels with its cell paragraph when
                                                // widow/orphan control changes the page split.
                                                marker_el.cell_paragraph_index = Some(cell_para_counter);
                                                marker_el.cell_row_index = Some(row_idx);
                                                marker_el.cell_col_index = Some(cell_idx);
                                                cell_elements.push(marker_el);
                                            }
                                        }
                                        for (
                                            frag_idx,
                                            (
                                                text,
                                                fs,
                                                tw,
                                                bold,
                                                italic,
                                                underline,
                                                underline_style,
                                                strikethrough,
                                                font_family,
                                                color,
                                                highlight,
                                                cs,
                                                ts,
                                                lrpb_before,
                                                _,
                                                s1626_frag_ruby, // S1312 ruby flag
                                                _source_style,
                                            ),
                                        ) in line.iter().enumerate()
                                        {
                                            let adj_w = *tw + frag_width_adj[frag_idx];
                                            // Incoming auto-space belongs before this fragment's first glyph.
                                            // Its width is already included in adj_w by the character walker.
                                            let natural_leading = frag_idx.checked_sub(1)
                                                .and_then(|i| line.get(i))
                                                .filter(|previous| {
                                                    let ruby_adjacent = std::env::var("OXI_S1316_DISABLE").is_err()
                                                        && (previous.16.ruby_field || _source_style.ruby_field);
                                                    !ruby_adjacent && previous.0.chars().last()
                                                        .zip(text.chars().next())
                                                        .is_some_and(|(a,b)| cell_aki_joint(a,b,
                                                            para.style.auto_space_de,para.style.auto_space_dn))
                                                })
                                                .map_or(0.0, |previous| {
                                                    self.natural_autospace_after(previous.0.chars().last().unwrap_or(' '),
                                                        &previous.16,&para.style,previous.1,previous.11)
                                                });
                                            let aki_leading = aki_plan.get(frag_idx)
                                                .map_or(natural_leading, |p| p.leading_gap);
                                            // Task P step 5 (2026-07-22, default ON, opt-out OXI_S982_DISABLE): a U+F8FE{index}
                                            // cell-inline OLE fragment → the registered &Image drawn on
                                            // the line. Step 5 places the object bottom at the line
                                            // bottom (content_h + lh); step 6 refines it to the text
                                            // baseline (mixed) / line bottom (solo). Gated on s982_cell;
                                            // a default line has no F8FE text so the strip_prefix never
                                            // matches — byte-identical either way, the flag is explicit.
                                            if s982_cell {
                                                if let Some(index) = text
                                                    .strip_prefix('\u{F8FE}')
                                                    .and_then(|s| s.parse::<usize>().ok())
                                                {
                                                    if let Some(img) =
                                                        cell_inline_objects.get(index).copied()
                                                    {
                                                        let oh = img.height;
                                                        let base_x = cell_x
                                                            + pad_l
                                                            + line_indent
                                                            + align_offset
                                                            + rx;
                                                        let obj_bottom = content_h + lh;
                                                        // S1238 (2026-08-27): a data-less
                                                        // flow-reservation placeholder with a
                                                        // visible wps frame renders it as a
                                                        // BoxRect at the flowed position
                                                        // (kyotei 労働保険番号 digit boxes).
                                                        // Box BOTTOM sits at line_bottom −
                                                        // effectExtent.b (kyotei row1: Word
                                                        // bottom 29.5 = 30.8 − 1.5 ± 0.2;
                                                        // plain line-bottom is 1.3 low,
                                                        // line-top 4.2 high, a 5pt-run
                                                        // baseline 2.6 high).
                                                        let s1238_top =
                                                            content_h + lh - oh - if img.data.is_empty() { img.effect_extent_b } else { 0.0 };
                                                        let s1238_content = if img.data.is_empty()
                                                            && std::env::var("OXI_S1238_DISABLE")
                                                                .is_err()
                                                        {
                                                            img.placeholder_outline.as_ref().map(
                                                                |(stroke, sw, fill)| {
                                                                    LayoutContent::BoxRect {
                                                                        fill: fill.clone(),
                                                                        stroke_color: Some(
                                                                            stroke.clone(),
                                                                        ),
                                                                        stroke_width: *sw,
                                                                        corner_radius: 0.0,
                                                                    }
                                                                },
                                                            )
                                                        } else {
                                                            None
                                                        };
                                                        if let Some(content) = s1238_content {
                                                            let mut e = LayoutElement::new(
                                                                base_x,
                                                                s1238_top,
                                                                img.width,
                                                                oh,
                                                                content,
                                                            );
                                                            e.paragraph_index = block_idx;
                                                            e.cell_paragraph_index =
                                                                Some(cell_para_counter);
                                                            e.cell_row_index = Some(row_idx);
                                                            e.cell_col_index = Some(cell_idx);
                                                            cell_elements.push(e);
                                                            rx += adj_w;
                                                            continue;
                                                        }
                                                        let mut e = LayoutElement::new(
                                                            base_x,
                                                            obj_bottom - oh,
                                                            img.width,
                                                            oh,
                                                            LayoutContent::Image {
                                                                data: img.data.clone(),
                                                                content_type: img
                                                                    .content_type
                                                                    .clone(),
                                                                crop: img.crop.as_ref().map(|c| {
                                                                    (
                                                                        c.top, c.right, c.bottom,
                                                                        c.left,
                                                                    )
                                                                }),
                                                            },
                                                        );
                                                        e.paragraph_index = block_idx;
                                                        e.cell_paragraph_index =
                                                            Some(cell_para_counter);
                                                        e.cell_row_index = Some(row_idx);
                                                        e.cell_col_index = Some(cell_idx);
                                                        if std::env::var("OXI_DBG_CELLOLE").is_ok()
                                                        {
                                                            eprintln!("[CELL-OLE] phase=emit enabled=1 index={} w={} h={} y={}",
                                                    index, img.width, oh, obj_bottom - oh);
                                                        }
                                                        cell_elements.push(e);
                                                        rx += adj_w;
                                                        continue;
                                                    }
                                                }
                                            }
                                            // S1252: a U+F8FD{index} cell fragment draws the
                                            // registered maths at the fragment position, its
                                            // baseline on the cell line's text baseline.
                                            if s982_cell {
                                                if let Some(index) = text
                                                    .strip_prefix('\u{F8FD}')
                                                    .and_then(|t| t.parse::<usize>().ok())
                                                {
                                                    if let Some(mb) =
                                                        cell_inline_math.get(index).copied()
                                                    {
                                                        let base_x = cell_x
                                                            + pad_l
                                                            + line_indent
                                                            + align_offset
                                                            + rx;
                                                        let bbox =
                                                            crate::layout::math::layout_math_block(
                                                                mb, *fs,
                                                            );
                                                        let asc = self
                                                            .registry
                                                            .default_metrics()
                                                            .win_ascent
                                                            * *fs;
                                                        let baseline =
                                                            content_h + cell_text_y_off + asc;
                                                        let (mut me, _) =
                                                            crate::layout::math::emit_math_block(
                                                                mb,
                                                                base_x,
                                                                baseline
                                                                    - bbox.ascent.max(*fs * 0.8),
                                                                *fs,
                                                            );
                                                        for e in me.iter_mut() {
                                                            e.paragraph_index = block_idx;
                                                            e.cell_paragraph_index =
                                                                Some(cell_para_counter);
                                                            e.cell_row_index = Some(row_idx);
                                                            e.cell_col_index = Some(cell_idx);
                                                        }
                                                        cell_elements.append(&mut me);
                                                        rx += adj_w;
                                                        continue;
                                                    }
                                                }
                                            }
                                            // S703c: a SENTINEL-encoded combine tuple → warichu (2
                                            // small rows + brackets), in place of one glyph element.
                                            if let Some(rest) = text.strip_prefix('\u{F8FF}') {
                                                let mut it = rest.splitn(2, '\u{F8FF}');
                                                let brk = it.next().unwrap_or("none");
                                                let realtext = it.next().unwrap_or("");
                                                let base_x = cell_x
                                                    + pad_l
                                                    + line_indent
                                                    + align_offset
                                                    + rx;
                                                let small = *fs * 0.5;
                                                let bsz = *fs * 0.8;
                                                let chars: Vec<char> = realtext.chars().collect();
                                                let half = (chars.len() + 1) / 2;
                                                let toprow: String = chars[..half].iter().collect();
                                                let botrow: String = chars[half..].iter().collect();
                                                let (lb, rb): (&str, &str) = match brk {
                                                    "round" => ("（", "）"),
                                                    "square" => ("〔", "〕"),
                                                    "angle" => ("〈", "〉"),
                                                    "curly" => ("｛", "｝"),
                                                    _ => ("", ""),
                                                };
                                                let mut wpush = |elements: &mut Vec<LayoutElement>, t: String, wx: f32, ydelta: f32, wsz: f32| {
                                        if t.is_empty() { return; }
                                        let mut e = LayoutElement::new(wx, content_h, wsz, lh, LayoutContent::Text {
                                            text: t, font_size: wsz, font_family: font_family.clone(),
                                            bold: *bold, italic: *italic, underline: false, underline_style: None,
                                            strikethrough: false, double_strikethrough: false, color: color.clone(),
                                            highlight: None, field_type: None, character_spacing: 0.0,
                                            text_scale: 100.0, is_vertical: false, effects: TextEffects::default(),
                                        });
                                        e.text_y_off = cell_text_y_off + ydelta;
                                        e.paragraph_index = block_idx;
                                        e.cell_paragraph_index = Some(cell_para_counter);
                                        e.cell_row_index = Some(row_idx);
                                        e.cell_col_index = Some(cell_idx);
                                        elements.push(e);
                                    };
                                                let mut wx = base_x;
                                                if !lb.is_empty() {
                                                    wpush(
                                                        &mut cell_elements,
                                                        lb.to_string(),
                                                        wx,
                                                        0.0,
                                                        bsz,
                                                    );
                                                    wx += bsz;
                                                }
                                                wpush(&mut cell_elements, toprow, wx, 0.0, small);
                                                wpush(&mut cell_elements, botrow, wx, small, small);
                                                wx += half as f32 * small;
                                                if !rb.is_empty() {
                                                    wpush(
                                                        &mut cell_elements,
                                                        rb.to_string(),
                                                        wx,
                                                        0.0,
                                                        bsz,
                                                    );
                                                }
                                                rx += adj_w;
                                                continue;
                                            }
                                            // 2026-04-19: Inject charSpace delta into GDI cs so TextOutW
                                            // renders at layout-correct advance (prevents glyph overlap
                                            // between fragments when pitch<natural).
                                            let grid_cs_adj = if let (Some(ratio), Some(pitch)) =
                                                (grid_char_cw_ratio, grid_char_pitch)
                                            {
                                                if ratio > 0.0 && pitch > 0.0 {
                                                    let default_fs = pitch / ratio;
                                                    let char_space_pt = pitch - default_fs;
                                                    // Only apply to fullwidth CJK content (halfwidth chars render naturally)
                                                    if text
                                                        .chars()
                                                        .any(|c| crate::font::is_fullwidth(c))
                                                    {
                                                        char_space_pt
                                                    } else {
                                                        0.0
                                                    }
                                                } else {
                                                    0.0
                                                }
                                            } else {
                                                0.0
                                            };
                                            // Session 75 Phase D: y is LINE BOX TOP; renderer adds cell_text_y_off.
                                            let mut cell_el = LayoutElement::new(
                                                cell_x + pad_l + line_indent + align_offset + rx + aki_leading,
                                                content_h,
                                                (adj_w - aki_leading).max(0.0),
                                                lh,
                                                LayoutContent::Text {
                                                    text: text.clone(),
                                                    font_size: *fs,
                                                    font_family: font_family.clone(),
                                                    bold: *bold,
                                                    italic: *italic,
                                                    underline: *underline,
                                                    underline_style: underline_style.clone(),
                                                    strikethrough: *strikethrough,
                                                    double_strikethrough: false,
                                                    color: color.clone(),
                                                    highlight: highlight.clone(),
                                                    character_spacing: *cs
                                                        + justify_char_spacing
                                                        + grid_cs_adj,
                                                    field_type: None,
                                                    text_scale: *ts,
                                                    is_vertical: false,
                                                    effects: TextEffects::default(),
                                                },
                                            );
                                            // Session 72 Phase A: populate text_y_off (y still includes it).
                                            cell_el.text_y_off = cell_text_y_off;
                                            cell_el.baseline_offset = shared_cell_baseline;
                                            self.set_cell_exact_baseline(&mut cell_el, effective_line_rule, effective_line_spacing.unwrap_or(lh));
                                            // Attribute to the table's source block index so diff tools
                                            // can localize cell text. Without this, para_idx is None and
                                            // docs with many tables produce unusable --dump-layout output.
                                            cell_el.paragraph_index = block_idx;
                                            // R7.32: also tag cell-internal paragraph index so the
                                            // matcher (aggregate_dump in measure_pagination_oxi.py)
                                            // can split cell paragraphs that share block_idx.
                                            cell_el.cell_paragraph_index = Some(cell_para_counter);
                                            // R7.44: tag (row, col) within the table so cells
                                            // sharing (block_idx, cpi=0) don't collapse.
                                            cell_el.cell_row_index = Some(row_idx);
                                            cell_el.cell_col_index = Some(cell_idx);
                                            // R7.56 (Day 34 part 25, 2026-05-13): mark the FIRST
                                            // text element of a paragraph whose run[0] carries
                                            // `<w:lastRenderedPageBreak/>`. The row-split logic
                                            // uses this to force a page break before this element
                                            // (mid-cell LRPB respect for e3c545 cpi=81/N/M).
                                            //
                                            // R7.64 (Day 37, 2026-05-14): exclude cell-first paragraphs
                                            // (cell_para_counter == 0). In a multi-cell row that
                                            // Word split mid-cell, each cell's first paragraph can
                                            // carry an LRPB indicating "this cell continues here
                                            // after page break", not "split before this element".
                                            // ed025c balance sheet row 1: cells 1, 3 first paragraphs
                                            // had LRPB at p0r0 alongside cell 0 p32 LRPB (genuine
                                            // mid-cell split). Without this gate, cells 1, 3 p0
                                            // elements at y=row_top pulled split_y to row_top → all
                                            // row 1 content pushed to next page. Mirrors R7.58 gate
                                            // (mod.rs:6166) which excludes (ci==0, first_para, ri==0)
                                            // — extends exclusion to ALL cells' first paragraphs.
                                            if s993_exact && cell_lrpb_enabled {
                                                // S993 (R2/R3, 2026-07-23): mark the EXACT fragment
                                                // whose originating run carried the mid-run LRPB —
                                                // the split consumer pulls that line to the next
                                                // page. The R7.73 next-paragraph approximation
                                                // (below) placed the anchor 1-2 lines late (or lost
                                                // it when the marker was in the cell's LAST
                                                // paragraph), stranding the continuation text on the
                                                // previous page. p0/r0 is already excluded via
                                                // s993_lrpb_pending (a continuation/row-start marker,
                                                // R7.64), so *lrpb_before is false there.
                                                if *lrpb_before {
                                                    cell_el.is_paragraph_start_with_lrpb = true;
                                                }
                                            } else if cell_lrpb_enabled && line_idx == 0
                                                && frag_idx == 0
                                                && cell_para_counter > 0
                                            {
                                                let para_has_lrpb_on_run0 = para
                                                    .runs
                                                    .first()
                                                    .map(|r| r.has_last_rendered_page_break)
                                                    .unwrap_or(false);
                                                // R7.73 (Day 37 session 58, 2026-05-15):
                                                // also mark when the IMMEDIATE PREVIOUS
                                                // paragraph in this cell had LRPB on a
                                                // non-run-0 (mid-paragraph LRPB indicates
                                                // Word split mid-paragraph; closest
                                                // paragraph-boundary approximation is to
                                                // pull-back at the start of this NEXT
                                                // paragraph). d4d126 wi=291 has LRPB on
                                                // run 1, wi=292 should be pulled to next
                                                // page. COM-confirmed wi=291 split is
                                                // current (not stale).
                                                if para_has_lrpb_on_run0
                                                    || prev_cell_para_had_mid_lrpb
                                                {
                                                    cell_el.is_paragraph_start_with_lrpb = true;
                                                }
                                            }
                                            // S714 (default ON, opt-out OXI_ASCELL_DISABLE): render the autoSpaceDE/DN aki
                                            // INSIDE a cell fragment. The cell wrap already put the fs/4
                                            // aki into adj_w (buf_w), but the fragment renders as ONE
                                            // monolithic element → glyphs tight, aki only trailing. Split
                                            // the element at DE/DN boundaries and redistribute the aki into
                                            // the gaps (total width unchanged → rx/wrap/pagination
                                            // unchanged; render-only). Matches the BODY fragment-widening.
                                            let single_byte_grid = if (std::env::var_os("OXI_CELL_GRID_SINGLE_BYTE").is_some()
                                            || std::env::var_os("OXI_S1449_DISABLE").is_none()) {
                                                match (grid_char_pitch, grid_char_cw_ratio) {
                                                    (Some(p), Some(r)) if p > 0.0 && r > 0.0 => Some(if _source_style.fit_text.is_some() {
                                                        (p / r, 1.0, self.balance_single_byte_double_byte_width, *ts / 100.0)
                                                    } else {
                                                        (p, r, self.balance_single_byte_double_byte_width, *ts / 100.0)
                                                    }),
                                                    _ => None,
                                                }
                                            } else { None };
                                            let grid_run_chars: Vec<char> = text.chars().collect();
                                            let grid_metric_owners: Vec<FontMetricsRef<'_>> = if single_byte_grid.is_some() {
                                                grid_run_chars.iter().copied().enumerate().map(|(i, ch)| self.metrics_for_char_in(ch, !self.s1052_cell_latin_quote(&grid_run_chars, i), _source_style, &para.style)).collect()
                                            } else { Vec::new() };
                                            let grid_metrics: Vec<&FontMetrics> = grid_metric_owners.iter().map(|metrics| &**metrics).collect();
                                            let grid_kern_active = std::env::var("OXI_KERNBREAK_DISABLE").is_err()
                                                && _source_style.kern.or_else(|| para.style.default_run_style.as_ref().and_then(|rs| rs.kern))
                                                    .map_or(false, |k| k > 0.0 && *fs >= k);
                                            let grid_latin_em = std::env::var("OXI_S869_DISABLE").is_err()
                                                && !grid_kern_active && std::env::var("OXI_LATINEM_DISABLE").is_err()
                                                && latinem_in_scope(self.doc_body_has_real_cjk);
                                            let grid_kern = std::env::var("OXI_S869_DISABLE").is_err()
                                                && grid_kern_active && std::env::var("OXI_S1017_DISABLE").is_err()
                                                && !self.doc_body_has_real_cjk;
                                            let ascell_split: Option<Vec<(String, f32, f32)>> =
                                                if std::env::var("OXI_ASCELL_DISABLE").is_err() {
                                                    let fm = if single_byte_grid.is_some() {
                                                        self.registry.get_with_style(font_family.as_deref().unwrap_or(""), *bold, *italic)
                                                    } else {
                                                        self.registry.get(font_family.as_deref().unwrap_or(""))
                                                    };
                                                    autospace_cell_segments(
                                                        text,
                                                        *fs,
                                                        &self.registry,
                                                        &fm,
                                                        *cs + justify_char_spacing + if single_byte_grid.is_some() { 0.0 } else { grid_cs_adj },
                                                        para.style.auto_space_de,
                                                        para.style.auto_space_dn,
                                                        None,
                                                        std::env::var_os("OXI_CELL_BALANCED_SPACE").is_some()
                                                            && self.balance_single_byte_double_byte_width,
                                                        frag_idx.checked_sub(1).and_then(|i| line.get(i))
                                                            .and_then(|f| f.0.chars().last()),
                                                        line.get(frag_idx + 1).and_then(|f| f.0.chars().next()),
                                                        single_byte_grid,
                                                        &grid_metrics, grid_latin_em, grid_kern,
                                                        &frag_char_width_adj[frag_idx],
                                                        &text.chars().map(|c| self.natural_autospace_after(c,
                                                            _source_style,&para.style,*fs,*cs)
                                                            * aki_plan.get(frag_idx).map_or(1.0, |p| p.gap_scale))
                                                            .collect::<Vec<_>>(),
                                                    )
                                                } else {
                                                    None
                                                };
                                            if let Some(segs) = ascell_split {
                                                let base_x = cell_x
                                                    + pad_l
                                                    + line_indent
                                                    + align_offset
                                                    + rx + aki_leading;
                                                let lrpb = cell_el.is_paragraph_start_with_lrpb;
                                                for (si, (seg_text, seg_dx, seg_w)) in
                                                    segs.iter().enumerate()
                                                {
                                                    let mut se = LayoutElement::new(
                                                        base_x + seg_dx,
                                                        content_h,
                                                        *seg_w,
                                                        lh,
                                                        LayoutContent::Text {
                                                            text: seg_text.clone(),
                                                            font_size: *fs,
                                                            font_family: font_family.clone(),
                                                            bold: *bold,
                                                            italic: *italic,
                                                            underline: *underline,
                                                            underline_style: underline_style
                                                                .clone(),
                                                            strikethrough: *strikethrough,
                                                            double_strikethrough: false,
                                                            color: color.clone(),
                                                            highlight: highlight.clone(),
                                                            character_spacing: *cs
                                                                + justify_char_spacing
                                                                + if single_byte_grid.is_some() { 0.0 } else { grid_cs_adj },
                                                            field_type: None,
                                                            text_scale: *ts,
                                                            is_vertical: false,
                                                            effects: TextEffects::default(),
                                                        },
                                                    );
                                                    se.text_y_off = cell_text_y_off;
                                                    se.baseline_offset = shared_cell_baseline;
                                                    se.paragraph_index = block_idx;
                                                    se.cell_paragraph_index =
                                                        Some(cell_para_counter);
                                                    se.cell_row_index = Some(row_idx);
                                                    se.cell_col_index = Some(cell_idx);
                                                    se.is_paragraph_start_with_lrpb =
                                                        lrpb && si == 0;
                                                    cell_elements.push(se);
                                                }
                                            } else {
                                                cell_elements.push(cell_el);
                                            }
                                            // S1626 (2026-10-01, default ON, opt-out OXI_S1626_DISABLE):
                                            // draw the ruby annotation of a CELL run. The body path
                                            // (Round 7) emits it; the cell path never did, so every
                                            // ruby in a table cell rendered bare (forms__01c5a769
                                            // furigana over the name label, 002abc3e, 160e800d). Placed
                                            // like the body: over the base run per rubyAlign, its
                                            // baseline `raise` above the base baseline the renderer
                                            // draws (y + text_y_off - 1.0 + ascent).
                                            if *s1626_frag_ruby
                                                && std::env::var_os("OXI_S1626_DISABLE").is_none()
                                                && s1626_ri < s1626_rubies.len()
                                            {
                                                let group = &s1626_rubies[s1626_ri];
                                                let run = group[0];
                                                let group_chars: usize = group.iter().map(|r| r.text.chars().count()).sum();
                                                if s1626_seen == 0 {
                                                    if let Some(ruby_ir) = run.ruby.as_ref() {
                                                        let base_pt = *fs;
                                                        let hps_pt = ruby_ir.hps_halfpt.map(|h| h as f32 / 2.0).unwrap_or(base_pt / 2.0);
                                                        let raise_pt = ruby_ir.hps_raise_halfpt.map(|h| h as f32 / 2.0)
                                                            .unwrap_or_else(|| ruby::default_hps_raise_pt(base_pt, hps_pt));
                                                        let base_m = match font_family.as_deref() {
                                                            Some(ff) => self.registry.get(ff),
                                                            None => self.registry.default_metrics(),
                                                        };
                                                        let ruby_family = ruby_ir.annotation_fonts.first().cloned()
                                                            .or_else(|| font_family.clone());
                                                        let ruby_m = match ruby_family.as_deref() {
                                                            Some(ff) => self.registry.get(ff),
                                                            None => self.registry.default_metrics(),
                                                        };
                                                        let ruby_text = ruby_ir.text.as_str();
                                                        let ruby_n = ruby_text.chars().count();
                                                        let ruby_w: f32 = ruby_text.chars()
                                                            .map(|c| self.registry.char_width_pt_with_fallback(c, hps_pt, &ruby_m))
                                                            .sum();
                                                        let base_w: f32 = group.iter().map(|r| {
                                                            let n = r.text.chars().count() as f32;
                                                            r.text.chars()
                                                                .map(|c| self.registry.char_width_pt_with_fallback(c, base_pt, &base_m))
                                                                .sum::<f32>()
                                                                + r.style.character_spacing.unwrap_or(0.0) * n
                                                        }).sum();
                                                        let (dx, ruby_cs) = ruby::ruby_position(base_w, ruby_w, ruby_n, ruby_ir.align);
                                                        let base_x = cell_x + pad_l + line_indent + align_offset + rx + aki_leading;
                                                        let base_baseline = content_h + cell_text_y_off - 1.0 + base_m.win_ascent * base_pt;
                                                        let ruby_baseline = base_baseline - raise_pt;
                                                        let mut ruby_el = LayoutElement::new(
                                                            base_x + dx,
                                                            content_h,
                                                            ruby_w + ruby_cs.max(0.0) * ruby_n as f32,
                                                            hps_pt * 1.2,
                                                            LayoutContent::Text {
                                                                text: ruby_text.to_string(),
                                                                font_size: hps_pt,
                                                                font_family: ruby_family,
                                                                bold: false,
                                                                italic: false,
                                                                underline: false,
                                                                underline_style: None,
                                                                strikethrough: false,
                                                                double_strikethrough: false,
                                                                color: color.clone(),
                                                                highlight: None,
                                                                character_spacing: ruby_cs,
                                                                field_type: None,
                                                                text_scale: 100.0,
                                                                is_vertical: false,
                                                                effects: TextEffects::default(),
                                                            },
                                                        );
                                                        ruby_el.text_y_off = ruby_baseline - content_h + 1.0 - ruby_m.win_ascent * hps_pt;
                                                        ruby_el.paragraph_index = block_idx;
                                                        ruby_el.cell_row_index = Some(row_idx);
                                                        ruby_el.cell_col_index = Some(cell_idx);
                                                        cell_elements.push(ruby_el);
                                                    }
                                                }
                                                s1626_seen += text.chars().count();
                                                if s1626_seen >= group_chars {
                                                    s1626_ri += 1;
                                                    s1626_seen = 0;
                                                }
                                            }
                                            rx += adj_w + frag_spacing[frag_idx];
                                        }
                                        if let Some(frame) = float_frame {
                                            float_row_height = float_row_height.max(lh);
                                            if frame.advance > 0.0 || line_idx + 1 == total_lines {
                                                content_h += float_row_height;
                                                float_row_height = 0.0;
                                            }
                                        } else { content_h += lh; }
                                    }
                                    content_h += effective_space_after.unwrap_or(0.0);
                                    // S427: record this paragraph's space_after so the next
                                    // cell paragraph can collapse its space_before against it.
                                    prev_cell_sa = Some(effective_space_after.unwrap_or(0.0));
                                    // R7.32: increment after each Paragraph block in the cell
                                    cell_para_counter += 1;
                                    // R7.73: track if THIS paragraph had LRPB on a non-run-0
                                    // run, so the NEXT paragraph in this cell can be marked
                                    // as a mid-cell row-split anchor.
                                    prev_cell_para_had_mid_lrpb = para
                                        .runs
                                        .iter()
                                        .enumerate()
                                        .any(|(i, r)| i > 0 && r.has_last_rendered_page_break);

                                    // PROBE (2026-07-01, OXI_CELLBUMP; default 0.0 = byte-identical):
                                    // an exact-lineRule cell line immediately FOLLOWED BY a default(auto)
                                    // line renders ~+2.4pt taller in Word (Word absolute-grid-snaps the
                                    // default line; the off-grid exact lines leave a gap the preceding
                                    // exact line absorbs). Oxi clips it to the exact value. tokyoshugyo
                                    // 割増賃金 fraction spacers (the empty line before each ×factor): ~+2.4
                                    // × 5 blocks ≈ +12.45 over page 51 → pushes ③深夜労働 off (= the −1×10
                                    // cascade origin). See [[tokyoshugyo_wrap_not_cellheight]] 2026-07-01.
                                    // SCOPE (2026-07-01): fire ONLY on the 賃金-style multi-page
                                    // SPLITTING single-cell regulation box (row taller than a page →
                                    // must split). This EXCLUDES the tokumei FORM family (multi-cell or
                                    // page-fitting fixed/atLeast rows) where the placement-only bump
                                    // (no matching row-height growth) shifts content within fixed rows
                                    // and regresses SSIM (31420 −0.1000, 15076/191cb5/order_09). The
                                    // 賃金 box is single-cell + row_height > content_height (spans pages).
                                    // S710 (2026-07-01) — RETIRED to default 0.0 (2026-07-02, S719b):
                                    // the flat +2.4/block bump approximated ONE missing 12pt line —
                                    // the whitespace-only exact-240 spacer the overflow-loop straddle
                                    // re-base displaced above page_top (and S570's trim()-empty
                                    // collapse dropped at Step-1 splits). Ink-chain decomposition:
                                    // Word STACKS these cell lines contiguously (exact→default
                                    // transition gap ≈ 0.05 on BOTH 月給制 [Σexact=36≡0 mod 18] and
                                    // 日給制 [Σexact=48≡12] blocks — the "absolute-grid-snap the
                                    // default line" model and g(Σ mod pitch) are both falsified);
                                    // the per-block bump over-pushed p50 by +2.4×6 = wi=1603 (+1).
                                    // OXI_CELLBUMP=<pt> kept as an A/B probe knob.
                                    if effective_line_rule == Some("exact")
                                        && row_line_pitch.is_some()
                                        && row.cells.len() == 1
                                        && row_height > content_height
                                    {
                                        let bump = std::env::var("OXI_CELLBUMP")
                                            .ok()
                                            .and_then(|v| v.parse::<f32>().ok())
                                            .unwrap_or(0.0);
                                        if bump != 0.0 {
                                            let next_default = cell
                                                .blocks
                                                .get(block_pos + 1)
                                                .map_or(false, |nb| match nb {
                                                    Block::Paragraph(np) => {
                                                        let nlr = if np
                                                            .style
                                                            .line_spacing_from_doc_defaults
                                                        {
                                                            None
                                                        } else {
                                                            np.style.line_spacing_rule.as_deref()
                                                        };
                                                        nlr != Some("exact")
                                                            && nlr != Some("atLeast")
                                                    }
                                                    _ => false,
                                                });
                                            if next_default {
                                                content_h += bump;
                                            }
                                        }
                                    }

                                    // Render shapes attached to this paragraph (e.g. bracketPair)
                                    // pos.y = offset from paragraph start (Word COM confirmed)
                                    for shape in &para.shapes {
                                        if let Some(ref pos) = shape.position {
                                            let content =
                                                shape_fill_boxrect(shape).unwrap_or_else(|| {
                                                    LayoutContent::PresetShape {
                                                        shape_type: shape.shape_type.clone(),
                                                        stroke_color: shape.stroke_color.clone(),
                                                        stroke_width: shape
                                                            .stroke_width
                                                            .unwrap_or(0.5),
                                                        flip_h: shape.flip_h,
                                                        flip_v: shape.flip_v,
                                                        arrow_head: shape.arrow_head,
                                                        arrow_tail: shape.arrow_tail,
                                                    }
                                                });
                                            // S711b: o:allowincell="f" shapes escape the cell — anchor x at
                                            // the page text column (start_x), not the cell content edge.
                                            // (注) gray box: Word = page-margin(85.05)+margin-left(50.2), NOT
                                            // cell_x+pad (measured Oxi was +11.52pt too far right).
                                            let escapes = shape.escapes_cell
                                                && std::env::var("OXI_VMLRECT_DISABLE").is_err();
                                            let sx = if escapes {
                                                start_x + pos.x
                                            } else {
                                                cell_x + pad_l + pos.x
                                            };
                                            // S711b vertical: an allowincell=f shape's margin-top is measured
                                            // from the paragraph TOP (before space_before), not from
                                            // para_content_start_h (which already added space_before). The (注)
                                            // box was +5.76pt too LOW (= the note para's 6pt space_before).
                                            let sy = if escapes {
                                                para_content_start_h - effective_space_before
                                                    + pos.y
                                            } else {
                                                para_content_start_h + pos.y
                                            };
                                            cell_elements.push(LayoutElement::new(
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
                            Block::Image(img) if cell_float_flow && img.position.is_some() => {
                                let (x, y) = fixed_float_positions.and_then(|positions| positions.get(block_pos))
                                    .and_then(|position| *position).unwrap_or_else(||
                                        LayoutEngine::cell_float_position(img, &float_tops, (cell_w - pad_l - pad_r).max(0.0)));
                                let mut element = LayoutElement::new(cell_x + pad_l + x, y,
                                    img.width, img.height, LayoutContent::Image {
                                        data: img.data.clone(), content_type: img.content_type.clone(),
                                        crop: img.crop.as_ref().map(|c| (c.top, c.right, c.bottom, c.left)),
                                    });
                                element.paragraph_index = block_idx;
                                element.cell_row_index = Some(row_idx);
                                element.cell_col_index = Some(cell_idx);
                                element.flow_line_height = Some(0.0);
                                // A positioned cell image stays with its reference origin.
                                // Preserve its painted offset when the span moves to a new page.
                                let is_margin_rel = img.position.as_ref().map_or(false, |p| p.v_relative.as_deref() == Some("margin"));
                                element.margin_float = is_margin_rel;
                                element.cell_float_row_bound = !img.allow_cell_overflow
                                    && matches!(img.wrap_type, Some(WrapType::Square | WrapType::Tight | WrapType::TopAndBottom));
                                // S1555: remember the anchor paragraph (cell-local block
                                // index) so the row split can ask whether it sits in the
                                // first fragment.
                                if is_margin_rel {
                                    element.cell_paragraph_index = cell.blocks.get(img.anchor_block_index)
                                        .filter(|b| matches!(b, Block::Paragraph(_)))
                                        .map(|_| cell.blocks[..img.anchor_block_index].iter()
                                            .filter(|b| matches!(b, Block::Paragraph(_))).count());
                                }
                                let origin = if is_margin_rel {
                                    float_replay.map_or(0.0, |r| r.origin(row_idx, cell_idx, img))
                                } else {
                                    float_tops.get(img.anchor_block_index).copied().unwrap_or(0.0)
                                };
                                element.flow_line_offset = (y - origin).max(0.0);
                                element.content_fit_height = Some(element.flow_line_offset + img.height);
                                cell_elements.push(element);
                            }
                            Block::Image(img) => {
                                let (image_sb, image_sa) = self.cell_image_spacing(
                                    img, table, row_line_pitch, &cell.blocks, block_pos,
                                    prev_cell_sa, s939_prev_r, s1075_prev_r,
                                );
                                content_h += image_sb;
                                // S533 (2026-06-10): place inline images inside table cells.
                                // The parser forwards cell-paragraph inline images as sibling
                                // Block::Image (S331, default ON as of S533); without this arm
                                // the image occupied no height and emitted no element, so an
                                // image-bearing cell collapsed to its text height (3a4f p34:
                                // the 321.75pt year-calendar EMF cell rendered ~28pt, pulling
                                // ~7 paragraphs up a page = the Phase-1 sole FAIL).
                                let effect_top = if img.position.is_none() {
                                    img.effect_extent_t.max(0.0)
                                } else { 0.0 };
                                cell_elements.push(LayoutElement::new(
                                    cell_x + pad_l,
                                    content_h + effect_top,
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
                                // S715: typed-grid cell image line snaps to whole grid cells
                                // (mirrors the pre-pass arm; see comment there).
                                let img_line =
                                    self.s971_image_line_h(img, 1.0e6, row_line_pitch, false);
                                let img_h_eff = if std::env::var("OXI_S715_DISABLE").is_err() {
                                    match row_line_pitch {
                                        Some(p) if p > 0.0 => (img_line / p).ceil() * p,
                                        _ => img_line,
                                    }
                                } else {
                                    img_line
                                };
                                if img.position.is_none() || std::env::var("OXI_ROW_CONTINUATION_FLOW").is_ok() {
                                    if let Some(element) = cell_elements.last_mut() {
                                        element.flow_line_offset = effect_top;
                                        // Fit the entire image line, but keep paragraph
                                        // after-spacing outside this atomic boundary.
                                        if img.position.is_none() {
                                            element.content_fit_height = Some(img_h_eff);
                                        }
                                        element.paragraph_index = block_idx;
                                        element.cell_row_index = Some(row_idx);
                                        element.cell_col_index = Some(cell_idx);
                                        element.flow_line_height = Some(if LayoutEngine::s1053_cell_float_no_reserve(img) {
                                            0.0
                                        } else { img_h_eff + image_sa });
                                    }
                                }
                                // S1053: a page-relative float still DRAWS (the element is
                                // pushed above) but reserves no flow height.
                                if !LayoutEngine::s1053_cell_float_no_reserve(img) {
                                    content_h += img_h_eff + image_sa;
                                }
                                prev_cell_sa = img.host_paragraph.as_ref().map(|_| image_sa);
                                s939_prev_r = img.host_paragraph.as_ref().map(|host| (host.style.contextual_spacing, host.style.style_id.as_deref()));
                                s1075_prev_r = img.host_paragraph.as_ref().map(|host| (host.style.after_autospacing, host.style.num_id.as_deref()));
                            }
                            _ => {}
                        } // match block
                    } // for block
                } // if !is_vmerge_continue

                if cell_float_flow {
                    let obstacles = if let Some(positions) = fixed_float_positions {
                        LayoutEngine::cell_float_obstacles_at(cell, positions)
                    } else {
                        LayoutEngine::cell_float_obstacles(cell, &float_tops, cell.blocks.len(), (cell_w - pad_l - pad_r).max(0.0))
                    };
                    let bottom = obstacles.iter().map(|obstacle| obstacle.bottom).fold(0.0_f32, f32::max);
                    content_h = content_h.max(bottom);
                }

                // Lay out the merged text across page bands before its final
                // row consumes the remaining height. Repeated headers occupy
                // the top of every continuation band.
                // S1433 (2026-09-16, default ON, opt-out OXI_S1433_DISABLE): the
                // checkpoint's OXI_VMERGE_PAGE_FLOW opt-in promoted. A vMerge
                // restart cell's text is paginated line by line against the
                // page bands (line bottom must fit; keepLines / spacing honoured)
                // instead of being cut by the row split of a later row -- the
                // legacy path kept a line whose TOP was above the bottom
                // (reports__1c313df3 p2: 「・入学前の児童」 line 1 at 752.25 with an
                // 18pt box against 756.85, Word moves the paragraph whole; the
                // misplaced line then shifted the continuation and the table
                // end by ~50pt). JA 189 -> 190, no PASS -> FAIL.
                let vmerge_page_flow = (std::env::var("OXI_VMERGE_PAGE_FLOW").as_deref() == Ok("1")
                        || std::env::var_os("OXI_S1433_DISABLE").is_none())
                    && cell.v_merge.as_deref() == Some("restart")
                    && !is_nested && row_footnotes.is_none()
                    && !matches!(cell.v_align.as_deref(), Some("center") | Some("bottom"))
                    && cell.blocks.iter().all(|b| matches!(b, Block::Paragraph(_)))
                    && content_height > s728_hdr_h + pad_t + pad_b;
                if vmerge_page_flow {
                    let origin = cursor.visual_y + pad_t;
                    let key = if std::env::var_os("OXI_S1192G_DISABLE").is_none() { cell_start_grid } else { cell_idx };
                    let mut flow = MergedCellTextFlow {
                        identity: std::sync::Arc::new(()), key, start_row: row_idx,
                        source_page: pages.len(), origin, page_top, page_height: content_height,
                        coordinate_stride: vmerge_coordinate_stride,
                        header: if s728_on { s728_hdr_h } else { 0.0 },
                        pad_top: pad_t, pad_bottom: pad_b, content_height: content_h,
                        element_count: cell_elements.len(), paragraphs: Vec::new(),
                        cuts: std::collections::BTreeMap::new(),
                    };
                    for (cpi, block) in cell.blocks.iter().enumerate() {
                        let Block::Paragraph(para) = block else { continue; };
                        let mut lines: Vec<MergedCellFlowLine> = Vec::new();
                        for (i, e) in cell_elements.iter().enumerate() {
                            if e.cell_paragraph_index != Some(cpi)
                                || !matches!(e.content, LayoutContent::Text { .. }) { continue; }
                            let y = e.y - e.flow_line_offset;
                            let fit = e.content_fit_height.unwrap_or(e.height);
                            if let Some(line) = lines.iter_mut().find(|line| (line.y - y).abs() < 0.01) {
                                line.fit = line.fit.max(fit);
                                line.elements.push((i, e.y));
                            } else {
                                lines.push(MergedCellFlowLine { y, fit, elements: vec![(i, e.y)] });
                            }
                        }
                        lines.sort_by(|a, b| a.y.total_cmp(&b.y));
                        flow.paragraphs.push(MergedCellFlowParagraph { lines,
                            keep_lines: para.style.keep_lines,
                            // Merged-cell widow protection is enabled by modern
                            // document compatibility. Legacy and unspecified modes
                            // allow a single line on either side of the page cut.
                            widow_control: para.style.widow_control
                                && self.compat_mode_explicit && self.compat_mode >= 15,
                            before: para.style.space_before.unwrap_or(0.0),
                            after: para.style.space_after.unwrap_or(0.0),
                        });
                    }
                    let (positions, end) = flow.paginate();
                    for (i, position) in positions.into_iter().enumerate() {
                        if let Some((page, y)) = position {
                            let offset = page - pages.len();
                            cell_elements[i].y = y + offset as f32 * vmerge_coordinate_stride - origin;
                            cell_elements[i].vmerge_destination_page = (offset > 0).then_some(page);
                            cell_elements[i].vmerge_flow_element = Some((flow.identity.clone(), i));
                        }
                    }
                    content_h = end - pages.len() as f32 * vmerge_coordinate_stride - origin - pad_b;
                    vmerge_absolute_ends.insert(key, end);
                    vmerge_text_flows.push(flow);
                }

                if std::env::var("OXI_CELL_TEXT_CLIP_DISABLE").is_err() && !self.doc_body_has_real_cjk {
                    for e in &mut cell_elements {
                        if matches!(e.content, LayoutContent::Text { .. }) {
                            let bounds = (cell_x - e.x, cell_x + cell_w - e.x);
                            e.horizontal_clip = Some(match e.horizontal_clip {
                                Some((l, r)) => (l.max(bounds.0), r.min(bounds.1)),
                                None => bounds,
                            });
                        }
                    }
                }

                // Track actual cell height for row_height correction.
                // content_h is the sum of per-paragraph layout heights; elements may
                // be positioned with a text_y_offset (vertical centering inside the
                // line box), which makes their bottom extend past content_h. Using
                // max(content_h, elem_bottom) double-counts this offset and inflates
                // row height. Trust content_h as the authoritative sum.
                let is_vmerge_restart = cell.v_merge.as_deref() == Some("restart");
                // S946 (2026-07-19, opt-out OXI_S946_DISABLE): an INK-LESS cell
                // (all blocks are paragraphs with only whitespace runs) never
                // drives the S648 row-height correction — its placement height
                // uses the S503 render line (GDI, e.g. 9.7 for a 7.5pt mark)
                // which exceeds the Word-correct estimate (mark hhea 8.6), and
                // nothing is painted to justify the growth. NDIS zones tables:
                // the blank enclave-column cells inflated every such row
                // +1.1 (Oxi 14.2 vs Word 13.15). Latin scope: the JP form
                // family's empty-cell heights are a calibrated stack.
                let s946_inkless = std::env::var("OXI_S946_DISABLE").is_err()
                    && !self.doc_body_has_real_cjk
                    && cell.blocks.iter().all(|b| {
                        matches!(b,
                        Block::Paragraph(p) if p.runs.iter().all(|r| r.text.trim().is_empty()))
                    });
                if !is_vmerge_continue && !is_exact_row && !is_vmerge_restart && !s946_inkless {
                    let actual = pad_t + content_h + pad_b;
                    if actual > max_actual_cell_h {
                        max_actual_cell_h = actual;
                    }
                }
                // S1192b: remember what the merged cell needs. It deliberately
                // does NOT grow its own row (the arm above excludes it, which is
                // Word's behaviour); the span's last row pays instead.
                if is_vmerge_restart && std::env::var("OXI_S1192_DISABLE").is_err() {
                    let need = pad_t + content_h + pad_b;
                    // S1192c (default ON since S1609, opt-out `OXI_S1192G_DISABLE`): key the pending
                    // need on the GRID COLUMN, not the cell's position in its
                    // row. A row's cell list is not a column identity once
                    // `gridSpan` or `gridBefore` are in play, and ed025's vMerge
                    // tables have rows of 4/5/6 and 6/7/11/13 cells. A census of
                    // the corpus says the shape is common — 44 of the 74 vMerge
                    // tables mix a span/gridBefore with rows of differing cell
                    // counts — but the grid key changed NOTHING on ed025, and it
                    // has never been through a corpus gate. Verify before making
                    // it the only path.
                    let key = if std::env::var_os("OXI_S1192G_DISABLE").is_none() {
                        cell_start_grid
                    } else {
                        cell_idx
                    };
                    if !vmerge_page_flow { vmerge_absolute_ends.remove(&key); }
                    s1192_pending.retain(|(c, _)| *c != key);
                    s1192_pending.push((key, need));
                    if std::env::var("OXI_DBG_S1192").is_ok() {
                        eprintln!("[S1192] REC row={} cell_idx={} grid={} need={:.2}",
                            row_idx, cell_idx, cell_start_grid, need);
                    }
                }

                // Apply vAlign offset
                // Session 79c (2026-05-17): use visual_row_h (pre-computed from
                // emit-equivalent estimate_para_height_emit) when greater than
                // row_height. visual_row_h reflects the actual emitted cell
                // content height (grid-snapped under adjustLineHeightInTable),
                // matching what Word centers within. row_height (page-break
                // logic) preserved to avoid 3a4f9f cascade. Also fall back to
                // max_actual_cell_h in case visual_row_h underestimated for
                // unusual cells (defensive).
                let mut effective_row_h = row_height.max(visual_row_h).max(max_actual_cell_h);
                // S503 (2026-06-08): include the render-line-height centering floor so a
                // vAlign=center cell that is laid out BEFORE a taller cell (col0 before
                // col1) centers within the FULL actual row content height, not the
                // under-counting estimate. Centering (v_offset) only — row_height (above,
                // pagination) is unchanged. Opt-in OXI_S503_ENABLE (default OFF).
                if s503_enable {
                    effective_row_h = effective_row_h.max(center_row_h);
                }
                // S217 (2026-05-23): vmerge=restart cells with vAlign should
                // center across the FULL merged span, not just row 0.
                // Look ahead to count span rows where the same grid column has
                // vmerge=continue. When ALL rows in the span have trHeight set,
                // use sum of trHeights as effective_row_h (covers 7ead's table).
                // Falls back to current row 0 behavior otherwise.
                // See session216_vmerge_valign_bug_confirmed.md.
                //
                // S218 (2026-05-23) DEFAULT ON: when some span rows lack
                // trHeight, compute natural row height for those rows via
                // estimate_table_row_natural_h (mirrors the main loop's
                // row-height pre-pass). Affects 459f05 p2 (4 cells,
                // matcher-detected +0.0110) and b5f706 p2 (11 cells, matcher-
                // invisible due to MIN_MATCH_LEN=2 on single-char "丸"
                // markers). Phase 1 53/55 unchanged, 0 IoU regressions.
                //
                // S236 (2026-05-23) removed OXI_LEGACY_VMERGE_VALIGN_ROW0 and
                // OXI_LEGACY_VMERGE_VALIGN_STRICT legacy env-var fallbacks
                // during hardening pass; both gates have been stable since
                // ship (~17-18 sessions).
                let is_vmerge_restart_for_valign = cell.v_merge.as_deref() == Some("restart");
                if is_vmerge_restart_for_valign && cell.v_align.is_some() {
                    let target_grid = cell_start_grid;
                    let mut span_count = 1usize;
                    for next_ri in (row_idx + 1)..num_rows {
                        let next_row = &table.rows[next_ri];
                        let mut next_grid = next_row.grid_before as usize;
                        let mut continues = false;
                        for next_cell in &next_row.cells {
                            let next_span = next_cell.grid_span.max(1) as usize;
                            if next_grid == target_grid {
                                if matches!(
                                    next_cell.v_merge.as_deref(),
                                    Some("continue") | Some("")
                                ) {
                                    continues = true;
                                }
                                break;
                            }
                            next_grid += next_span;
                        }
                        if continues {
                            span_count += 1;
                        } else {
                            break;
                        }
                    }
                    if span_count > 1 {
                        let mut span_h: f32 = 0.0;
                        let mut all_have_h = true;
                        for ri in row_idx..(row_idx + span_count) {
                            if let Some(h) = table.rows[ri].height {
                                span_h += h;
                            } else {
                                all_have_h = false;
                                break;
                            }
                        }
                        if all_have_h {
                            // S305 (2026-05-26) [opt-in via OXI_S305_ENABLE]:
                            // for the trHeight-declared path, use natural row
                            // height as an atLeast floor when content overflow
                            // is detected on any subsequent row. Pre-fix path
                            // (default) sums declared trHeight only.
                            //
                            // Wins (when enabled): 31420af1a08f mean_iou
                            // 0.8697 → 0.9471 (+0.0774) — cell (7,1)/(9,1)
                            // "物理的管理措置"/"技術的管理措置" vMerge=restart
                            // headers move from cell top to merged-span center
                            // matching Word.
                            //
                            // Losses (why kept opt-in): 3a4f9fbe1a83 and
                            // ed025cbecffb regress Phase 1 pagination (PASS
                            // → FAIL, 39 + 1 paragraph page_delta=-1) when
                            // the gate fires on their tables — the relaxed
                            // span height moves a row's vMerge content far
                            // enough that downstream layout cascades push a
                            // few paragraphs to earlier pages. Need a tighter
                            // discriminator before flipping default.
                            let opt_in = std::env::var("OXI_S305_ENABLE").is_ok();
                            const OVERFLOW_GATE_PT: f32 = 20.0;
                            let mut should_relax = false;
                            if opt_in {
                                for ri in (row_idx + 1)..(row_idx + span_count) {
                                    let r = &table.rows[ri];
                                    if let (Some(h), rule) = (r.height, r.height_rule.as_deref()) {
                                        if rule == Some("exact") {
                                            continue;
                                        }
                                        let nat = self.estimate_table_row_natural_h(
                                            r,
                                            &col_widths,
                                            default_pad_l,
                                            default_pad_r,
                                            default_pad_t,
                                            default_pad_b,
                                            table,
                                            table_grid_pitch,
                                            grid_char_pitch,
                                            grid_char_cw_ratio,
                                        );
                                        if nat > h + OVERFLOW_GATE_PT {
                                            should_relax = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            if should_relax {
                                let mut relaxed_span_h: f32 = 0.0;
                                for ri in row_idx..(row_idx + span_count) {
                                    let r = &table.rows[ri];
                                    let eff_h = if ri == row_idx {
                                        effective_row_h
                                    } else {
                                        let nat = self.estimate_table_row_natural_h(
                                            r,
                                            &col_widths,
                                            default_pad_l,
                                            default_pad_r,
                                            default_pad_t,
                                            default_pad_b,
                                            table,
                                            table_grid_pitch,
                                            grid_char_pitch,
                                            grid_char_cw_ratio,
                                        );
                                        match (r.height, r.height_rule.as_deref()) {
                                            (Some(h), Some("exact")) => h,
                                            // ROWBOX2: binding atLeast = trH + bw
                                            (Some(h), _) => {
                                                nat.max(h + self.rowbox2_trh_bw(table, r))
                                            }
                                            (None, _) => nat,
                                        }
                                    };
                                    relaxed_span_h += eff_h;
                                }
                                effective_row_h = effective_row_h.max(relaxed_span_h);
                            } else {
                                effective_row_h = effective_row_h.max(span_h);
                            }
                        } else {
                            // S218 relax: compute span height when trHeight missing.
                            // S222 (2026-05-23): for row_idx, use the already-computed
                            // `effective_row_h` (= max of row_height, visual_row_h,
                            // max_actual_cell_h — matches emit). Pre-S222 used
                            // `row_height` alone (pre-pass natural, non-grid-snap),
                            // which underestimated by ~14pt for 2-line grid-snapped
                            // cells (b5f706 p2 row 1: 21.75 vs 36.5). For future
                            // span rows, helper now also uses
                            // `estimate_para_height_emit` so its natural h matches
                            // emit grid-snap. S220 attempted this but was blocked
                            // by derive_oxi_heights noise; S221 resolved that.
                            let mut relaxed_span_h: f32 = 0.0;
                            for ri in row_idx..(row_idx + span_count) {
                                let r = &table.rows[ri];
                                let eff_h = if ri == row_idx {
                                    effective_row_h
                                } else {
                                    let nat = self.estimate_table_row_natural_h(
                                        r,
                                        &col_widths,
                                        default_pad_l,
                                        default_pad_r,
                                        default_pad_t,
                                        default_pad_b,
                                        table,
                                        table_grid_pitch,
                                        grid_char_pitch,
                                        grid_char_cw_ratio,
                                    );
                                    match (r.height, r.height_rule.as_deref()) {
                                        (Some(h), Some("exact")) => h,
                                        // ROWBOX2: binding atLeast = trH + bw
                                        (Some(h), _) => nat.max(h + self.rowbox2_trh_bw(table, r)),
                                        (None, _) => nat,
                                    }
                                };
                                relaxed_span_h += eff_h;
                            }
                            effective_row_h = effective_row_h.max(relaxed_span_h);
                        }
                    }
                }
                // S647 (2026-06-23) tested+reverted: a vAlign=center glyph-vs-line-box
                // correction (Word centers the glyph cell ~1.06×fs, Oxi centers content_h
                // = line box ~1.29×fs) gained tokumei only +0.0014 (peak LEAD=0.05) — most
                // form content is vAlign=TOP, not center, so this is a tiny component, not
                // the form-family ~1pt drift. See [[tokumei_form_family_ssim]].
                let v_offset = match cell.v_align.as_deref() {
                    Some("center") => {
                        ((effective_row_h - pad_t - pad_b - content_h) / 2.0).max(0.0)
                    }
                    Some("bottom") => (effective_row_h - pad_t - pad_b - content_h).max(0.0),
                    _ => 0.0, // top (default)
                };

                // Emit cell elements with absolute Y positions
                let dy = cursor.visual_y + pad_t + v_offset;
                if dump_table {
                    let valign = cell.v_align.as_deref().unwrap_or("(top)");
                    eprintln!(
                        "[TBL_DUMP]   row={} cell={} cursor_y={:.3} pad_t={:.3} pad_b={:.3} content_h={:.3} v_align={} v_offset={:.3} dy={:.3} row_h={:.3}",
                        row_idx, cell_idx, cursor.cursor_y, pad_t, pad_b, content_h, valign, v_offset, dy, row_height
                    );
                }
                if fragment_valign && cell.v_merge.is_none() && cell.cell_text_boxes.is_empty() {
                    let factor = match cell.v_align.as_deref() {
                        Some("center") => 0.5,
                        Some("bottom") => 1.0,
                        _ => 0.0,
                    };
                    let start = elements.len() - elements_before_row;
                    fragment_valign_cells.push((start..start + cell_elements.len(),
                        cursor.visual_y + pad_t, content_h, pad_b, factor));
                }
                let is_vmerge_restart = cell.v_merge.as_deref() == Some("restart");
                for mut elem in cell_elements {
                    if cell_float_flow && elem.cell_ancestor_path.is_empty()
                        && matches!(elem.content, LayoutContent::Text { .. }) {
                        if let Some(replay) = float_replay {
                            elem.cell_float_fragment = Some((replay.identity.clone(), elem.y - elem.flow_line_offset));
                        }
                    }
                    elem.y += dy;
                    // Also update y-coords inside TableBorder content (nested tables)
                    if let LayoutContent::TableBorder {
                        ref mut y1,
                        ref mut y2,
                        ..
                    } = elem.content
                    {
                        *y1 += dy;
                        *y2 += dy;
                    }
                    // R7.61 (Day 36 part 8): mark vMerge=restart cell text content
                    // that overflows the page bottom. Post-paginate sweep moves
                    // these to next page (a1d6 ※２/※３ on row 13 cell[0]).
                    // Only text elements (skip borders/shading). cell_paragraph_index
                    // > 0 condition prevents the cell's first paragraph from being
                    // moved (it anchors the cell to its row's page).
                    if is_vmerge_restart
                        && matches!(&elem.content, LayoutContent::Text { .. })
                        && elem.vmerge_destination_page.is_none()
                        && elem.y > page_bottom + 0.5
                        && elem.cell_paragraph_index.map_or(false, |cpi| cpi > 0)
                    {
                        elem.vmerge_restart_overflow_to_next_page = true;
                    }
                    if v_offset > 0.0 && cell.v_merge.is_none() && cell.cell_text_boxes.is_empty() {
                        split_valign_offsets.push((elements.len() - elements_before_row, v_offset));
                    }
                    elements.push(elem);
                }

                // Draw cell borders if table has borders OR cell has its own borders
                let has_cell_borders = cell.borders.as_ref().map_or(false, |b| {
                    b.top.is_some() || b.bottom.is_some() || b.left.is_some() || b.right.is_some()
                });
                if table.style.border || has_cell_borders
                    || inherited_top_rule.is_some() || inherited_bottom_rule.is_some() {
                    let bx = cell_x;
                    let by = cursor.visual_y;

                    // Resolve border color, width and S480 style from cell borders,
                    // falling back to table style.
                    let resolve_border =
                        |side: Option<&BorderDef>,
                         table_width: f32|
                         -> (Option<String>, f32, Option<String>) {
                            if let Some(b) = side {
                                // S482: explicit w:val="nil"/"none" cell edge SUPPRESSES
                                // the border (do NOT fall through to the table border).
                                if b.style == "none" {
                                    return (None, 0.0, None);
                                }
                                let c = b.color.as_ref().map(|c| {
                                    if c.starts_with('#') {
                                        c.clone()
                                    } else {
                                        format!("#{}", c)
                                    }
                                });
                                (c, b.width, Some(b.style.clone()))
                            } else if table.style.border {
                                // Table-level borders: use table style color, default to black
                                let c = Some(
                                    table
                                        .style
                                        .border_color
                                        .as_ref()
                                        .map(|c| {
                                            if c.starts_with('#') {
                                                c.clone()
                                            } else {
                                                format!("#{}", c)
                                            }
                                        })
                                        .unwrap_or_else(|| "#000000".to_string()),
                                );
                                (c, table_width, table.style.border_style.clone())
                            } else {
                                (None, 0.4, None)
                            }
                        };

                    let cell_borders = cell.borders.as_ref();
                    // S921: a direct tblBorders frame can coexist with thinner
                    // insideH/insideV edges inherited from the table style.
                    let outer_width = table.style.border_width.unwrap_or(0.4);
                    let s921_inner_widths = std::env::var("OXI_S921_DISABLE").is_err();
                    let horizontal_width = if row_idx + 1 == num_rows || !s921_inner_widths {
                        outer_width
                    } else {
                        table
                            .style
                            .inside_horizontal_border
                            .as_ref()
                            .filter(|b| b.style != "none")
                            .map(|b| b.width)
                            .unwrap_or(outer_width)
                    };
                    let vertical_width = if cell_idx + 1 == row.cells.len() || !s921_inner_widths {
                        outer_width
                    } else {
                        table
                            .style
                            .inside_vertical_border
                            .as_ref()
                            .filter(|b| b.style != "none")
                            .map(|b| b.width)
                            .unwrap_or(outer_width)
                    };
                    let (top_color, top_width, top_style) = resolve_border(
                        cell_borders.and_then(|b| b.top.as_ref()).or(inherited_top_rule.as_ref()),
                        if row_idx == 0 {
                            outer_width
                        } else {
                            horizontal_width
                        },
                    );
                    let (bot_color, bot_width, bot_style) = resolve_border(
                        cell_borders.and_then(|b| b.bottom.as_ref()).or_else(|| {
                            if row_idx + 1 == num_rows { inherited_bottom_rule.as_ref() } else { None }
                        }),
                        horizontal_width,
                    );
                    let (left_color, left_width, left_style) = resolve_border(
                        cell_borders.and_then(|b| b.left.as_ref()),
                        if cell_idx == 0 {
                            outer_width
                        } else {
                            vertical_width
                        },
                    );
                    let (right_color, right_width, right_style) =
                        resolve_border(cell_borders.and_then(|b| b.right.as_ref()), vertical_width);

                    // When cells have their own borders (tcBorders), draw each side per cell.
                    // When using table-level borders, use collapsed model to avoid double-drawing.
                    let use_collapsed = table.style.border && !has_cell_borders;

                    // Top — skip for vMerge continue cells (internal to merged range)
                    // S585c: draw the box borders at the clamped width (eff_cell_w) so the
                    // right border matches Word's position (content_right + cellMar).
                    if !is_vmerge_continue
                        && top_color.is_some()
                    {
                        let edge = LayoutElement::new(
                            bx,
                            by - s1621_lift,
                            eff_cell_w,
                            0.0,
                            LayoutContent::TableBorder {
                                x1: bx,
                                y1: by - s1621_lift,
                                x2: bx + eff_cell_w,
                                y2: by - s1621_lift,
                                color: top_color,
                                width: top_width,
                                style: top_style,
                            },
                        );
                        // Collapsed borders are painted by the preceding row,
                        // but still define this row's continuation-page edge.
                        row_declared_top_edges.push(edge.clone());
                        let inherited_only = cell_borders.and_then(|b| b.top.as_ref()).is_none()
                            && inherited_top_rule.is_some();
                        if (!use_collapsed || row_idx == 0)
                            && (!inherited_only || row_idx == 0 || pages.len() > inherited_outer_page_before_row) {
                            elements.push(edge);
                        }
                    }
                    // Bottom — skip for vMerge continue cells unless next row is not continue
                    let next_is_continue = if row_idx + 1 < num_rows {
                        table.rows[row_idx + 1]
                            .cells
                            .get(cell_idx)
                            .map_or(false, |nc| {
                                nc.v_merge.as_deref() == Some("continue")
                                    || nc.v_merge.as_deref() == Some("")
                            })
                    } else {
                        false
                    };
                    if bot_color.is_some() && !next_is_continue {
                        elements.push(LayoutElement::new(
                            bx,
                            by + row_height,
                            eff_cell_w,
                            0.0,
                            LayoutContent::TableBorder {
                                x1: bx,
                                y1: by + row_height,
                                x2: bx + eff_cell_w,
                                y2: by + row_height,
                                color: bot_color,
                                width: bot_width,
                                style: bot_style,
                            },
                        ));
                    }
                    // Left
                    if left_color.is_some() && (!use_collapsed || cell_idx == 0) {
                        elements.push(LayoutElement::new(
                            bx,
                            by - s1621_lift,
                            0.0,
                            row_height + s1621_lift,
                            LayoutContent::TableBorder {
                                x1: bx,
                                y1: by - s1621_lift,
                                x2: bx,
                                y2: by + row_height,
                                color: left_color,
                                width: left_width,
                                style: left_style,
                            },
                        ));
                    }
                    // Right
                    if right_color.is_some() {
                        elements.push(LayoutElement::new(
                            bx + eff_cell_w,
                            by - s1621_lift,
                            0.0,
                            row_height + s1621_lift,
                            LayoutContent::TableBorder {
                                x1: bx + eff_cell_w,
                                y1: by - s1621_lift,
                                x2: bx + eff_cell_w,
                                y2: by + row_height,
                                color: right_color,
                                width: right_width,
                                style: right_style,
                            },
                        ));
                    }
                }

                // S488 (CLASS E step 3): emit in-cell floating text boxes with the
                // COM-derived anchor model (replaces S487's naive cell-origin +
                // posOffset). Measured on 1636d28 (tools/metrics/_s488c_anchor_clean.py):
                //   relH="column"/"character" → cell CONTENT-left (cell_x + pad_l) + posX
                //   relH="margin"             → page left margin + posX
                //   relH="page"               → posX
                //   relV="paragraph"/"line"   → anchoring paragraph's absolute top + posY
                //   relV="margin"             → page top margin + posY
                //   relV="page"               → posY
                // The paragraph top = cell_block_tops[anchor_block_index] + dy (dy
                // is the cell-content absolute base computed above). S487's bug was
                // using cell_x (border-left, not content-left) for X and
                // cursor.visual_y (cell top, not the anchor paragraph) for Y.
                // Opt-IN OXI_S487_ENABLE (default OFF until gate-validated).
                if !cell.cell_text_boxes.is_empty() && std::env::var("OXI_S487_ENABLE").is_ok() {
                    let cell_content_left = cell_x + pad_l;
                    for tb in &cell.cell_text_boxes {
                        let (px, py) = tb
                            .position
                            .as_ref()
                            .map(|p| (p.x, p.y))
                            .unwrap_or((0.0, 0.0));
                        let h_rel = tb.position.as_ref().and_then(|p| p.h_relative.as_deref());
                        let v_rel = tb.position.as_ref().and_then(|p| p.v_relative.as_deref());
                        let abs_x = match h_rel {
                            Some("page") => px,
                            Some("margin") => page.margin.left + px,
                            // column / character / default → cell content-left
                            _ => cell_content_left + px,
                        };
                        let abs_y = match v_rel {
                            Some("page") => py,
                            Some("margin") => page.margin.top + py,
                            // paragraph / line / default → anchor paragraph top
                            _ => {
                                let para_top_rel = cell_block_tops
                                    .get(tb.anchor_block_index)
                                    .copied()
                                    .unwrap_or(0.0);
                                para_top_rel + dy + py
                            }
                        };
                        let tb_elems = self.layout_text_box_at(tb, page, &[], &[], Some((abs_x, abs_y)));
                        // Defer (not `elements.extend`) so the box paints on top of the
                        // whole table grid — see deferred_cell_textboxes declaration.
                        deferred_cell_textboxes.extend(tb_elems);
                    }
                }

                cell_x += cell_w;
                grid_idx += span;
            }

            if dump_table {
                eprintln!(
                    "[TBL_DUMP] row={} pre_correction row_height={:.3} max_actual_cell_h={:.3}",
                    row_idx, row_height, max_actual_cell_h
                );
            }
            // S1192: the LAST row of a vMerge span must also cover whatever the
            // merged cell still needs. Folded into `max_actual_cell_h` rather
            // than into `row_height` directly, so the border fix-up right below
            // carries this row's already-emitted rules down with it — the
            // estimator hook alone changed nothing precisely because the emit
            // pass computes its own height here.
            if std::env::var("OXI_S1192_DISABLE").is_err() {
                let is_cont = |c: &TableCell| {
                    matches!(c.v_merge.as_deref(), Some("continue") | Some(""))
                };
                let bygrid = std::env::var_os("OXI_S1192G_DISABLE").is_none();
                for (ci, remaining) in s1192_pending.iter() {
                    let here_cont = LayoutEngine::s1192_cell_at(row, *ci, bygrid).map_or(false, is_cont);
                    let next_cont = table
                        .rows
                        .get(row_idx + 1)
                        .and_then(|r| LayoutEngine::s1192_cell_at(r, *ci, bygrid))
                        .map_or(false, is_cont);
                    if std::env::var("OXI_DBG_S1192").is_ok() {
                        eprintln!("[S1192] FOLD row={} key={} rem={:.2} here={} next={}",
                            row_idx, ci, remaining, here_cont, next_cont);
                    }
                    if here_cont && !next_cont {
                        let remaining = vmerge_absolute_ends.get(ci).map(|end| {
                            end - (pages.len() as f32 * vmerge_coordinate_stride + cursor.visual_y)
                        }).unwrap_or(*remaining);
                        max_actual_cell_h = max_actual_cell_h.max(remaining);
                    }
                }
            }
            // If actual content exceeds estimated row_height, fix border elements
            if max_actual_cell_h > row_height + 0.01 {
                let old_h = row_height;
                row_height = max_actual_cell_h;
                let by = cursor.visual_y;
                let old_bottom = by + old_h;
                let new_bottom = by + row_height;
                for elem in elements[elements_before_row..].iter_mut() {
                    match &mut elem.content {
                        LayoutContent::TableBorder { y1, y2, .. } => {
                            if (*y1 - old_bottom).abs() < 0.5 {
                                *y1 = new_bottom;
                            }
                            if (*y2 - old_bottom).abs() < 0.5 {
                                *y2 = new_bottom;
                            }
                            // S648 (2026-06-23): keep elem.y / elem.height in sync
                            // with the corrected content y1/y2. Both renderers draw
                            // borders from y1/y2 (GDI main.rs:393, DWrite main.rs:387),
                            // so the RENDER was already correct — but elem.y was left
                            // STALE at the pre-correction row bottom. The --dump-layout
                            // JSON emits elem.y, so any border-position diagnostic read
                            // the stale value and mis-concluded "rows render too short"
                            // (the row-height correction had in fact moved the rendered
                            // border to the right place). Render-neutral (renderers use
                            // y1/y2); pagination uses text elements only; element_iou
                            // is paragraph-derived — so this fixes the dump diagnostic
                            // ONLY, with zero gate impact.
                            elem.y = (*y1).min(*y2);
                            elem.height = (*y2 - *y1).abs();
                        }
                        LayoutContent::CellShading { .. } => {
                            if (elem.height - old_h).abs() < 0.5 {
                                elem.height = row_height;
                            }
                        }
                        _ => {}
                    }
                }
            }

            // Row splitting across pages: when the row content extends beyond
            // the current page bottom, split elements between current and next page.
            // This handles single-cell rows with many paragraphs (e.g. list boxes).
            // S1565 (2026-09-26, default ON, opt-out OXI_S1565_DISABLE): a split
            // fragment in a Latin document whose table has no table-level border
            // but whose cells declare a BOTTOM border still closes with that rule
            // -- the row-bottom law (rowfoot_pdf: boundary + bw/2 <= body bottom)
            // applied to the fragment. educational__00116bbe p16: a 7-line TNR-12
            // (x1.15) cell paragraph whose cell has tcBorders bottom sz=18 (2.25pt,
            // white); Word PDF keeps 6 lines (last baseline 764.26) and starts p17
            // with the 7th, although the 7th's box ends at 784.4 on a 785.2 page:
            // 784.4 + 1.125 > 785.2. Oxi cut at 784.79 and kept all seven, which
            // S1559's correct empty-line height then exposed (the 2pt-too-tall
            // empty above had been pushing the line off by itself).
            let s1565_cell_bottom = std::env::var_os("OXI_S1565_DISABLE").is_none()
                && !separate_outer_edges
                && !self.doc_body_has_real_cjk
                && self.s1188_on()
                && !table.style.border
                && row.cells.iter().any(|c| c.borders.as_ref().and_then(|b| b.bottom.as_ref())
                    .map_or(false, |d| d.style != "nil" && d.style != "none"));
            let fragment_bottom_width = if separate_outer_edges {
                self.table_fragment_bottom_width(table, Some(row))
            } else if s1191_foot > 0.0 && std::env::var_os("OXI_S1482_DISABLE").is_none() {
                s1191_foot
            } else if s1565_cell_bottom {
                self.table_fragment_bottom_width(table, Some(row)) * 0.5
            } else { 0.0 };
            let fragment_top_width = if separate_outer_edges {
                self.table_fragment_top_width(table, row)
            } else { 0.0 };
            let fragment_content_bottom = page_bottom - fragment_bottom_width;
            let row_bottom = cursor.cursor_y + row_height;
            if !keep_float_whole && row_bottom > fragment_content_bottom + row_fit_epsilon && !row.cant_split && !first_row_forced {
                // R7.56 (Day 34 part 25, 2026-05-13): respect mid-cell LRPB markers.
                // If any element in this row carries `is_paragraph_start_with_lrpb`,
                // force the split before the FIRST such element above page_bottom
                // (i.e., pull split_y back to that element's y so it goes to next page).
                // e3c545 cpi=81 LRPB at y=765.62: without this pull-back, the element
                // bottom (777.24) fits split_y=785.2 (page_bottom) and stays on current
                // page; with pull-back, split_y becomes 765.62 → element goes overflow.
                let mut row_elements = elements.split_off(elements_before_row);
                // A split row lays out each cell from the fragment's top.
                // Remove alignment against the unbroken row before testing
                // line fit and widow limits. Preserve unsplit rows as emitted.
                let split_valign = fragment_valign || (!self.doc_body_has_real_cjk
                    && std::env::var("OXI_SPLIT_CELL_VALIGN_DISABLE").is_err()
                    && row.cells.iter().any(|c| c.v_merge.is_none() && c.cell_text_boxes.is_empty()));
                if split_valign {
                    for &(index, offset) in &split_valign_offsets {
                        let elem = &mut row_elements[index];
                        elem.y -= offset;
                        if let LayoutContent::TableBorder { y1, y2, .. } = &mut elem.content {
                            *y1 -= offset;
                            *y2 -= offset;
                        }
                    }
                }
                // S1527 (2026-09-24, opt-out OXI_S1527_DISABLE): the notes referenced
                // by the lines that STAY on a page occupy its bottom, so the row's
                // later lines split above them. policies__0097185c: one row (the
                // FSP eligibility cell) runs p4..p6 and references notes 4/5 in
                // lines at the top of p5; Word draws both notes on p5 (separator
                // at 630.7) and moves "e) Is living in housing" to p6. S740 v1
                // commits a row's notes where the NEXT row begins, so Oxi drew
                // them on p6 and kept two more lines on p5. A ref whose line would
                // itself fall below the shrunk bottom keeps v1 (it travels with
                // its line). Per-note height is the row total shared equally.
                // Cell text elements carry no run_index; the paragraph's first
                // line stands for the ref's line (every corpus ref sits in its
                // paragraph's first line). Applied at the first split here and
                // at every continuation page in the loop below.
                let s1527_on = std::env::var_os("OXI_S1527_DISABLE").is_none();
                let (s1527_refs, s1527_heights): (Vec<(usize, usize, String, u32)>, Vec<(u32, f32)>) = match row_footnotes {
                    Some(rf) if s1527_on && !rf[row_idx].0.is_empty() => {
                        let (ids_all, _h_all, hs_all) = &rf[row_idx];
                        let mut refs = Vec::new();
                        for (ci, cell) in row.cells.iter().enumerate() {
                            let mut pi = 0usize;
                            for b in &cell.blocks {
                                if let Block::Paragraph(p) = b {
                                    for r in &p.runs {
                                        if let Some(id) = r.footnote_ref {
                                            refs.push((ci, pi, r.text.trim().to_string(), id));
                                        }
                                    }
                                    pi += 1;
                                }
                            }
                        }
                        (refs, ids_all.iter().copied().zip(hs_all.iter().copied()).collect())
                    }
                    _ => (Vec::new(), Vec::new()),
                };
                // (bottom, page_has_notes, elements, already placed) -> (new bottom, kept ids)
                let s1527_reserve = |bottom: f32, has_notes: bool, els: &[LayoutElement], early: &[u32]| -> (f32, Vec<u32>) {
                    // (line_top, line_bot, note height, id) per reference whose
                    // line is among `els`; the marker element (rendered note
                    // number in a superscript-sized font, same cell paragraph)
                    // names the line, else the paragraph's first line stands in.
                    let mut cands: Vec<(f32, f32, f32, u32)> = Vec::new();
                    for (ci, pi, marker, id) in s1527_refs.iter() {
                        let (ci, pi, id) = (*ci, *pi, *id);
                        if early.contains(&id) || cands.iter().any(|c| c.3 == id) {
                            continue;
                        }
                        let nh = s1527_heights.iter().find(|(i, _)| *i == id).map(|(_, h)| *h).unwrap_or(0.0);
                        let para_els: Vec<&LayoutElement> = els
                            .iter()
                            .filter(|e| e.cell_col_index == Some(ci) && e.cell_paragraph_index == Some(pi)
                                && e.cell_ancestor_path.is_empty()
                                && matches!(e.content, LayoutContent::Text { .. }))
                            .collect();
                        if para_els.is_empty() {
                            continue;
                        }
                        let para_fs = para_els
                            .iter()
                            .filter_map(|e| match &e.content { LayoutContent::Text { font_size, .. } => Some(*font_size), _ => None })
                            .fold(0.0f32, f32::max);
                        let marker_els: Vec<&&LayoutElement> = if marker.is_empty() {
                            Vec::new()
                        } else {
                            para_els.iter().filter(|e| match &e.content {
                                LayoutContent::Text { text, font_size, .. } =>
                                    text.trim() == marker.as_str() && *font_size < para_fs * 0.85,
                                _ => false,
                            }).collect()
                        };
                        let (top, bot) = if !marker_els.is_empty() {
                            marker_els.iter().fold((f32::INFINITY, f32::NEG_INFINITY), |(t, b), e| {
                                (t.min(e.y - e.flow_line_offset), b.max(e.y + e.height))
                            })
                        } else {
                            // first line of the paragraph
                            let t = para_els.iter().map(|e| e.y - e.flow_line_offset).fold(f32::INFINITY, f32::min);
                            let b = para_els.iter().filter(|e| (e.y - e.flow_line_offset - t).abs() < 0.5)
                                .map(|e| e.y + e.height).fold(f32::NEG_INFINITY, f32::max);
                            (t, b)
                        };
                        cands.push((top, bot, nh, id));
                    }
                    cands.sort_by(|x, y| x.0.partial_cmp(&y.0).unwrap_or(std::cmp::Ordering::Equal));
                    let mut bottom = bottom;
                    let mut kept: Vec<u32> = Vec::new();
                    for (top, bot, nh, id) in cands {
                        if top >= bottom {
                            break;
                        }
                        let sep = if kept.is_empty() && !has_notes { fn_sep } else { 0.0 };
                        let fits = bot <= bottom - nh - sep + 0.5;
                        if std::env::var("OXI_DBG_SPLIT").is_ok() {
                            eprintln!("[SPLIT-S1527-REF] id={} line {:.2}..{:.2} nh={:.2} sep={:.2} bottom={:.2} fits={}", id, top, bot, nh, sep, bottom, fits);
                        }
                        if fits {
                            bottom -= nh + sep;
                            kept.push(id);
                        } else {
                            // the line moves to the next page with its note; so
                            // does everything below it. The cut sits just under
                            // the line's top so the line ABOVE (whose bottom
                            // coincides with this top within rounding) stays.
                            bottom = bottom.min(top + 0.4);
                            break;
                        }
                    }
                    (bottom, kept)
                };
                let fragment_content_bottom = if !s1527_refs.is_empty() {
                    let (bottom, kept) = s1527_reserve(fragment_content_bottom, s740_page_has_notes, &row_elements, &s1527_early);
                    if !kept.is_empty() {
                        if std::env::var("OXI_DBG_SPLIT").is_ok() {
                            eprintln!("[SPLIT-S1527] row={} kept_notes={:?} bottom {:.2} -> {:.2}", row_idx, kept, fragment_content_bottom, bottom);
                        }
                        s740_page_has_notes = true;
                        let off = pages.len() - s740_entry_pages;
                        while s740_fn_pages.len() <= off {
                            s740_fn_pages.push(Vec::new());
                        }
                        for id in &kept {
                            if !s740_fn_pages[off].contains(id) {
                                s740_fn_pages[off].push(*id);
                            }
                        }
                        s1527_early.extend(kept.iter().copied());
                    }
                    bottom
                } else {
                    fragment_content_bottom
                };
                // R7.70 (Day 37 session 58, 2026-05-15): pick the FIRST LRPB-marked
                // element in document order (= cell-render order), not the min elem.y.
                // ed025c row has 3 LRPB elements: (8) at y=761.5 (cell 0, correct
                // break point), "× × ×" at y=743.5 (number-cell at (7)-position, in
                // a different cell of same row), "１" at y=1463.5 (much later). The
                // previous min-y rule picked × × × at 743.5 → split_y pulled below
                // page_bottom → (7) overflowed mistakenly. Document-order picks
                // cell 0's (8) at 761.5 first → split_y = page_bottom fallback (since
                // 761.5 > page_bottom 760.5) → (7) stays. e3c545 cpi=81 path is
                // unaffected because there it was the only LRPB in the row (single
                // element → first == min).
                let lrpb_split_y = row_elements
                    .iter()
                    .find(|e| e.is_paragraph_start_with_lrpb && e.y > cursor.cursor_y + 0.5)
                    .map(|e| e.y)
                    .unwrap_or(f32::INFINITY);
                let split_y = if lrpb_split_y.is_finite() && lrpb_split_y < fragment_content_bottom {
                    lrpb_split_y
                } else {
                    fragment_content_bottom
                };
                // S1168 (2026-08-19, default ON, opt-out OXI_S1168_DISABLE):
                // a row-split line is ONE horizontal line and an image is
                // ATOMIC, so the line may not cross an image. Word's rule,
                // derived by `_pb_cellimgtail_gen.py` (2-cell row, cell =
                // 220pt image + tail lines, filler swept a line at a time,
                // image coordinates read back with PDF get_image_info):
                //   fill26 img 445.1-665.1 keeps 3 tail lines
                //   fill27 img 458.9-678.9 keeps 2
                //   fill28 img 472.7-692.7 keeps 1
                //   fill29 img 486.5-706.5 keeps 0 -- the image still stays
                //   fill30 img bottom 720.3 (0.3pt over) -- the image moves
                // i.e. the image stays whenever its BOTTOM fits and the tail
                // packs in whatever is left (zero lines is a legal outcome);
                // when it does not fit the line must move above it. Pulling
                // up can expose a second image, so iterate to a fixed point.
                //
                // Landing at the row top means no line can be drawn at all,
                // which is Word's whole-row push (educational__00161422 p4:
                // Word leaves 113pt blank because the next row's tallest cell
                // holds a 161pt image and only 113.5 is free). That falls out
                // of the partition below for free -- every element's bottom is
                // then past the line, so the entire row moves -- so this needs
                // no separate branch, and the vertical borders / shading self-
                // suppress because both test strictly against the line.
                //
                // This replaces the `image_atomic_push` whole-row veto, which
                // was S533's stand-in for exactly this geometry (see the veto
                // site). S1130b tried pulling the line back WITHOUT the row-top
                // case and scored 0.3524 -- "cut higher up" is not the rule,
                // "if no line can be drawn, move the row" is.
                let s1168 = std::env::var("OXI_S1168_DISABLE").is_err();
                let row_top = cursor.cursor_y;
                let split_y = if s1168 {
                    let mut cand = split_y;
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        for e in row_elements.iter().filter(|e| matches!(e.content, LayoutContent::Image { .. })) {
                            eprintln!("[SPLIT-S1168-IMG] y={:.2} h={:.2} off={:.2} fit={:?} top={:.2} row_top={:.2} cand={:.2} margin={}",
                                e.y, e.height, e.flow_line_offset, e.content_fit_height, e.y - e.flow_line_offset, row_top, cand, e.margin_float);
                        }
                    }
                    loop {
                        let crossed = row_elements
                            .iter()
                            .filter(|e| matches!(e.content, LayoutContent::Image { .. })
                                || (std::env::var_os("OXI_EMPTY_CELL_LINE_SPLIT").is_some()
                                    && matches!(&e.content, LayoutContent::Text { text, .. } if text.is_empty())))
                            // S1555 (2026-09-25, default ON, opt-out OXI_S1555_DISABLE): a
                            // MARGIN-relative cell float (cell_float_flow gives it origin 0,
                            // so flow_line_offset = its absolute y) is not row content: its
                            // computed span 0..y+h "crosses" every candidate and the split
                            // collapsed to the row top (correspondence__101d483d: a 1-row
                            // 8-paragraph table with 7 margin-anchored photos left page 2
                            // empty; Word splits the row with three paragraphs on it).
                            // OPT-IN (OXI_S1555=1) until the Word probe decides: with it ON
                            // correspondence__101d483d 0.9123->0.9649 but forms__005a5d91 and
                            // educational__00161422 PASS->FAIL (their cell floats straddle the
                            // page bottom the same way and Word moves the row). All three
                            // are layoutInCell=1, so that attribute is not the discriminator.
                            // S1555 v2 (2026-09-25, default ON, opt-out OXI_S1555_DISABLE): a
                            // cell-relative (margin) float counts as a straddling image only
                            // when its ANCHOR paragraph's line lies in the first fragment
                            // (above the candidate split). Word probe cellfloat_split (18
                            // arms): a picture straddling the page bottom moves the whole
                            // row (compat 14) / splits the row after the anchor line
                            // (compat 15) — but correspondence__101d483d's straddling photo
                            // is anchored to a paragraph that begins below the split, so
                            // Word splits the row by its text and the photo follows its
                            // paragraph. Oxi moved the split to the row top (+1 page).
                            .filter(|e| std::env::var_os("OXI_S1555_DISABLE").is_some()
                                || !e.margin_float
                                || e.cell_paragraph_index.map_or(true, |ap| {
                                    row_elements.iter()
                                        .filter(|t| matches!(t.content, LayoutContent::Text { .. })
                                            && t.cell_col_index == e.cell_col_index
                                            && t.cell_paragraph_index == Some(ap))
                                        .map(|t| t.y)
                                        .fold(f32::INFINITY, f32::min)
                                        < cand - 0.1
                                }))
                            .filter(|e| e.y - e.flow_line_offset < cand - 0.1
                                && e.y - e.flow_line_offset
                                    + e.content_fit_height.unwrap_or(e.height + e.flow_line_offset)
                                    > cand + 0.1)
                            // Effects offset the painted image inside its line.
                            // Splitting at the paint top misclassifies a whole-row
                            // move as a continuation and adds continuation space.
                            .map(|e| {
                                let origin = e.y - e.flow_line_offset;
                                if e.margin_float && float_replay.is_some() {
                                    let anchor = row_elements.iter()
                                        .filter(|t| matches!(t.content, LayoutContent::Text { .. })
                                            && t.cell_col_index == e.cell_col_index
                                            && t.cell_paragraph_index == e.cell_paragraph_index)
                                        .map(|t| t.y - t.flow_line_offset)
                                        .fold(f32::INFINITY, f32::min);
                                    if anchor.is_finite() { origin.max(anchor) } else { origin }
                                } else { origin }
                            })
                            .fold(f32::INFINITY, f32::min);
                        if !crossed.is_finite() {
                            break cand;
                        }
                        if crossed <= row_top + 0.1 {
                            break row_top;
                        }
                        cand = crossed;
                    }
                } else {
                    split_y
                };
                // S819 (2026-07-13, default ON, opt-out OXI_S819_DISABLE, Latin
                // exact-cell bundle): Word's split-row fill keeps a text line only when
                // line_box_bottom + tcMar_b + border_width fits the content
                // bottom — the split reserves the cell's BOTTOM frame at the
                // page bottom (_pb_rowsplit_gen: base Q=5.97=5.25+0.72 in
                // [5.93,6.13); tcMar_b=15tw variant Q=1.47 in [1.48,1.68);
                // sa=0 variant unchanged → sa excluded). The p1-side mirror
                // of the S817 continuation tail. TEXT elements only; the
                // border/shading clip geometry keeps split_y. S402 history:
                // a GLOBAL tighten here is catastrophic (ed025 0.9986→0.80)
                // — Latin-scoped, opt-in.
                // S819 applies to the NATURAL page-bottom split only: an
                // LRPB-pulled split_y is Word's own recorded break position,
                // which already encodes the bottom-frame decision —
                // subtracting q again double-counts and drops one more line
                // (uklocal p21 row 1: saved split 708.8, line-1 bottom 707 —
                // Word keeps 1 line, the doubled q pushed the row whole).
                let empty_line_boundary = std::env::var_os("OXI_EMPTY_CELL_LINE_SPLIT").is_some()
                    && split_y < page_bottom
                    && row_elements.iter().any(|e| {
                        matches!(&e.content, LayoutContent::Text { text, .. } if text.is_empty())
                            && (e.y - split_y).abs() < 0.01
                    });
                let s819_natural_split = !(lrpb_split_y.is_finite() && lrpb_split_y < page_bottom)
                    && !empty_line_boundary;
                // S1606 (2026-09-29): the cell-frame reservation is not Latin-only.
                // `_pb_vmerge_colbottom_gen.py` (faithful slice of blind-G JA
                // policies__1e87d3e6, tcMar_b 43tw, sz4 borders, exact 11.5 lines):
                // Word splits row 14 while line_bottom + 2.15 + bw <= cbot and moves
                // it one 0.25pt step later, identically with vMerge and vAlign
                // removed and in a one-column section; Oxi without S819 kept it
                // 2.5pt longer.
                // Default ON with S1618 (Bug A's width moved under the table); opt-out
                // OXI_S1606_DISABLE.
                let s1606 = std::env::var_os("OXI_S1606_DISABLE").is_none();
                let s819_q = if (!self.doc_body_has_real_cjk || s1606)
                    && s819_natural_split
                    && std::env::var("OXI_S819_DISABLE").is_err()
                {
                    let pad_b = row
                        .cells
                        .first()
                        .and_then(|c| c.margins.as_ref().and_then(|m| m.bottom))
                        .unwrap_or(row_default_pad_b);
                    let bw = if separate_outer_edges { fragment_bottom_width } else { row
                        .cells
                        .first()
                        .and_then(|c| c.borders.as_ref())
                        .and_then(|b| b.bottom.as_ref())
                        .map(|bd| bd.width)
                        .unwrap_or_else(|| {
                            if table.style.border {
                                table.style.border_width.unwrap_or(0.4)
                            } else {
                                0.0
                            }
                        }) };
                    pad_b + bw
                } else {
                    0.0
                };
                // The natural cut already reserves fragment_bottom_width.
                // Reserve only the remaining cell frame here, including when
                // the edge came from cell borders rather than table borders.
                // Word boundary controls: a 1pt bottom edge keeps two lines;
                // a 2pt edge moves the paragraph. Charging the edge twice
                // incorrectly moves it with the 1pt edge as well.
                let s819_fit_q = if separate_outer_edges
                    || std::env::var_os("OXI_CELL_FRAME_RESERVE_DISABLE").is_none()
                {
                    (s819_q - fragment_bottom_width).max(0.0)
                } else { s819_q };
                // S1092 (2026-08-07, opt-out OXI_S1092_DISABLE): a split-row fragment must
                // contain the LAST paragraph's OWN `space_after`. DERIVED on
                // policies__0028d1be (Letter, one 5-page table row) with a
                // bottom-margin ladder + a causal after=0 arm (Word COM):
                //   orig    cbot 720.00  box_bot 708.30  room 11.70 → PUSH
                //   after=0 cbot 720.00  box_bot 708.30  room 11.70 → KEEP
                //   after=0 cbot 712.00  box_bot 708.30  room  3.70 → KEEP
                //   before=0 (control)   room 11.70            → PUSH
                // and the fine ladder puts the required room in (13.2, 13.7] =
                // the paragraph's 14.0 auto after within COM's 0.75 quantum.
                // The COLLAPSED inter-paragraph gap is NOT the quantity:
                // zeroing this paragraph's after leaves the gap at
                // max(0, next.before) = 14 yet flips the verdict, so the
                // fragment closes on the paragraph's OWN after. S819 already
                // reserves tcMar_b + bw (the cell frame); this is the content
                // term that sits above it. Latin-scoped + natural split only,
                // mirroring S819.
                let s1092 = !self.doc_body_has_real_cjk
                    && s819_natural_split
                    && std::env::var("OXI_S1092_DISABLE").is_err();
                let mut s1092_after: std::collections::HashMap<(usize, usize), f32> =
                    Default::default();
                let mut s1092_last: std::collections::HashMap<(usize, usize), f32> =
                    Default::default();
                if s1092 {
                    let tps = table.style.para_style.as_ref();
                    let s952_tbl =
                        tps.map_or(false, |ts| ts.before_autospacing || ts.after_autospacing);
                    for (ci, cell) in row.cells.iter().enumerate() {
                        let paras: Vec<&Paragraph> = cell
                            .blocks
                            .iter()
                            .filter_map(|b| match b {
                                Block::Paragraph(p) => Some(p),
                                _ => None,
                            })
                            .collect();
                        for (pi, para) in paras.iter().enumerate() {
                            let (sb, sa) = self.cell_para_spacing(para, tps, table_grid_pitch);
                            let sa = if para.style.before_autospacing
                                || para.style.after_autospacing
                                || para.style.contextual_spacing
                                || s952_tbl
                            {
                                self.cell_effective_spacing(
                                    para,
                                    tps,
                                    matches!(cell.blocks.iter().find(|b| LayoutEngine::is_cell_spacing_paragraph(b)),
                                        Some(Block::Paragraph(p)) if std::ptr::eq(p, *para)),
                                    matches!(cell.blocks.iter().rev().find(|b| LayoutEngine::is_cell_spacing_paragraph(b)),
                                        Some(Block::Paragraph(p)) if std::ptr::eq(p, *para)),
                                    sb,
                                    sa,
                                )
                                .1
                            } else {
                                sa
                            };
                            // Contextual after spacing belongs to the boundary
                            // with the following paragraph. A page cut does not
                            // restore a boundary that same-style adjacency removed.
                            let same_style_next = cell.blocks.windows(2).any(|pair| {
                                matches!((&pair[0], &pair[1]),
                                    (Block::Paragraph(current), Block::Paragraph(next))
                                    if std::ptr::eq(current, *para)
                                        && current.style.style_id == next.style.style_id)
                            });
                            let sa = if !self.preserve_same_style_cell_spacing
                                && std::env::var("OXI_S939_DISABLE").is_err()
                                && para.style.contextual_spacing && same_style_next
                            { 0.0 } else { sa };
                            // The trailing after and the cell's bottom frame
                            // (tcMar_b + border = S819's q) OVERLAP rather than
                            // stack: Word's fragment closes at
                            //   last_line_bottom + max(after, tcMar_b + bw)
                            // (uklocalspending p40 render-truth: the row's bottom
                            // border sits at last_line_bottom + 5.97 = q, NOT
                            // + after 6.0 + q). Only the EXCESS over q is new
                            // reservation; policies__0028d1be has q = 0 so its
                            // whole 14pt after applies.
                            let extra = (sa - s819_q).max(0.0);
                            if extra > 0.01 {
                                s1092_after.insert((ci, pi), extra);
                            }
                        }
                    }
                    for e in row_elements.iter() {
                        if matches!(
                            e.content,
                            LayoutContent::TableBorder { .. } | LayoutContent::CellShading { .. }
                        ) {
                            continue;
                        }
                        if let (Some(ci), Some(pi)) = (e.cell_col_index, e.cell_paragraph_index) {
                            let b = e.y + e.height;
                            let cur = s1092_last.entry((ci, pi)).or_insert(b);
                            if b > *cur {
                                *cur = b;
                            }
                        }
                    }
                }
                // S1246 (2026-08-28, default ON, opt-out OXI_S1246_DISABLE):
                // a paragraph inside a splitting table row does not leave its
                // LAST line alone on the continuation page.
                //
                // DERIVED, `tools/metrics/_pb_widow_{gen,read}.py` (Word PDF truth,
                // the same 5-line paragraph laid out once in the body and once in a
                // 2-cell row, 4 shapes x 12 filler counts):
                //     fill      BODY   BODYOFF      CELL   CELLOFF
                //       58       3/2       4/1       3/2       4/1
                //       59       3/2       3/2       3/2       3/2
                //       60       2/3       2/3       2/3       2/3
                //       61       0/5       1/4       0/5       1/4
                // Word's CELL column IS its BODY column, and w:widowControl w:val="0"
                // restores the natural split in both. Oxi matched Word in the body and
                // scored 4/1 and 1/4 in the cell -- CELL and CELLOFF identical at every
                // arm, i.e. the flag never reached the row split.
                //
                // PER CELL, NOT PER ROW (`_pb_rowwidow_*`, a short cell A beside a
                // taller cell B, Word PDF): the adjustment moves ONE cell's line, not
                // the row's cut. MED fill58 (A 5 lines, B 9) reads
                //     p1  A01 B01 / A02 B02 / A03 B03 / __  B04
                //     p2  A04 B05 / A05 B06 / __  B07 ...
                // -- A stops a line above B on the same page, and both cells resume at
                // the SAME continuation y (which S1093 already produces). So the rule
                // caps each paragraph's kept-line count; it must not pull the row's
                // shared `split_y`, which is what an earlier version of this did (it
                // dragged B to 3/6 and cost Phase 1 a document).
                //
                // SCOPE 1 -- natural splits only, mirroring S819 and S1092. An
                // LRPB-pulled `split_y` is a break position Word itself recorded in the
                // file, so it already carries whatever Word decided about widows.
                // uklocalspending p21 row 1 is the witness the S819 note describes:
                // saved split 708.8, one line above it, and Word keeps that lone line.
                //
                // SCOPE 2 -- the last-line adjustment applies only where 2 lines
                // can still be KEPT (n >= 4). Where the
                // pull would leave 1 or 0, Word's behaviour is measured but NOT
                // resolved, so this rule declines rather than guess:
                //   `_pb_rowwidow` SHORT fill60 -- A is 3 lines splitting 2/1, and Word
                //     moves the WHOLE ROW to p2 (A 0/3, B 0/9);
                //   uklocalspending p36 row 2 -- 3-line cells splitting 2/1 beside a
                //     15-line cell, and Word keeps the 2/1, leaving the lone line.
                // Same shape, opposite outcome, with COM reporting WidowControl=-1 on
                // the uklocalspending cells. PRESHORT/PRETWO rule out "the row is the
                // table's first row"; the row is ~263pt so it is not "taller than a
                // page"; the remaining candidates (space before/after, lineRule
                // atLeast, cell count, following rows) are untested. Until one of them
                // separates the two, the n<=3 region keeps its last-line behaviour.
                let mut short_widow_moves_row = false;
                let s1246_limit: std::collections::HashMap<CellFlowKey, f32> =
                    if s819_natural_split
                        && self.compat_mode >= 15 && self.compat_mode_explicit
                        && std::env::var("OXI_S1246_DISABLE").is_err()
                    {
                        // Lines of each cell paragraph, as (top, bottom) per distinct
                        // baseline -- a line can be several elements (one run each, and
                        // a justified line is one element per word).
                        let mut plines: std::collections::HashMap<CellFlowKey, Vec<(f32, f32)>> =
                            Default::default();
                        for e in row_elements.iter() {
                            if !matches!(e.content, LayoutContent::Text { .. }) {
                                continue;
                            }
                            if let Some(key) = e.cell_flow_key()
                            {
                                let v = plines.entry(key).or_default();
                                match v.iter_mut().find(|(t, _)| (*t - e.y).abs() < 0.1) {
                                    Some(l) => l.1 = l.1.max(e.y + e.height),
                                    None => v.push((e.y, e.y + e.height)),
                                }
                            }
                        }
                        // A nested row whose FIRST paragraph would leave one of two
                        // lines behind moves whole, as a top-level row does.
                        let nested_orphan_on = std::env::var_os("OXI_NESTED_ORPHAN_ROW_DISABLE").is_none();
                        let mut nested_row_top: std::collections::HashMap<(Vec<(usize, usize, usize)>, usize), f32> =
                            Default::default();
                        if nested_orphan_on {
                            for (key, lines) in plines.iter() {
                                if key.0.is_empty() { continue; }
                                let t = lines.iter().map(|l| l.0).fold(f32::INFINITY, f32::min);
                                let e = nested_row_top.entry((key.0.clone(), key.1)).or_insert(t);
                                if t < *e { *e = t; }
                            }
                        }
                        let all_keys: Vec<CellFlowKey> = plines.keys().cloned().collect();
                        let mut nested_rows_moved: Vec<(Vec<(usize, usize, usize)>, usize)> = Vec::new();
                        let mut out = std::collections::HashMap::new();
                        for (key, mut lines) in plines {
                            if lines.len() < 2 {
                                continue;
                            }
                            let mut source_row = Some(row);
                            for (depth, &(_, ci, bi)) in key.0.iter().enumerate() {
                                source_row = source_row.and_then(|r| r.cells.get(ci))
                                    .and_then(|c| c.blocks.get(bi))
                                    .and_then(|b| match b {
                                        Block::Table(t) => t.rows.get(key.0.get(depth + 1)
                                            .map_or(key.1, |p| p.0)),
                                        _ => None,
                                    });
                            }
                            let source_cell = source_row.and_then(|r| r.cells.get(key.2));
                            let widow_on = source_cell.and_then(|c| c.blocks.iter()
                                .filter_map(|b| match b { Block::Paragraph(p) => Some(p), _ => None })
                                .nth(key.3)).map_or(false, |p| p.style.widow_control);
                            if !widow_on {
                                continue;
                            }
                            lines.sort_by(|a, b| {
                                a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal)
                            });
                            // The same keep test the partition below applies, so this
                            // count is the count that would actually be kept.
                            let n = lines.len();
                            let k = lines
                                .iter()
                                .filter(|l| {
                                    let extra = if s1092 && key.0.is_empty()
                                        && s1092_last
                                            .get(&(key.2, key.3))
                                            .map_or(false, |b| (l.1 - *b).abs() < 0.01)
                                    {
                                        *s1092_after.get(&(key.2, key.3)).unwrap_or(&0.0)
                                    } else {
                                        0.0
                                    };
                                    l.1 + extra <= split_y + row_fit_epsilon - s819_fit_q
                                })
                                .count();
                            if nested_orphan_on && !key.0.is_empty() && key.1 > 0 && key.3 == 0
                                && n == 2 && k == 1
                                && self.compat_mode >= 15 && self.compat_mode_explicit
                            {
                                nested_rows_moved.push((key.0.clone(), key.1));
                                continue;
                            }
                            // Once earlier paragraphs of this cell have fitted,
                            // keep the next paragraph's first two lines together.
                            // Moving a row's first paragraph needs a separate
                            // whole-row decision and retains its existing rule.
                            // A three-line first paragraph cannot leave either
                            // one or two lines here while keeping two on each page.
                            if ((n == 3 && k == 2) || (n >= 2 && k == 1
                                && std::env::var("OXI_CELL_FIRST_ORPHAN_DISABLE").is_err()))
                                && key.3 == 0 && key.0.is_empty()
                                && self.compat_mode >= 15 && self.compat_mode_explicit
                                && row_height <= page_bottom - page_top
                                && source_cell.map_or(false, |cell|
                                    !cell.blocks.iter().any(|b| matches!(b, Block::Table(_))))
                            {
                                // A three-line first paragraph cannot keep two lines
                                // on both pages. Modern Word moves the complete row.
                                short_widow_moves_row = true;
                            }
                            if (k == 1 || (n == 3 && k == 2)) && key.3 > 0
                                // A three-line paragraph cannot keep two lines
                                // on both sides; move it after preceding cell content.
                                // Legacy and settings-less documents allow a
                                // lone first line in a split table paragraph.
                                && self.compat_mode >= 15 && self.compat_mode_explicit
                                && std::env::var("OXI_CELL_ORPHAN_DISABLE").is_err()
                                // Nested cells reuse local paragraph indices;
                                // their lines must not be joined to this paragraph.
                                && source_cell.map_or(false, |cell|
                                    !cell.blocks.iter().any(|b| matches!(b, Block::Table(_))))
                            {
                                out.insert(key, lines[0].0);
                            } else if n >= 4 && k == n - 1 {
                                out.insert(key, lines[n - 2].0);
                            }
                        }
                        for (path, r) in nested_rows_moved {
                            if let Some(&top) = nested_row_top.get(&(path.clone(), r)) {
                                for k2 in all_keys.iter().filter(|k2| k2.0 == path && k2.1 == r) {
                                    out.insert(k2.clone(), top);
                                }
                            }
                        }
                        out
                    } else {
                        Default::default()
                    };
                let split_y = if short_widow_moves_row { row_top } else { split_y };
                // Use exactly the same fit decision for the fragment's natural
                // content extent and its final element partition. Retain only
                // floating-point roundoff, as in whole-row fit: 0.1pt of slack
                // incorrectly keeps a last line plus after-spacing 0.096pt
                // beyond the page bottom in the short-cell regression fixture.
                let fits_at = |elem: &LayoutElement, split_y: f32| {
                    let bottom = if matches!(elem.content, LayoutContent::Image { .. }) {
                        elem.y - elem.flow_line_offset
                            + elem.content_fit_height.unwrap_or(elem.height + elem.flow_line_offset)
                    } else { elem.y + elem.height };
                    let after = if s1092 {
                        match (elem.cell_col_index, elem.cell_paragraph_index) {
                            (Some(ci), Some(pi)) if s1092_last.get(&(ci, pi))
                                .map_or(false, |b| (bottom - *b).abs() < 0.01) =>
                                *s1092_after.get(&(ci, pi)).unwrap_or(&0.0),
                            _ => 0.0,
                        }
                    } else { 0.0 };
                    let widow_limit = elem.cell_flow_key()
                        .and_then(|k| s1246_limit.get(&k)).copied()
                        .unwrap_or(f32::INFINITY);
                    bottom + after <= split_y + row_fit_epsilon - s819_fit_q
                        && elem.y < widow_limit - 0.1
                };
                // S1607 (2026-09-29): when no line of the row (empty lines and images count) fits above the
                // cut, the row moves WHOLE -- cut at the row top, the S1168 row-top
                // case, so it keeps its top cell margin and its own height. The
                // continuation path instead re-anchored the first line to the page
                // top (row 14 at 81.5, Word 84.75) and closed the fragment on the
                // lowest element, which for a vMerge-restart row includes the merged
                // cells' text centred over the rows below (row 15 at 132.55, Word
                // 101.25; `_pb_vmerge_colbottom_gen.py`, blind-G policies__1e87d3e6).
                let split_y = if std::env::var_os("OXI_S1607_DISABLE").is_none()
                    && split_y > row_top + 0.1
                    && row_elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { .. } | LayoutContent::Image { .. }))
                    && !row_elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { .. } | LayoutContent::Image { .. })
                        && fits_at(e, split_y))
                {
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        eprintln!("[SPLIT-S1607] row={} no text line fits above {:.2} -> whole-row move", row_idx, split_y);
                    }
                    row_top
                } else { split_y };
                // S1607: R7.61 flagged a vMerge-restart cell's later paragraphs
                // that were laid out past the old page bottom; after a whole-row
                // move they sit on the next page already, and the post-paginate
                // sweep would move them a second page (row 14's "１回" at y -601.6).
                if std::env::var_os("OXI_S1607_DISABLE").is_none() && split_y <= row_top + 0.1 {
                    let shift = split_y - page_top;
                    for e in row_elements.iter_mut() {
                        if e.vmerge_restart_overflow_to_next_page && e.y - shift <= page_bottom + 0.5 {
                            e.vmerge_restart_overflow_to_next_page = false;
                        }
                    }
                }
                let fits_fragment = |elem: &LayoutElement| fits_at(elem, split_y);
                // If the fit rules move the entire row, it is no longer a
                // split fragment: retain its original vertical alignment.
                if split_valign && split_y <= row_top + 0.1 {
                    for &(index, offset) in &split_valign_offsets {
                        let elem = &mut row_elements[index];
                        elem.y += offset;
                        if let LayoutContent::TableBorder { y1, y2, .. } = &mut elem.content {
                            *y1 += offset;
                            *y2 += offset;
                        }
                    }
                }

                if fragment_valign && split_y > row_top + 0.1 {
                    // Short cells retain their alignment within the first page
                    // fragment; longer cells flow from its top without a gap.
                    // A partially filled last line leaves unused page space.
                    // Alignment follows the fitted content, not that unused space.
                    let mut paragraph_bottoms: std::collections::HashMap<CellFlowKey, f32> =
                        std::collections::HashMap::new();
                    for elem in &row_elements {
                        if let Some(key) = elem.cell_flow_key() {
                            let bottom = elem.y + elem.height;
                            paragraph_bottoms.entry(key).and_modify(|b| *b = b.max(bottom))
                                .or_insert(bottom);
                        }
                    }
                    let paragraph_after = |elem: &LayoutElement| {
                        if !elem.cell_ancestor_path.is_empty() || elem.cell_row_index != Some(row_idx) {
                            return 0.0;
                        }
                        let is_last = elem.cell_flow_key().and_then(|key| paragraph_bottoms.get(&key))
                            .map_or(false, |b| (*b - elem.y - elem.height).abs() < 0.01);
                        if is_last {
                            elem.cell_col_index.zip(elem.cell_paragraph_index)
                                .and_then(|key| fragment_paragraph_after.get(&key)).copied().unwrap_or(0.0)
                        } else { 0.0 }
                    };
                    let fragment_bottom = fragment_valign_cells.iter()
                        .flat_map(|(range, _, _, padding, _)| {
                            row_elements[range.clone()].iter()
                                .filter(|e| !matches!(&e.content,
                                    LayoutContent::TableBorder { .. } | LayoutContent::CellShading { .. }))
                                .filter(|e| fits_fragment(e))
                                .map(|e| e.y + e.height + padding.max(paragraph_after(e)))
                        }).fold(row_top, f32::max);
                    for (range, origin, height, bottom_padding, factor) in &fragment_valign_cells {
                        let offset = (fragment_bottom - origin - bottom_padding - height).max(0.0) * factor;
                        for elem in &mut row_elements[range.clone()] {
                            elem.y += offset;
                            if let LayoutContent::TableBorder { y1, y2, .. } = &mut elem.content {
                                *y1 += offset;
                                *y2 += offset;
                            }
                        }
                    }
                }

                // Keep the declared row top edges for the continuation fragment.
                // A full fragment still has a top edge; it is independent of
                // whether the continuation bottom needs to be extended.
                let split_top_edges = row_declared_top_edges;
                let split_bottom_edges: Vec<LayoutElement> = row_elements.iter()
                    .filter(|e| matches!(&e.content,
                        LayoutContent::TableBorder { y1, y2, .. }
                        if (*y1 - *y2).abs() < 0.1 && (*y1 - row_bottom).abs() <= 1.0))
                    .cloned().collect();
                // Partition elements: those fitting on current page vs overflow
                let mut current_page_elems: Vec<LayoutElement> = Vec::new();
                let mut next_page_elems: Vec<LayoutElement> = Vec::new();

                for elem in row_elements {
                    let _elem_top = elem.y;
                    match &elem.content {
                        LayoutContent::TableBorder {
                            y1,
                            y2,
                            x1,
                            x2,
                            ref color,
                            width,
                            ref style,
                        } => {
                            // Horizontal borders: keep on their respective page
                            if (y1 - y2).abs() < 0.1 {
                                // Horizontal line
                                if *y1 <= split_y + 0.5 {
                                    current_page_elems.push(elem);
                                } else {
                                    // Shift to next page
                                    let shift = split_y - page_top;
                                    let mut e = elem;
                                    e.y -= shift;
                                    if let LayoutContent::TableBorder {
                                        ref mut y1,
                                        ref mut y2,
                                        ..
                                    } = e.content
                                    {
                                        *y1 -= shift;
                                        *y2 -= shift;
                                    }
                                    next_page_elems.push(e);
                                }
                            } else {
                                // Vertical border: split at page boundary
                                // Current page portion
                                let vy_top = *y1;
                                let vy_bot = *y2;
                                if vy_top < split_y {
                                    current_page_elems.push(LayoutElement::new(
                                        elem.x,
                                        elem.y,
                                        elem.width,
                                        split_y - vy_top,
                                        LayoutContent::TableBorder {
                                            x1: *x1,
                                            y1: vy_top,
                                            x2: *x2,
                                            y2: split_y,
                                            color: color.clone(),
                                            width: *width,
                                            style: style.clone(),
                                        },
                                    ));
                                }
                                // Next page portion
                                if vy_bot > split_y {
                                    let shift = split_y - page_top;
                                    let new_y1 = page_top;
                                    let new_y2 = vy_bot - shift;
                                    next_page_elems.push(LayoutElement::new(
                                        elem.x,
                                        new_y1,
                                        elem.width,
                                        new_y2 - new_y1,
                                        LayoutContent::TableBorder {
                                            x1: *x1,
                                            y1: new_y1,
                                            x2: *x2,
                                            y2: new_y2,
                                            color: color.clone(),
                                            width: *width,
                                            style: style.clone(),
                                        },
                                    ));
                                }
                            }
                        }
                        LayoutContent::CellShading { ref color } => {
                            // Split shading across pages
                            let shade_bottom = elem.y + elem.height;
                            if elem.y < split_y {
                                let clip_h = (split_y - elem.y).min(elem.height);
                                current_page_elems.push(LayoutElement::new(
                                    elem.x,
                                    elem.y,
                                    elem.width,
                                    clip_h,
                                    LayoutContent::CellShading {
                                        color: color.clone(),
                                    },
                                ));
                            }
                            if shade_bottom > split_y {
                                let shift = split_y - page_top;
                                let new_y = (elem.y - shift).max(page_top);
                                let new_h = shade_bottom - shift - new_y;
                                next_page_elems.push(LayoutElement::new(
                                    elem.x,
                                    new_y,
                                    elem.width,
                                    new_h.max(0.0),
                                    LayoutContent::CellShading {
                                        color: color.clone(),
                                    },
                                ));
                            }
                        }
                        _ => {
                            // Text and other elements. Step 1 (2026-04-22):
                            // use element BOTTOM vs split_y, not top. A line at
                            // y=761 with lh=18 has bottom=779; if split_y=771,
                            // the line's bottom overflows and must move to the
                            // next page. Previously used elem_top < split_y,
                            // which kept the overflow-bottom line on current
                            // page. This matches d77a p6/p7 cell-paragraph split.
                            //
                            // Session 75 Phase D (2026-05-17): elem.y is now
                            // LINE BOX TOP (was glyph_top = LBT + text_y_off
                            // pre-Phase-D). So elem.y + elem.height = line_box
                            // bottom directly, no recovery needed. Replaces the
                            // R7.69 text_y_off_recovered workaround.
                            if fits_fragment(&elem) {
                                current_page_elems.push(elem);
                            } else {
                                let shift = split_y - page_top;
                                let mut e = elem;
                                e.y -= shift;
                                next_page_elems.push(e);
                            }
                        }
                    }
                }

                // Step 1 (2026-04-22): re-anchor overflow text so the FIRST
                // overflow line lands at page_top, preserving relative spacing
                // between subsequent lines. The original `shift = split_y -
                // page_top` assumed overflow starts exactly at split_y, which
                // is wrong when the line's top is below split_y but its bottom
                // straddles. Compute the actual minimum y of overflow text and
                // re-shift.
                // S570 (2026-06-14): collapse a LEADING EMPTY line at the row-split
                // continuation top. A cell empty paragraph that straddles the page
                // boundary lands a full-height (16.5pt) blank line at the continuation
                // top; Word COLLAPSES it (RENDER-TRUTH harassbun: Word's first p2 line
                // is content at y=51.9, Oxi had an empty text line at y=48 + content at
                // 64.5 = a +16.5pt offset). Anchor to the first NON-EMPTY text and drop
                // the leading empty-text lines above it. Opt-out OXI_S570_DISABLE.
                let s570 = std::env::var("OXI_S570_DISABLE").is_err();
                // S719 (2026-07-02, default ON, opt-out OXI_S719_DISABLE): the S570
                // leading-line collapse applies to TRULY EMPTY paragraphs (no runs,
                // element text == "") ONLY — a WHITESPACE-run line is CONTENT Word
                // keeps at the continuation top. Render-truth tokyoshugyo p50/51:
                // the 賃金 box split lands just before a whitespace-only exact-240
                // spacer («␣␣＋U+3000×28», line=240 exact); Word renders it as the
                // p51 continuation's first line (the 日給 numerator box at 111.55 =
                // 99.55 + 12), while the trim()-empty test dropped it. harassbun's
                // S570 case (a TRUE no-run empty, text == "") still collapses.
                let s719_true_empty = std::env::var("OXI_S719_DISABLE").is_err();
                let s719_collapsible = |text: &str| -> bool {
                    if s719_true_empty {
                        text.is_empty()
                    } else {
                        text.trim().is_empty()
                    }
                };
                // S998 (2026-07-25): when this row is the interior-image case
                // (real content before AND after the image), the re-anchor must
                // also anchor the IMAGE, not just text. The generic re-anchor
                // (Step 1) moves only TEXT to page_top, leaving a straddling
                // image at its raw shift position — so the image and the text
                // AFTER it decouple and overlap (technical__0061c884: image at
                // y=4.18, VOSpace at 109.38 instead of ~415.5, the S533 stranding
                // recurrence the whole-push veto papered over). Including the
                // image in min_overflow and in the adjust keeps the image + its
                // following text coupled (image lands at page_top+tcMar_t, the
                // suffix text lands after it with the full image height reserved).
                // S1168 widens this to EVERY split that carried an image over,
                // not just S998's interior case: with the whole-row veto gone,
                // any image-bearing row can now split, and a text-only re-anchor
                // leaves the image at its raw shift while the text jumps to the
                // page top -- the S533 stranding (S1130 saw it as an image at
                // y=-26.9 on the continuation page).
                let s998_reanchor_img = (s998_interior_image || s1168)
                    && next_page_elems
                        .iter()
                        .any(|e| matches!(e.content, LayoutContent::Image { .. }));
                let anchors = |e: &LayoutElement| -> bool {
                    matches!(&e.content,
                        LayoutContent::Text { text, .. } if !s570 || !s719_collapsible(text))
                        || (s998_reanchor_img
                            && matches!(&e.content, LayoutContent::Image { .. }))
                };
                let min_overflow_text_y = next_page_elems
                    .iter()
                    .filter(|e| anchors(e))
                    .map(|e| e.y - e.flow_line_offset)
                    .fold(f32::INFINITY, f32::min);
                // S1431: when the first overflow line is the FIRST line of its
                // paragraph (no line of that paragraph stayed on this page), the
                // paragraph's space_before rides along above it.
                let s1431_sb = if std::env::var_os("OXI_S1431_DISABLE").is_none()
                    && min_overflow_text_y.is_finite()
                {
                    next_page_elems
                        .iter()
                        .filter(|e| anchors(e) && (e.y - e.flow_line_offset - min_overflow_text_y).abs() < 0.01)
                        .filter_map(|e| {
                            let key = (e.cell_col_index?, e.cell_paragraph_index?);
                            let stayed = current_page_elems.iter().any(|c| {
                                matches!(c.content, LayoutContent::Text { .. })
                                    && c.cell_row_index == Some(row_idx)
                                    && c.cell_col_index == Some(key.0)
                                    && c.cell_paragraph_index == Some(key.1)
                            });
                            if stayed { None } else { s1431_cell_para_sb.get(&key).copied() }
                        })
                        .fold(0.0f32, f32::max)
                } else {
                    0.0
                };
                let min_overflow_text_y = min_overflow_text_y - s1431_sb;
                // S1093 (2026-08-07, opt-out OXI_S1093_DISABLE): Word restarts
                // EACH CELL's remaining content at the continuation cell top —
                // the re-anchor above takes ONE global minimum and shifts the
                // whole overflow by it, which lands a cell whose first overflow
                // line sat lower than another cell's below the continuation top.
                // DERIVED (`tools/metrics/_pb_cellanchor_gen.py`, 4 arms: a
                // 2-cell row split with DIFFERENT font sizes per cell so the two
                // cells have different line phases; Word PDF baselines minus the
                // TNR ascent 0.891*fs):
                //   EQ  12/12pt  A 83.90  B 83.90            -> top 73.21 / 73.21
                //   AB1 12/ 8pt  A 83.90  B 80.30            -> top 73.21 / 73.17
                //   AB2 12/16pt  A 83.90  B 87.62            -> top 73.21 / 73.36
                //   AB3 16/ 8pt  A 87.62  B 80.30            -> top 73.36 / 73.17
                // i.e. all 12 cells restart at the SAME continuation top even
                // though their pre-split y and their baselines differ.  Real-doc
                // specimen: uk_local_spending p46/p47 row 11 — Word's ink bands
                // are col2 75.08..83.72 | 86.60..93.32 and col3 75.08..81.80, so
                // the two cells' continuations share the first line; Oxi put them
                // one line apart.  A single-cell row has one group so its adjust
                // is unchanged (harassbun S570 / tokyoshugyo S719b are 1x1), and
                // a multi-cell row whose cells share a line phase has one common
                // minimum -- both are byte-identical by construction.  Latin
                // scope, matching S817/S819/S1092 in this same region.
                let s1093 = !self.doc_body_has_real_cjk
                    && std::env::var("OXI_S1093_DISABLE").is_err();
                let mut s1093_col_min: std::collections::HashMap<usize, f32> = Default::default();
                if s1093 {
                    for e in next_page_elems.iter().filter(|e| anchors(e)) {
                        if let Some(ci) = e.cell_col_index {
                            let line_top = e.y - e.flow_line_offset;
                            let m = s1093_col_min.entry(ci).or_insert(line_top);
                            if line_top < *m {
                                *m = line_top;
                            }
                        }
                    }
                }
                if s570 && min_overflow_text_y.is_finite() {
                    next_page_elems.retain(|e| {
                        !matches!(&e.content,
                        LayoutContent::Text { text, .. }
                            if s719_collapsible(text) && e.y < min_overflow_text_y - 0.1)
                    });
                }
                // S817 (2026-07-13): Word re-applies the cell TOP margin at the
                // row-split continuation top — uklocal rt.pdf p38: continuation
                // first line ink 77.66 = page-top border 72.14 + tcMar_t 5.25
                // (the continuation region then obeys the full rowbox formula:
                // tcMar_t + n×line + after + tcMar_b + bw = 51.75 = measured
                // 51.72). Oxi anchored the first line at the raw page_top.
                // Latin scope: the JP split calibration (S570 harassbun,
                // S719b tokyoshugyo) anchors at page_top. Opt-out
                // OXI_S817_DISABLE.
                let s817_cont_pad =
                    if !self.doc_body_has_real_cjk && std::env::var("OXI_S817_DISABLE").is_err() {
                        row.cells
                            .first()
                            .and_then(|c| c.margins.as_ref().and_then(|m| m.top))
                            .unwrap_or(row_default_pad_t)
                    } else {
                        0.0
                    };
                let s817_cont_pad = s817_cont_pad + fragment_top_width;
                // S817 tail (companion): Word closes the continuation box like
                // a normal row bottom — last line + space_after + tcMar_b
                // (uklocal rt.pdf p37: row-2 close 252.74 = last line box
                // bottom 241.8 + 6 + 5.25). Applied to the Step-3 border
                // re-close AND the post-split cursor branches below. 0.0 for
                // CJK docs = byte-identical.
                let s817_tail =
                    if std::env::var_os("OXI_CELL_EMPTY_LINES").is_some() {
                        0.0
                    } else if !self.doc_body_has_real_cjk && std::env::var("OXI_S817_DISABLE").is_err() {
                        let pad_b = row
                            .cells
                            .first()
                            .and_then(|c| c.margins.as_ref().and_then(|m| m.bottom))
                            .unwrap_or(row_default_pad_b);
                        // S1528 (2026-09-24, opt-out OXI_S1528_DISABLE): the tail
                        // is the cell's RESOLVED space_after, i.e. after Word's
                        // in-cell reset of inherited spacing, not the raw style
                        // value. reports__00870bdf p32: every cell paragraph has a
                        // direct `w:spacing w:line=276` and inherits docDefaults
                        // after=200; between paragraphs Oxi already spaces 14.0
                        // like Word, but the split row's continuation closed 10pt
                        // below its last line (Word: MOSTI row at 172.6, Oxi
                        // 182.6) and the whole page ran one line long.
                        // The continuation box closes with the after of the paragraph
                        // that actually ENDS the fragment (default ON, opt-out
                        // OXI_CONT_TAIL_LIVE_DISABLE): only cells with text in the
                        // continuation count, and a nested table's S716 stub is not that
                        // paragraph -- the nested table's last row supplies the after.
                        // iiitg_fac_information row 13: Word tail 0, the finished
                        // siblings' after=3 pushed p5 down 3pt (unitG repro A..H).
                        fn tail_after(eng: &LayoutEngine, cell: &TableCell,
                                      tps: Option<&ParagraphStyle>, grid: Option<f32>) -> f32 {
                            let end = if eng.nested_table_stub_pos(cell).is_some() {
                                cell.blocks.len().saturating_sub(1)
                            } else {
                                cell.blocks.len()
                            };
                            match cell.blocks[..end].last() {
                                Some(Block::Paragraph(p)) => eng.cell_para_spacing(p, tps, grid).1,
                                Some(Block::Table(t)) => t.rows.last().map_or(0.0, |r| {
                                    r.cells.iter()
                                        .map(|c| tail_after(eng, c, t.style.para_style.as_ref(), grid))
                                        .fold(0.0_f32, f32::max)
                                }),
                                _ => 0.0,
                            }
                        }
                        let live_cells: std::collections::HashSet<usize> = next_page_elems
                            .iter()
                            .filter(|e| matches!(e.content, LayoutContent::Text { .. }))
                            .filter_map(|e| match e.cell_ancestor_path.first() {
                                Some(&(r, c, _)) => (r == row_idx).then_some(c),
                                None if e.cell_row_index == Some(row_idx) => e.cell_col_index,
                                None => None,
                            })
                            .collect();
                        let after_last = if std::env::var_os("OXI_CONT_TAIL_LIVE_DISABLE").is_none()
                            && !live_cells.is_empty()
                            && std::env::var_os("OXI_S1528_DISABLE").is_none()
                        {
                            let eng: &LayoutEngine = self;
                            row.cells
                                .iter()
                                .enumerate()
                                .filter(|(ci, _)| live_cells.contains(ci))
                                .map(|(_, c)| tail_after(eng, c, table.style.para_style.as_ref(), table_grid_pitch))
                                .fold(0.0_f32, f32::max)
                        } else {
                            row
                            .cells
                            .iter()
                            .filter_map(|c| {
                                c.blocks.iter().rev().find_map(|b| match b {
                                    Block::Paragraph(p) => Some(
                                        if std::env::var_os("OXI_S1528_DISABLE").is_none() {
                                            self.cell_para_spacing(p, table.style.para_style.as_ref(), table_grid_pitch).1
                                        } else {
                                            p.style.space_after.unwrap_or(0.0)
                                        },
                                    ),
                                    _ => None,
                                })
                            })
                            .fold(0.0_f32, f32::max)
                        };
                        pad_b + after_last
                    } else {
                        0.0
                    };
                // S1607: a whole-row move keeps the row's own geometry (top cell
                // margin included); only a genuine continuation re-anchors.
                let s1607_whole = std::env::var_os("OXI_S1607_DISABLE").is_none()
                    && split_y <= row_top + 0.1;
                if min_overflow_text_y.is_finite() && !s1607_whole {
                    let original_shift = split_y - page_top;
                    let correct_shift =
                        (min_overflow_text_y + original_shift) - (page_top + s817_cont_pad);
                    let adjust = correct_shift - original_shift;
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        let first_after = min_overflow_text_y - adjust;
                        eprintln!("[REANCHOR] page_top={:.2} split_y={:.2} min_overflow_text_y={:.2} orig_shift={:.2} correct_shift={:.2} adjust={:.2} -> first_overflow_line_y={:.2}",
                            page_top, split_y, min_overflow_text_y, original_shift, correct_shift, adjust, first_after);
                        for e in next_page_elems.iter() {
                            if let LayoutContent::Text { text, .. } = &e.content {
                                if !text.trim().is_empty() {
                                    eprintln!("    [REANCHOR-EL] y={:.2} -> {:.2} row={:?} col={:?} {:?}",
                                        e.y, e.y - adjust - original_shift + original_shift,
                                        e.cell_row_index, e.cell_col_index,
                                        text.chars().take(18).collect::<String>());
                                }
                            }
                        }
                    }
                    // S1093 fires only when the split genuinely spans MULTIPLE
                    // columns and every shifted element carries a cell index —
                    // otherwise per-cell and global adjusts would decouple
                    // elements of the same cell (a single-column row has nothing
                    // to restart independently, and keeps the S570/S719b path).
                    let s1093_ok = s1093
                        && s1093_col_min.len() >= 2
                        && next_page_elems
                            .iter()
                            .filter(|e| {
                                matches!(e.content, LayoutContent::Text { .. })
                                    || matches!(e.content, LayoutContent::Image { .. })
                            })
                            .all(|e| {
                                e.cell_col_index
                                    .map(|ci| s1093_col_min.contains_key(&ci))
                                    .unwrap_or(false)
                            });
                    for e in next_page_elems.iter_mut() {
                        if matches!(e.content, LayoutContent::Text { .. })
                            || (s998_reanchor_img
                                && matches!(e.content, LayoutContent::Image { .. }))
                        {
                            // S1093: each cell restarts at the continuation top,
                            // so use that cell's own first overflow line.
                            let cell_adjust = if s1093_ok {
                                e.cell_col_index
                                    .and_then(|ci| s1093_col_min.get(&ci))
                                    .map(|m| *m - (page_top + s817_cont_pad))
                                    .unwrap_or(adjust)
                            } else {
                                adjust
                            };
                            e.y -= cell_adjust;
                        }
                    }
                }

                let terminal_spacing = |e: &LayoutElement| -> f32 {
                    if e.cell_row_index != Some(row_idx) {
                        return 0.0;
                    }
                    e.cell_col_index.zip(e.cell_paragraph_index)
                        .and_then(|key| cell_terminal_spacing.get(&key).copied())
                        .unwrap_or(0.0)
                };
                // Step 3 (2026-04-23): On continuation page, re-close the box
                // to match actual overflow content. The shifted row_bottom lands
                // at a position that doesn't reflect the continuation line's
                // actual bottom (Oxi's row_height undersizes by one overflow line).
                // Word draws (a) top horizontal border at page_top AND (b) bottom
                // border at continuation content bottom. COM-verified on d77a p.7:
                // top=71.04 (=page_top), bottom=89.28 (=line_top 71 + line_height 18).
                {
                    let max_cont_text_bottom = next_page_elems
                        .iter()
                        .filter_map(|e| match &e.content {
                            LayoutContent::Text { .. } => Some(e.y + e.height + terminal_spacing(e)),
                            _ => None,
                        })
                        .fold(f32::NEG_INFINITY, f32::max);
                    // S817 tail (hoisted above Step 1): border re-close lands at
                    // content bottom + tail.
                    let s817_close = if max_cont_text_bottom.is_finite() {
                        max_cont_text_bottom + s817_tail
                    } else {
                        max_cont_text_bottom
                    };
                    // S942: the continuation band honors the atLeast trHeight
                    // (band = max(content, trH + bw); see the cursor branch).
                    let s817_close = if std::env::var("OXI_S942_DISABLE").is_err()
                        && (std::env::var("OXI_S940T_DISABLE").is_err()
                            || std::env::var("OXI_S1025_DISABLE").is_err())
                        // S1430 (2026-09-16, default ON, opt-out OXI_S1430_DISABLE): CJK
                        // bodies too -- `_pb_trhsplit_gen.py` n7/n9: the continuation
                        // band runs page_top 56.7 -> row 2 at 130.5 = trH 72 + bw.
                        && (!self.doc_body_has_real_cjk
                            || std::env::var_os("OXI_CELL_EMPTY_LINES").is_some()
                            || std::env::var_os("OXI_S1430_DISABLE").is_none())
                        && row.height_rule.as_deref() != Some("exact")
                    {
                        match row.height {
                            Some(trh) => {
                                s817_close.max(page_top + (trh + self.rowbox2_trh_bw(table, row)).min(content_height))
                            }
                            None => s817_close,
                        }
                    } else {
                        s817_close
                    };

                    // Find a horizontal bottom border in next_page_elems (the
                    // shifted row_bottom from the split row).
                    let bot_border_idx = next_page_elems.iter().position(|e| {
                        matches!(&e.content,
                            LayoutContent::TableBorder { y1, y2, .. }
                                if (*y1 - *y2).abs() < 0.1)
                    });

                    if let Some(bi) = bot_border_idx {
                        if max_cont_text_bottom.is_finite() {
                            // Only apply when border is above content bottom
                            // (the broken-box-top-of-page case).
                            let cur_border_y = match &next_page_elems[bi].content {
                                LayoutContent::TableBorder { y1, .. } => *y1,
                                _ => f32::INFINITY,
                            };
                            if cur_border_y < s817_close - 0.5 {
                                // Move bottom horizontal border down to content
                                // bottom (+ the S817 tail; tail=0 keeps the old
                                // content-bottom condition byte-identical).
                                if let LayoutContent::TableBorder { y1, y2, .. } =
                                    &mut next_page_elems[bi].content
                                {
                                    *y1 = s817_close;
                                    *y2 = s817_close;
                                }
                                next_page_elems[bi].y = s817_close;

                                // Extend the sides to the continuation bottom.
                                // Its top edges are repeated from their declarations below.
                                for e in next_page_elems.iter_mut() {
                                    if let LayoutContent::TableBorder { y1, y2, .. } =
                                        &mut e.content
                                    {
                                        if (*y1 - *y2).abs() >= 0.1 && *y2 < s817_close {
                                            *y2 = s817_close;
                                            e.height = *y2 - *y1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Repeat the declared top edges even when the existing bottom
                // already encloses the continuation text.
                if current_page_elems.iter().any(|e| matches!(e.content, LayoutContent::Text { .. }))
                    && next_page_elems.iter().any(|e| matches!(e.content, LayoutContent::Text { .. }))
                {
                    for mut edge in split_top_edges {
                        if let LayoutContent::TableBorder { x1, x2, y1, y2, .. } = &mut edge.content {
                            let covered = next_page_elems.iter().any(|e| matches!(&e.content,
                                LayoutContent::TableBorder { x1: a, x2: b, y1: c, y2: d, .. }
                                if (*c - *d).abs() < 0.1 && (*c - page_top).abs() < 0.1
                                    && *a <= *x1 + 0.1 && *b >= *x2 - 0.1));
                            if covered { continue; }
                            *y1 = page_top;
                            *y2 = page_top;
                            edge.y = page_top;
                            next_page_elems.push(edge);
                        }
                    }
                }

                // Close at the fragment content edge, including a completed
                // cell paragraph's trailing spacing. Word retains this rule
                // even when less than ten points remain below the last line.
                {
                    let last_text_y = current_page_elems
                        .iter()
                        .filter_map(|e| match &e.content {
                            LayoutContent::Text { .. } => Some(e.y),
                            _ => None,
                        })
                        .fold(f32::NEG_INFINITY, f32::max);

                    if last_text_y.is_finite() {
                        let line_eps = 2.0;
                        let mut max_nat: f32 = 0.0;
                        for e in current_page_elems.iter() {
                            if let LayoutContent::Text {
                                font_size,
                                font_family,
                                ..
                            } = &e.content
                            {
                                if (e.y - last_text_y).abs() < line_eps {
                                    let mut rpr = crate::ir::RunStyle::default();
                                    rpr.font_family = font_family.clone();
                                    rpr.font_size = Some(*font_size);
                                    let para_style = crate::ir::ParagraphStyle::default();
                                    let metrics = &*self.metrics_for_text("", &rpr, &para_style);
                                    let h = metrics.word_ascent_pt(*font_size)
                                        + metrics.word_descent_pt(*font_size);
                                    if h > max_nat {
                                        max_nat = h;
                                    }
                                }
                            }
                        }
                        let close_y = if !self.doc_body_has_real_cjk {
                            current_page_elems.iter().filter_map(|e| {
                                if !matches!(e.content, LayoutContent::Text { .. }) { return None; }
                                let extra = e.cell_col_index.zip(e.cell_paragraph_index)
                                    .filter(|key| s1092_last.get(key)
                                        .map_or(false, |bottom| e.y + e.height >= *bottom - 0.1))
                                    .and_then(|key| s1092_after.get(&key)).copied().unwrap_or(0.0);
                                Some(e.y + e.height + s819_q + extra)
                            }).fold(last_text_y + max_nat, f32::max)
                        } else {
                            // Exact/grid leading is part of the cell line box.
                            // Closing at glyph height alone can strike through
                            // the last line when its text has a leading offset.
                            current_page_elems.iter().filter_map(|e| {
                                matches!(e.content, LayoutContent::Text { .. })
                                    .then_some(e.y + e.height)
                            }).fold(last_text_y + max_nat, f32::max)
                        };
                        if close_y <= split_y + 0.5 {
                            let border_style = |e: &LayoutElement| match &e.content {
                                LayoutContent::TableBorder { x1, x2, y1, y2, color, width, style }
                                    if (*y1 - *y2).abs() < 0.1 =>
                                    Some((*x1, *x2, color.clone(), *width, style.clone())),
                                _ => None,
                            };
                            // A top-only cell remains open below each fragment.
                            // A continuation's top edge is not a bottom template.
                            let templates: Vec<_> =
                                split_bottom_edges.iter().filter_map(border_style).collect();
                            for (bx1, bx2, color, bw, bstyle) in templates {
                                for e in current_page_elems.iter_mut() {
                                    if let LayoutContent::TableBorder { y1, y2, .. } =
                                        &mut e.content
                                    {
                                        if (*y1 - *y2).abs() >= 0.1 {
                                            if *y2 > close_y {
                                                *y2 = close_y;
                                            }
                                            if *y1 > close_y {
                                                *y1 = close_y;
                                            }
                                        }
                                    }
                                }
                                current_page_elems.push(LayoutElement::new(
                                    bx1,
                                    close_y,
                                    bx2 - bx1,
                                    0.0,
                                    LayoutContent::TableBorder {
                                        x1: bx1,
                                        y1: close_y,
                                        x2: bx2,
                                        y2: close_y,
                                        color,
                                        width: bw,
                                        style: bstyle,
                                    },
                                ));
                            }
                        }
                    }
                }

                // Push current page elements
                if std::env::var("OXI_DBG_SPLIT").is_ok() {
                    let cur_txt = current_page_elems
                        .iter()
                        .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                        .count();
                    let nxt_txt = next_page_elems
                        .iter()
                        .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                        .count();
                    let nxt_maxy = next_page_elems
                        .iter()
                        .map(|e| match &e.content {
                            LayoutContent::TableBorder { y2, .. } => *y2,
                            _ => e.y + e.height,
                        })
                        .fold(0.0_f32, f32::max);
                    eprintln!("[SPLIT] split_y={:.1} ptop={:.1} pbot={:.1} pages_pre={} | cur_elems={} (txt {}) | next_elems={} (txt {}) next_maxy={:.1}",
                        split_y, page_top, page_bottom, pages.len(), current_page_elems.len(), cur_txt, next_page_elems.len(), nxt_txt, nxt_maxy);
                }
                elements.extend(current_page_elems);
                current_elements.extend(std::mem::take(&mut elements));
                dbg_page_push(pages.len(), 0);
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });

                page_bottom += std::mem::take(&mut first_page_fit_offset);
                // S1527: the finished page's note reserve belongs to that page;
                // the continuation page starts clean and S1527 subtracts only the
                // notes whose referencing lines land on it (S740 v1 carried the
                // entry page's reserve across every continuation page).
                if row_footnotes.is_none() || std::env::var_os("OXI_S1527_DISABLE").is_none() {
                    page_bottom += std::mem::take(&mut s740_reserve);
                }
                page_bottom += advance_table_page_geometry(
                    page_geometry, pages.len() + 1, &mut page_top,
                    &mut content_height, &mut next_page_elems,
                );

                // Handle multi-page overflow: if next_page_elems still overflow,
                // keep splitting into additional pages.
                //
                // S485 (TRIED + REVERTED, finding only): Word repeats the box TOP
                // border at each page's content top for a bordered table spanning
                // 3+ pages; Oxi's continuation fragments render with an OPEN top
                // (confirmed e3c545 p5/p8: content + side borders match Word, top
                // missing). A synth_top closure added a top horizontal at page_top
                // to this_page/remaining here — instrumented (OXI_S485_DEBUG) it
                // DID fire & ADD (vt=true, has_top=false) but the render was
                // byte-identical (delta 0.00000 all 12 pages): the overflow-loop
                // fragments are NOT the final rendered pages — `elements`/`this_page`
                // get further processed downstream and the synth'd border is
                // dropped/repositioned. The correct synthesis point is the
                // final-fragment render path, which needs flow-tracing through the
                // post-loop `elements` handling (S269/Day34 multi-layer split).
                // Deferred — multi-session. Reverted (byte-identical).
                let mut remaining = next_page_elems;
                // S754b: replay the captured tblHeader at the top of EVERY
                // split-continuation page (Word repeats it above the continued
                // row content — probethdr: each continuation missing the header
                // packed ~1 header-height too much → −1×3/−3). The continuation
                // content shifts down by hdr_h and the header clones sit at
                // page_top; the loop's fit test then accounts for the header
                // automatically. Applied after the first split and after each
                // loop push (headers already placed stay above next_split so
                // they never re-enter overflow).
                let s754_hdr_replay = std::env::var("OXI_S754_DISABLE").is_err()
                    && s728_on
                    && s728_capture_done
                    && !s728_hdr_elems.is_empty()
                    && !row.header;
                let s754_apply_hdr = |els: &mut Vec<LayoutElement>,
                                      hdr: &Vec<LayoutElement>,
                                      hdr_h: f32,
                                      pt: f32| {
                    for e in els.iter_mut() {
                        e.y += hdr_h;
                        if let LayoutContent::TableBorder {
                            ref mut y1,
                            ref mut y2,
                            ..
                        } = e.content
                        {
                            *y1 += hdr_h;
                            *y2 += hdr_h;
                        }
                    }
                    let y0 = hdr.iter().map(|e| e.y).fold(f32::INFINITY, f32::min);
                    if y0.is_finite() {
                        let dy = pt - y0;
                        for el in hdr {
                            let mut c = el.clone();
                            c.y += dy;
                            if let LayoutContent::TableBorder {
                                ref mut y1,
                                ref mut y2,
                                ..
                            } = c.content
                            {
                                *y1 += dy;
                                *y2 += dy;
                            }
                            els.push(c);
                        }
                    }
                };
                if s754_hdr_replay {
                    s754_apply_hdr(&mut remaining, &s728_hdr_elems, s728_hdr_h + self.repeated_header_border_delta(table, row_idx), page_top);
                }
                let continuation_pad_b = if separate_outer_edges
                    && std::env::var("OXI_S819_DISABLE").is_err() {
                    row.cells.first().and_then(|c| c.margins.as_ref().and_then(|m| m.bottom))
                        .unwrap_or(row_default_pad_b)
                } else { 0.0 };
                // 2026-09-25: hard cap on continuation pages for ONE row. The
                // loop's progress guarantee (overflow shifts up by a page each
                // pass) can be broken by a re-anchor that shifts it back down
                // (S1530 before S1530b); without a cap the renderer allocated
                // pages until the machine's memory was gone. 4096 pages for a
                // single row is far beyond any real document.
                let mut split_loop_pages = 0usize;
                loop {
                    split_loop_pages += 1;
                    if split_loop_pages > 4096 {
                        eprintln!("[SPLIT-GUARD] row {} continuation split made no progress after {} pages; placing the rest on the current page", row_idx, split_loop_pages - 1);
                        break;
                    }
                    // Find the maximum Y in remaining elements.
                    // R7.77 (Session 62, 2026-05-16): exclude PresetShape elements
                    // from the max_y check. PresetShapes (e.g. 3a4f9f Shape A
                    // cy=686.6pt with wrap=wrapNone, H position off-page) are
                    // overlays — they don't reserve text flow space in Word. When
                    // their height exceeds page content_height (657pt), they cause
                    // the row-split loop to iterate indefinitely (or push extra
                    // pages until the shape "fits"), driving Sub-jump 3b in 3a4f9f
                    // (wi=1042→1045 +1 page cascade). The shape itself is still
                    // partitioned and rendered; only its height is excluded from
                    // the page-fit determination.
                    let max_y = remaining
                        .iter()
                        .map(|e| {
                            match &e.content {
                                LayoutContent::TableBorder { y2, .. } => *y2,
                                LayoutContent::PresetShape { .. } => e.y, // ignore height
                                LayoutContent::Text { .. } => e.y + e.height + continuation_pad_b,
                                _ => e.y + e.height,
                            }
                        })
                        .fold(0.0_f32, f32::max);

                    let continuation_bottom = page_bottom - fragment_bottom_width;
                    if max_y <= continuation_bottom + row_fit_epsilon {
                        // Everything fits on this page
                        break;
                    }

                    // Need another split at page_bottom (or earlier if a mid-cell
                    // LRPB marker exists at y < page_bottom).
                    // R7.56 (Day 34 part 25): pull split back to first LRPB-marked
                    // element above page_top. Same logic as the first-split path.
                    let lrpb_next_split = remaining
                        .iter()
                        .filter(|e| e.is_paragraph_start_with_lrpb && e.y > page_top + 0.5)
                        .map(|e| e.y)
                        .fold(f32::INFINITY, f32::min);
                    // S565 (2026-06-14): page-half-full gate on the overflow-loop
                    // LRPB pull-back (mirror of S563 for the body s391 path). A
                    // STALE lastRenderedPageBreak inside a multi-page row cell
                    // (harassbun: 1 table / 1 row / 1 cell, an LRPB at y=64.5 only
                    // 16.5pt below page_top) pulled the continuation split to ~1
                    // line, spawning a near-empty page (p2) — same class as S563
                    // but in the row-split overflow loop (R7.56), which S564
                    // missed (S564 gated the FIRST-split path at 11348, but the
                    // first split here falls back to page_bottom because
                    // lrpb_split_y 807.5 > page_bottom; the real stale LRPB is in
                    // THIS loop). Only honour the LRPB once the continuation page
                    // is at least half full. Opt-out OXI_S565_DISABLE.
                    let s565_half_full = std::env::var("OXI_S565_DISABLE").is_ok()
                        || lrpb_next_split > page_top + content_height * 0.5;
                    // S1527: notes referenced by the lines landing on THIS
                    // continuation page shrink its bottom (see the first split).
                    let continuation_bottom = if !s1527_refs.is_empty() {
                        let (bottom, kept) = s1527_reserve(continuation_bottom, false, &remaining, &s1527_early);
                        if !kept.is_empty() {
                            if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                eprintln!("[SPLIT-S1527] row={} loop kept_notes={:?} bottom {:.2} -> {:.2}", row_idx, kept, continuation_bottom, bottom);
                            }
                            let off = pages.len() - s740_entry_pages;
                            while s740_fn_pages.len() <= off {
                                s740_fn_pages.push(Vec::new());
                            }
                            for id in &kept {
                                if !s740_fn_pages[off].contains(id) {
                                    s740_fn_pages[off].push(*id);
                                }
                            }
                            s1527_early.extend(kept.iter().copied());
                        }
                        bottom
                    } else {
                        continuation_bottom
                    };
                    let next_split = if lrpb_next_split.is_finite()
                        && lrpb_next_split < continuation_bottom
                        && s565_half_full
                    {
                        lrpb_next_split
                    } else {
                        continuation_bottom
                    };
                    let mut this_page: Vec<LayoutElement> = Vec::new();
                    let mut overflow: Vec<LayoutElement> = Vec::new();

                    // S1092: recompute the per-paragraph LAST-line bottom from the
                    // elements still to place (they are re-shifted each iteration,
                    // so the first-split map's absolute values are stale here).
                    let mut s1092_ovlast: std::collections::HashMap<(usize, usize), f32> =
                        Default::default();
                    if s1092 {
                        for e in remaining.iter() {
                            if matches!(
                                e.content,
                                LayoutContent::TableBorder { .. }
                                    | LayoutContent::CellShading { .. }
                            ) {
                                continue;
                            }
                            if let (Some(ci), Some(pi)) =
                                (e.cell_col_index, e.cell_paragraph_index)
                            {
                                let bt = e.y + e.height;
                                let cur = s1092_ovlast.entry((ci, pi)).or_insert(bt);
                                if bt > *cur {
                                    *cur = bt;
                                }
                            }
                        }
                    }

                    for elem in remaining {
                        let _elem_top = elem.y;
                        match &elem.content {
                            LayoutContent::TableBorder {
                                y1,
                                y2,
                                x1,
                                x2,
                                ref color,
                                width,
                                ref style,
                            } => {
                                if (y1 - y2).abs() < 0.1 {
                                    if *y1 <= next_split + 0.5 {
                                        this_page.push(elem);
                                    } else {
                                        let shift = next_split - page_top;
                                        let mut e = elem;
                                        e.y -= shift;
                                        if let LayoutContent::TableBorder {
                                            ref mut y1,
                                            ref mut y2,
                                            ..
                                        } = e.content
                                        {
                                            *y1 -= shift;
                                            *y2 -= shift;
                                        }
                                        overflow.push(e);
                                    }
                                } else {
                                    let vy_top = *y1;
                                    let vy_bot = *y2;
                                    if vy_top < next_split {
                                        this_page.push(LayoutElement::new(
                                            elem.x,
                                            elem.y,
                                            elem.width,
                                            next_split - vy_top,
                                            LayoutContent::TableBorder {
                                                x1: *x1,
                                                y1: vy_top,
                                                x2: *x2,
                                                y2: next_split,
                                                color: color.clone(),
                                                width: *width,
                                                style: style.clone(),
                                            },
                                        ));
                                    }
                                    if vy_bot > next_split {
                                        let shift = next_split - page_top;
                                        let new_y1 = page_top;
                                        let new_y2 = vy_bot - shift;
                                        overflow.push(LayoutElement::new(
                                            elem.x,
                                            new_y1,
                                            elem.width,
                                            new_y2 - new_y1,
                                            LayoutContent::TableBorder {
                                                x1: *x1,
                                                y1: new_y1,
                                                x2: *x2,
                                                y2: new_y2,
                                                color: color.clone(),
                                                width: *width,
                                                style: style.clone(),
                                            },
                                        ));
                                    }
                                }
                            }
                            LayoutContent::CellShading { ref color } => {
                                let shade_bottom = elem.y + elem.height;
                                if elem.y < next_split {
                                    let clip_h = (next_split - elem.y).min(elem.height);
                                    this_page.push(LayoutElement::new(
                                        elem.x,
                                        elem.y,
                                        elem.width,
                                        clip_h,
                                        LayoutContent::CellShading {
                                            color: color.clone(),
                                        },
                                    ));
                                }
                                if shade_bottom > next_split {
                                    let shift = next_split - page_top;
                                    let new_y = (elem.y - shift).max(page_top);
                                    let new_h = shade_bottom - shift - new_y;
                                    overflow.push(LayoutElement::new(
                                        elem.x,
                                        new_y,
                                        elem.width,
                                        new_h.max(0.0),
                                        LayoutContent::CellShading {
                                            color: color.clone(),
                                        },
                                    ));
                                }
                            }
                            _ => {
                                // Day 34 part 24 (2026-05-13): use element BOTTOM
                                // vs next_split, NOT top. Mirrors the same fix
                                // applied to the first split at line 6807 on
                                // 2026-04-22 (Step 1). For rows that span 3+ pages,
                                // the multi-page loop here was still using top-only
                                // check, so a line whose top fits but bottom
                                // overflows incorrectly stayed on the current page.
                                // e3c545 cpi=82 at y=777.25 h=11.62 (bottom=788.87)
                                // on page 5 with split_y=785.2: top-check kept it
                                // on p5 (777.25<785.2), bottom-check moves to p6
                                // (788.87>785.3). Fixes 4 -1 outliers in e3c545.
                                let elem_bottom = elem.y + elem.height;
                                // S1092: the LAST line of a cell paragraph must also
                                // fit that paragraph's own space_after.
                                let s1092_extra = if s1092 {
                                    match (elem.cell_col_index, elem.cell_paragraph_index) {
                                        (Some(ci), Some(pi)) => {
                                            if s1092_ovlast
                                                .get(&(ci, pi))
                                                .map_or(false, |b| (elem_bottom - *b).abs() < 0.01)
                                            {
                                                *s1092_after.get(&(ci, pi)).unwrap_or(&0.0)
                                            } else {
                                                0.0
                                            }
                                        }
                                        _ => 0.0,
                                    }
                                } else {
                                    0.0
                                };
                                if elem_bottom + s1092_extra <= next_split + row_fit_epsilon - continuation_pad_b {
                                    this_page.push(elem);
                                } else {
                                    let shift = next_split - page_top;
                                    let mut e = elem;
                                    e.y -= shift;
                                    overflow.push(e);
                                }
                            }
                        }
                    }

                    // S1530 (2026-09-24, opt-out OXI_S1530_DISABLE): widow/orphan
                    // control for a cell paragraph cut by the continuation split.
                    // Word never leaves a single line of a widowControl paragraph
                    // on either side of the cut: one line staying -> the whole
                    // paragraph moves; one line moving -> a second line goes with
                    // it (and when that would leave one line, everything moves).
                    // policies__0097185c p5: "e) Is living in housing" (4 lines)
                    // had one line left above the footnote reserve; Word starts
                    // it on p6 (PDF 93.1) while Oxi kept line 1 at 603.4. The
                    // first-split path has its own widow handling; this loop
                    // partitioned line by line with no look at the paragraph.
                    if std::env::var_os("OXI_S1530_DISABLE").is_none() {
                        let line_key = |e: &LayoutElement| ((e.y - e.flow_line_offset) * 10.0).round() as i64;
                        let shift = next_split - page_top;
                        let is_txt = |e: &LayoutElement| matches!(e.content, LayoutContent::Text { .. });
                        // S1530b (2026-09-25): a paragraph is the host cell's OWN
                        // paragraph only when its ancestor path is empty.
                        // `cell_paragraph_index` restarts at 0 inside a nested
                        // table, so keying on (col, para) alone merged a nested
                        // table's lines with the host cell's paragraphs
                        // (administrative__0001ce58b20a6729: the Zoom-invite
                        // nested table and the "Agenda" paragraph both carried
                        // (0, 1)). The phantom orphan sat at the page TOP, was
                        // moved into the overflow, re-anchored to the same y, and
                        // the loop never advanced — the renderer allocated pages
                        // until the machine's memory was gone (three unclean
                        // reboots 2026-09-24/25). Nested paragraphs are left to
                        // the nested table's own layout.
                        let direct = |e: &LayoutElement| e.cell_ancestor_path.is_empty() && e.cell_row_index == Some(row_idx);
                        let mut moved_any = false;
                        loop {
                            let mut moved = false;
                            let mut keys: Vec<(usize, usize)> = Vec::new();
                            for e in overflow.iter() {
                                if let (Some(ci), Some(pi)) = (e.cell_col_index, e.cell_paragraph_index) {
                                    if is_txt(e) && direct(e) && !keys.contains(&(ci, pi)) {
                                        keys.push((ci, pi));
                                    }
                                }
                            }
                            // The topmost text line of the page never moves: a
                            // paragraph that starts at the page top has nowhere
                            // higher to go, and emptying the page makes no
                            // progress.
                            let page_min_key = this_page.iter().filter(|e| is_txt(e)).map(|e| line_key(e)).min();
                            for (ci, pi) in keys {
                                let widow_on = row.cells.get(ci).map_or(true, |c| {
                                    c.blocks
                                        .iter()
                                        .filter_map(|b| match b { Block::Paragraph(p) => Some(p), _ => None })
                                        .nth(pi)
                                        .map_or(true, |p| p.style.widow_control)
                                });
                                if !widow_on {
                                    continue;
                                }
                                let mut stay: Vec<i64> = Vec::new();
                                for e in this_page.iter() {
                                    if e.cell_col_index == Some(ci) && e.cell_paragraph_index == Some(pi) && is_txt(e) && direct(e) {
                                        let k = line_key(e);
                                        if !stay.contains(&k) {
                                            stay.push(k);
                                        }
                                    }
                                }
                                if stay.is_empty() {
                                    continue;
                                }
                                let mut go: Vec<i64> = Vec::new();
                                for e in overflow.iter() {
                                    if e.cell_col_index == Some(ci) && e.cell_paragraph_index == Some(pi) && is_txt(e) && direct(e) {
                                        let k = line_key(e);
                                        if !go.contains(&k) {
                                            go.push(k);
                                        }
                                    }
                                }
                                let move_all = stay.len() == 1 || (go.len() == 1 && stay.len() == 2);
                                let move_last = !move_all && go.len() == 1 && stay.len() >= 3;
                                if !move_all && !move_last {
                                    continue;
                                }
                                if move_all && page_min_key == stay.iter().min().copied() {
                                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                        eprintln!("[SPLIT-S1530] row={} cell={} para={} starts at the page top: kept", row_idx, ci, pi);
                                    }
                                    continue;
                                }
                                let last_key = *stay.iter().max().unwrap();
                                let mut i = 0;
                                while i < this_page.len() {
                                    let e = &this_page[i];
                                    let hit = e.cell_col_index == Some(ci)
                                        && e.cell_paragraph_index == Some(pi)
                                        && is_txt(e)
                                        && direct(e)
                                        && (move_all || line_key(e) == last_key);
                                    if hit {
                                        let mut e = this_page.remove(i);
                                        e.y -= shift;
                                        overflow.push(e);
                                        moved = true;
                                    } else {
                                        i += 1;
                                    }
                                }
                                if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                    eprintln!("[SPLIT-S1530] row={} cell={} para={} stay={} go={} move_all={} move_last={}", row_idx, ci, pi, stay.len(), go.len(), move_all, move_last);
                                }
                            }
                            moved_any |= moved;
                            if !moved {
                                break;
                            }
                        }
                        if moved_any {
                            // The moved lines were re-based by the split shift and
                            // now sit above the continuation top; slide every
                            // overflow line down so the first one starts there.
                            let min_ov_y = overflow.iter().filter(|e| is_txt(e)).map(|e| e.y).fold(f32::INFINITY, f32::min);
                            let cont_top = page_top + s817_cont_pad;
                            if min_ov_y.is_finite() && min_ov_y < cont_top - 0.1 {
                                let adjust = cont_top - min_ov_y;
                                for e in overflow.iter_mut() {
                                    if is_txt(e) || (float_replay.is_some() && e.margin_float) {
                                        e.y += adjust;
                                    }
                                }
                                if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                    eprintln!("[SPLIT-S1530] re-anchored overflow text +{:.2}", adjust);
                                }
                            }
                        }
                    }
                    // S719b (2026-07-02, default ON, opt-out OXI_S719_DISABLE): the
                    // overflow loop lacked Step-1's re-anchor — a text line STRADDLING
                    // the split boundary (top < next_split, bottom > next_split) gets
                    // re-based by `next_split − page_top` and lands ABOVE page_top
                    // (into the margin, over the box border). tokyoshugyo p50/51: the
                    // whitespace exact-240 spacer (top 747.5, bottom 759.5, split
                    // 756.85) landed at y=90.15 above the border 99.5 → the +12 line
                    // Word renders at the continuation top was lost → the S710 flat
                    // +2.4/block bump compensated (now retired to 0). Port the Step-1
                    // re-anchor for the straddle case only (min text y < page_top):
                    // collapse leading TRUE-empty lines (S570/S719 semantics), then
                    // shift Text elements down so the first content line sits at
                    // page_top. Straddle-only keeps non-straddling rounds byte-
                    // identical (the tuned corpus).
                    if s719_true_empty {
                        let min_ov_y = overflow.iter()
                            .filter(|e| matches!(&e.content,
                                LayoutContent::Text { text, .. } if !s570 || !s719_collapsible(text)))
                            .map(|e| e.y)
                            .fold(f32::INFINITY, f32::min);
                        let continuation_text_top = page_top + s817_cont_pad;
                        if min_ov_y.is_finite() && min_ov_y < continuation_text_top - 0.1 {
                            if s570 {
                                overflow.retain(|e| {
                                    !matches!(&e.content,
                                    LayoutContent::Text { text, .. }
                                        if s719_collapsible(text) && e.y < min_ov_y - 0.1)
                                });
                            }
                            let adjust = continuation_text_top - min_ov_y;
                            for e in overflow.iter_mut() {
                                if matches!(e.content, LayoutContent::Text { .. })
                                    || (float_replay.is_some() && e.margin_float) {
                                    e.y += adjust;
                                }
                            }
                            if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                eprintln!("[SPLIT-REANCHOR] min_ov_y={:.2} < page_top={:.2} -> shifted overflow text +{:.2}",
                                    min_ov_y, page_top, adjust);
                            }
                        }
                    }
                    // A fresh paragraph carries the same resolved before gap
                    // on every continuation, including the third and later pages.
                    // Nested tables have their own flow identities and carry map.
                    if std::env::var_os("OXI_S1431_DISABLE").is_none() {
                        let first_top = overflow.iter().filter(|e| anchors(e))
                            .map(|e| e.y - e.flow_line_offset)
                            .fold(f32::INFINITY, f32::min);
                        let carry = overflow.iter()
                            .filter(|e| anchors(e)
                                && e.cell_ancestor_path.is_empty()
                                && e.cell_row_index == Some(row_idx)
                                && (e.y - e.flow_line_offset - first_top).abs() < 0.01)
                            .filter_map(|e| {
                                let key = (e.cell_col_index?, e.cell_paragraph_index?);
                                let stayed = this_page.iter().any(|c| {
                                    anchors(c) && c.cell_ancestor_path.is_empty()
                                        && c.cell_row_index == Some(row_idx)
                                        && c.cell_col_index == Some(key.0)
                                        && c.cell_paragraph_index == Some(key.1)
                                });
                                if stayed { None } else { s1431_cell_para_sb.get(&key).copied() }
                            }).fold(0.0f32, f32::max);
                        let target_top = page_top + s817_cont_pad + carry;
                        if carry > 0.0 && first_top.is_finite() && first_top < target_top - 0.01 {
                            let adjust = target_top - first_top;
                            for e in overflow.iter_mut() {
                                if matches!(e.content, LayoutContent::Text { .. })
                                    || (s998_reanchor_img && matches!(e.content, LayoutContent::Image { .. }))
                                    || (float_replay.is_some() && e.margin_float)
                                { e.y += adjust; }
                            }
                        }
                    }
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        let tp_txt = this_page
                            .iter()
                            .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                            .count();
                        let ov_txt = overflow
                            .iter()
                            .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                            .count();
                        eprintln!("[SPLIT-LOOP] next_split={:.1} -> pushed this_page txt={} | overflow txt={}", next_split, tp_txt, ov_txt);
                    }
                    dbg_page_push(pages.len(), 0);
                    pages.push(LayoutPage {
                        width: page_width,
                        height: page_height,
                        elements: this_page,
                    });
                    page_bottom += std::mem::take(&mut first_page_fit_offset);
                // S1527: the finished page's note reserve belongs to that page;
                // the continuation page starts clean and S1527 subtracts only the
                // notes whose referencing lines land on it (S740 v1 carried the
                // entry page's reserve across every continuation page).
                if row_footnotes.is_none() || std::env::var_os("OXI_S1527_DISABLE").is_none() {
                    page_bottom += std::mem::take(&mut s740_reserve);
                }
                page_bottom += advance_table_page_geometry(
                        page_geometry, pages.len() + 1, &mut page_top,
                        &mut content_height, &mut overflow,
                    );
                    remaining = overflow;
                    // S754b: header on the next continuation page too.
                    if s754_hdr_replay {
                        s754_apply_hdr(&mut remaining, &s728_hdr_elems, s728_hdr_h + self.repeated_header_border_delta(table, row_idx), page_top);
                    }
                }

                if std::env::var("OXI_DBG_SPLIT").is_ok() {
                    let rem_txt = remaining
                        .iter()
                        .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                        .count();
                    let rem_maxy = remaining
                        .iter()
                        .map(|e| match &e.content {
                            LayoutContent::TableBorder { y2, .. } => *y2,
                            _ => e.y + e.height,
                        })
                        .fold(0.0_f32, f32::max);
                    eprintln!(
                        "[SPLIT-END] pages_now={} | remaining(=p_cont) elems={} txt={} maxy={:.1}",
                        pages.len(),
                        remaining.len(),
                        rem_txt,
                        rem_maxy
                    );
                }
                elements = remaining;
                // A spanning cell can continue through later rows. Its remaining
                // height does not set the end of this row's independent cells.
                // A spanning cell's continuation flow follows modern document
                // compatibility, the same boundary that gates its widow
                // protection below. Legacy modes keep the row-height model.
                let row_continuation_flow = std::env::var_os("OXI_ROW_CONTINUATION_FLOW_DISABLE").is_none()
                    && self.compat_mode_explicit && self.compat_mode >= 15
                    && !is_nested && row_footnotes.is_none();
                let mut independent_continuation_end = f32::NEG_INFINITY;
                let mut independent_continuation_ink = false;
                if row_continuation_flow {
                    for e in &elements {
                        if e.cell_row_index != Some(row_idx) { continue; }
                        let Some(cell) = e.cell_col_index.and_then(|ci| row.cells.get(ci)) else { continue; };
                        if cell.v_merge.is_some() { continue; }
                        let bottom = cell.margins.as_ref().and_then(|m| m.bottom)
                            .unwrap_or(row_default_pad_b) + self.rowbox2_trh_bw(table, row);
                        match &e.content {
                            LayoutContent::Text { text, .. } => {
                                independent_continuation_end = independent_continuation_end
                                    .max(e.y + e.height + terminal_spacing(e) + bottom);
                                independent_continuation_ink |= !text.trim().is_empty();
                            }
                            LayoutContent::Image { .. } => {
                                let advance = e.flow_line_height.unwrap_or(e.height);
                                if advance > 0.0 {
                                    independent_continuation_end = independent_continuation_end.max(e.y - e.flow_line_offset + advance + bottom);
                                    independent_continuation_ink = true;
                                }
                            }
                            _ => {}
                        }
                    }
                }
                let independent_continuation_end = if row.cells.iter().any(|cell| cell.v_merge.is_some()) {
                    independent_continuation_end
                } else { f32::NEG_INFINITY };
                // Empty paragraphs in one column do not erase a sibling's text.
                let other_cell_empty_continues = std::env::var_os("OXI_EMPTY_TAIL_CELL_SCOPE").is_some()
                    && elements.iter().any(|e| {
                        e.cell_row_index == Some(row_idx)
                            && matches!(&e.content, LayoutContent::Text { text, .. } if text.trim().is_empty())
                            && e.cell_col_index.and_then(|ci| row.cells.get(ci)).map_or(true, |cell| {
                                cell.blocks.iter().rev().take_while(|block| {
                                    matches!(block, Block::Paragraph(p) if p.runs.iter().all(|r| r.text.is_empty()))
                                }).count() < 2
                            })
                    });
                // A sibling's trailing blank paragraphs cannot collapse actual
                // content on this continuation. This is independent of the
                // compatibility rule for vertically merged cell extents.
                let continuation_has_ink = std::env::var_os("OXI_CELL_CONTINUATION_INK_DISABLE").is_none()
                    && elements.iter().any(|e| {
                        if e.cell_row_index != Some(row_idx) { return false; }
                        let Some(cell) = e.cell_col_index.and_then(|ci| row.cells.get(ci)) else { return false; };
                        if cell.v_merge.is_some() { return false; }
                        match &e.content {
                            LayoutContent::Text { text, .. } => !text.trim().is_empty(),
                            LayoutContent::Image { .. } => e.flow_line_height.unwrap_or(e.height) > 0.0,
                            _ => false,
                        }
                    });
                let s864_empty_tail_split = s864_empty_tail_split
                    && !independent_continuation_ink && !other_cell_empty_continues
                    && !continuation_has_ink;
                // S269 Pattern A fix (default ON since S269 part 7): replace
                // geometric overflow with structural line_pitch snap matching
                // Word's measured formula `body_y = last_cont_top + lh ×
                // (1 + trailing_empty)`. 4 real-doc splits (d77a t5/t8/t10 +
                // e3c545 t2) + CR_6 minimal repro confirm formula (residuals
                // ≤ 1pt). Original geometric formula undercount ~1 line_pitch
                // per wrap caused -15pt/wrap drift (S264 d77a) cascading
                // through subsequent body paragraphs.
                //
                // Phase 1+2+SSIM verify all met before flipping default
                // (commit dda9a58 + S269 part 6 SSIM measurement on multi-page
                // baseline). OXI_PATTERN_A_DISABLE=1 opt-out for diagnostic.
                //
                // trailing_empty_count = max across row.cells of consecutive
                // trailing empty paragraphs. v3 data shows boolean 0/1 (no doc
                // observed with 2+ trailing empties), but counting handles future
                // cases. d77a t8/t10 + e3c545 t2 each have 1 trailing empty
                // (formula ×2); CR_6 has 0 (formula ×1).
                // S269 part 5: gate fix on (single-column rows) OR (no-border tables).
                // Multi-col bordered tables (ed025 10x4 / b35123 13x2 etc.) show
                // -0.29/-0.33 IoU regression with fix because cell-wise trailing_empty
                // in shorter cells doesn't translate to row-bottom advance — longer
                // cells already determine the row geometry. The structural formula
                // `last_cont_top + lh × (1+te)` was derived from 1x1 (d77a t5/t8/t10 +
                // e3c545 t2) and generalizes cleanly to:
                //   (a) single-column N-row tables where each row's cell determines bottom
                //   (b) no-border layout tables (d4d126 31x4 border=false) where the row's
                //       bottom is similarly determined by the longest cell's content
                let allow_fix = row.cells.len() == 1 || !table.style.border;
                let fix_disabled = std::env::var("OXI_PATTERN_A_DISABLE").is_ok()
                    || (std::env::var("OXI_TYPED_GRID_CONTINUATION_DISABLE").is_err()
                        && !self.doc_body_has_real_cjk
                        && table_grid_pitch.is_some()
                        && !page.doc_grid_no_type);
                if !fix_disabled && allow_fix {
                    let last_cont_top = elements
                        .iter()
                        .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                        .map(|e| e.y)
                        .fold(f32::NEG_INFINITY, f32::max);
                    let te_of = |cell: &TableCell| {
                        let te = cell
                            .blocks
                            .iter()
                            .rev()
                            .take_while(|b| {
                                matches!(b, Block::Paragraph(p)
                                if p.runs.iter().all(|r| r.text.is_empty()))
                            })
                            .count();
                        // S716: the post-nested-table stub para is collapsed
                        // to ~0 by Word — exclude it from the split-continuation
                        // formula (the stub is Some only when te == 1 and the
                        // preceding block is a nested table).
                        if self.nested_table_stub_pos(cell).is_some()
                            || self.s1311_hidemark_tail_pos(cell).is_some()
                        {
                            te.saturating_sub(1)
                        } else {
                            te
                        }
                    };
                    // S908: which cells' trailing empties push the split-
                    // continuation cursor. Two real-doc pins:
                    //   uklocal Annex row 5 — sibling = TEXT + trailing empty:
                    //     the empty is placed right after its text on the first
                    //     page → contributes NOTHING (max-over-ALL credited it →
                    //     a +11.5 phantom line; Word close = last line + after +
                    //     tcMar_b = the S817 B3 model, rt.pdf 123.9 EXACT).
                    //   d4d126 row 22 — sibling = an ENTIRELY-EMPTY cell (its
                    //     single empty para IS the cell): Word DOES count it
                    //     (word_png/rt.pdf p5 (１)提供媒体 110.7 = the te=1
                    //     cursor; te=0 was −14.6 = the v1/v2 −0.0729 regression).
                    // Rule: count (a) cells with text in the continuation (their
                    // trailing empties follow the text below the split) and
                    // (b) entirely-empty cells (the unplaced ¶ lands in the
                    // continuation); a sibling's post-text empties do not.
                    // The d77a/e3c545 derivation was 1×1 tables where all
                    // variants coincide.
                    let trailing_empty_count = if std::env::var("OXI_S908_DISABLE").is_err() {
                        let mut cols: Vec<usize> = elements
                            .iter()
                            .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                            .filter_map(|e| e.cell_col_index)
                            .collect();
                        cols.sort_unstable();
                        cols.dedup();
                        let all_empty = |cell: &TableCell| {
                            cell.blocks.iter().all(|b| {
                                matches!(b, Block::Paragraph(p)
                                    if p.runs.iter().all(|r| r.text.is_empty()))
                            })
                        };
                        // S1440: an all-empty cell's empty lines were placed from the
                        // row top; the continuation only owes the ones past the split.
                        let s1440_placed: usize = if std::env::var_os("OXI_S1440_DISABLE").is_none() {
                            let lh = table_grid_pitch.unwrap_or_else(|| {
                                elements
                                    .iter()
                                    .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                                    .filter(|e| (e.y - last_cont_top).abs() < 0.5)
                                    .map(|e| e.height)
                                    .fold(0.0_f32, f32::max)
                            });
                            if lh > 0.0 {
                                (((split_y - row_top).max(0.0) / lh) + 0.01).floor() as usize
                            } else {
                                0
                            }
                        } else {
                            0
                        };
                        row.cells
                            .iter()
                            .enumerate()
                            .filter(|(ci, cell)| cols.binary_search(ci).is_ok() || all_empty(cell))
                            .map(|(_, cell)| {
                                if all_empty(cell) {
                                    te_of(cell).saturating_sub(s1440_placed)
                                } else {
                                    te_of(cell)
                                }
                            })
                            .max()
                            .unwrap_or_else(|| row.cells.iter().map(te_of).max().unwrap_or(0))
                    } else {
                        row.cells.iter().map(te_of).max().unwrap_or(0)
                    };
                    // With explicit non-painting line boxes, the last continuation
                    // line already includes each cell's empty paragraphs. Cells run
                    // in parallel: do not append a sibling's blank lines again.
                    let trailing_empty_count = if std::env::var_os("OXI_CELL_EMPTY_LINES").is_some() {
                        0
                    } else { trailing_empty_count };
                    if last_cont_top.is_finite() {
                        if let Some(lh) = table_grid_pitch {
                            cursor.set(
                                last_cont_top
                                    + lh * (1.0 + trailing_empty_count as f32)
                                    + s817_tail,
                            );
                        } else {
                            // S304 (2026-05-26): no-docGrid extension of Pattern A.
                            // When docGrid is absent, derive `lh` from the last text
                            // element's own height. Same formula shape — cursor lands
                            // at last_cont_top + lh × (1 + trailing_empty) — so the
                            // body content that follows the table starts at last
                            // text's bottom + trailing-empty space (if any).
                            //
                            // The pre-fix geometric formula at `row_bottom - split_y`
                            // undercounts when many wrap lines overflow to the next
                            // page (`row_height` derived from cell content stays in
                            // line-pitch units while `row_bottom - split_y` collapses
                            // page-fold geometry). e3c545 LOD code listing (1×1 table,
                            // 70+ lines, no docGrid) showed a uniform -11pt cursor
                            // drift on p6 → cascades 12-15pt across 24 body
                            // paragraphs (S304 diagnosis).
                            //
                            // OXI_PATTERN_A_DISABLE (parent block guard) disables
                            // this together with the docGrid path. Allow_fix gate
                            // (row.cells.len() == 1 || !table.style.border) confines
                            // the change to the same cell topologies as S269.
                            let last_cont_h: f32 = elements
                                .iter()
                                .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                                .filter(|e| (e.y - last_cont_top).abs() < 0.5)
                                .map(|e| e.height)
                                .fold(0.0_f32, f32::max);
                            if last_cont_h > 0.0 {
                                cursor.set(
                                    last_cont_top
                                        + last_cont_h * (1.0 + trailing_empty_count as f32)
                                        + s817_tail,
                                );
                            } else {
                                let overflow_on_next = row_bottom - split_y;
                                let pages_used =
                                    ((overflow_on_next) / content_height).floor() as usize;
                                cursor.set(
                                    page_top + overflow_on_next
                                        - (pages_used as f32 * content_height),
                                );
                            }
                        }
                    } else {
                        // No text on final page (border-only): geometric fallback.
                        let overflow_on_next = row_bottom - split_y;
                        let pages_used = ((overflow_on_next) / content_height).floor() as usize;
                        cursor.set(
                            page_top + overflow_on_next - (pages_used as f32 * content_height),
                        );
                    }
                    if std::env::var_os("OXI_CELL_EMPTY_LINES").is_some() {
                        let content_end = elements.iter()
                            .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                            .map(|e| e.y + table_grid_pitch.unwrap_or(e.height) + terminal_spacing(e))
                            .fold(f32::NEG_INFINITY, f32::max);
                        let content_end = if independent_continuation_end.is_finite() {
                            independent_continuation_end
                        } else { content_end };
                        if content_end.is_finite() {
                            cursor.set(content_end);
                        }
                    }
                    // A continuation may end with an inline image below its last
                    // text line. Keep subsequent rows below that image.
                    if std::env::var("OXI_S1250_DISABLE").is_err() {
                        let cont_image_bottom = elements
                            .iter()
                            .filter(|e| matches!(&e.content, LayoutContent::Image { .. }))
                            .map(|e| {
                                if e.cell_float_row_bound {
                                    let pad = e.cell_col_index.and_then(|i| row.cells.get(i))
                                        .and_then(|c| c.margins.as_ref()).and_then(|m| m.bottom)
                                        .unwrap_or(row_default_pad_b);
                                    e.y + e.height + pad + self.rowbox2_trh_bw(table, row)
                                } else {
                                    e.y - e.flow_line_offset + e.flow_line_height.unwrap_or(e.height)
                                }
                            })
                            .fold(f32::NEG_INFINITY, f32::max);
                        if cont_image_bottom.is_finite() && cursor.cursor_y < cont_image_bottom {
                            cursor.set(cont_image_bottom);
                        }
                    }
                } else if ((s754_split && !has_lrpb_mid_row)
                    || (has_lrpb_mid_row && s754_hdr_replay))
                    && !is_single_cell_row
                {
                    // Replayed headers occupy space even when a saved page-break
                    // hint selected the split. Measure every continuation header
                    // and its line reanchoring instead of using the original row
                    // geometry, which contains none of that repeated content.
                    // S754: multi-cell content split — the geometric fallback
                    // (row_bottom − split_y) lands the cursor ABOVE the actual
                    // continuation bottom (probethdr: row 11 at y=93 OVERLAPPED
                    // row 10's continuation ending 112 — the formula knows
                    // neither the replayed header (+hdr_h) nor line granularity).
                    // Use the measured continuation extent instead.
                    let cont_max_y = elements
                        .iter()
                        .map(|e| match &e.content {
                            LayoutContent::TableBorder { y2, .. } => *y2,
                            LayoutContent::PresetShape { .. } => e.y,
                            _ => e.y + e.height,
                        })
                        .fold(f32::NEG_INFINITY, f32::max);
                    let cont_max_y = if independent_continuation_end.is_finite()
                        && row.cells.iter().any(|cell| cell.v_merge.is_some()) {
                        independent_continuation_end
                    } else { cont_max_y };
                    // S942 (2026-07-19, bundle member with OXI_S940; opt-out
                    // OXI_S942_DISABLE): a split atLeast-trHeight row's
                    // CONTINUATION still honors the declared minimum — the
                    // continuation band = max(content, trH + bw). uklocal
                    // rt.pdf wp47: row 11 (trH 63.75) continues with only 2
                    // text lines (~23pt) yet Word's band runs page_top 72.1 →
                    // border 136.6 = trH 63.75 + bw 0.75 (the ROWBOX2
                    // "atLeast = trH + bw" rule applied per FRAGMENT). Oxi's
                    // content-anchored band (27.2) started the next row 37pt
                    // early and cascaded the whole template (+the S941 flip).
                    let s942_floor = if std::env::var("OXI_S942_DISABLE").is_err()
                        && (std::env::var("OXI_S940T_DISABLE").is_err()
                            || std::env::var("OXI_S1025_DISABLE").is_err())
                        && (!self.doc_body_has_real_cjk
                            || std::env::var_os("OXI_CELL_EMPTY_LINES").is_some()
                            || std::env::var_os("OXI_S1430_DISABLE").is_none()) /* S1430 */
                        && row.height_rule.as_deref() != Some("exact")
                    {
                        row.height
                            .map(|trh| page_top + (trh + self.rowbox2_trh_bw(table, row)).min(content_height))
                    } else {
                        None
                    };
                    let cont_max_y = if let Some(fl) = s942_floor {
                        cont_max_y.max(fl)
                    } else {
                        cont_max_y
                    };
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        eprintln!("[SPLIT-CURSOR] branch=s754 cont_max_y={:.2}", cont_max_y);
                        let mut tops: Vec<(f32, f32, String)> = elements
                            .iter()
                            .map(|e| {
                                let b = match &e.content {
                                    LayoutContent::TableBorder { y2, .. } => *y2,
                                    LayoutContent::PresetShape { .. } => e.y,
                                    _ => e.y + e.height,
                                };
                                let k = match &e.content {
                                    LayoutContent::Text { text, .. } => format!("text:{}", text.chars().take(12).collect::<String>()),
                                    LayoutContent::TableBorder { .. } => "border".to_string(),
                                    _ => "other".to_string(),
                                };
                                (b, e.y, k)
                            })
                            .collect();
                        tops.sort_by(|a, b| b.0.partial_cmp(&a.0).unwrap());
                        for t in tops.iter().take(4) {
                            eprintln!("[SPLIT-CURSOR-EL] bottom={:.2} y={:.2} {}", t.0, t.1, t.2);
                        }
                    }
                    if s864_empty_tail_split {
                        // Only the non-painting empty tail crossed the page.
                        // Word collapses that continuation to the page top, so
                        // the following block may begin at the same Y. A split
                        // from a titlePg first page must use the DEFAULT header's
                        // body top, not the first-page body top passed to the
                        // table at entry.
                        let next_page_top = page
                            .margin
                            .top
                            .max(self.s755_header_bottom(&page.header, page));
                        // S1477 (2026-09-18, default ON, opt-out
                        // OXI_S1477_DISABLE): when a page_override is in force
                        // for the page we just moved onto, the content top is
                        // NOT the section's own body top — a floating table's
                        // exclusion band has pushed it down, and
                        // `advance_table_page_geometry` has already put that
                        // value in `page_top`. policies__00602e8a page 4: the
                        // float reserves 121.55..365.79, the split shifts the
                        // row-2 tail to 365.8 and sets cont_max_y 435.29, and
                        // then THIS branch threw it away and sent row 3 back to
                        // the section top 110.92 -- its box stayed at 435.3
                        // while its text drew at 111.4. The titlePg case this
                        // recomputation exists for has no page_override, so it
                        // is untouched.
                        let next_page_top = if std::env::var("OXI_S1477_DISABLE").is_err()
                            && page_geometry
                                .and_then(|g| g.page_override)
                                .is_some_and(|(p, _, _)| p == pages.len() + 1)
                        {
                            page_top
                        } else {
                            next_page_top
                        };
                        cursor.set(next_page_top);
                    } else if cont_max_y.is_finite() {
                        cursor.set(cont_max_y);
                    } else {
                        let overflow_on_next = row_bottom - split_y;
                        let pages_used = ((overflow_on_next) / content_height).floor() as usize;
                        cursor.set(
                            page_top + overflow_on_next - (pages_used as f32 * content_height),
                        );
                    }
                } else {
                    let overflow_on_next = row_bottom - split_y;
                    let pages_used = ((overflow_on_next) / content_height).floor() as usize;
                    cursor.set(page_top + overflow_on_next - (pages_used as f32 * content_height));
                    // S1526 (2026-09-24, opt-out OXI_S1526_DISABLE): the geometric
                    // remainder (row_bottom − split_y) knows no line granularity.
                    // A cell whose line straddles the split carries that WHOLE
                    // line to the next page, so the continuation is taller than
                    // the remainder. educational__0056be35 row 11 (3 cells, an
                    // LRPB mid-row so the S754 branch above is skipped): the
                    // remainder is 19.9 while cell 2 continues with two lines
                    // (Word PDF: 74.9 / 89.5, rule at 101.8, row 12 at 104.7);
                    // Oxi set row 12 at 92.4 over the second line, −11pt for the
                    // rest of the table and the three −1 paragraphs downstream.
                    // Floor the cursor at the measured continuation text bottom
                    // when the continuation stays on this one page.
                    if std::env::var_os("OXI_S1526_DISABLE").is_none() && pages_used == 0 {
                        let cont_text_bottom = elements
                            .iter()
                            .filter(|e| matches!(&e.content, LayoutContent::Text { text, .. } if !text.trim().is_empty()))
                            .map(|e| e.y + e.height)
                            .fold(f32::NEG_INFINITY, f32::max);
                        // A continuation owns its resolved bottom padding and spacing.
                        // Its successor must start after that tail, not inside it.
                        let cont_row_bottom = cont_text_bottom + if std::env::var_os("OXI_CELL_TAIL_FLOOR_DISABLE").is_none() { s817_tail } else { 0.0 };
                        if cont_text_bottom.is_finite()
                            && cont_row_bottom <= page_top + content_height + 0.5
                            && cursor.cursor_y < cont_row_bottom
                        {
                            if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                eprintln!("[SPLIT-CURSOR] s1526 floor {:.2} -> {:.2}", cursor.cursor_y, cont_row_bottom);
                            }
                            cursor.set(cont_row_bottom);
                        }
                    }
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        eprintln!("[SPLIT-CURSOR] branch=geom cursor={:.2} (row_bottom={:.2} split_y={:.2})",
                            cursor.cursor_y, row_bottom, split_y);
                    }
                }
                // S1168: the line landed at the row top, so nothing stayed
                // behind -- this is a whole-row PUSH wearing a split's clothes,
                // not a continuation. Every branch above prices a continuation
                // (S817's tail, the measured cont_max_y, the pages_used
                // geometry), and on educational__00161422 that left the cursor
                // 24.7pt low: the next row slid down, its last line no longer
                // fit, and it landed alone on a page of its own -- an empty p6
                // and +1 for the rest of the document. A row that moved whole
                // advances exactly like one laid out at the top of a fresh page.
                if s1168 && split_y <= row_top + 0.1 {
                    // A replayed header already shifted the pushed row's content.
                    // Advance past the same header when positioning its successor.
                    let header_height = if s754_hdr_replay
                        && std::env::var_os("OXI_CELL_HEADER_PUSH_DISABLE").is_none() {
                        s728_hdr_h + self.repeated_header_border_delta(table, row_idx)
                    } else { 0.0 };
                    cursor.set(page_top + header_height + row_height);
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        eprintln!(
                            "[SPLIT-CURSOR] branch=s1168-wholepush cursor={:.2}",
                            cursor.cursor_y
                        );
                    }
                }
                // Preserve the terminal spacing in branches that use geometric
                // extents instead of the last text line. Whole-row pushes and
                // collapsed empty tails retain their separate cursor rules.
                if std::env::var_os("OXI_CELL_EMPTY_LINES").is_some()
                    && !s864_empty_tail_split
                    && !(s1168 && split_y <= row_top + 0.1)
                {
                    let content_end = elements.iter()
                        .filter(|e| e.cell_row_index == Some(row_idx))
                        .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                        .map(|e| e.y + e.height + terminal_spacing(e))
                        .fold(f32::NEG_INFINITY, f32::max);
                    let content_end = if independent_continuation_end.is_finite() {
                        independent_continuation_end
                    } else { content_end };
                    if content_end.is_finite() && cursor.cursor_y < content_end {
                        cursor.set(content_end);
                    }
                }
                // S942 post-clamp (all split cursor branches): the continuation
                // band of a split atLeast-trHeight row honors the declared
                // minimum — cursor ≥ page_top + trH + bw (uklocal row 11:
                // Word band 72.1→136.6 = trH 63.75 + bw with only 2 content
                // lines; the mid-LRPB/geometric branches bypassed the s754
                // in-branch floor). The S864B empty-tail split is EXCLUDED:
                // Word collapses a continuation that carries only the
                // non-painting empty tail to the page top (administrative__
                // 0001ce58 — clamping it to trH broke its PASS).
                if std::env::var("OXI_S942_DISABLE").is_err()
                    && (std::env::var("OXI_S940T_DISABLE").is_err()
                        || std::env::var("OXI_S1025_DISABLE").is_err())
                    && (!self.doc_body_has_real_cjk
                        || std::env::var_os("OXI_CELL_EMPTY_LINES").is_some()
                        || std::env::var_os("OXI_S1430_DISABLE").is_none()) /* S1430 */
                    && !s864_empty_tail_split
                    && row.height_rule.as_deref() != Some("exact")
                {
                    if let Some(trh) = row.height {
                        // A continued row's minimum starts below its repeated
                        // header, just like the row's text and padding do.
                        let continuation_header_h = if s754_hdr_replay {
                            s728_hdr_h + self.repeated_header_border_delta(table, row_idx)
                        } else { 0.0 };
                        let continuation_top = page_top + continuation_header_h;
                        let fl = continuation_top + (trh + self.rowbox2_trh_bw(table, row))
                            .min((content_height - continuation_header_h).max(0.0));
                        if cursor.cursor_y < fl {
                            if std::env::var("OXI_DBG_SPLIT").is_ok() {
                                eprintln!(
                                    "[SPLIT-CURSOR] s942 clamp {:.2} -> {:.2}",
                                    cursor.cursor_y, fl
                                );
                            }
                            cursor.set(fl);
                        }
                    }
                }
            } else {
                // S200 (2026-05-22): visual/cursor decoupling for Word's per-row
                // +0.5pt table row pitch overhead with sparse-content narrow gate.
                // COM matrix M01-M14: when docGrid is present AND row content fits
                // within one linePitch (sparse cells), Word's row pitch = linePitch + 0.5pt.
                // When row content is multi-line / fills the grid (content-driven),
                // Oxi's row_height already > linePitch and Word's existing logic
                // gives correct positions; +0.5pt would over-correct.
                // Discriminator: |row_height - linePitch| < 0.5pt (sparse cell).
                // Using advance_split: cursor_y advances by row_height (preserves
                // Phase 1 pagination 53/55), visual_y advances by row_height + 0.5pt
                // (corrects element positions).
                // S236 (2026-05-23) removed OXI_LEGACY_NO_TBL_ROW_PLUS_HALF
                // legacy env-var fallback during hardening pass; the gate
                // has been stable since S200 (~35 sessions).
                // S477 (2026-06-02) PROBED + REVERTED: hypothesized the S200 gate
                // misses d4d126's drift because trHeight inflates row_height above
                // linePitch on sparse rows. FALSIFIED by instrumentation — d4d126's
                // DOMINANT 14 rows are CONTENT-driven (row_h≈visual_row_h≈21.26,
                // insideH=true), NOT trHeight-sparse (only 6 trHeight=20.25 rows, and
                // those have insideH=FALSE). So the drift is the content-row CJK
                // line-height (21.26pt, the S467/S468 over-snap / killed-VSNAP regime),
                // NOT the trHeight insideH-border. The trHeight+border rule IS real
                // (repro rowh_border) but does not apply to d4d126's structure. d4d126
                // = confirmed killed-VSNAP/multi-cell dead-end. Gate kept as-is.
                let apply_plus_half = table_grid_pitch
                    .map(|p| (row_height - p).abs() < 0.5)
                    .unwrap_or(false);
                // S661 (2026-06-24, default ON, opt-out OXI_S661_DISABLE): the S200 +0.5
                // row-pitch overhead ALSO applies to trHeight-BOUND SPARSE rows that the
                // `≈linePitch` gate misses. perrow_drift.py (Word PDF horizontal lines vs Oxi
                // rendered borders, 4 tokumei docs): binding-atLeast rows (trHeight binds,
                // content < row) render a CONSISTENT ~+0.5pt taller in Word — the SAME +0.5
                // sparse-row overhead, just at trHeight (~22pt) not linePitch (16.8). The
                // per-row +0.5 accumulates (31420 cumulative border drift +3.3 at the page
                // bottom); +0.5 on these rows cuts it to +1.3 and aligns the whole family.
                // Fire when docGrid present AND content is SPARSE (visual_row_h, the actual
                // cell content, notably < row_height = trHeight binds) AND the row is small
                // (< 3 cells). NOTE: max_actual_cell_h is init to row_height (~11124) so it
                // is always ≥ row_height — use visual_row_h (init 0). Render-only
                // (advance_split → cursor_y/pagination unchanged) → Phase-1 preserved by
                // construction. Full corpus SSIM A/B (on top of S660): net +0.7720, 16
                // improve (ed025 +0.18, a1d6 +0.14, de6e32/d4d126 +0.09, 6514 +0.08, the
                // whole tokumei/order/index form family), 6 over-fire regress (a47e/2ea81a/
                // bd90b00 ~−0.024 — sparse rows that don't drift in Word; the residual gate-
                // refinement). lib 142/0/6.
                // CELLPAIR: the raw heights make rows content-correct; the S661
                // sparse +0.5 was calibrated for the FLOORED heights and double-counts
                // under raw (b35123 r3/r7/r10 inset 1.0, +0.5..+0.7/row; with S661
                // excluded all insets = 0.5 and rows land on Word within +-0.3).
                // Within the CELLPAIR scope, raise the sparse margin 0.4 -> 1.5: the
                // raw-height estimate skews row_height ~+0.5 over the visual content
                // (b35123 content-bound rows falsely read "sparse" and double-counted
                // +0.5), while GENUINE trHeight-bound sparse rows (191cb labels,
                // trHeight - content ~6pt) still qualify and keep Word's +0.5.
                // (margin experiment superseded): within the CELLPAIR scope, S661
                // requires an EXPLICIT trHeight (row.height) -- its own derivation
                // ("sparse-trHeight rows") -- so b35123's content-bound checkbox rows
                // (trPr=None; the raw-estimate skew ~+0.5 falsely read "sparse") are
                // excluded while 191cb's genuine trHeight-bound labels (trHeight=724)
                // keep Word's +0.5.
                // S760 (2026-07-06, opt-out OXI_S760_DISABLE): hRule=EXACT rows
                // get NO +0.5 — the rowbox sweep (_rowbox_sweep.py, Word PDF
                // border truth) shows exact rows render at trHeight EXACTLY
                // (45.00; the border eats into the box) while atLeast rows are
                // trH + bw (45.48). Oxi's cursor pitch was already exact-correct
                // (45.0, TBL_DUMP) — the S661 sparse +0.5 was firing on the
                // VISUAL track for exact rows too (painted pitch 45.5).
                let s760_exact = row.height_rule.as_deref() == Some("exact")
                    && std::env::var("OXI_S760_DISABLE").is_err();
                let s661_sparse_trheight = std::env::var("OXI_S661_DISABLE").is_err()
                    && !s760_exact
                    && (!self.cellpair_active() || row.height.is_some())
                    && table_grid_pitch
                        .map(|p| {
                            visual_row_h > 0.1
                                && visual_row_h + 0.4 < row_height
                                && row_height < p * 3.0
                        })
                        .unwrap_or(false);
                // S666 (2026-06-25, default ON, opt-out OXI_S666_DISABLE): render-only
                // +0.5 border-box overhead for content rows in cell-tcBorder docGrid
                // tables (Word adds the inside border to the row height; Oxi's
                // has_inside_h path misses cell-level borders). advance_split → visual_y
                // only → pagination/element.y unchanged (pagination-safe by construction).
                // Gated to docGrid tables without table-level insideH that carry cell
                // horizontal borders. OR'd with S200/S661 so a row gets +0.5 ONCE.
                // Gate to CONTENT-BOUND rows (visual_row_h ≈ row_height): the border-box
                // +0.5 applies where the cell CONTENT (+border) determines the height,
                // complementary to S661 (which handles trHeight-bound rows, vrh < rh).
                // 15076 (08_09)'s large cell-bordered rows are trHeight-bound (vrh < rh)
                // or overflow (vrh > rh) and do NOT drift +0.5 in Word — firing on them
                // over-corrects (−0.047). The 08_01 family's drifting rows (small AND large,
                // 22-125pt) are all content-bound (vrh ≈ rh), so this keeps them.
                let s666_cellborder = std::env::var("OXI_S666_DISABLE").is_err()
                    && row_cell_hborder
                    && table_grid_pitch.is_some()
                    && (row_height - visual_row_h).abs() < 0.5;
                if std::env::var("OXI_DBG_S661").is_ok() && s661_sparse_trheight {
                    eprintln!("[S661] rule={:?} rh={:.2} vrh={:.2} gap={:.2} pitch={:?} hborder={} cy={:.1}",
                        row.height_rule.as_deref().unwrap_or("none"), row_height, visual_row_h,
                        row_height - visual_row_h, table_grid_pitch, row_cell_hborder, cursor.cursor_y);
                }
                // S682 (2026-06-28, default ON, opt-out OXI_BBOX_DISABLE): LARGE trHeight-
                // bound rows (explicit trHeight >= pitch*3) in a table with inside-H borders
                // miss the border-box overhead. Word renders each such row at trHeight +
                // top_bw + bottom_bw (border-box); Oxi renders EXACTLY trHeight (the unused
                // `_border_overhead` above). S661's pitch*3 gate EXCLUDES these large rows
                // and only adds +0.5 (half the border-box) to the small ones. 7ead52b: 56.7pt
                // trHeight rows render at 56.7 in Oxi but ~58.0 in Word (+1.3, accumulating
                // −9pt over 9 rows). Render-only (advance_split → visual_y) ⇒ pagination
                // byte-identical. ★Gate on the EXPLICIT trHeight (row.height) binding a large
                // row, NOT the computed row_height (which includes multi-line content) — this
                // EXCLUDES content-tall rows (b35123 max trHeight 30.75pt, whose content makes
                // row_height >= pitch*3; firing on them regressed b35123 −0.0258, a bottom-N
                // floor doc). Amount = 2*border_width (the top+bottom border-box). Direct
                // per-page SSIM over the 20 word_png docs with large trHeight + inside-H (the
                // complete surface): net +0.0907, ed025 +0.0934 (16pg), a47e +0.0190, 191cb5
                // +0.0122, 31420 +0.0104, 7ead52b +0.0082, f16f228 +0.0078; regress e8caed
                // −0.0574 (its large rows don't drift in Word — no docx discriminator vs the
                // same-structure f16f228 that improves; the documented per-doc over-fire wall),
                // 29dc6e/4a36b small. Bottom-N floor (b35123/1ec1/15076) byte-identical →
                // Phase-3 gate (bottom-N protected, mean up) passes. Canaries (3a4f/459f/0e7af/
                // 2ea81a/tokumei_08_01) byte-identical. OXI_BBOX overrides the amount (tuning).
                // NOTE: extending to cell-level tcBorders (OXI_BBOX_CELL, tested 2026-06-28)
                // helps the tokumei_08_01 family (+0.0266) but REGRESSES 15076_tokumei_08_09
                // −0.0429 (a BOTTOM-N FLOOR doc whose large cell-border rows render ~at
                // trHeight in Word — the per-doc wall, no signal vs de6e32). Net −0.0163,
                // lowers the floor → NOT shippable. Kept table-level insideH only.
                let bbox_large = std::env::var("OXI_BBOX_DISABLE").is_err()
                    && table.style.has_inside_h
                    && table_grid_pitch
                        .map(|p| {
                            row.height.map(|h| h >= p * 3.0).unwrap_or(false)
                                && visual_row_h > 0.1
                                && visual_row_h + 0.4 < row_height
                        })
                        .unwrap_or(false);
                // S1192b: this row is now final — draw it down from every span
                // it passes through, and forget spans that ended here.
                if std::env::var("OXI_S1192_DISABLE").is_err() && !s1192_pending.is_empty() {
                    let is_cont = |c: &TableCell| {
                        matches!(c.v_merge.as_deref(), Some("continue") | Some(""))
                    };
                    let bygrid = std::env::var_os("OXI_S1192G_DISABLE").is_none();
                    for (ci, remaining) in s1192_pending.iter_mut() {
                        if LayoutEngine::s1192_cell_at(row, *ci, bygrid).is_some() {
                            *remaining -= row_height;
                        }
                    }
                    s1192_pending.retain(|(ci, remaining)| {
                        let next_cont = table
                            .rows
                            .get(row_idx + 1)
                            .and_then(|r| LayoutEngine::s1192_cell_at(r, *ci, bygrid))
                            .map_or(false, is_cont);
                        if std::env::var("OXI_DBG_S1192").is_ok() {
                            eprintln!("[S1192] DRAW row={} key={} rh={:.2} rem_after={:.2} keep={}",
                                row_idx, ci, row_height, remaining, next_cont && *remaining > 0.01);
                        }
                        next_cont && *remaining > 0.01
                    });
                }
                if std::env::var("OXI_DBG_ROWADV").is_ok() {
                    eprintln!("[ROWADV] r={} rh={:.3} vrh={:.3} trH={:?} cy={:.2}",
                        row_idx, row_height, visual_row_h, row.height, cursor.cursor_y);
                }
                if self.rowbox2_active() {
                    // ROWBOX2: the border-box is already IN row_height
                    // (cursor-real); the render-only visual pluses
                    // (S200/S661/S666/S682) would double-count.
                    cursor.advance(row_height);
                } else if apply_plus_half || s661_sparse_trheight || s666_cellborder {
                    cursor.advance_split(row_height, row_height + 0.5);
                } else if bbox_large {
                    if std::env::var("OXI_DBG_BBOX").is_ok() {
                        eprintln!(
                            "[BBOX] trH={:.1} rh={:.2} vrh={:.2} content={:.2} bw={:.2} cy={:.1}",
                            row.height.unwrap_or(0.0),
                            row_height,
                            visual_row_h,
                            visual_row_h,
                            table.style.border_width.unwrap_or(0.5),
                            cursor.cursor_y
                        );
                    }
                    // S682b (2026-06-28): amount = 1× border_width, NOT 2×. The per-row
                    // PITCH overhead is ONE inside-H border (shared between adjacent rows:
                    // N rows have N+1 horizontal borders → per-row ≈ 1 bw), so 2*bw
                    // DOUBLE-counted. The SSIM audit (per-DOC-mean, not per-page sum) showed
                    // 2*bw over-corrected the small 1-page docs (191cb5/f16f228/e8caed/4a36b)
                    // → per-doc-mean NEUTRAL (−0.0007); 1*bw recovers them → per-doc-mean
                    // +0.0464. (per-page sum is misleading — it over-weights the 16-page ed025.)
                    let amt = std::env::var("OXI_BBOX")
                        .ok()
                        .and_then(|v| v.parse::<f32>().ok())
                        .filter(|v| *v > 0.01)
                        .unwrap_or_else(|| table.style.border_width.unwrap_or(0.5));
                    cursor.advance_split(row_height, row_height + amt);
                } else {
                    cursor.advance(row_height);
                }
            }

            // S728: capture the LEADING w:tblHeader row(s)' emitted elements +
            // advanced height for continuation-page replay (see the pre-loop
            // decl). Word repeats only the table's FIRST consecutive header
            // rows; a non-header row (or a header-flagged later row) ends the
            // capture window. Guarded against a mid-row page push having
            // flushed `elements` (slice would be stale — skip capture then;
            // such a straddling header row is not replayed).
            if s728_on && !s728_capture_done {
                if std::env::var("OXI_DBG728").is_ok() {
                    eprintln!(
                        "[S728] row_idx={} header={} seen={} elems_range={}..{} rh={:.1}",
                        row_idx,
                        row.header,
                        s728_hdr_rows_seen,
                        elements_before_row,
                        elements.len(),
                        row_height
                    );
                }
                if row.header && row_idx == s728_hdr_rows_seen {
                    if let Some(slice) = elements.get(elements_before_row..) {
                        // S1587 (2026-09-27, default ON, opt-out OXI_S1587_DISABLE): a
                        // header row captured on a LATER page than the rows before it
                        // is stacked under them, so the replay keeps the rows in
                        // order. reference__0096d8c9 p156: header rows 0-1 captured
                        // on p155, row 2 pushed to p156 -- the replay aligned on
                        // row 2's small y and drew rows 0-1 185pt down, under the
                        // body (Word repeats all three at the top) and 57GD went to
                        // p157.
                        let s1587_page_moved = std::env::var_os("OXI_S1587_DISABLE").is_none()
                            && !s728_hdr_elems.is_empty()
                            && s728_capture_page.map_or(false, |p0| p0 != pages.len());
                        if s1587_page_moved {
                            let prev_bottom = s728_hdr_elems.iter().map(|e| e.y + e.height).fold(f32::NEG_INFINITY, f32::max);
                            let slice_top = slice.iter().map(|e| e.y).fold(f32::INFINITY, f32::min);
                            if prev_bottom.is_finite() && slice_top.is_finite() {
                                let dy = prev_bottom - slice_top;
                                for el in slice.iter() {
                                    let mut c = el.clone();
                                    c.y += dy;
                                    if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = c.content {
                                        *y1 += dy;
                                        *y2 += dy;
                                    }
                                    s728_hdr_elems.push(c);
                                }
                            }
                        } else {
                            s728_hdr_elems.extend(slice.iter().cloned());
                        }
                        if s728_capture_page.is_none() {
                            s728_capture_page = Some(pages.len());
                        }
                        s728_hdr_h += row_height;
                        s728_hdr_rows_seen += 1;
                    } else {
                        s728_capture_done = true;
                    }
                    if !table.rows.get(row_idx + 1).map_or(false, |r| r.header) {
                        s728_capture_done = true;
                    }
                } else {
                    s728_capture_done = true;
                }
            }
        }

        // S1191: the table's bottom rule is drawn BELOW the last row's box, and
        // Word starts the following block under it (see s1191_foot_bw).
        if (separate_outer_edges || (self.s1191_on() && self.s1191_table_needs_foot(table))) {
            let foot = self.s1191_foot_bw(table);
            if std::env::var_os("OXI_DBG_TBLFOOT").is_some() {
                eprintln!("[TBLFOOT] cy={:.2} vy={:.2} foot={:.2} sep_outer={}", cursor.cursor_y, cursor.visual_y, foot, separate_outer_edges);
            }
            if foot > 0.0 {
                cursor.advance(foot);
            }
        }
        if s1618_foot > 0.0 {
            cursor.advance(s1618_foot);
        }

        // S740: final flush — trailing page transitions + the LAST row's notes,
        // then hand the per-page ids to the caller.
        if let Some(rf) = row_footnotes {
            if pages.len() != s740_pages_len {
                while s740_fn_pages.len() < pages.len() - s740_entry_pages + 1 {
                    s740_fn_pages.push(Vec::new());
                }
                s740_reserve = 0.0;
                s740_page_has_notes = false;
            }
            if let Some(prev) = s740_pending_commit.take() {
                let (ids_all, _h_all, hs_all) = &rf[prev];
                // S1527: leave out the ids the split already placed.
                let ids: Vec<u32> = ids_all.iter().copied().filter(|id| !s1527_early.contains(id)).collect();
                let h: f32 = ids_all.iter().zip(hs_all.iter()).filter(|(id, _)| ids.contains(id)).map(|(_, h)| *h).sum();
                if !ids.is_empty() {
                    if !s740_page_has_notes {
                        s740_reserve += fn_sep;
                        s740_page_has_notes = true;
                    }
                    s740_reserve += h;
                    if let Some(last) = s740_fn_pages.last_mut() {
                        for id in &ids {
                            if !last.contains(id) {
                                last.push(*id);
                            }
                        }
                    }
                }
            }
            let _ = s740_reserve;
            if let Some(out) = fn_pages_out.as_deref_mut() {
                *out = s740_fn_pages;
            }
        }

        // S487: paint cell-anchored floating text boxes last (on top of the grid).
        elements.extend(deferred_cell_textboxes);
        elements
    }
}
