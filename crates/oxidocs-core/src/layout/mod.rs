// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

mod kinsoku;
pub mod math;
mod ruby;
/// Session 117: per-char compression algorithm for jc=both / distribute paragraphs.
/// Not yet integrated into the layout pipeline — see docs/design/jc_both_per_char_compression.md.
#[allow(dead_code)]
pub mod jc_both_compress;

use crate::font::{FontMetrics, FontMetricsRegistry};
use crate::ir::*;

/// Pre-allocated single-character strings to avoid heap allocation in hot loops.
const TAB_STRING: &str = "\t";
const SPACE_STRING: &str = " ";

/// Convert a char to a String with pre-sized buffer (avoids realloc for multi-byte chars).
#[inline]
fn char_to_string(ch: char) -> String {
    let mut s = String::with_capacity(ch.len_utf8());
    s.push(ch);
    s
}

/// Characters that allow a line break AFTER them (English punctuation).
/// Word treats these as breakable opportunities similar to spaces.
fn is_break_after(ch: char) -> bool {
    matches!(ch, '-' | '/' | '\\' | ')' | ']' | '}' | '>' | '!' | '?' | ';' | ':' | ',')
}

/// Word's default 8-color rotation for tracked-change author tints. The
/// author's `color_index` in `Document.authors` selects a slot here.
///
/// COM-confirmed against Word 16.0 (see
/// `docs/spec/comments_tracked_changes/com_measurements/PIXEL_PASS_README.md`):
///   - Slot 0 → #D03337 (Alice in fixture_05/06/07/10, 2026-04-25)
///   - Slot 1 → #5B2C90 (Bob in fixture_10, 2026-04-25)
///   - Slot 2 → #478103 (Carol in fixture_12, R65 2026-04-29 — distinct
///     from the move-revision green #2B6033; the earlier conflation has
///     been corrected, see PHASE_2_CLOSEOUT.md item 9).
///
/// Slots 3..7 are Word's documented rotation defaults but not yet COM-confirmed
/// — they need a 4+author fixture. The list below is the Word/Office reviewing
/// palette as published by Microsoft; if a future measurement contradicts a
/// specific slot, replace just that entry.
const REVISION_AUTHOR_PALETTE: [&str; 8] = [
    "#D03337", // 0 — confirmed (Alice)
    "#5B2C90", // 1 — confirmed (Bob)
    "#478103", // 2 — confirmed (Carol, R65)
    "#ED7D31", // 3 — orange
    "#4472C4", // 4 — blue
    "#843C0C", // 5 — brown
    "#C00000", // 6 — dark red
    "#00B050", // 7 — teal
];

/// Word renders `<w:moveFrom>` / `<w:moveTo>` in a fixed green regardless of
/// author (COM-confirmed 2026-04-25 in fixture_08). The author-color rotation
/// does NOT apply to moves.
const REVISION_MOVE_COLOR: &str = "#2B6033";

/// Per-author comment-range highlight tint. These are lightened (≈12% author
/// color + 88% white) versions of the revision palette — Word uses them as
/// the in-line background tint for text inside `commentRangeStart/End`, as
/// well as for the unresolved balloon background.
///
/// Slot 0 (#FAE6E7) is the COM-confirmed Alice tint from the pixel pass
/// (fixture_01 balloon background, 2026-04-25). Slots 1-7 are computed via
/// the 12/88 white-blend formula off the same palette; pixel-pass
/// confirmation for the other slots awaits a fixture with 3+ authors.
const COMMENT_HIGHLIGHT_TINT_PALETTE: [&str; 8] = [
    "#FAE6E7", // 0 — Alice, COM-confirmed
    "#EFEAF4", // 1 — Bob (derived from #5B2C90)
    "#E9F0E1", // 2 — derived from #478103 (R65 base correction, re-derived
               //     in b2cedc6 via the 12/88 white-blend formula)
    "#FCEEE0", // 3 — derived from #ED7D31
    "#E8ECF6", // 4 — derived from #4472C4
    "#F2EAE4", // 5 — derived from #843C0C
    "#F6E0E0", // 6 — derived from #C00000
    "#E1F3E9", // 7 — derived from #00B050
];

/// Public resolver: given an author's palette index and whether the comment
/// is resolved, return the hex color string a renderer should use as the
/// balloon background fill (R-05g) or the in-line range highlight tint (R-04 /
/// R-09 in-line). Slots are clamped via modulo so any caller-supplied index
/// works.
pub fn comment_balloon_fill(author_color_index: usize, resolved: bool) -> &'static str {
    let palette = if resolved {
        &COMMENT_HIGHLIGHT_RESOLVED_PALETTE
    } else {
        &COMMENT_HIGHLIGHT_TINT_PALETTE
    };
    palette[author_color_index % palette.len()]
}

/// Resolved variant of the comment tint palette (R-09). When a comment has
/// `Comment.resolved == true` (`<w15:done="1"/>` in `commentsExtended.xml`),
/// Word desaturates the in-line range tint AND the balloon background by
/// blending the unresolved tint with grey at ~75/25 — chroma drops to ~5
/// while lightness stays the same.
///
/// Slot 0 (#F1EDEC) is the COM-confirmed Alice resolved tint from the pixel
/// pass (fixture_03 balloon background, 2026-04-25). Slots 1-7 are derived
/// by applying the same 25% tint + 75% grey blend per slot; awaits 3+ author
/// confirmation.
const COMMENT_HIGHLIGHT_RESOLVED_PALETTE: [&str; 8] = [
    "#F1EDEC", // 0 — Alice resolved, COM-confirmed
    "#EFEEF1", // 1 — Bob resolved
    "#EBEDE9", // 2 — derived from #E9F0E1 (R65 base correction)
    "#F2EDE6", // 3 — derived orange
    "#EBEDF1", // 4 — derived blue
    "#EFECEA", // 5 — derived brown
    "#F1E9E9", // 6 — derived dark red
    "#E8EFEA", // 7 — derived teal
];

/// Pre-pass: apply Word's default tracked-change visual styling to runs.
///
/// For runs whose `tracked_change` is set, mutate `run.style` to add the
/// underline / strikethrough / color the renderer would normally need to
/// special-case at the layout site. This keeps the rest of the layout pipeline
/// style-only and matches Word's "All markup" view (the default).
///
/// Notes:
/// - `Run::style` is a small struct (clone is cheap) and only revision-bearing
///   runs are touched.
/// - The styling is non-destructive in the sense that a run's *original*
///   `tracked_change` field is preserved on the IR (`doc_resolved` is a clone)
///   so future passes / tools can still inspect the revision metadata.
/// - Recurses into table cells. Walks body, headers, footers, footnotes,
///   endnotes, and textbox content via `for_each_block_tree`.
fn apply_revision_styling(doc: &mut Document) {
    use std::collections::HashMap;

    let author_to_idx: HashMap<String, usize> = doc
        .authors
        .iter()
        .map(|a| (a.display.clone(), a.color_index))
        .collect();

    for_each_block_tree(doc, |blocks| {
        for block in blocks.iter_mut() {
            apply_revision_styling_to_block(block, &author_to_idx);
        }
    });
}

/// Iterate every top-level `Vec<Block>` in the document — body
/// (`page.blocks`), headers, footers, footnote bodies, endnote bodies,
/// and textbox contents. Each pre-pass that wants to operate on the full
/// document can call this once and process the yielded slice.
///
/// Skips `Page.shapes` / `Page.floating_images` because those don't contain
/// runs (just geometry).
fn for_each_block_tree<F: FnMut(&mut Vec<Block>)>(doc: &mut Document, mut f: F) {
    for page in &mut doc.pages {
        f(&mut page.blocks);
        f(&mut page.header);
        f(&mut page.footer);
        for footnote in &mut page.footnotes {
            f(&mut footnote.blocks);
        }
        for endnote in &mut page.endnotes {
            f(&mut endnote.blocks);
        }
        for tb in &mut page.text_boxes {
            f(&mut tb.blocks);
        }
    }
}

fn apply_revision_styling_to_block(
    block: &mut Block,
    author_to_idx: &std::collections::HashMap<String, usize>,
) {
    match block {
        Block::Paragraph(p) => {
            for run in &mut p.runs {
                if let Some(tc) = run.tracked_change.clone() {
                    apply_revision_styling_to_run(run, &tc, author_to_idx);
                }
            }
        }
        Block::Table(t) => {
            for row in &mut t.rows {
                for cell in &mut row.cells {
                    for b in &mut cell.blocks {
                        apply_revision_styling_to_block(b, author_to_idx);
                    }
                }
            }
        }
        Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
    }
}

/// R-04 pre-pass: apply in-line comment-range highlight tint.
///
/// Walks runs in document order, maintains a set of currently-open comment
/// ids (pushed on `comment_range_start`, popped on `comment_range_end`), and
/// stamps `style.highlight` on every run *between* the start and end markers.
///
/// The marker attachment convention after the parser fix (2026-04-25) is:
/// - `comment_range_start` on run R means "range starts AFTER R"; R itself is
///   outside the range.
/// - `comment_range_end` on run R means "range ends AFTER R"; R IS the last
///   run inside the range.
/// So the walk applies highlight *before* processing markers — R gets the
/// tint only if the open set was already non-empty at visit time.
///
/// Resolved comments (`Comment.resolved = true`) still get highlighted but
/// with a desaturated tint — R-09 will refine this; for now the in-line
/// highlight is identical whether resolved or not. The visual "resolved"
/// signal lives primarily on the balloon (R-05 / R-09), not the in-line
/// range.
fn apply_comment_range_highlighting(doc: &mut Document) {
    use std::collections::{HashMap, HashSet};

    // Build comment_id → author tint hex. If a comment has no author or the
    // author isn't in the palette, fall back to slot 0.
    let author_color_index: HashMap<String, usize> = doc
        .authors
        .iter()
        .map(|a| (a.display.clone(), a.color_index))
        .collect();
    let comment_tint: HashMap<String, String> = doc
        .comments
        .iter()
        .map(|c| {
            let idx = c
                .author
                .as_deref()
                .and_then(|a| author_color_index.get(a).copied())
                .unwrap_or(0);
            // R-09: resolved comments (`<w15:done="1"/>`) use the desaturated
            // palette. Same author slot, lower chroma. The palette is keyed on
            // the same color_index so a reply (which inherits the parent's
            // author) gets the same slot regardless of resolved state.
            let palette = if c.resolved {
                &COMMENT_HIGHLIGHT_RESOLVED_PALETTE
            } else {
                &COMMENT_HIGHLIGHT_TINT_PALETTE
            };
            (c.id.clone(), palette[idx % palette.len()].to_string())
        })
        .collect();

    // Nothing to do if there are no comments at all.
    if comment_tint.is_empty() {
        return;
    }

    let mut open: HashSet<String> = HashSet::new();
    for_each_block_tree(doc, |blocks| {
        for block in blocks.iter_mut() {
            apply_comment_highlight_to_block(block, &comment_tint, &mut open);
        }
    });
}

fn apply_comment_highlight_to_block(
    block: &mut Block,
    comment_tint: &std::collections::HashMap<String, String>,
    open: &mut std::collections::HashSet<String>,
) {
    match block {
        Block::Paragraph(p) => {
            for run in &mut p.runs {
                // 1. Apply highlight if any comment range is currently open.
                if !open.is_empty() && run.style.highlight.is_none() {
                    // Use the most-recently-opened comment's tint. `open` is a
                    // HashSet so order isn't stable — pick any deterministic
                    // member. `min()` gives deterministic output across runs.
                    if let Some(first_open) = open.iter().min() {
                        if let Some(tint) = comment_tint.get(first_open) {
                            run.style.highlight = Some(tint.clone());
                        }
                    }
                }
                // 2. Process comment_range_end — the current run was the LAST
                //    inside; remove so the next run is outside.
                for id in &run.comment_range_end {
                    open.remove(id);
                }
                // 3. Process comment_range_start — range opens AFTER this run,
                //    so next run will be inside.
                for id in &run.comment_range_start {
                    open.insert(id.clone());
                }
            }
        }
        Block::Table(t) => {
            for row in &mut t.rows {
                for cell in &mut row.cells {
                    for b in &mut cell.blocks {
                        apply_comment_highlight_to_block(b, comment_tint, open);
                    }
                }
            }
        }
        Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
    }
}

/// R-05c: emit one `LayoutContent::Balloon` per visible comment on this
/// LayoutPage, anchored to the rendered Y of its `commentRangeStart`.
///
/// Per `r05_balloon_design.md`:
/// - Balloon column right edge is `page_width − 4pt`.
/// - Balloon width is 293.8pt for unresolved, 190.1pt for resolved
///   (COM-confirmed 2026-04-25 from fixture_01 / fixture_03 pixel pass).
/// - Anchor Y is the rendered Y of the FIRST `commentRangeStart` marker for
///   the comment found on this page. (If a comment's scope starts on a prior
///   page, no balloon emits for it on later pages — Word renders the balloon
///   only on the page where the scope begins.)
/// - This iteration emits balloons in document order; per-balloon stacking
///   to prevent overlap is R-05d.
fn emit_balloons_for_layout_page(
    layout_page: &mut LayoutPage,
    doc: &Document,
    ir_page_idx: usize,
) {
    use std::collections::HashMap;

    // Map comment.id → comment for fast lookup.
    let id_to_comment: HashMap<&str, &Comment> = doc
        .comments
        .iter()
        .map(|c| (c.id.as_str(), c))
        .collect();
    if id_to_comment.is_empty() {
        return;
    }
    // Map author display → palette index (slot 0 fallback for unknown authors).
    let author_to_idx: HashMap<&str, usize> = doc
        .authors
        .iter()
        .map(|a| (a.display.as_str(), a.color_index))
        .collect();

    let ir_page = match doc.pages.get(ir_page_idx) {
        Some(p) => p,
        None => return,
    };

    // First pass: build `paragraph_index → first-rendered (x, y)` map by
    // walking LayoutElements. This handles the case where a comment's
    // `commentRangeStart` is attached to an empty-text marker run that
    // emits no Text LayoutElement — we still need an anchor Y, so we use
    // the paragraph's first visible element instead.
    let mut para_first_xy: HashMap<usize, (f32, f32)> = HashMap::new();
    for el in &layout_page.elements {
        if !matches!(&el.content, LayoutContent::Text { .. }) {
            continue;
        }
        if let Some(pi) = el.paragraph_index {
            para_first_xy.entry(pi).or_insert((el.x, el.y));
        }
    }

    // Second pass: scan IR paragraphs (in document order) for any run that
    // carries a `commentRangeStart` id. Anchor Y is taken from the
    // paragraph's first-rendered LayoutElement (built above), so empty
    // marker runs don't lose their anchor.
    let mut anchors: Vec<(String, f32, f32)> = Vec::new();
    let mut seen_ids: std::collections::HashSet<String> = std::collections::HashSet::new();

    for (pi, block) in ir_page.blocks.iter().enumerate() {
        if let Block::Paragraph(p) = block {
            for run in &p.runs {
                for cid in &run.comment_range_start {
                    if seen_ids.insert(cid.clone()) {
                        if let Some(&(x, y)) = para_first_xy.get(&pi) {
                            anchors.push((cid.clone(), x, y));
                        }
                    }
                }
            }
        }
    }

    if anchors.is_empty() {
        return;
    }

    // Compute balloon column geometry once per page.
    let page_w = layout_page.width;
    let balloon_right_inset = 4.0;
    let balloon_width_unresolved = 293.8;
    let balloon_width_resolved = 190.1;

    // Pre-compute every balloon's natural (anchor-aligned) Y + height. We
    // emit *after* applying R-05d stacking, so adjacent balloons can be
    // pushed down to avoid overlap.
    struct PendingBalloon {
        cid: String,
        author: String,
        author_color_index: usize,
        resolved: bool,
        body: String,
        replies: Vec<BalloonReply>,
        anchor_x: f32,
        anchor_y: f32,
        balloon_left: f32,
        balloon_width: f32,
        balloon_height: f32,
        /// Resolved Y after stacking — initialised to anchor_y.
        y: f32,
    }

    let mut pending: Vec<PendingBalloon> = Vec::new();

    for (cid, anchor_x, anchor_y) in &anchors {
        let comment = match id_to_comment.get(cid.as_str()) {
            Some(c) => *c,
            None => continue, // commentRangeStart with no matching <w:comment> body
        };

        let color_idx = comment
            .author
            .as_deref()
            .and_then(|a| author_to_idx.get(a).copied())
            .unwrap_or(0);

        let body = comment_body_text(&comment.blocks);

        // R-05f: fold any reply comments (those whose `parent_para_id`
        // matches this comment's `para_id`) into the parent balloon's
        // `replies` Vec. Replies don't get their own standalone Balloon
        // because they share the parent's range — Word renders them
        // indented inside the same balloon.
        let replies: Vec<BalloonReply> = if let Some(parent_pid) = comment.para_id.as_deref() {
            doc.comments
                .iter()
                .filter(|c| c.parent_para_id.as_deref() == Some(parent_pid))
                .map(|reply| BalloonReply {
                    author: reply.author.clone().unwrap_or_default(),
                    author_color_index: reply
                        .author
                        .as_deref()
                        .and_then(|a| author_to_idx.get(a).copied())
                        .unwrap_or(0),
                    body: comment_body_text(&reply.blocks),
                })
                .collect()
        } else {
            Vec::new()
        };

        // Pick balloon width based on resolved state.
        let balloon_width = if comment.resolved {
            balloon_width_resolved
        } else {
            balloon_width_unresolved
        };
        let balloon_left = (page_w - balloon_right_inset - balloon_width).max(0.0);

        // Estimate balloon height. Word actually wraps at the balloon width;
        // for v1 use a quick line estimate based on character count and a
        // rough average glyph width. Replies indent ~10pt inside but we
        // estimate them as full-width text for the height calc — minor
        // over-estimate that R-05g will refine.
        let avg_glyph_pt = 5.0; // ~5pt per glyph at 11pt Calibri (approximate)
        let max_chars_per_line = ((balloon_width - 8.0) / avg_glyph_pt).max(1.0) as usize;
        let body_est_lines = body
            .lines()
            .map(|line| (line.chars().count().max(1) + max_chars_per_line - 1) / max_chars_per_line)
            .sum::<usize>()
            .max(1);
        let reply_est_lines = replies
            .iter()
            .map(|r| {
                r.body
                    .lines()
                    .map(|line| (line.chars().count().max(1) + max_chars_per_line - 1) / max_chars_per_line)
                    .sum::<usize>()
                    .max(1)
                    + 1 // +1 for the author header chip on each reply
            })
            .sum::<usize>();
        // Each text section (header chip, body, reply chip, reply body)
        // contributes its own height + an inter-section pad in the renderer.
        // The rendering loop in `tools/oxi-gdi-renderer/src/main.rs` adds
        // (header_fs + pad) after header, (body_h + pad) after body, and
        // (header_fs + pad) + (body_h + pad) per reply. Mirror those costs
        // here so the bounding box doesn't truncate.
        let line_height = 14.0; // ~10pt body + small leading
        let chip_h = 14.0; // ~9pt header + leading
        let section_pad = 4.0; // matches the renderer's inter-section pad
        let outer_pad = 8.0;
        let body_h = (body_est_lines as f32) * line_height;
        let replies_h = replies.iter().map(|r| {
            let r_body_lines = r
                .body
                .lines()
                .map(|line| (line.chars().count().max(1) + max_chars_per_line - 1) / max_chars_per_line)
                .sum::<usize>()
                .max(1);
            chip_h + section_pad + (r_body_lines as f32) * line_height + section_pad
        }).sum::<f32>();
        let balloon_height = outer_pad
            + chip_h
            + section_pad
            + body_h
            + section_pad
            + replies_h
            + outer_pad;
        let _ = reply_est_lines; // kept for future per-reply line accounting

        pending.push(PendingBalloon {
            cid: cid.clone(),
            author: comment.author.clone().unwrap_or_default(),
            author_color_index: color_idx,
            resolved: comment.resolved,
            body,
            replies,
            anchor_x: *anchor_x,
            anchor_y: *anchor_y,
            balloon_left,
            balloon_width,
            balloon_height,
            y: *anchor_y,
        });
    }

    // R-05d: stack to prevent overlap. Sort by anchor Y ascending; pure
    // helper handles the per-element Y resolution. See `stack_balloon_ys`
    // below for the algorithm.
    pending.sort_by(|a, b| {
        a.anchor_y
            .partial_cmp(&b.anchor_y)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    {
        let mut ys: Vec<(f32, f32)> = pending
            .iter()
            .map(|pb| (pb.anchor_y, pb.balloon_height))
            .collect();
        stack_balloon_ys(&mut ys, BALLOON_STACK_GAP);
        for (pb, (resolved_y, _)) in pending.iter_mut().zip(ys.iter()) {
            pb.y = *resolved_y;
        }
    }

    for pb in pending {
        // R-05e: connector line from the inline anchor to the balloon's left
        // edge. Drawn before the balloon so the balloon's background (when
        // R-05g lands) overdraws any portion that crosses the balloon's
        // bounding box.
        let connector_color = COMMENT_HIGHLIGHT_TINT_PALETTE
            [pb.author_color_index % COMMENT_HIGHLIGHT_TINT_PALETTE.len()]
        .to_string();
        // Connector sticks slightly into the balloon (~5pt below balloon top)
        // so it visually meets the balloon's vertical centerline of its
        // first text row, matching Word's appearance.
        let to_x = pb.balloon_left;
        let to_y = pb.y + 5.0;
        layout_page.elements.push(LayoutElement::new(
            pb.anchor_x.min(to_x),
            pb.anchor_y.min(to_y),
            (to_x - pb.anchor_x).abs(),
            (to_y - pb.anchor_y).abs(),
            LayoutContent::BalloonConnector {
                from_x: pb.anchor_x,
                from_y: pb.anchor_y,
                to_x,
                to_y,
                color_hex: connector_color,
            },
        ));
        layout_page.elements.push(LayoutElement::new(
            pb.balloon_left,
            pb.y,
            pb.balloon_width,
            pb.balloon_height,
            LayoutContent::Balloon {
                comment_id: pb.cid,
                author: pb.author,
                author_color_index: pb.author_color_index,
                resolved: pb.resolved,
                body: pb.body,
                replies: pb.replies,
                anchor_x: pb.anchor_x,
                anchor_y: pb.anchor_y,
            },
        ));
    }
}

/// R-12: build a "Formatted: …" balloon body describing what changed
/// between an rPrChange's prior run style and the run's current style.
/// Returns `None` when no axis in the supported set differs (the caller
/// then suppresses the balloon — empty bodies would still anchor a
/// balloon, which is not Word's behaviour for no-op rPrChanges).
///
/// Each axis appends a label to `axes`; the final body is
/// `"Formatted: " + axes.join(", ")`. Per Word's "Formatted:" markup
/// vocabulary, label phrasing is human-readable rather than literal
/// OOXML names (e.g., "Bold" not "w:b"). Axis labels are visible to
/// end users in the margin balloon.
///
/// Supported axes (extended each ship — comma-join already in place):
/// S254 — Bold.
/// S255 — Font, Font Size, All Caps, Character Spacing, Superscript
///        (vertical_align), Shading.
/// S256 — Outline, Emboss, Imprint, Shadow, Hidden (vanish),
///        Double Strikethrough.
/// S257 — Highlight, Position, Emphasis Mark.
pub fn describe_rpr_diff(prior: &crate::ir::RunStyle, current: &crate::ir::RunStyle) -> Option<String> {
    let mut axes: Vec<String> = Vec::new();
    if prior.bold != current.bold {
        axes.push("Bold".to_string());
    }
    if prior.font_family != current.font_family {
        if let Some(name) = current.font_family.as_deref() {
            axes.push(format!("Font: {name}"));
        } else {
            axes.push("Font".to_string());
        }
    }
    if prior.font_size != current.font_size {
        if let Some(sz) = current.font_size {
            // Word renders integer font sizes plainly; only show decimals
            // when the value isn't a whole number.
            let label = if sz.fract().abs() < f32::EPSILON {
                format!("Font Size: {}pt", sz as i32)
            } else {
                format!("Font Size: {sz}pt")
            };
            axes.push(label);
        } else {
            axes.push("Font Size".to_string());
        }
    }
    if prior.all_caps != current.all_caps {
        axes.push("All Caps".to_string());
    }
    if prior.character_spacing != current.character_spacing {
        axes.push("Character Spacing".to_string());
    }
    if prior.vertical_align != current.vertical_align {
        use crate::ir::VerticalAlign;
        let label = match current.vertical_align {
            Some(VerticalAlign::Superscript) => "Superscript",
            Some(VerticalAlign::Subscript) => "Subscript",
            Some(VerticalAlign::Baseline) | None => "Vertical Alignment: Baseline",
        };
        axes.push(label.to_string());
    }
    if prior.shading != current.shading {
        if let Some(hex) = current.shading.as_deref() {
            axes.push(format!("Shading: {hex}"));
        } else {
            axes.push("Shading".to_string());
        }
    }
    // S256: bool decoration toggles (outline, emboss, imprint, shadow,
    // vanish→Hidden, double_strikethrough→Double Strikethrough). Each is
    // a single rPr bool; we only report a toggle (not direction), so
    // turn-off shows the same label as turn-on. Word's "Formatted:"
    // balloon does the same — it reports the property name, not the
    // before/after values.
    if prior.outline != current.outline {
        axes.push("Outline".to_string());
    }
    if prior.emboss != current.emboss {
        axes.push("Emboss".to_string());
    }
    if prior.imprint != current.imprint {
        axes.push("Imprint".to_string());
    }
    if prior.shadow != current.shadow {
        axes.push("Shadow".to_string());
    }
    if prior.vanish != current.vanish {
        axes.push("Hidden".to_string());
    }
    if prior.double_strikethrough != current.double_strikethrough {
        axes.push("Double Strikethrough".to_string());
    }
    if prior.highlight != current.highlight {
        if let Some(name) = current.highlight.as_deref() {
            axes.push(format!("Highlight: {name}"));
        } else {
            axes.push("Highlight".to_string());
        }
    }
    if prior.position != current.position {
        if let Some(pos) = current.position {
            // Word renders integer points plainly; show decimals when needed.
            let label = if pos.fract().abs() < f32::EPSILON {
                format!("Position: {}pt", pos as i32)
            } else {
                format!("Position: {pos}pt")
            };
            axes.push(label);
        } else {
            axes.push("Position".to_string());
        }
    }
    if prior.emphasis_mark != current.emphasis_mark {
        if let Some(mark) = current.emphasis_mark.as_deref() {
            axes.push(format!("Emphasis Mark: {mark}"));
        } else {
            axes.push("Emphasis Mark".to_string());
        }
    }
    if axes.is_empty() {
        None
    } else {
        Some(format!("Formatted: {}", axes.join(", ")))
    }
}

/// R-12 v2: paragraph-level companion to `describe_rpr_diff`. Compares
/// the pPrChange's prior paragraph style (and prior alignment) against
/// the paragraph's current style + alignment. Returns the same
/// "Formatted: …" string shape so the renderer doesn't need to know
/// which kind of change produced it.
///
/// Supported axes (extended each ship):
/// S254 — Indent Left.
/// S255 — Alignment (from `prior_alignment`, which is stored outside
///        `prior_paragraph_style` because `Paragraph.alignment` is a
///        top-level IR field), Paragraph Shading.
/// S257 — Keep With Next, Page Break Before, Widow/Orphan Control
///        (inverted phrasing on turn-off), Right-to-Left (bidi),
///        Text Alignment.
/// S258 — Borders Added, Tab Stops Added, Numbering (num_id).
pub fn describe_ppr_diff(change: &crate::ir::PropertyChange, current: &crate::ir::Paragraph) -> Option<String> {
    let mut axes: Vec<String> = Vec::new();
    if let Some(prior_pstyle) = change.prior_paragraph_style.as_deref() {
        if prior_pstyle.indent_left != current.style.indent_left {
            axes.push("Indent Left".to_string());
        }
        if prior_pstyle.shading != current.style.shading {
            if let Some(hex) = current.style.shading.as_deref() {
                axes.push(format!("Paragraph Shading: {hex}"));
            } else {
                axes.push("Paragraph Shading".to_string());
            }
        }
        if prior_pstyle.keep_next != current.style.keep_next {
            axes.push("Keep With Next".to_string());
        }
        if prior_pstyle.page_break_before != current.style.page_break_before {
            axes.push("Page Break Before".to_string());
        }
        if prior_pstyle.widow_control != current.style.widow_control {
            // Inverted phrasing — fixture_26 turns widow_control OFF and
            // expects "Not Widow/Orphan Control". The ON direction is
            // (informally) the no-op default, so we always label the
            // OFF state explicitly.
            if !current.style.widow_control {
                axes.push("Not Widow/Orphan Control".to_string());
            } else {
                axes.push("Widow/Orphan Control".to_string());
            }
        }
        if prior_pstyle.bidi != current.style.bidi {
            axes.push("Right-to-Left".to_string());
        }
        if prior_pstyle.text_alignment != current.style.text_alignment {
            if let Some(val) = current.style.text_alignment.as_deref() {
                axes.push(format!("Text Alignment: {val}"));
            } else {
                axes.push("Text Alignment".to_string());
            }
        }
        // S258: borders/tab_stops/numPr. Word reports these as
        // side-summary additions in the "Formatted:" balloon rather than
        // enumerating each border edge / tab position; the test fixtures
        // assert the summary phrasing. ParagraphBorders / TabStop / etc.
        // don't derive PartialEq (Vec<TabStop> + Option<ParagraphBorders>
        // with nested option types would require a wide derive cascade),
        // so we use presence/absence as the discriminator — only "Added"
        // and "Removed" cases are surfaced. A "Changed" case requires
        // PartialEq derives across the IR border/tab structs; defer
        // until a fixture needs it.
        if prior_pstyle.borders.is_none() && current.style.borders.is_some() {
            axes.push("Borders Added".to_string());
        } else if prior_pstyle.borders.is_some() && current.style.borders.is_none() {
            axes.push("Borders Removed".to_string());
        }
        if prior_pstyle.tab_stops.is_empty() && !current.style.tab_stops.is_empty() {
            axes.push("Tab Stops Added".to_string());
        } else if !prior_pstyle.tab_stops.is_empty() && current.style.tab_stops.is_empty() {
            axes.push("Tab Stops Removed".to_string());
        }
        if prior_pstyle.num_id != current.style.num_id {
            if let Some(id) = current.style.num_id.as_deref() {
                axes.push(format!("Numbering: list {id}"));
            } else {
                axes.push("Numbering Removed".to_string());
            }
        }
    }
    if let Some(prior_align) = change.prior_alignment {
        if prior_align != current.alignment {
            use crate::ir::Alignment;
            let label = match current.alignment {
                Alignment::Left => "Left",
                Alignment::Center => "Centered",
                Alignment::Right => "Right",
                Alignment::Justify => "Justified",
                Alignment::Distribute => "Distributed",
            };
            axes.push(format!("Alignment: {label}"));
        }
    }
    if axes.is_empty() {
        None
    } else {
        Some(format!("Formatted: {}", axes.join(", ")))
    }
}

/// R-12: emit one narrow "Formatted: …" balloon per rPrChange / pPrChange
/// found on this layout page. Mirrors `emit_balloons_for_layout_page` but
/// the anchor is the paragraph itself (not a `commentRangeStart` marker)
/// and the comment_id uses a synthetic prefix so the renderer / tests
/// can distinguish R-12 balloons from R-05 comment balloons:
///
///   - `rprchange:<sequence>` for run-level changes
///   - `pprchange:<sequence>` for paragraph-level changes
///
/// Resolved=true is hard-coded — per fixture_09's COM-confirmed
/// expectation, R-12 balloons use the narrow grey geometry regardless
/// of whether the underlying revision is resolved. The R-05 comment
/// balloon's resolved/unresolved distinction is about
/// `<w15:commentEx done="1"/>`, which doesn't apply to rPrChange.
fn emit_property_change_balloons_for_layout_page(
    layout_page: &mut LayoutPage,
    doc: &Document,
    ir_page_idx: usize,
) {
    use std::collections::HashMap;

    let ir_page = match doc.pages.get(ir_page_idx) {
        Some(p) => p,
        None => return,
    };

    // Same anchor-resolution strategy as the comment balloon pass: map
    // paragraph_index → first rendered (x, y) so an empty / marker-only
    // paragraph still has a stable anchor.
    let mut para_first_xy: HashMap<usize, (f32, f32)> = HashMap::new();
    for el in &layout_page.elements {
        if !matches!(&el.content, LayoutContent::Text { .. }) {
            continue;
        }
        if let Some(pi) = el.paragraph_index {
            para_first_xy.entry(pi).or_insert((el.x, el.y));
        }
    }

    let author_to_idx: HashMap<&str, usize> = doc
        .authors
        .iter()
        .map(|a| (a.display.as_str(), a.color_index))
        .collect();

    #[derive(Debug)]
    struct PendingPCBalloon {
        cid: String,
        author: String,
        author_color_index: usize,
        body: String,
        anchor_x: f32,
        anchor_y: f32,
        balloon_left: f32,
        balloon_width: f32,
        balloon_height: f32,
        y: f32,
    }

    let page_w = layout_page.width;
    let balloon_right_inset = 4.0;
    // R-12 always uses narrow geometry (resolved width).
    let balloon_width = 190.1;
    let balloon_left = (page_w - balloon_right_inset - balloon_width).max(0.0);
    let avg_glyph_pt = 5.0;
    let max_chars_per_line = ((balloon_width - 8.0) / avg_glyph_pt).max(1.0) as usize;

    let mut pending: Vec<PendingPCBalloon> = Vec::new();
    let mut next_sequence: usize = 0;

    for (pi, block) in ir_page.blocks.iter().enumerate() {
        let Block::Paragraph(p) = block else { continue };
        let Some(&(anchor_x, anchor_y)) = para_first_xy.get(&pi) else { continue };

        // pPrChange (paragraph-level).
        if let Some(change) = &p.ppr_change {
            if let Some(body) = describe_ppr_diff(change, p) {
                let cid = format!("pprchange:{next_sequence}");
                next_sequence += 1;
                let author = change.author.clone().unwrap_or_default();
                let color_idx = author_to_idx
                    .get(author.as_str())
                    .copied()
                    .unwrap_or(0);
                let body_lines = body
                    .lines()
                    .map(|l| (l.chars().count().max(1) + max_chars_per_line - 1) / max_chars_per_line)
                    .sum::<usize>()
                    .max(1);
                // Same height accounting as R-05 balloons (chip + body + pads).
                let line_height = 14.0;
                let chip_h = 14.0;
                let section_pad = 4.0;
                let outer_pad = 8.0;
                let balloon_height =
                    outer_pad + chip_h + section_pad + body_lines as f32 * line_height + outer_pad;
                pending.push(PendingPCBalloon {
                    cid,
                    author,
                    author_color_index: color_idx,
                    body,
                    anchor_x,
                    anchor_y,
                    balloon_left,
                    balloon_width,
                    balloon_height,
                    y: anchor_y,
                });
            }
        }

        // rPrChange (run-level). Multiple runs may carry rpr_change; each
        // becomes its own balloon. fixture_09 has 1, fixture_14 has 1,
        // multi-run fixtures fan out to N.
        for run in &p.runs {
            let Some(change) = &run.rpr_change else { continue };
            let Some(prior_rstyle) = change.prior_run_style.as_deref() else { continue };
            let Some(body) = describe_rpr_diff(prior_rstyle, &run.style) else { continue };
            let cid = format!("rprchange:{next_sequence}");
            next_sequence += 1;
            let author = change.author.clone().unwrap_or_default();
            let color_idx = author_to_idx
                .get(author.as_str())
                .copied()
                .unwrap_or(0);
            let body_lines = body
                .lines()
                .map(|l| (l.chars().count().max(1) + max_chars_per_line - 1) / max_chars_per_line)
                .sum::<usize>()
                .max(1);
            let line_height = 14.0;
            let chip_h = 14.0;
            let section_pad = 4.0;
            let outer_pad = 8.0;
            let balloon_height =
                outer_pad + chip_h + section_pad + body_lines as f32 * line_height + outer_pad;
            pending.push(PendingPCBalloon {
                cid,
                author,
                author_color_index: color_idx,
                body,
                anchor_x,
                anchor_y,
                balloon_left,
                balloon_width,
                balloon_height,
                y: anchor_y,
            });
        }
    }

    if pending.is_empty() {
        return;
    }

    // Stack to prevent overlap, same algorithm as R-05d.
    pending.sort_by(|a, b| {
        a.anchor_y
            .partial_cmp(&b.anchor_y)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    {
        let mut ys: Vec<(f32, f32)> = pending
            .iter()
            .map(|pb| (pb.anchor_y, pb.balloon_height))
            .collect();
        stack_balloon_ys(&mut ys, BALLOON_STACK_GAP);
        for (pb, (resolved_y, _)) in pending.iter_mut().zip(ys.iter()) {
            pb.y = *resolved_y;
        }
    }

    for pb in pending {
        let connector_color = COMMENT_HIGHLIGHT_TINT_PALETTE
            [pb.author_color_index % COMMENT_HIGHLIGHT_TINT_PALETTE.len()]
        .to_string();
        let to_x = pb.balloon_left;
        let to_y = pb.y + 5.0;
        layout_page.elements.push(LayoutElement::new(
            pb.anchor_x.min(to_x),
            pb.anchor_y.min(to_y),
            (to_x - pb.anchor_x).abs(),
            (to_y - pb.anchor_y).abs(),
            LayoutContent::BalloonConnector {
                from_x: pb.anchor_x,
                from_y: pb.anchor_y,
                to_x,
                to_y,
                color_hex: connector_color,
            },
        ));
        layout_page.elements.push(LayoutElement::new(
            pb.balloon_left,
            pb.y,
            pb.balloon_width,
            pb.balloon_height,
            LayoutContent::Balloon {
                comment_id: pb.cid,
                author: pb.author,
                author_color_index: pb.author_color_index,
                resolved: true, // R-12 balloons are narrow regardless of underlying revision state
                body: pb.body,
                replies: Vec::new(),
                anchor_x: pb.anchor_x,
                anchor_y: pb.anchor_y,
            },
        ));
    }
}

/// 6pt vertical gap between stacked balloons (R-05d). Approximate; Word's
/// actual gap looks closer to 4-8pt depending on density. Pixel-tune in
/// R-05g once GDI render lands and we can A/B compare.
const BALLOON_STACK_GAP: f32 = 6.0;

/// Pure helper for R-05d. Given a slice of `(anchor_y, height)` pairs sorted
/// ascending by `anchor_y`, mutate each pair's first element to its
/// post-stacking Y so no two balloons overlap. The first balloon stays at
/// its natural anchor; each subsequent balloon's Y is at least
/// `prev_y + prev_height + gap`.
///
/// Pure function (no I/O, no allocations beyond the input slice) — unit-
/// testable without touching the rest of layout. Test lives in `mod tests`
/// at the bottom of this file.
fn stack_balloon_ys(positions: &mut [(f32, f32)], gap: f32) {
    if positions.len() < 2 {
        return;
    }
    for i in 1..positions.len() {
        let prev_bottom = positions[i - 1].0 + positions[i - 1].1;
        let floor = prev_bottom + gap;
        if positions[i].0 < floor {
            positions[i].0 = floor;
        }
    }
}

/// Flatten a comment's `blocks` (paragraph runs) into a plain string for the
/// balloon body. Replies / nested formatting are not preserved here — R-05g
/// renderer can re-walk the structured form when it needs to.
fn comment_body_text(blocks: &[Block]) -> String {
    let mut out = String::new();
    for block in blocks {
        if let Block::Paragraph(p) = block {
            if !out.is_empty() {
                out.push('\n');
            }
            for run in &p.runs {
                out.push_str(&run.text);
            }
        }
    }
    out
}

/// S-02 (Simple mode): strip the parser's pre-applied tracked-change
/// styling (underline + red on `<w:ins>` runs, strikethrough + red on
/// `<w:del>` runs) WITHOUT removing `tracked_change` itself. Keeps R-10's
/// margin change bar firing while letting the in-line text render plain.
fn strip_parser_revision_styling(doc: &mut Document) {
    fn visit(blocks: &mut Vec<Block>) {
        for block in blocks.iter_mut() {
            match block {
                Block::Paragraph(p) => {
                    for run in &mut p.runs {
                        if let Some(tc) = run.tracked_change.as_ref() {
                            match tc.change_type.as_str() {
                                "insert" | "moveTo" => {
                                    run.style.underline = false;
                                    run.style.underline_style = None;
                                }
                                "delete" | "moveFrom" => {
                                    run.style.strikethrough = false;
                                }
                                _ => {}
                            }
                            if run.style.color.as_deref() == Some("FF0000") {
                                run.style.color = None;
                            }
                        }
                    }
                }
                Block::Table(t) => {
                    for row in &mut t.rows {
                        for cell in &mut row.cells {
                            visit(&mut cell.blocks);
                        }
                    }
                }
                Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
            }
        }
    }
    for_each_block_tree(doc, |blocks| visit(blocks));
}

/// S-02: filter / clear tracked-change runs in-place for `Original` /
/// `Final` view modes.
///
/// `final_view = true` keeps `insert` / `moveTo` runs (clears their
/// `tracked_change` so they render as normal text) and DROPS
/// `delete` / `moveFrom` runs from the paragraph.
///
/// `final_view = false` (== Original) does the inverse: drops `insert`
/// / `moveTo`, keeps `delete` / `moveFrom`.
///
/// Recurses into table cells. After this pass `apply_revision_styling`
/// is intentionally NOT called — surviving runs render as normal text.
fn filter_runs_for_show_revisions(doc: &mut Document, final_view: bool) {
    fn visit(blocks: &mut Vec<Block>, final_view: bool) {
        for block in blocks.iter_mut() {
            match block {
                Block::Paragraph(p) => {
                    p.runs.retain_mut(|run| match run.tracked_change.as_ref() {
                        None => true,
                        Some(tc) => {
                            let drop_in_final = matches!(tc.change_type.as_str(), "delete" | "moveFrom");
                            let drop_in_original = matches!(tc.change_type.as_str(), "insert" | "moveTo");
                            let drop = if final_view { drop_in_final } else { drop_in_original };
                            if drop {
                                false
                            } else {
                                // Run survives — strip the parser's
                                // pre-applied tracked-change styling
                                // (underline + red for ins, strike + red
                                // for del) and the IR marker, so the run
                                // renders as plain body text.
                                let kind = tc.change_type.clone();
                                run.tracked_change = None;
                                if kind == "insert" || kind == "moveTo" {
                                    run.style.underline = false;
                                    run.style.underline_style = None;
                                    if run.style.color.as_deref() == Some("FF0000") {
                                        run.style.color = None;
                                    }
                                } else if kind == "delete" || kind == "moveFrom" {
                                    run.style.strikethrough = false;
                                    if run.style.color.as_deref() == Some("FF0000") {
                                        run.style.color = None;
                                    }
                                }
                                true
                            }
                        }
                    });
                }
                Block::Table(t) => {
                    for row in &mut t.rows {
                        for cell in &mut row.cells {
                            visit(&mut cell.blocks, final_view);
                        }
                    }
                }
                Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
            }
        }
    }
    for_each_block_tree(doc, |blocks| visit(blocks, final_view));
}

fn apply_revision_styling_to_run(
    run: &mut Run,
    tc: &TrackedChange,
    author_to_idx: &std::collections::HashMap<String, usize>,
) {
    // Choose the ink color: hard-coded green for moves, otherwise the author's
    // palette slot. If the author isn't in the palette (defensive fallback —
    // I-03 builds the palette from the same source authors so this should be
    // unreachable in practice), fall back to slot 0.
    let color_hex = match tc.change_type.as_str() {
        "moveFrom" | "moveTo" => REVISION_MOVE_COLOR.to_string(),
        _ => {
            let idx = tc
                .author
                .as_deref()
                .and_then(|a| author_to_idx.get(a).copied())
                .unwrap_or(0);
            REVISION_AUTHOR_PALETTE[idx % REVISION_AUTHOR_PALETTE.len()].to_string()
        }
    };

    match tc.change_type.as_str() {
        "insert" => {
            run.style.underline = true;
            if run.style.underline_style.is_none() {
                run.style.underline_style = Some("single".to_string());
            }
            run.style.color = Some(color_hex);
        }
        "delete" => {
            run.style.strikethrough = true;
            run.style.color = Some(color_hex);
        }
        "moveFrom" => {
            // R-11 v2 (R66, 2026-04-29): Word default is double-strikethrough
            // in green. COM-confirmed by pixel sampling fixture_08 (two
            // full-width green strike lines 1pt apart on the moved runs).
            run.style.strikethrough = true;
            run.style.double_strikethrough = true;
            run.style.color = Some(color_hex);
        }
        "moveTo" => {
            // R-11 v2 (R66): Word default is double-underline in green.
            // COM-confirmed by pixel sampling fixture_08.
            run.style.underline = true;
            run.style.underline_style = Some("double".to_string());
            run.style.color = Some(color_hex);
        }
        _ => {
            // Unknown change_type — leave the run alone. This branch keeps
            // forward-compatibility with future change kinds without forcing
            // a code change.
        }
    }
}

/// Result of layout: positioned elements across pages
pub struct LayoutResult {
    pub pages: Vec<LayoutPage>,
}

pub struct LayoutPage {
    pub width: f32,
    pub height: f32,
    pub elements: Vec<LayoutElement>,
}

pub struct LayoutElement {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub content: LayoutContent,
    /// Source paragraph index in the document body (for hit testing / editing).
    /// For table cell elements this is the TABLE's page-level block_idx
    /// (shared across all cells of the table).
    pub paragraph_index: Option<usize>,
    /// Source run index within the paragraph
    pub run_index: Option<usize>,
    /// Character offset within the run's text where this fragment starts
    pub char_offset: Option<usize>,
    /// R7.32 (Day 33 part 72, 2026-05-13): paragraph index within a table cell
    /// (0-based, counts only Block::Paragraph blocks in document order).
    /// `paragraph_index` alone is shared across all cells of the table, so
    /// without this field aggregate_dump (matcher) cannot distinguish which
    /// cell paragraph an element comes from and misattributes diff matches.
    /// None for non-cell elements. See e3c545 # プレフィックス case.
    pub cell_paragraph_index: Option<usize>,
    /// R7.44 (Day 34 part 13, 2026-05-13): row and column index within the
    /// table block, so aggregate_dump can distinguish cells that share
    /// (paragraph_index, cell_paragraph_index). Without these, all cells'
    /// first paragraphs share key (block_idx, cpi=0) and collapse into one
    /// aggregate record — the matcher sees one "千円千円千円千円" instead of
    /// four separate "千円" cells. See 04b88e w_i=99 / b5f706 6(9) cases.
    /// None for non-cell elements.
    pub cell_row_index: Option<usize>,
    pub cell_col_index: Option<usize>,
    /// R7.56 (Day 34 part 25, 2026-05-13): true on the FIRST text element of
    /// a cell paragraph whose run[0] carries `<w:lastRenderedPageBreak/>`.
    /// The row-split logic uses this to force a page break before this
    /// element (mid-cell LRPB respect, analogous to R7.45/R7.47 for body
    /// and row-first cases). e3c545 cpi=81/cpi=N LRPB markers caused
    /// remaining -1 outliers w_i=202, 314, 483 because the existing
    /// R7.47 row-LRPB check only inspects `cell.blocks.first()`.
    pub is_paragraph_start_with_lrpb: bool,
    /// R7.61 (Day 36 part 8, 2026-05-14): true on text elements emitted from
    /// vMerge="restart" cell content whose final Y exceeds the page bottom.
    /// Post-paginate sweep moves these to the next page so pagination reports
    /// them as page N+1 rather than page N (matches Word's split-mid-cell for
    /// vMerge restart overflow). Narrow scope: only set in cell render path
    /// for vMerge=restart cells past page_bottom — body text and other cells
    /// are never marked. a1d6 row 13 ※２/※３: w_i=179/180 are on Word p4 but
    /// Oxi renders them visually past page 3's bottom; this flag lifts them
    /// to page 4 in the pagination output.
    pub vmerge_restart_overflow_to_next_page: bool,
    /// Session 72 Phase A (2026-05-17): vertical offset from LINE BOX TOP
    /// to glyph top. Currently `y` already includes this offset (y = line_box_top
    /// + text_y_off). After Session 75, `y` will be LINE BOX TOP and renderers
    /// will add this at draw time. Populated by body line emit (mod.rs ~4197)
    /// and table cell emit (mod.rs ~6840) for text elements; 0.0 for non-text
    /// elements. See [[session71-y-convention-refactor-design]].
    pub text_y_off: f32,
}

impl LayoutElement {
    /// Create a non-text element (border, shading, image, etc.) with no source indices.
    fn new(x: f32, y: f32, width: f32, height: f32, content: LayoutContent) -> Self {
        Self { x, y, width, height, content, paragraph_index: None, run_index: None, char_offset: None, cell_paragraph_index: None, cell_row_index: None, cell_col_index: None, is_paragraph_start_with_lrpb: false, vmerge_restart_overflow_to_next_page: false, text_y_off: 0.0 }
    }

    /// Create a text element with source location for hit testing.
    #[allow(dead_code)]
    fn text(x: f32, y: f32, width: f32, height: f32, content: LayoutContent,
            para_idx: usize, run_idx: usize, char_offset: usize) -> Self {
        Self { x, y, width, height, content,
               paragraph_index: Some(para_idx), run_index: Some(run_idx), char_offset: Some(char_offset),
               cell_paragraph_index: None, cell_row_index: None, cell_col_index: None, is_paragraph_start_with_lrpb: false, vmerge_restart_overflow_to_next_page: false, text_y_off: 0.0 }
    }
}

pub enum LayoutContent {
    Text {
        text: String,
        font_size: f32,
        font_family: Option<String>,
        bold: bool,
        italic: bool,
        underline: bool,
        underline_style: Option<String>,
        strikethrough: bool,
        /// R-11 v2 (R66): w:dstrike emits double strikethrough.
        /// Propagated from RunStyle.double_strikethrough; renderers may
        /// fall back to single strikethrough in v1.
        double_strikethrough: bool,
        color: Option<String>,
        highlight: Option<String>,
        field_type: Option<FieldType>,
        /// Pixel-snapped character spacing in points (0.0 = no extra spacing)
        character_spacing: f32,
        /// Horizontal font scale percentage (100 = default, <100 compresses glyphs)
        /// OOXML w:w value. Renderer applies via CreateFontW lfWidth.
        text_scale: f32,
        /// Session 132 (2026-05-20): true if this text element should be
        /// rendered with 90° CW rotation (textDirection="tbRlV" cells).
        /// GDI: set LOGFONT.lfEscapement = -900. DWrite: apply per-run
        /// transform matrix. Default false (horizontal flow).
        is_vertical: bool,
    },
    Image {
        data: Vec<u8>,
        content_type: Option<String>,
    },
    TableBorder {
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: Option<String>,
        width: f32,
        /// S480: OOXML border art style (w:val) — "single"/"dashed"/"dotted"/
        /// "dashDotStroked"/etc. None = solid. Renderers map this to a dash
        /// pattern; unknown/None/single = solid stroke.
        style: Option<String>,
    },
    CellShading {
        color: String,
    },
    /// A filled/stroked rectangle, optionally with rounded corners.
    BoxRect {
        fill: Option<String>,
        stroke_color: Option<String>,
        stroke_width: f32,
        corner_radius: f32,
    },
    /// A preset shape outline (e.g. bracketPair, brace, etc.)
    PresetShape {
        shape_type: String,
        stroke_color: Option<String>,
        stroke_width: f32,
        /// Connector diagonal direction (a:xfrm flipH/flipV). Default false.
        flip_h: bool,
        flip_v: bool,
        /// Connector arrowheads (a:ln headEnd/tailEnd). Default false.
        arrow_head: bool,
        arrow_tail: bool,
    },
    /// Begin a clipping region. All subsequent elements until ClipEnd are clipped to this rect.
    ClipStart,
    /// End the current clipping region (restore graphics state).
    ClipEnd,
    /// A right-margin comment balloon (R-05). Renderer is expected to draw a
    /// rounded-rect background filled with the resolved-or-unresolved tint
    /// for `author_color_index`, then lay the comment `body` and any
    /// `replies` inside the bounding box `(x, y, width, height)`. See
    /// `docs/spec/comments_tracked_changes/r05_balloon_design.md`.
    Balloon {
        comment_id: String,
        author: String,
        author_color_index: usize,
        resolved: bool,
        body: String,
        replies: Vec<BalloonReply>,
        /// X of the inline anchor on the body (for renderer-side connector
        /// drawing or hit-testing). Set to the same value as the
        /// `BalloonConnector.from_x` if the renderer prefers reading from
        /// the balloon directly.
        anchor_x: f32,
        /// Y of the inline anchor on the body — typically the rendered Y of
        /// the comment's `commentRangeStart` line.
        anchor_y: f32,
    },
    /// Dotted connector line from the inline comment anchor to its balloon
    /// (R-07). Drawn separately so the renderer can apply a dashed pen
    /// without bundling the geometry into `Balloon`.
    BalloonConnector {
        from_x: f32,
        from_y: f32,
        to_x: f32,
        to_y: f32,
        color_hex: String,
    },
}

/// One reply inside a balloon (R-08). Renderer indents these by ~10pt
/// inside the parent balloon's body area.
#[derive(Debug, Clone)]
pub struct BalloonReply {
    pub author: String,
    pub author_color_index: usize,
    pub body: String,
}

/// Two-track cursor for the layout pipeline.
///
/// Session 99 (2026-05-18) decoupling refactor (see
/// `docs/spec/decoupling_refactor_design.md`). Splits the running y
/// position into:
///   - `cursor_y` — used by page-break gates (row-fit check at
///     mod.rs:6255, LRPB threshold at mod.rs:6276, space_before
///     suppression at mod.rs:3388). MUST remain stable across the
///     refactor to preserve Phase 1 pagination correctness.
///   - `visual_y` — used by visual emit sites (text lines, table
///     borders, cell shading, images). Initially mirrors `cursor_y`
///     exactly (Phase A1 = no behavior change). Phase B can let it
///     diverge per-quirk (e.g. Word's atLeast row snap visual = trH
///     rounded to 0.75pt, while page-break uses unsnapped trH).
///
/// Phase A1 invariant: every advance mirrors to both fields. No
/// reads of `visual_y` yet — all visual emit sites still read
/// `cursor_y`. Verified by Phase 1 pagination_diff (53/55 unchanged)
/// and SSIM verify (0 regressed).
#[derive(Debug, Clone, Copy)]
pub struct LayoutCursor {
    pub cursor_y: f32,
    pub visual_y: f32,
    /// S494 (2026-06-04): un-rounded "ideal" twip position for the current LM2
    /// docGrid line run. Word advances grid lines by the EXACT grid pitch and
    /// snaps the ABSOLUTE position to the 96dpi device pixel (0.75pt), tracking
    /// a fractional pitch (357tw=17.85pt) instead of the integer-rounded 18.0pt.
    /// Carries the un-rounded accumulation so the rounding error doesn't compound
    /// across paragraphs. 0.0 = uninitialized → resync to cursor_y on next use.
    /// Consulted on the empty-line device-snap path (opt-out OXI_S494_DISABLE);
    /// otherwise inert.
    pub lm2_ideal_y: f32,
}

impl LayoutCursor {
    pub fn new(y: f32) -> Self {
        Self { cursor_y: y, visual_y: y, lm2_ideal_y: 0.0 }
    }

    /// Advance both tracks by the same amount (Phase A default).
    pub fn advance(&mut self, dy: f32) {
        self.cursor_y += dy;
        self.visual_y += dy;
    }

    /// Advance the two tracks by DIFFERENT amounts (Phase B divergence).
    ///
    /// Session 103 (2026-05-18) Phase B1. First use case: Word's atLeast
    /// row snap (S98 finding) advances visual position by
    /// `ceil((trH + bw) / 0.75) × 0.75` while page-break cursor advances
    /// by the unsnapped `max(content, trH)`. Keeping `cursor_y` stable
    /// preserves Phase 1 pagination while `visual_y` corrects pixel
    /// positions. See `docs/spec/decoupling_refactor_design.md` Phase B1.
    pub fn advance_split(&mut self, page_dy: f32, visual_dy: f32) {
        self.cursor_y += page_dy;
        self.visual_y += visual_dy;
    }

    /// Set both tracks to the same value (used at page boundary reset).
    pub fn set(&mut self, y: f32) {
        self.cursor_y = y;
        self.visual_y = y;
    }
}

pub struct LayoutEngine {
    default_font_size: f32,
    default_font_family: Option<String>,
    default_font_family_east_asia: Option<String>,
    registry: FontMetricsRegistry,
    /// Compatibility: adjustLineHeightInTable=true disables grid snap in table cells.
    adjust_line_height_in_table: bool,
    /// Document-level default tab stop interval (from w:settings/w:defaultTabStop)
    default_tab_stop: f32,
    /// Compatibility mode: 14=Word 2010 (table cells no grid snap), 15=Word 2013+ (grid snap)
    compat_mode: u32,
    /// S545: whether compatibilityMode was explicitly present in settings.xml.
    /// Absent = legacy (Word ≤2010) document — Word applies ≤14 layout
    /// behaviors (jc=left demand oikomi) even though compat_mode reports 15.
    compat_mode_explicit: bool,
    /// w:characterSpacingControl: enable yakumono (CJK punctuation) compression
    /// True when value is "compressPunctuation" or "compressPunctuationAndJapaneseKana".
    /// False (default) when "doNotCompress" or absent.
    compress_punctuation: bool,
    /// w:doNotExpandShiftReturn: don't justify lines ending with soft break (Shift+Enter)
    do_not_expand_shift_return: bool,
    /// w:balanceSingleByteDoubleByteWidth: when set, character_spacing is doubled
    /// for CJK fullwidth chars (Word "balance single/double byte widths" mode).
    /// Derived from V19 minimal repro vs real 1636 (Session 56 Finding 3).
    balance_single_byte_double_byte_width: bool,
    /// R-05b: when the document has any comments, the body's available width
    /// is reduced by this many points to make room for the right-margin
    /// balloon column. 0.0 when the document has no comments. Set in
    /// `for_document` from `doc.comments.is_empty()`.
    balloon_column_width: f32,
    /// S-01: render-time toggle for the comment family — balloons,
    /// connectors, in-line range highlight, AND body-width compression. When
    /// `false`, the LayoutResult contains zero `Balloon` / `BalloonConnector`
    /// elements, no comment-range highlight tints, and the body uses the full
    /// width. Default `true` (Word's "All Markup" view).
    show_comments: bool,
    /// S-02: which view of revisions to render. Mirrors Word's "Display for
    /// Review" dropdown via the existing `ir::ShowRevisions` enum:
    ///
    /// - `All` (default): every revision rendered with markup (current
    ///   behavior — pre-S-02).
    /// - `Simple`: change bar in left margin only (R-10 fires); ins/del
    ///   text renders without color or underline/strike.
    /// - `Original`: pre-edit view — `<w:ins>` and `<w:moveTo>` runs are
    ///   removed; `<w:del>` and `<w:moveFrom>` render as normal text.
    /// - `Final`: post-edit view — `<w:del>` and `<w:moveFrom>` runs are
    ///   removed; `<w:ins>` and `<w:moveTo>` render as normal text.
    show_revisions: ShowRevisions,
    /// S463 (2026-05-31): true when ANY body text in the document contains a
    /// CJK character. The Latin-table border-overhead correction
    /// (OXI_S463_LATIN_BORDER) is scoped to Latin-context documents: in CJK
    /// documents the table-row deficit is masked by a separate compensating
    /// error below the table, so applying the (otherwise-correct) overhead
    /// there regresses SSIM (gen2 Japanese-template family −0.01..−0.035).
    doc_body_has_cjk: bool,
}

/// S463: recursive CJK scan over a block (paragraph runs + nested table cells).
fn block_has_cjk(block: &Block) -> bool {
    match block {
        Block::Paragraph(p) => p.runs.iter().any(|r| r.text.chars().any(kinsoku::is_cjk)),
        Block::Table(t) => t.rows.iter().any(|row| {
            row.cells.iter().any(|c| c.blocks.iter().any(block_has_cjk))
        }),
        _ => false,
    }
}

/// Word's default heading font sizes (in points)
fn heading_default_font_size(level: u8) -> f32 {
    // Word default heading sizes (half-points in styles.xml → points)
    match level {
        1 => 14.0,  // sz=28
        2 => 13.0,  // sz=26
        3 => 11.0,  // sz=22 (default body)
        4 => 11.0,
        _ => 11.0,
    }
}

/// S546 (2026-06-11): autoSpaceDE/DN gap = exactly fontSize/4 in Word's
/// layout space (2.625 @10.5pt). The old per-fontSize table
/// (((fs/2)+0.5).floor()*0.5: 2.5 @9-10.5, 3.0 @11-12, 3.5 @14 — COM 2026-04-08)
/// was derived from painted advances, which carry the 96dpi cumulative px-snap
/// (true 13.125 paints as 13.5 → "3.0 before"; cluster end lands exact →
/// "2.25 after"). The S546 fs-sweep (repro_s546_digit_sweep.py, fs 9/10.5/12/14,
/// MS Mincho + Century digits/letters/kana boundaries) is explained to the
/// decimal by gap = fs/4 in true space, both sides; per-cluster total = fs/2.
/// Opt-out: OXI_S546_DISABLE (shared with the fs/2 halfwidth fix).
fn s546_autospace_extra(font_size: f32) -> f32 {
    if crate::font::s546_exact_halfwidth() {
        font_size / 4.0
    } else {
        ((font_size / 2.0) + 0.5).floor() * 0.5
    }
}

/// Snap character spacing to pixel grid (DPI=96 fixed).
/// Character spacing pixel-snap: Word converts twips→pixels at 96 DPI
/// using round-to-nearest integer division, then back to points.
///
/// Derived from COM measurement: comparing Word's actual character
/// positions (Range.Information) against input spacing values.
/// Example: cs=-0.45pt → -9tw → round(-9*96/1440) = -1px → -0.75pt
fn snap_character_spacing(cs_pt: f32) -> f32 {
    // 2026-05-04 (Day 11-17): Word applies char_spacing at twip precision (1/20pt),
    // NOT GDI pixel precision. Old code snapped to 96 DPI pixels which destroyed
    // precision: cs in [-1, -19]tw all snapped to -0.75pt; cs=-20tw also to -0.75pt.
    // 6-variant cs/grid isolation matrix verified the rule:
    //   v0/v3 (no cs): Oxi matches Word exactly
    //   v1 cs=-9: Word -0.444pt, Oxi (old) -0.75pt → over-compress by 0.31pt/char
    //   v2 cs=-20: Word -1.000pt, Oxi (old) -0.75pt → under-compress by 0.25pt/char
    // Replace with twip-precision rounding.
    // Full-baseline verify: net +0.008 (4 improvements, 5 regressions <0.005pt).
    (cs_pt * 20.0).round() / 20.0
}

impl LayoutEngine {
    pub fn new() -> Self {
        Self {
            default_font_size: 11.0,
            default_font_family: None,
            default_font_family_east_asia: None,
            registry: FontMetricsRegistry::load(),
            adjust_line_height_in_table: false,
            default_tab_stop: 36.0,
            compat_mode: 15,
            compat_mode_explicit: true,
            compress_punctuation: false,
            do_not_expand_shift_return: false,
            balance_single_byte_double_byte_width: false,
            balloon_column_width: 0.0,
            show_comments: true,
            show_revisions: ShowRevisions::All,
            doc_body_has_cjk: false,
        }
    }

    /// S-01 setter: toggle whether comments / balloons / range highlight render.
    pub fn with_show_comments(mut self, show: bool) -> Self {
        self.show_comments = show;
        self
    }

    /// S-02 setter: pick the revision-display mode (`ShowRevisions::{All,
    /// Simple, Original, Final}`). See the field doc-comment for what each
    /// mode does.
    pub fn with_show_revisions(mut self, mode: ShowRevisions) -> Self {
        self.show_revisions = mode;
        self
    }

    /// Create a LayoutEngine with document-specific defaults from docDefaults
    pub fn for_document(doc: &Document) -> Self {
        let default_font_size = doc.styles.doc_default_run_style
            .as_ref()
            .and_then(|s| s.font_size)
            .unwrap_or(11.0);
        let default_font_family = doc.styles.doc_default_run_style
            .as_ref()
            .and_then(|s| s.font_family.clone());
        let default_font_family_east_asia = doc.styles.doc_default_run_style
            .as_ref()
            .and_then(|s| s.font_family_east_asia.clone());
        Self {
            default_font_size,
            default_font_family,
            default_font_family_east_asia,
            registry: FontMetricsRegistry::load(),
            adjust_line_height_in_table: doc.adjust_line_height_in_table,
            default_tab_stop: doc.default_tab_stop.unwrap_or(36.0),
            compat_mode: doc.compat_mode,
            compat_mode_explicit: doc.compat_mode_explicit,
            compress_punctuation: doc.compress_punctuation,
            do_not_expand_shift_return: doc.do_not_expand_shift_return,
            balance_single_byte_double_byte_width: doc.balance_single_byte_double_byte_width,
            // R-05b: 293.8pt balloon column + 24pt buffer between body and
            // balloon = 317.8pt total. Width is COM-confirmed (fixture_01
            // pixel pass, 2026-04-25); buffer is approximate and refined as
            // R-05c+ pixel-tests narrow it.
            balloon_column_width: if doc.comments.is_empty() { 0.0 } else { 293.8 + 24.0 },
            show_comments: true,
            show_revisions: ShowRevisions::All,
            doc_body_has_cjk: doc.pages.iter().any(|pg| pg.blocks.iter().any(block_has_cjk)),
        }
    }

    pub fn layout(&self, doc: &Document) -> LayoutResult {
        // Pre-pass: resolve fitText runs using actual font metrics
        let mut doc_resolved = doc.clone();
        for page in &mut doc_resolved.pages {
            self.resolve_fit_text_page(page);
        }

        // S-02: filter revisions by ShowRevisions mode BEFORE applying the
        // styling pass, so Original/Final actually drop runs from the IR
        // rather than just hiding them at render time.
        match self.show_revisions {
            ShowRevisions::All => {
                // Default — keep every run, apply markup styling.
                apply_revision_styling(&mut doc_resolved);
            }
            ShowRevisions::Simple => {
                // Change bar only — keep tracked_change so R-10 still fires
                // a margin bar per revision-bearing line, but DON'T apply the
                // color / underline / strike pre-pass. Also strip the
                // parser's pre-applied underline+red on ins / strike+red on
                // del so the surviving runs read as plain body text.
                strip_parser_revision_styling(&mut doc_resolved);
            }
            ShowRevisions::Original => {
                // Pre-edit view — drop `ins` / `moveTo` runs entirely. Keep
                // `del` / `moveFrom` runs but clear their `tracked_change`
                // so they render as normal text (no strikethrough, no
                // margin bar).
                filter_runs_for_show_revisions(&mut doc_resolved, false);
            }
            ShowRevisions::Final => {
                // Post-edit view — drop `del` / `moveFrom` runs. Keep `ins`
                // / `moveTo` runs but clear their `tracked_change` so they
                // render as normal text.
                filter_runs_for_show_revisions(&mut doc_resolved, true);
            }
        }

        // Pre-pass: apply in-line comment-range highlight tint (R-04). Must
        // run after `apply_revision_styling` so revision-bearing runs still
        // get the tint on top of their underline/strikethrough color. Gated
        // by S-01.
        if self.show_comments {
            apply_comment_range_highlighting(&mut doc_resolved);
        }

        let mut pages = Vec::new();
        // R-05c: track which IR page each LayoutPage came from so the
        // post-pass can resolve `paragraph_index` → source `Run` for balloon
        // emission. A single IR page may produce multiple LayoutPages
        // (pagination), all sharing the same IR index.
        let mut layout_to_ir_page: Vec<usize> = Vec::new();

        for (ir_idx, page) in doc_resolved.pages.iter().enumerate() {
            let laid_out = self.layout_page(page);
            for _ in &laid_out {
                layout_to_ir_page.push(ir_idx);
            }
            pages.extend(laid_out);
        }

        // R-05c post-pass: emit one Balloon LayoutElement per visible
        // comment, anchored to the rendered Y of its `commentRangeStart`.
        // Gated by S-01.
        if self.show_comments && !doc_resolved.comments.is_empty() {
            for (layout_idx, layout_page) in pages.iter_mut().enumerate() {
                let ir_idx = layout_to_ir_page[layout_idx];
                emit_balloons_for_layout_page(layout_page, &doc_resolved, ir_idx);
            }
        }

        // R-12 post-pass: emit narrow "Formatted: …" balloons for
        // rPrChange / pPrChange revisions. Gated by show_comments (the
        // margin-balloon column visibility toggle) — these are visually
        // peer to comment balloons.
        if self.show_comments {
            for (layout_idx, layout_page) in pages.iter_mut().enumerate() {
                let ir_idx = layout_to_ir_page[layout_idx];
                emit_property_change_balloons_for_layout_page(layout_page, &doc_resolved, ir_idx);
            }
        }

        // Post-layout pass: substitute PAGE and NUMPAGES field placeholders
        let total_pages = pages.len();
        for (page_idx, page) in pages.iter_mut().enumerate() {
            // S534 (2026-06-10): apply the section's pgNumType format to the
            // PAGE field. 3a4f's footer section sets `<w:pgNumType
            // w:fmt="numberInDash"/>` → Word renders "- 34 -"; Oxi rendered
            // the bare number. layout_to_ir_page maps this LayoutPage back to
            // its IR section page, which carries page_number_format.
            let page_num_fmt = layout_to_ir_page.get(page_idx)
                .and_then(|&ir| doc_resolved.pages.get(ir))
                .and_then(|p| p.page_number_format.clone());
            for elem in &mut page.elements {
                if let LayoutContent::Text { text, field_type: Some(ft), font_size, .. } = &mut elem.content {
                    let new_text = match ft {
                        FieldType::Page => match page_num_fmt.as_deref() {
                            Some(fmt) => crate::parser::numbering::format_number(
                                (page_idx + 1) as u32, fmt),
                            None => format!("{}", page_idx + 1),
                        },
                        FieldType::NumPages => format!("{}", total_pages),
                    };
                    if &new_text != text {
                        // Estimate new width: each digit is ~size/2 pt for typical fonts.
                        // Old width is `elem.width`. The element's text already reflects
                        // the placeholder text length (typically "1") used during layout.
                        // 2026-05-03: per spec_page_field_width_gap_2026_05_02.md.
                        let old_text = text.clone();
                        let fs = *font_size;
                        // Approximate digit width ≈ fs * 0.5 (for digits 0-9 in typical fonts)
                        let old_w = elem.width;
                        let digit_w = if !old_text.is_empty() {
                            old_w / old_text.chars().count() as f32
                        } else {
                            fs * 0.5
                        };
                        let new_w = digit_w * new_text.chars().count() as f32;
                        let dw = new_w - old_w;
                        // For centered text, shift x by half the width diff to keep
                        // the visual center aligned. This relies on the placeholder
                        // having been positioned with intended centering.
                        elem.x -= dw / 2.0;
                        elem.width = new_w;
                        *text = new_text;
                    }
                }
            }
        }

        LayoutResult { pages }
    }

    /// Resolve font size for a run, considering paragraph style defaults and heading level
    fn resolve_font_size(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> f32 {
        let base = if let Some(fs) = run_style.font_size {
            fs
        } else if let Some(ref drs) = para_style.default_run_style {
            if let Some(fs) = drs.font_size {
                fs
            } else if let Some(level) = para_style.heading_level {
                heading_default_font_size(level)
            } else {
                self.default_font_size
            }
        } else if let Some(level) = para_style.heading_level {
            heading_default_font_size(level)
        } else {
            self.default_font_size
        };
        // Word auto-shrinks superscript/subscript to 2/3 of base size
        // when no explicit font_size is set on the run.
        if run_style.font_size.is_none() {
            if matches!(run_style.vertical_align, Some(VerticalAlign::Superscript) | Some(VerticalAlign::Subscript)) {
                return (base * 2.0 / 3.0 * 2.0).round() / 2.0; // round to 0.5pt
            }
        }
        base
    }

    /// Resolve font family for a run.
    /// For CJK text, prefer font_family_east_asia over font_family.
    fn resolve_font_family<'a>(&'a self, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        if let Some(ref ff) = run_style.font_family {
            return Some(ff.as_str());
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref ff) = drs.font_family {
                return Some(ff.as_str());
            }
        }
        // Fallback to document default font (docDefaults rPrDefault)
        self.default_font_family.as_deref()
    }

    /// Resolve font family considering East Asian font for CJK characters.
    fn resolve_font_family_for_text<'a>(&'a self, text: &str, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        let has_cjk = text.chars().any(|c| kinsoku::is_cjk(c));
        if has_cjk {
            // Prefer East Asian font for CJK text: run → paragraph → docDefaults
            if let Some(ref ff) = run_style.font_family_east_asia {
                return Some(ff.as_str());
            }
            if let Some(ref drs) = para_style.default_run_style {
                if let Some(ref ff) = drs.font_family_east_asia {
                    return Some(ff.as_str());
                }
            }
            // Fall back to document-level default East Asian font (from docDefaults/theme)
            if let Some(ref ff) = self.default_font_family_east_asia {
                return Some(ff.as_str());
            }
        }
        self.resolve_font_family(run_style, para_style)
    }

    /// Get font metrics for a run (uses registry with font-family resolution).
    /// Considers bold to look up Bold variant when applicable.
    fn metrics_for(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        match self.resolve_font_family(run_style, para_style) {
            Some(family) => self.registry.get_with_bold(family, self.resolve_bold(run_style, para_style)),
            None => self.registry.default_metrics(),
        }
    }

    /// Get font metrics considering East Asian font for CJK text.
    fn metrics_for_text(&self, text: &str, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        match self.resolve_font_family_for_text(text, run_style, para_style) {
            Some(family) => self.registry.get_with_bold(family, self.resolve_bold(run_style, para_style)),
            None => self.registry.default_metrics(),
        }
    }

    /// Get font metrics for the paragraph mark (¶) / empty paragraph.
    /// COM-confirmed: Word uses the East Asian font for empty paragraph line height
    /// in CJK documents (0e7a: MS 明朝 13.5pt, not Calibri 15.5pt).
    fn metrics_for_para_mark(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        if let Some(m) = self.metrics_for_cjk(run_style, para_style) {
            return m;
        }
        self.metrics_for(run_style, para_style)
    }

    /// Get font metrics for a single character, using East Asian font for CJK.
    fn metrics_for_char(&self, ch: char, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        if kinsoku::is_cjk(ch) {
            if let Some(m) = self.metrics_for_cjk(run_style, para_style) {
                return m;
            }
        }
        self.metrics_for(run_style, para_style)
    }

    /// Get East Asian font metrics if an east-asia font family is specified.
    /// Returns None if no east-asia font is set (caller should fall back to latin metrics).
    fn metrics_for_cjk(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> Option<&FontMetrics> {
        if let Some(ref ff) = run_style.font_family_east_asia {
            return Some(self.registry.get(ff.as_str()));
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref ff) = drs.font_family_east_asia {
                return Some(self.registry.get(ff.as_str()));
            }
        }
        // Fall back to document-level default East Asian font
        if let Some(ref ff) = self.default_font_family_east_asia {
            return Some(self.registry.get(ff.as_str()));
        }
        None
    }

    /// Resolve bold for a run, considering paragraph style defaults
    fn resolve_bold(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> bool {
        if run_style.bold {
            return true;
        }
        if let Some(ref drs) = para_style.default_run_style {
            if drs.bold {
                return true;
            }
        }
        if let Some(level) = para_style.heading_level {
            return level <= 2;
        }
        false
    }

    fn resolve_color<'a>(&self, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        if let Some(ref c) = run_style.color {
            return Some(c.as_str());
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref c) = drs.color {
                return Some(c.as_str());
            }
        }
        None
    }

    fn resolve_italic(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> bool {
        if run_style.italic {
            return true;
        }
        if let Some(ref drs) = para_style.default_run_style {
            if drs.italic {
                return true;
            }
        }
        false
    }

    /// Default font metrics for the document (uses docDefaults font if set, otherwise Calibri).
    fn doc_default_metrics(&self) -> &FontMetrics {
        match self.default_font_family.as_deref() {
            Some(ff) => self.registry.get(ff),
            None => self.registry.default_metrics(),
        }
    }

    /// Resolve fitText runs: calculate actual character widths using font metrics,
    /// then set character_spacing (expand) or text_scale (compress) to match target.
    fn resolve_fit_text_page(&self, page: &mut Page) {
        self.resolve_fit_text_blocks(&mut page.blocks);
        for note in &mut page.footnotes {
            self.resolve_fit_text_blocks(&mut note.blocks);
        }
        self.resolve_fit_text_blocks(&mut page.header);
        self.resolve_fit_text_blocks(&mut page.footer);
    }

    fn resolve_fit_text_blocks(&self, blocks: &mut [Block]) {
        for block in blocks.iter_mut() {
            match block {
                Block::Paragraph(para) => {
                    self.resolve_fit_text_runs(&mut para.runs, &para.style);
                }
                Block::Table(table) => {
                    for row in table.rows.iter_mut() {
                        for cell in row.cells.iter_mut() {
                            self.resolve_fit_text_blocks(&mut cell.blocks);
                        }
                    }
                }
                _ => {}
            }
        }
    }

    fn resolve_fit_text_runs(&self, runs: &mut Vec<Run>, para_style: &ParagraphStyle) {
        let mut i = 0;
        while i < runs.len() {
            if let (Some(target_w), Some(group_id)) = (runs[i].style.fit_text, runs[i].style.fit_text_id) {
                let start = i;
                while i < runs.len() && runs[i].style.fit_text_id == Some(group_id) {
                    i += 1;
                }
                let char_count: usize = runs[start..i].iter()
                    .map(|r| r.text.chars().count()).sum();
                if char_count == 0 { continue; }
                // 2026-04-19: If docx already has explicit text_scale set for
                // this fit group, trust Word's pre-computed values and skip the
                // recomputation (which may disagree due to CJK pitch/yakumono
                // differences, causing overflow on b35 事務処理体制 row).
                if runs[start..i].iter().any(|r| r.style.text_scale.is_some()) {
                    continue;
                }

                // Get ACTUAL natural width by calling break_into_lines with cs=0.
                // This accounts for autoSpaceDE, yakumono, etc.
                let saved_cs: Vec<Option<f32>> = runs[start..i].iter()
                    .map(|r| r.style.character_spacing).collect();
                for run in &mut runs[start..i] {
                    run.style.character_spacing = Some(0.0);
                }
                let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> =
                    runs[start..i].iter().enumerate()
                    .map(|(ri, run)| (run.text.as_str(), &run.style, None, ri, 0))
                    .collect();
                let lines = self.break_into_lines(&fragments, 1e6, 0.0, para_style, None, None, true, false, true, false, false);
                let natural_w: f32 = lines.iter()
                    .flat_map(|l| l.fragments.iter())
                    .map(|f| f.width)
                    .sum();

                // Restore original cs before overriding
                for (idx, run) in runs[start..i].iter_mut().enumerate() {
                    run.style.character_spacing = saved_cs[idx];
                }

                if natural_w > 0.01 {
                    if natural_w <= target_w {
                        // 2026-04-20: Word's fitText distributes expansion
                        // proportional to char fullwidth/halfwidth (a fullwidth CJK
                        // char gets ~2× the cs that a halfwidth ASCII digit gets).
                        // Per-char cs = per_em_cs × char_em_width where
                        //   per_em_cs = (target − natural) / total_em
                        //   char_em_width = 1.0 for fullwidth, 0.5 for halfwidth
                        // N-1 semantics preserved by setting last char's cs=0.
                        // b837 p1 meta "平成29年5月30日" required uniform cs first
                        // (af3c790) to fix non-uniform spread, then N-1 (ae609ef)
                        // for trailing space. Per-char proportional refines '2'/'3'
                        // position at CJK→halfwidth-digit boundaries.
                        let total_em: f32 = runs[start..i].iter()
                            .flat_map(|r| r.text.chars())
                            .map(|c| if crate::font::is_fullwidth(c) { 1.0 } else { 0.5 })
                            .sum();
                        let denom_em = (total_em - runs[start..i].iter()
                            .flat_map(|r| r.text.chars()).last()
                            .map(|c| if crate::font::is_fullwidth(c) { 1.0 } else { 0.5 })
                            .unwrap_or(0.0)).max(0.5);
                        let per_em_cs = (target_w - natural_w) / denom_em;
                        for run in &mut runs[start..i] {
                            // cs applies per-char in break_into_lines. For a mixed run
                            // (e.g. "平成2"), we need per-char variable cs which isn't
                            // directly representable. Compromise: set cs to the AVERAGE
                            // per-char cs for this run's own em mix.
                            let run_chars: Vec<char> = run.text.chars().collect();
                            if run_chars.is_empty() { continue; }
                            let run_em: f32 = run_chars.iter()
                                .map(|&c| if crate::font::is_fullwidth(c) { 1.0 } else { 0.5 })
                                .sum();
                            let avg_cs = per_em_cs * run_em / run_chars.len() as f32;
                            run.style.character_spacing = Some(avg_cs);
                        }
                        // Split last run: last char carries cs=0 so no trailing advance.
                        let last_idx = i - 1;
                        let last_chars: Vec<char> = runs[last_idx].text.chars().collect();
                        if last_chars.len() > 1 {
                            let body: String = last_chars[..last_chars.len()-1].iter().collect();
                            let tail: String = last_chars[last_chars.len()-1..].iter().collect();
                            runs[last_idx].text = body;
                            // Recompute body's avg_cs based on its own em mix
                            let body_chars: Vec<char> = runs[last_idx].text.chars().collect();
                            let body_em: f32 = body_chars.iter()
                                .map(|&c| if crate::font::is_fullwidth(c) { 1.0 } else { 0.5 })
                                .sum();
                            let body_avg_cs = per_em_cs * body_em / body_chars.len() as f32;
                            runs[last_idx].style.character_spacing = Some(body_avg_cs);
                            let mut tail_run = runs[last_idx].clone();
                            tail_run.text = tail;
                            tail_run.style.character_spacing = Some(0.0);
                            runs.insert(last_idx + 1, tail_run);
                            i += 1;
                        } else if last_chars.len() == 1 {
                            runs[last_idx].style.character_spacing = Some(0.0);
                        }
                    } else {
                        let scale = target_w / natural_w * 100.0;
                        for run in &mut runs[start..i] {
                            run.style.text_scale = Some(scale);
                            run.style.character_spacing = Some(0.0);
                        }
                    }
                }
            } else {
                i += 1;
            }
        }
    }

    #[allow(unused_assignments)]
    fn layout_page(&self, page: &Page) -> Vec<LayoutPage> {
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
        let header_bottom = if !page.header.is_empty() {
            let header_y = page.header_distance.unwrap_or(36.0);
            let mut hdr_h = 0.0_f32;
            for block in &page.header {
                if let Block::Paragraph(para) = block {
                    let fs = para.runs.first()
                        .and_then(|r| r.style.font_size)
                        .unwrap_or(self.default_font_size);
                    let metrics = para.runs.first()
                        .map(|r| self.metrics_for(&r.style, &para.style))
                        .unwrap_or_else(|| self.doc_default_metrics());
                    let lh = metrics.word_line_height(fs, 96.0);
                    hdr_h += lh;
                    hdr_h += para.style.space_after.unwrap_or(0.0);
                }
            }
            header_y + hdr_h
        } else {
            0.0
        };
        let start_y = page.margin.top.max(header_bottom);

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
                            let first_run_style = first_para.runs.first().map(|r| &r.style).unwrap_or(&default_run_style);
                            let fs = first_para.style.ppr_rpr.as_ref()
                                .and_then(|r| r.font_size)
                                .unwrap_or_else(|| self.resolve_font_size(first_run_style, &first_para.style));
                            let metrics = first_para.runs.first()
                                .map(|r| self.metrics_for(&r.style, &first_para.style))
                                .unwrap_or_else(|| {
                                    let rpr_ref = first_para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                    self.metrics_for_para_mark(&rpr_ref, &first_para.style)
                                });
                            // LM0 base line height (Round 9 lookup if available).
                            let lm0_lh = self.registry
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
        let footer_reserved = if !page.footer.is_empty() {
            let footer_dist = page.footer_distance.unwrap_or(36.0);
            let cw = page.size.width - page.margin.left - page.margin.right;
            let gp = page.grid_line_pitch;
            let mut footer_h: f32 = 0.0;
            for block in &page.footer {
                if let Block::Paragraph(p) = block {
                    // Day 33 part 18 (2026-05-10): skip framePr-wrapped paragraphs.
                    // framePr means floating-positioned frame (vAnchor/hAnchor set,
                    // wrap=around) — Word excludes these from inline footer height
                    // because they're positioned independently of inline flow.
                    // COM-confirmed via 3-variant minimal repro (FP_A/B/C with
                    // fs=80pt framePr para → all identical break boundaries).
                    // Affects 备考 cluster: d4d126 (3.6pt over-reservation) plus
                    // 6514, a1d6 candidates.
                    if p.style.frame_pr.is_some() {
                        continue;
                    }
                    // estimate_para_height uses word_line_height_table_cell for
                    // empty paragraphs, which under-estimates footer empty lines.
                    // Override for empty footer paragraphs: use no-grid line height
                    // matching Word's actual footer rendering.
                    let h = if p.runs.is_empty() || p.runs.iter().all(|r| r.text.is_empty()) {
                        let empty_fs = p.style.ppr_rpr.as_ref()
                            .and_then(|r| r.font_size)
                            .unwrap_or_else(|| self.resolve_font_size(&RunStyle::default(), &p.style));
                        let rpr_ref = p.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        let metrics = self.metrics_for_para_mark(&rpr_ref, &p.style);
                        // Day 33 part 11 (2026-05-10): always use natural line height
                        // for empty footer paragraphs. Previous grid-snap (added
                        // 2026-04-17 for 04b88e) over-reserves footer area when
                        // grid pitch > natural line height (bd90b00: 16.5pt pitch,
                        // ~13pt natural → 2.3pt over-reservation pushes 備考 to
                        // page 2). The max-with-bottom-margin guard below ensures
                        // we still reserve at least bottom_margin space.
                        metrics.word_line_height_no_grid(empty_fs)
                    } else {
                        self.estimate_para_height(p, cw, gp, None, false, None, None)
                    };
                    footer_h += h;
                }
            }
            (footer_dist + footer_h).max(page.margin.bottom)
        } else {
            page.margin.bottom
        };
        let content_height = page.size.height - start_y - footer_reserved;
        // Round 29 (2026-04-08): per-page dynamic footnote reservation.
        // Footnotes are reserved at the bottom of the page where their reference
        // appears. The amount reserved varies per page based on which footnotes
        // are referenced. Tracked dynamically as the body layout progresses:
        // when a paragraph contains a footnoteReference, the corresponding note's
        // estimated body height is added to the running reservation; on page
        // break the reservation resets. The body's effective overflow check uses
        // (content_height - footnote_reserved_current_page).
        // Helper to estimate one footnote body height by id.
        let estimate_footnote_h = |id: u32| -> f32 {
            if let Some(note) = page.footnotes.iter().find(|n| n.number == id) {
                let cw = page.size.width - page.margin.left - page.margin.right;
                let mut h: f32 = 0.0;
                let mut first_para = true;
                for nb in &note.blocks {
                    if let Block::Paragraph(p) = nb {
                        if first_para {
                            // Footnote rendering prepends a seq number to the first
                            // paragraph, which increases its width and may add a line.
                            // Clone and add prefix to match actual rendering.
                            let mut p2 = p.clone();
                            let seq = page.footnotes.iter()
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
                            let ph = self.estimate_para_height(&p2, cw, None, None, false, None, None);
                            // 2026-05-05 Track A continuation: removed +2.0pt
                            // per-fn marker overhead. Empirically (b837 spill data
                            // 25 fns) Oxi's est = Word actual + exactly 2.0pt for
                            // every fn — the marker renders inline, no extra
                            // line-height. Over-reservation by 10pt per page (5
                            // fns × 2pt) prevented para 70 from fitting on p5.
                            h += ph;
                            first_para = false;
                        } else {
                            h += self.estimate_para_height(p, cw, None, None, false, None, None);
                        }
                    }
                }
                h
            } else {
                0.0
            }
        };

        // Multi-column layout: compute column X positions and widths
        // COM-confirmed: col_x = margin + Σ(prev_width + prev_spacing)
        // S560 (2026-06-13): factored into a closure so per-section column
        // layouts (page.column_runs, populated when `continuous` section breaks
        // merge sections with DIFFERENT column counts) can be recomputed at
        // each section boundary inside the block loop below.
        let margin_left = page.margin.left;
        let compute_cols = |cols: &Option<crate::ir::ColumnLayout>| -> (usize, Vec<f32>, Vec<f32>) {
            let num = cols.as_ref().map(|c| c.num.max(1) as usize).unwrap_or(1);
            let mut xs: Vec<f32> = Vec::with_capacity(num);
            let mut ws: Vec<f32> = Vec::with_capacity(num);
            if num > 1 {
                if let Some(ref c) = cols {
                    if !c.columns.is_empty() {
                        // Unequal width columns: use explicit definitions
                        let mut x = margin_left;
                        for col_def in &c.columns {
                            xs.push(x);
                            ws.push(col_def.width);
                            x += col_def.width + col_def.space.unwrap_or(0.0);
                        }
                    } else {
                        // Equal width columns
                        let spacing = c.space.unwrap_or(36.0); // default 36pt
                        let col_w = (total_content_width - spacing * (num - 1) as f32) / num as f32;
                        let mut x = margin_left;
                        for _ in 0..num {
                            xs.push(x);
                            ws.push(col_w);
                            x += col_w + spacing;
                        }
                    }
                }
            }
            if xs.is_empty() {
                xs.push(margin_left);
                ws.push(total_content_width);
            }
            (xs.len(), xs, ws)
        };

        let (mut num_columns, mut col_x_positions, mut col_widths) = compute_cols(&page.columns);

        // S560: per-section column runs. Switch per-section ONLY when the
        // merged page has HETEROGENEOUS column counts (e.g. kyotei36spec: a
        // 1-col form table + a continuous 2-col 記載心得 instruction block).
        // When all runs share one column count (the entire 269-doc baseline is
        // num=1), `heterogeneous` is false and the loop never switches → the
        // pre-S560 single-layout path runs byte-identically.
        let col_runs: Vec<(usize, usize, Vec<f32>, Vec<f32>)> = page.column_runs.iter()
            .map(|(start, cols)| {
                let (n, xs, ws) = compute_cols(cols);
                (*start, n, xs, ws)
            })
            .collect();
        let heterogeneous = {
            let mut it = col_runs.iter().map(|r| r.1);
            match it.next() {
                Some(first) => it.any(|n| n != first),
                None => false,
            }
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

        let mut current_column: usize = 0;
        let mut start_x = col_x_positions[0];
        let mut content_width = col_widths[0];
        // S560: lowest column-bottom reached on the current page, so a
        // following column-section flows below ALL columns of the one it
        // succeeds. Only read on the heterogeneous (per-section column) path.
        let mut section_max_y = start_y;
        let mut section_prev_page = 0usize;

        let grid_pitch = page.grid_line_pitch;
        let mut mult_cumul_raw: f32 = 0.0;
        let mut pages: Vec<LayoutPage> = Vec::new();
        let mut elements: Vec<LayoutElement> = Vec::new();
        let mut cursor = LayoutCursor::new(start_y);
        let mut lm2_cells: usize = 0;
        let mut prev_para_style_id: Option<String> = None;
        let mut prev_contextual_spacing: bool = false;
        let mut prev_space_after: f32 = 0.0;
        // Track Y position and layout page index for each block (for paragraph-relative TextBox positioning)
        let mut block_y_positions: Vec<f32> = Vec::with_capacity(page.blocks.len());
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

        // R7.60 (Day 36 part 6, 2026-05-14): track floating-table Y ranges per page
        // for body-position (vertAnchor="page", tblpY > top_margin) full-width tables.
        // When a new floating table would Y-overlap with an already-placed one on the
        // current page, push it to the next page. Word's behavior for 459f05: both
        // floating tables span content width, so they cannot share a page; second
        // table's anchor effectively rolls to next page.
        // Header-position floats (tblpY <= top_margin, e.g. 1ec1/2ea81a) and
        // vertAnchor="text" floats (e.g. 3a4f9f/ed025c) are NOT tracked — they
        // co-locate with body content normally.
        let mut floating_tables_per_page: Vec<Vec<(f32, f32)>> = vec![Vec::new()];

        for (block_idx, block) in page.blocks.iter().enumerate() {
            // S560: on a fresh page the section-bottom tracker resets to the
            // top content origin (the deep value belongs to the prior page).
            if heterogeneous && current_page_idx != section_prev_page {
                section_max_y = start_y;
                section_prev_page = current_page_idx;
            }
            // S560: switch column geometry at a section boundary within a
            // merged continuous-section page (only when heterogeneous, i.e.
            // the page mixes column counts — kyotei36spec's 1-col form table
            // followed by a continuous 2-col 記載心得 block). Word flows the
            // new section continuously below the previous section's content;
            // a 1-col section must NOT inherit the trailing 2-col geometry.
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
                if run_cols != num_columns {
                    // New section continues below ALL columns of the section
                    // it succeeds (continuous flow, same page if room).
                    section_max_y = section_max_y.max(cursor.cursor_y);
                    cursor.set(section_max_y);
                    let run = &col_runs[active_run_idx];
                    num_columns = run.1;
                    col_x_positions = run.2.clone();
                    col_widths = run.3.clone();
                    current_column = 0;
                    start_x = col_x_positions[0];
                    content_width = col_widths[0];
                    section_max_y = cursor.cursor_y;
                }
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
                        if let Some(ref pos) = tb.position {
                            let anchor_y = block_y_positions.get(tb.anchor_block_index).copied().unwrap_or(0.0);
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
            // S469: record the NATURAL (pre-wrap) Y for anchor resolution by
            // subtracting any accumulated wrap-below advance on this page.
            block_y_positions.push(cursor.cursor_y - anchor_flow_offset);
            block_page_indices.push(current_page_idx);
            match block {
                Block::Paragraph(para) => {
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
                                        let sep_extra: f32 = std::env::var("OXI_FN_SEP_GAP_EXTRA")
                                            .ok()
                                            .and_then(|v| v.parse().ok())
                                            .unwrap_or(6.0);
                                        full += 6.0 + sep_extra;
                                        delta += 6.0 + sep_extra;
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
                    let effective_content_h = (content_height - (footnote_reserve_current + delta_if_current)).max(0.0);
                    let effective_content_h_new_page = (content_height - full_if_new).max(0.0);
                    let _ = effective_content_h_new_page;

                    // Helper closure: commit this paragraph's footnotes to the
                    // current page's running reservation (called once we've
                    // decided which page the paragraph lands on).
                    let commit_para_footnotes = |reserve: &mut f32, ids: &mut Vec<u32>, page_i: usize, blk_i: usize| {
                        if page.footnotes.is_empty() { return; }
                        for r in &para.runs {
                            if let Some(id) = r.footnote_ref {
                                if !ids.contains(&id) {
                                    // First footnote: separator line (2pt + 4pt padding).
                                    // S160 env gate OXI_FN_SEP_GAP_EXTRA adds leading
                                    // above separator (Word measurement: ~21pt missing).
                                    if ids.is_empty() {
                                        // S160: see estimate-path comment near line 1934.
                                        // S240 (2026-05-23): removed OXI_LEGACY_FN_SEP_GAP
                                        // legacy env-var fallback during hardening pass.
                                        // OXI_FN_SEP_GAP_EXTRA tuning knob preserved.
                                        let sep_extra: f32 = std::env::var("OXI_FN_SEP_GAP_EXTRA")
                                            .ok()
                                            .and_then(|v| v.parse().ok())
                                            .unwrap_or(6.0);
                                        *reserve += 6.0 + sep_extra;
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
                    let has_lrpb_at_start = para.runs.first()
                        .map(|r| r.has_last_rendered_page_break).unwrap_or(false);
                    let lrpb_should_break = if has_lrpb_at_start && !elements.is_empty() {
                        let est_h = self.estimate_para_height(para, content_width, grid_pitch, None, false, None, None);
                        let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                        let consumed = cursor.cursor_y - start_y;
                        let half_page = effective_content_h * 0.5;
                        est_h <= remaining && consumed > half_page
                    } else {
                        false
                    };

                    // pageBreakBefore: force a new page (not just next column)
                    if (para.style.page_break_before || lrpb_should_break) && !elements.is_empty() {
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        cursor.set(start_y);
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        lm2_cells = 0; current_page_idx += 1;
                        lm2_cells = 0; // Reset cumul line index for new page
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                        commit_para_footnotes(&mut footnote_reserve_current, &mut footnote_ids_current_page, current_page_idx, block_idx);
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
                        let est_h = self.estimate_para_height(para, content_width, grid_pitch, None, false, None, None);
                        let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                        if est_h > remaining && est_h <= effective_content_h {
                            if num_columns > 1 && current_column + 1 < num_columns {
                                current_column += 1;
                                start_x = col_x_positions[current_column];
                                content_width = col_widths[current_column];
                                cursor.set(start_y);
                            } else {
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: std::mem::take(&mut elements),
                                });
                                cursor.set(start_y);
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                                lm2_cells = 0; current_page_idx += 1;
                                // Round 29: page push moves this paragraph (and
                                // its footnote refs) to the new page.
                                footnote_reserve_current = 0.0;
                                footnote_ids_current_page.clear();
                                commit_para_footnotes(&mut footnote_reserve_current, &mut footnote_ids_current_page, current_page_idx, block_idx);
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
                    if para.style.keep_next && !elements.is_empty() {
                        if let Some(Block::Paragraph(next_para)) = page.blocks.get(block_idx + 1) {
                            let this_h = self.estimate_para_height(para, content_width, grid_pitch, None, false, None, None);
                            let next_h = self.estimate_para_height(next_para, content_width, grid_pitch, None, false, None, None);
                            let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                            if this_h + next_h > remaining && this_h > remaining && this_h + next_h <= effective_content_h {
                                if num_columns > 1 && current_column + 1 < num_columns {
                                    current_column += 1;
                                    start_x = col_x_positions[current_column];
                                    content_width = col_widths[current_column];
                                    cursor.set(start_y);
                                } else {
                                    pages.push(LayoutPage {
                                        width: page.size.width,
                                        height: page.size.height,
                                        elements: std::mem::take(&mut elements),
                                    });
                                    cursor.set(start_y);
                                    current_column = 0;
                                    start_x = col_x_positions[0];
                                    content_width = col_widths[0];
                                    lm2_cells = 0; current_page_idx += 1;
                                    footnote_reserve_current = 0.0;
                                    footnote_ids_current_page.clear();
                                    commit_para_footnotes(&mut footnote_reserve_current, &mut footnote_ids_current_page, current_page_idx, block_idx);
                                }
                                *block_page_indices.last_mut().unwrap() = current_page_idx;
                                *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                            }
                        }
                    }

                    // Multi-column pre-check: advance column if paragraph won't fit
                    if num_columns > 1 {
                        let est_h = self.estimate_para_height(para, content_width, grid_pitch, None, false, None, None);
                        let remaining = (start_y + effective_content_h) - cursor.cursor_y;
                        if est_h > remaining && est_h <= effective_content_h {
                            if current_column + 1 < num_columns {
                                current_column += 1;
                                start_x = col_x_positions[current_column];
                                content_width = col_widths[current_column];
                                cursor.set(start_y);
                            } else {
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: std::mem::take(&mut elements),
                                });
                                cursor.set(start_y);
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                                lm2_cells = 0; current_page_idx += 1;
                                footnote_reserve_current = 0.0;
                                footnote_ids_current_page.clear();
                                commit_para_footnotes(&mut footnote_reserve_current, &mut footnote_ids_current_page, current_page_idx, block_idx);
                            }
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        }
                    }

                    let pages_before = pages.len();
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
                    let is_empty_p = |blk: &Block| matches!(blk, Block::Paragraph(p) if p.runs.iter().all(|r| r.text.is_empty()));
                    let is_content_para = |blk: &Block| matches!(blk, Block::Paragraph(p) if p.runs.iter().any(|r| !r.text.is_empty()));
                    let this_empty = para.runs.iter().all(|r| r.text.is_empty());
                    // prev_2_empty: the 2 blocks immediately before are empty AND block_idx-3 is content paragraph
                    let prev_2_empty = block_idx >= 3
                        && is_empty_p(&page.blocks[block_idx - 1])
                        && is_empty_p(&page.blocks[block_idx - 2])
                        && is_content_para(&page.blocks[block_idx - 3]);
                    // next_2_empty: the 2 blocks immediately after are empty AND block_idx+3 is content paragraph
                    let next_2_empty = block_idx + 3 < page.blocks.len()
                        && is_empty_p(&page.blocks[block_idx + 1])
                        && is_empty_p(&page.blocks[block_idx + 2])
                        && is_content_para(&page.blocks[block_idx + 3]);
                    let adjacent_to_empty_run = !this_empty && (prev_2_empty || next_2_empty);

                    // Step 0: bucket for per-page fn refs actually rendered by this
                    // paragraph. Used by the post-layout reserve-correction (Step 1).
                    let mut para_fn_refs_per_page: Vec<Vec<u32>> = Vec::new();
                    let (para_elements, sa) = self.layout_paragraph(
                        para,
                        start_x,
                        &mut cursor,
                        content_width,
                        effective_content_h,
                        start_y,
                        page,
                        &mut pages,
                        &mut elements,
                        grid_pitch,
                        prev_para_style_id.as_deref(), prev_contextual_spacing, false,
                        prev_space_after,
                        Some(block_idx),
                        lm2_param,
                        Some(&mut mult_cumul_raw),
                        adjacent_to_empty_run,
                        Some(&mut para_fn_refs_per_page),
                        // R7.53: first-line lenient — add back this para's
                        // own fn reserve delta so line 0 fits if it would
                        // without the para's footnotes.
                        delta_if_current,
                        // S168: per-fn heights for per-line lenient calculation.
                        &para_fn_heights_map,
                    );
                    prev_space_after = sa;
                    elements.extend(para_elements);
                    if std::env::var("OXI_FN_PROBE").is_ok() && !para_fn_refs_per_page.is_empty() {
                        let any_ref = para_fn_refs_per_page.iter().any(|v| !v.is_empty());
                        if any_ref {
                            eprintln!("[FN_LINE_REFS] block_idx={} per_page={:?}",
                                block_idx, para_fn_refs_per_page);
                        }
                    }

                    // Track page/column breaks that happened inside layout_paragraph
                    let pages_added = pages.len() - pages_before;
                    // Step 1 partial: attribute per-line fn refs to the page each
                    // line actually rendered on. para_fn_refs_per_page[i] holds
                    // the ids on the (start_page + i)-th page. Always runs —
                    // pages_added==0 has 1 bucket == current page's slot.
                    let start_page_for_para = current_page_idx;
                    for (offset, refs) in para_fn_refs_per_page.iter().enumerate() {
                        let page_i = start_page_for_para + offset;
                        while page_fn_refs.len() <= page_i { page_fn_refs.push(Vec::new()); }
                        for id in refs {
                            if !page_fn_refs[page_i].contains(id) {
                                page_fn_refs[page_i].push(*id);
                            }
                        }
                    }
                    if pages_added > 0 {
                        // Multi-column: a "page break" inside layout_paragraph may actually
                        // be a column break. Check if we can advance to the next column.
                        if num_columns > 1 && current_column < num_columns - 1 {
                            // Move to next column instead of creating a new page.
                            // The page was already pushed by layout_paragraph — undo it
                            // by popping and re-merging elements.
                            // Actually, layout_paragraph already pushed the page.
                            // We update column state for subsequent blocks.
                            current_column += 1;
                            start_x = col_x_positions[current_column];
                            content_width = col_widths[current_column];
                            // cursor_y was already reset to start_y by layout_paragraph
                        } else if num_columns > 1 {
                            // All columns exhausted: reset to column 0 for new page
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                        }
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
                        if let Some(new_page_refs) = para_fn_refs_per_page.get(pages_added) {
                            for id in new_page_refs {
                                if !footnote_ids_current_page.contains(id) {
                                    if footnote_ids_current_page.is_empty() {
                                        // S160: see estimate-path comment near line 1934.
                                        // S240 (2026-05-23): removed OXI_LEGACY_FN_SEP_GAP
                                        // legacy env-var fallback during hardening pass.
                                        // OXI_FN_SEP_GAP_EXTRA tuning knob preserved.
                                        let sep_extra: f32 = std::env::var("OXI_FN_SEP_GAP_EXTRA")
                                            .ok()
                                            .and_then(|v| v.parse().ok())
                                            .unwrap_or(6.0);
                                        footnote_reserve_current += 6.0 + sep_extra;
                                    }
                                    footnote_ids_current_page.push(*id);
                                    footnote_reserve_current += estimate_footnote_h(*id);
                                }
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
                                        // S240 (2026-05-23): removed OXI_LEGACY_FN_SEP_GAP
                                        // legacy env-var fallback during hardening pass.
                                        // OXI_FN_SEP_GAP_EXTRA tuning knob preserved.
                                        let sep_extra: f32 = std::env::var("OXI_FN_SEP_GAP_EXTRA")
                                            .ok()
                                            .and_then(|v| v.parse().ok())
                                            .unwrap_or(6.0);
                                        footnote_reserve_current += 6.0 + sep_extra;
                                    }
                                    footnote_ids_current_page.push(*id);
                                    footnote_reserve_current += estimate_footnote_h(*id);
                                }
                            }
                        }
                    }
                    // Round 30: render shapes attached to this paragraph (e.g.
                    // bracketPair preset frame around the date block in
                    // b837808d0555). The shape's anchor reference uses the
                    // paragraph's start Y position; pos.y is the offset from
                    // the paragraph start in points.
                    let para_anchor_y = block_y_positions
                        .get(block_idx)
                        .copied()
                        .unwrap_or(start_y);
                    for shape in &para.shapes {
                        if let Some(ref pos) = shape.position {
                            // h_relative=column: x = margin_left + offset
                            // v_relative=paragraph: y = anchor_y + offset
                            let sx = page.margin.left + pos.x;
                            let sy = para_anchor_y + pos.y;
                            elements.push(LayoutElement::new(
                                sx, sy, shape.width, shape.height,
                                LayoutContent::PresetShape {
                                    shape_type: shape.shape_type.clone(),
                                    stroke_color: shape.stroke_color.clone(),
                                    stroke_width: shape.stroke_width.unwrap_or(0.75), flip_h: shape.flip_h, flip_v: shape.flip_v, arrow_head: shape.arrow_head, arrow_tail: shape.arrow_tail,
                                },
                            ));
                        }
                    }

                    // page_break_after: render the (typically empty) paragraph
                    // on the current page, then force a new page for the NEXT
                    // block. Used for the inline-br-in-empty-paragraph pattern;
                    // see `project_empty_br_para_stub.md`.
                    if para.style.page_break_after && !elements.is_empty() {
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        cursor.set(start_y);
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        current_page_idx += 1;
                        lm2_cells = 0;
                        footnote_reserve_current = 0.0;
                        footnote_ids_current_page.clear();
                    }

                    prev_para_style_id = para.style.style_id.clone();
                    prev_contextual_spacing = para.style.contextual_spacing;
                }
                Block::Table(table) => {
                    // COM-confirmed: prev paragraph's space_after is always added before table
                    cursor.advance(prev_space_after);
                    prev_space_after = 0.0;

                    let is_floating = table.style.position.is_some();
                    let saved_cursor_y = cursor.cursor_y;

                    // Floating table (tblpPr): position relative to anchor
                    let mut candidate_y_top: f32 = 0.0;
                    let mut is_body_floating: bool = false;
                    if let Some(ref pos) = table.style.position {
                        candidate_y_top = match pos.v_anchor.as_deref() {
                            Some("page") => pos.y,
                            Some("margin") => start_y + pos.y,
                            _ => cursor.cursor_y + pos.y, // "text": offset from anchor para bottom
                        };
                        cursor.set(candidate_y_top);
                        // R7.60 body-floating eligibility: vertAnchor="page" AND
                        // table positioned below top margin (in body region).
                        is_body_floating = pos.v_anchor.as_deref() == Some("page")
                            && candidate_y_top > start_y + 0.1;
                    }
                    let pages_before = pages.len();
                    let table_elements = self.layout_table(
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
                    );
                    let candidate_y_bottom = cursor.cursor_y;

                    // R7.60: for body-position vertAnchor=page floating tables,
                    // check overlap with previously-placed floating tables on the
                    // current page. If overlap, push table elements to next page.
                    let mut target_page = current_page_idx;
                    if is_body_floating && is_floating {
                        while floating_tables_per_page
                            .get(target_page)
                            .map_or(false, |ranges| ranges.iter().any(|(t, b)|
                                !(candidate_y_bottom <= *t || candidate_y_top >= *b)
                            ))
                        {
                            target_page += 1;
                        }
                    }

                    if target_page > current_page_idx {
                        // Finalize current page; advance to target page.
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        current_page_idx += 1;
                        while current_page_idx < target_page {
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
                    elements.extend(table_elements);
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
                        let v_anchor_text = table.style.position.as_ref()
                            .map_or(false, |p| p.v_anchor.as_deref() == Some("text"));
                        let pos_x_zero = table.style.position.as_ref()
                            .map_or(false, |p| p.x.abs() < 0.5);
                        let h_anchor_page = table.style.position.as_ref()
                            .map_or(false, |p| p.h_anchor.as_deref() == Some("page"));
                        let wide_table = table_w_pt > content_width - 30.0;
                        let needs_wrap_below = v_anchor_text
                            && wide_table
                            && (pos_x_zero || h_anchor_page);

                        if needs_wrap_below && pages_added > 0 {
                            current_page_idx += pages_added;
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            cursor.set(candidate_y_bottom + 1.5);
                            *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                        } else if needs_wrap_below {
                            // S469: the cursor advances below the table so body
                            // TEXT wraps under it, but objects anchored to the
                            // following paragraph keep the natural (pre-wrap) Y.
                            // Record the advance so block_y_positions can undo it.
                            if s469_enabled {
                                anchor_flow_offset += (candidate_y_bottom + 1.5) - saved_cursor_y;
                            }
                            cursor.set(candidate_y_bottom + 1.5);
                        } else {
                            // Original behavior: floating tables don't advance text flow
                            cursor.set(saved_cursor_y);
                        }
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
                    prev_space_after = 0.0;
                }
                Block::Image(img) => {
                    // S549 (2026-06-12, opt-out OXI_S549_DISABLE): in a docGrid
                    // lines section the image-only paragraph's line occupies a
                    // WHOLE number of grid cells — ceil(extent/pitch)×pitch.
                    // COM repro (_s549_img_grid.py, pitch 18): extent 185→198
                    // (11 cells), 100→108, 90→90, 36→36 (exact multiples pass
                    // through); docGrid none → extent EXACTLY (S537 model
                    // unchanged). Live 3a4f "/" figures (extent 185 → Word
                    // block 198) were leaving every downstream para 13pt high
                    // → the last 3 Phase-1 delta=-1 boundary paras.
                    let img_adv = match page.grid_line_pitch {
                        Some(p) if p > 0.1 && std::env::var("OXI_S549_DISABLE").is_err() => {
                            (img.height / p).ceil() * p
                        }
                        _ => img.height,
                    };
                    if cursor.cursor_y + img_adv > start_y + content_height {
                        if num_columns > 1 && current_column + 1 < num_columns {
                            current_column += 1;
                            start_x = col_x_positions[current_column];
                            content_width = col_widths[current_column];
                            cursor.set(start_y);
                        } else {
                            pages.push(LayoutPage {
                                width: page.size.width,
                                height: page.size.height,
                                elements: std::mem::take(&mut elements),
                            });
                            cursor.set(start_y);
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            lm2_cells = 0; current_page_idx += 1;
                        }
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        *block_y_positions.last_mut().unwrap() = cursor.cursor_y;
                    }
                    elements.push(LayoutElement::new(start_x, cursor.visual_y, img.width, img.height, LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                    }));
                    cursor.advance(img_adv);
                    prev_para_style_id = None;
                }
                Block::UnsupportedElement(_) => {
                    // Skip unsupported elements in layout
                }
                Block::Math(math_block) => {
                    // Phase 3: emit positioned LayoutElements for math primitives.
                    // Fraction/Sup/Sub/SubSup render stacked; other primitives
                    // fall back to flat text for now.
                    let math_font_size: f32 = 10.5;
                    // S524 (coverage, 2026-06-09): apply the display equation's jc
                    // (default Center) — Word CENTERS display math (oMathPara) at the
                    // page center; Oxi previously hard-coded the left margin. PDF-confirmed
                    // on a/b, x^2, x_i, sqrt(x), nested (all Word-centered at page mid).
                    // Compute the bbox width first, then position by jc.
                    let content_w = (page.size.width - page.margin.left - page.margin.right).max(0.0);
                    let bbox_pre = crate::layout::math::layout_math_block(math_block, math_font_size);
                    let math_jc = match math_block {
                        crate::ir::MathBlock::Display { jc, .. } => *jc,
                        _ => crate::ir::MathAlignment::Left,
                    };
                    let x = match math_jc {
                        crate::ir::MathAlignment::Center | crate::ir::MathAlignment::CenterGroup =>
                            page.margin.left + ((content_w - bbox_pre.advance) * 0.5).max(0.0),
                        crate::ir::MathAlignment::Right =>
                            page.margin.left + (content_w - bbox_pre.advance).max(0.0),
                        crate::ir::MathAlignment::Left => page.margin.left,
                    };
                    let (math_elems, bbox) = crate::layout::math::emit_math_block(
                        math_block, x, cursor.cursor_y, math_font_size,
                    );
                    if !math_elems.is_empty() {
                        elements.extend(math_elems);
                        // Advance by full bbox height + a line of descent leeway.
                        let advance = bbox.height().max(math_font_size * 1.2)
                            + math_font_size * 0.3;
                        cursor.advance(advance);
                    }
                }
            }
            // S560: record the deepest column-bottom reached on this page so a
            // following column-section (heterogeneous path) starts below it.
            if heterogeneous {
                section_max_y = section_max_y.max(cursor.cursor_y);
            }
        }

        // Final page
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
                let (keep, overflow): (Vec<_>, Vec<_>) = take.into_iter()
                    .partition(|e| !e.vmerge_restart_overflow_to_next_page);
                pages[i].elements = keep;
                if overflow.is_empty() { continue; }
                let next_idx = i + 1;
                if next_idx >= pages.len() {
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: Vec::new(),
                    });
                }
                let shift = page_bottom_sweep - page_top;
                for mut elem in overflow {
                    elem.y -= shift;
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
        // KNOWN UNTESTED EDGES (zero instances in the current corpus, so not
        // handled here — confirmed by the S478 13-doc structural audit):
        //   (a) behind_doc=1 text boxes: a true behindDoc object should render
        //       BEHIND the body-text layer; this sort only orders floats among
        //       themselves (all text boxes are emitted after body text here).
        //       If a behindDoc watermark with a high relativeHeight ever appears,
        //       stratify (emit behind_doc=1 first, then behind_doc=0, each
        //       relHeight-sorted) — or emit it before the body loop for true
        //       behind-text. No such object exists in the corpus to verify against.
        //   (b) standalone text-bearing VML (relativeHeight=0): all corpus VML is
        //       an mc:Fallback alternate (never rendered — Oxi takes mc:Choice
        //       DrawingML, which always carries a relativeHeight). A future
        //       standalone VML text box would parse relHeight=0 and be forced to
        //       the bottom of every overlap; give it a doc-order-derived key then.
        let s478_zorder = std::env::var("OXI_S478_DISABLE").is_err();
        let tb_order: Vec<usize> = {
            let mut idx: Vec<usize> = (0..page.text_boxes.len()).collect();
            if s478_zorder {
                idx.sort_by_key(|&i| page.text_boxes[i].relative_height);
            }
            idx
        };
        for &tbi in &tb_order {
            let text_box = &page.text_boxes[tbi];
            let target_page = block_page_indices
                .get(text_box.anchor_block_index)
                .copied()
                .unwrap_or(0);
            // S535: env-gated anchor-resolution tracing (figure-collage overlap).
            if std::env::var("OXI_DEBUG_TB").is_ok() {
                let (rx, ry) = self.resolve_textbox_position(text_box, page, &block_y_positions);
                let anchor_in_range = text_box.anchor_block_index < block_y_positions.len();
                let preview: String = text_box.blocks.iter().filter_map(|b| match b {
                    Block::Paragraph(p) => Some(p.runs.iter().flat_map(|r| r.text.chars()).take(8).collect::<String>()),
                    _ => None,
                }).find(|s| !s.is_empty()).unwrap_or_default();
                eprintln!("[TB] tbi={} anchor={} in_range={} tgt_page={} resolved=({:.1},{:.1}) wh=({:.0},{:.0}) vrel={:?} text={:?}",
                    tbi, text_box.anchor_block_index, anchor_in_range, target_page, rx, ry,
                    text_box.width, text_box.height,
                    text_box.position.as_ref().and_then(|p| p.v_relative.clone()),
                    preview);
            }
            let tb_elements = self.layout_text_box(text_box, page, &block_y_positions);
            if let Some(lp) = pages.get_mut(target_page) {
                lp.elements.extend(tb_elements);
            }
        }

        // Layout floating images and add to the correct layout page
        for img in &page.floating_images {
            if let Some(ref _pos) = img.position {
                let (abs_x, abs_y) = self.resolve_floating_image_position(img, page, &block_y_positions);
                // Use the same page as the anchor block
                let target_page = block_page_indices
                    .get(img.anchor_block_index)
                    .copied()
                    .unwrap_or(0);
                let el = LayoutElement::new(abs_x, abs_y, img.width, img.height, LayoutContent::Image {
                        data: img.data.clone(),
                        content_type: img.content_type.clone(),
                });
                if let Some(lp) = pages.get_mut(target_page) {
                    lp.elements.push(el);
                } else if let Some(lp) = pages.last_mut() {
                    lp.elements.push(el);
                }
            } else {
                // No position info — treat as inline at end of last page
                if let Some(lp) = pages.last_mut() {
                    lp.elements.push(LayoutElement::new(start_x, 0.0, img.width, img.height, LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                    }));
                }
            }
        }

        // Layout header/footer on each layout page
        // Header y = headerDistance (from page top edge), default 36pt (0.5in)
        // Footer y = pageHeight - footerDistance - footerContentHeight
        let header_y = page.header_distance.unwrap_or(36.0);
        let footer_dist = page.footer_distance.unwrap_or(36.0);
        let hdr_x = page.margin.left;
        let hdr_width = content_width;
        for (page_idx, lp) in pages.iter_mut().enumerate() {
            if !page.header.is_empty() {
                let mut cy = LayoutCursor::new(header_y);
                for block in &page.header {
                    if let Block::Paragraph(para) = block {
                        let empty_fn_h_hdr = std::collections::HashMap::new();
                        let (hdr_elements, _) = self.layout_paragraph(
                            para, hdr_x, &mut cy, hdr_width, page.size.height,
                            header_y, page, &mut Vec::new(), &mut Vec::new(),
                            grid_pitch, None, false,
                            false, 0.0, None, None, None,
                            false, None,
                            0.0,
                            &empty_fn_h_hdr,
                        );
                        lp.elements.extend(hdr_elements);
                    }
                }
            }
            if !page.footer.is_empty() {
                // Estimate footer content height first.
                // Day 33 part 18: skip framePr paragraphs (floating frames) —
                // they're positioned independently of inline flow, so they
                // should not shift footer_top.
                let mut footer_h: f32 = 0.0;
                for block in &page.footer {
                    if let Block::Paragraph(para) = block {
                        if para.style.frame_pr.is_some() { continue; }
                        footer_h += self.estimate_para_height(para, hdr_width, grid_pitch, None, false, None, None);
                    }
                }
                let footer_top = page.size.height - footer_dist - footer_h;
                let mut cy = LayoutCursor::new(footer_top);
                for block in &page.footer {
                    if let Block::Paragraph(para) = block {
                        let empty_fn_h_ftr = std::collections::HashMap::new();
                        let (ftr_elements, _) = self.layout_paragraph(
                            para, hdr_x, &mut cy, hdr_width, page.size.height,
                            footer_top, page, &mut Vec::new(), &mut Vec::new(),
                            grid_pitch, None, false,
                            false, 0.0, None, None, None,
                            false, None,
                            0.0,
                            &empty_fn_h_ftr,
                        );
                        lp.elements.extend(ftr_elements);
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
                                        if !out.contains(&id) { out.push(id); }
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
                        if !referenced_ids.contains(id) { referenced_ids.push(*id); }
                    }
                }
                for (i, b) in page.blocks.iter().enumerate() {
                    if block_page_indices.get(i).copied().unwrap_or(0) == page_idx {
                        if let Block::Table(_) = b {
                            collect_footnote_refs(std::slice::from_ref(b), &mut referenced_ids);
                        }
                    }
                }

                if !referenced_ids.is_empty() {
                    // Resolve referenced footnotes (preserve order, dedup already done).
                    let notes: Vec<&Footnote> = referenced_ids.iter()
                        .filter_map(|id| page.footnotes.iter().find(|n| n.number == *id))
                        .collect();

                    if !notes.is_empty() {
                        // Footnote area bottom: just above the footer, or at the
                        // bottom margin if no footer is present.
                        let footnote_bottom = if !page.footer.is_empty() {
                            // Recompute footer top here (mirrors lines 832-839 above).
                            // Day 33 part 18: skip framePr paragraphs (floating frames).
                            let mut footer_h: f32 = 0.0;
                            for block in &page.footer {
                                if let Block::Paragraph(para) = block {
                                    if para.style.frame_pr.is_some() { continue; }
                                    footer_h += self.estimate_para_height(para, hdr_width, grid_pitch, None, false, None, None);
                                }
                            }
                            page.size.height - footer_dist - footer_h - 4.0
                        } else {
                            page.size.height - page.margin.bottom
                        };

                        // Find the last body element Y on this page to avoid overlap
                        let body_bottom_y = lp.elements.iter()
                            .map(|e| e.y + e.height)
                            .fold(0.0_f32, f32::max);

                        // Calculate footnote heights — grid-snap per paragraph to
                        // match actual render (layout_paragraph stacks lines at
                        // grid pitch when grid_pitch is Some). estimate_para_height
                        // returns natural height which diverges from the render.
                        // COM-derived 2026-04-20 from 6 minimal repros: render uses
                        // grid_pitch × line_count per paragraph.
                        let grid_snap_para = |p: &Paragraph| -> (f32, usize) {
                            // Natural estimated height (may include space_before/after)
                            let nat = self.estimate_para_height(p, hdr_width, grid_pitch, None, false, None, None);
                            // Per-line natural height (used to derive line_count)
                            let line_fs = self.resolve_font_size(
                                p.runs.first().map(|r| &r.style).unwrap_or(&RunStyle::default()),
                                &p.style,
                            );
                            let metrics = p.runs.first()
                                .map(|r| self.metrics_for_text(&r.text, &r.style, &p.style))
                                .unwrap_or_else(|| {
                                    let rpr = p.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                    self.metrics_for_para_mark(&rpr, &p.style)
                                });
                            let line_nat = metrics.word_line_height_no_grid(line_fs).max(0.01);
                            let line_count = ((nat / line_nat).round() as usize).max(1);
                            // Only grid-snap if the paragraph opts in (snapToGrid default=true).
                            // b837's FootnoteText style has snapToGrid=0 → Word uses natural.
                            let height = if let Some(pitch) = grid_pitch {
                                if pitch > 0.0 && p.style.snap_to_grid {
                                    line_count as f32 * pitch
                                } else {
                                    line_count as f32 * line_nat
                                }
                            } else { line_count as f32 * line_nat };
                            (height, line_count)
                        };
                        let mut note_heights: Vec<f32> = Vec::new();
                        for note in &notes {
                            let mut nh: f32 = 0.0;
                            for nb in &note.blocks {
                                if let Block::Paragraph(p) = nb {
                                    let (h, _) = grid_snap_para(p);
                                    nh += h;
                                }
                            }
                            note_heights.push(nh);
                        }
                        // Word anchors the LAST footnote line's BOTTOM to
                        // page_h - margin.bottom (derived 2026-04-20 from 6 minimal
                        // repros). Only INNER lines stack at grid pitch; the last
                        // line's height is natural (word_line_height_no_grid).
                        // Subtract (grid_pitch - natural_last) once to compensate.
                        let last_line_adjust: f32 = if let (Some(pitch), Some(last_note)) = (grid_pitch, notes.last()) {
                            if pitch > 0.0 {
                                if let Some(Block::Paragraph(last_p)) = last_note.blocks.last() {
                                    // Only applicable when the last paragraph grid-snaps.
                                    // snapToGrid=0 footnote lines are already natural-height.
                                    if !last_p.style.snap_to_grid {
                                        0.0
                                    } else {
                                        let text_run = last_p.runs.iter().rev().find(|r| !r.text.is_empty())
                                            .or_else(|| last_p.runs.first());
                                        let fs = self.resolve_font_size(
                                            text_run.map(|r| &r.style).unwrap_or(&RunStyle::default()),
                                            &last_p.style,
                                        );
                                        let metrics = text_run
                                            .map(|r| self.metrics_for_text(&r.text, &r.style, &last_p.style))
                                            .unwrap_or_else(|| {
                                                let rpr = last_p.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                                self.metrics_for_para_mark(&rpr, &last_p.style)
                                            });
                                        let natural_last = metrics.word_line_height_no_grid(fs);
                                        // Word centers the last footnote line in the grid-pitch
                                        // overshoot: adjustment = (pitch - natural) / 2 — confirmed
                                        // 2026-04-20 via 6 minimal repros (fn top at page_h -
                                        // margin.bottom - 14.85pt = pitch - 3.15 for 9pt MS Mincho).
                                        ((pitch - natural_last) * 0.5).max(0.0)
                                    }
                                } else { 0.0 }
                            } else { 0.0 }
                        } else { 0.0 };

                        let separator_h_pre: f32 = 2.0;
                        let separator_pad_pre: f32 = 4.0;
                        // Determine how many notes fit: add notes one by one from the
                        // bottom; stop when area_top would overlap body content.
                        let mut total_h: f32 = 0.0;
                        let mut fit_count = notes.len();
                        for i in 0..notes.len() {
                            let candidate = total_h + note_heights[i] + separator_h_pre + separator_pad_pre;
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
                        let area_top = footnote_bottom - total_h - separator_pad - separator_h + last_line_adjust;

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
                            let seq = page.footnotes.iter()
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
                                                first_run.text = format!("{}{}", prefix, first_run.text);
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
                                    let footnote_width = page.size.width - page.margin.left - page.margin.right;
                                    let empty_fn_h_note = std::collections::HashMap::new();
                                    let (note_elements, _) = self.layout_paragraph(
                                        &para_to_render, page.margin.left, &mut cy, footnote_width, footnote_page_height_huge,
                                        footnote_page_top, page,
                                        &mut Vec::new(), &mut Vec::new(),
                                        grid_pitch, None, false,
                                        false, 0.0, None, None, None,
                                        false, None,
                                        0.0,
                                        &empty_fn_h_note,
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
                    let anchor_y = block_y_positions.get(shape.anchor_block_index).copied().unwrap_or(start_y);
                    let anchor_page = block_page_indices.get(shape.anchor_block_index).copied().unwrap_or(0);

                    // Only render on the correct page
                    if anchor_page == page_idx {
                        // h_relative="column": x = margin_left + offset
                        // v_relative="paragraph": y = anchor_paragraph_y + offset
                        let sx = start_x + pos.x;
                        let sy = anchor_y + pos.y;
                        lp.elements.push(LayoutElement::new(sx, sy, shape.width, shape.height, LayoutContent::PresetShape {
                                shape_type: shape.shape_type.clone(),
                                stroke_color: shape.stroke_color.clone(),
                                stroke_width: shape.stroke_width.unwrap_or(0.75), flip_h: shape.flip_h, flip_v: shape.flip_v, arrow_head: shape.arrow_head, arrow_tail: shape.arrow_tail,
                        }));
                    }
                }
            }
        }

        pages
    }

    /// Resolve absolute (x, y) position for a text box based on its anchor references.
    fn resolve_textbox_position(&self, text_box: &TextBox, page: &Page, block_y_positions: &[f32]) -> (f32, f32) {
        let pos = match &text_box.position {
            Some(p) => p,
            None => return (page.margin.left, page.margin.top),
        };

        let content_width = page.size.width - page.margin.left - page.margin.right;

        // Horizontal: alignment takes precedence over offset
        let abs_x = if let Some(ref align) = pos.h_align {
            let ref_left;
            let ref_width;
            match pos.h_relative.as_deref() {
                Some("page") => { ref_left = 0.0; ref_width = page.size.width; }
                Some("margin") | Some("column") | _ => { ref_left = page.margin.left; ref_width = content_width; }
            }
            match align.as_str() {
                "left" => ref_left,
                "center" => ref_left + (ref_width - text_box.width) / 2.0,
                "right" => ref_left + ref_width - text_box.width,
                _ => ref_left,
            }
        } else {
            match pos.h_relative.as_deref() {
                Some("page") => pos.x,
                Some("margin") | Some("column") | Some("character") => page.margin.left + pos.x,
                Some("leftMarginArea") => pos.x,
                Some("rightMarginArea") => (page.size.width - page.margin.right) + pos.x,
                _ => page.margin.left + pos.x,
            }
        };

        // Vertical: paragraph-relative uses anchor block Y position
        let abs_y = if let Some(ref align) = pos.v_align {
            let ref_top;
            let ref_height;
            match pos.v_relative.as_deref() {
                Some("page") => { ref_top = 0.0; ref_height = page.size.height; }
                Some("margin") | _ => { ref_top = page.margin.top; ref_height = page.size.height - page.margin.top - page.margin.bottom; }
            }
            match align.as_str() {
                "top" => ref_top,
                "center" => ref_top + (ref_height - text_box.height) / 2.0,
                "bottom" => ref_top + ref_height - text_box.height,
                _ => ref_top,
            }
        } else {
            match pos.v_relative.as_deref() {
                Some("page") => pos.y,
                Some("paragraph") | Some("line") => {
                    let anchor_y = block_y_positions
                        .get(text_box.anchor_block_index)
                        .copied()
                        .unwrap_or(page.margin.top);
                    anchor_y + pos.y
                }
                Some("margin") => page.margin.top + pos.y,
                Some("topMarginArea") => pos.y,
                Some("bottomMarginArea") => (page.size.height - page.margin.bottom) + pos.y,
                _ => page.margin.top + pos.y,
            }
        };

        // Clamp TextBox to page boundaries (prevent overflow beyond page edge)
        let abs_y = if abs_y + text_box.height > page.size.height {
            (page.size.height - text_box.height).max(0.0)
        } else {
            abs_y
        };
        let abs_x = if abs_x + text_box.width > page.size.width {
            (page.size.width - text_box.width).max(0.0)
        } else {
            abs_x
        };

        (abs_x, abs_y)
    }

    /// Resolve absolute (x, y) position for a floating image.
    fn resolve_floating_image_position(&self, img: &Image, page: &Page, block_y_positions: &[f32]) -> (f32, f32) {
        let pos = match &img.position {
            Some(p) => p,
            None => return (page.margin.left, page.margin.top),
        };

        let content_width = page.size.width - page.margin.left - page.margin.right;

        let abs_x = if let Some(ref align) = pos.h_align {
            let (ref_left, ref_width) = match pos.h_relative.as_deref() {
                Some("page") => (0.0, page.size.width),
                Some("leftMargin") => (0.0, page.margin.left),
                Some("rightMargin") => (page.size.width - page.margin.right, page.margin.right),
                Some("margin") | Some("column") | _ => (page.margin.left, content_width),
            };
            match align.as_str() {
                "left" => ref_left,
                "center" => ref_left + (ref_width - img.width) / 2.0,
                "right" => ref_left + ref_width - img.width,
                _ => ref_left,
            }
        } else {
            match pos.h_relative.as_deref() {
                Some("page") => pos.x,
                Some("margin") | Some("column") => page.margin.left + pos.x,
                Some("leftMargin") | Some("leftMarginArea") => pos.x,
                Some("rightMargin") | Some("rightMarginArea") => (page.size.width - page.margin.right) + pos.x,
                _ => page.margin.left + pos.x,
            }
        };

        let abs_y = if let Some(ref align) = pos.v_align {
            let (ref_top, ref_height) = match pos.v_relative.as_deref() {
                Some("page") => (0.0, page.size.height),
                _ => (page.margin.top, page.size.height - page.margin.top - page.margin.bottom),
            };
            match align.as_str() {
                "top" => ref_top,
                "center" => ref_top + (ref_height - img.height) / 2.0,
                "bottom" => ref_top + ref_height - img.height,
                _ => ref_top,
            }
        } else {
            match pos.v_relative.as_deref() {
                Some("page") => pos.y,
                Some("paragraph") | Some("line") => {
                    let anchor_y = block_y_positions
                        .get(img.anchor_block_index)
                        .copied()
                        .unwrap_or(page.margin.top);
                    anchor_y + pos.y
                }
                Some("margin") => page.margin.top + pos.y,
                _ => page.margin.top + pos.y,
            }
        };

        // Clamp to page boundaries (floating images can extend into margins)
        let abs_y_clamped = if abs_y + img.height > page.size.height {
            (page.size.height - img.height).max(0.0)
        } else {
            abs_y
        };

        (abs_x, abs_y_clamped)
    }

    /// Layout a single text box: background, borders, and inner content.
    fn layout_text_box(&self, text_box: &TextBox, page: &Page, block_y_positions: &[f32]) -> Vec<LayoutElement> {
        self.layout_text_box_at(text_box, page, block_y_positions, None)
    }

    /// S487/S488 (CLASS E step 2/3): layout a text box, optionally with an
    /// explicit ALREADY-RESOLVED absolute top-left for IN-CELL text boxes. A
    /// cell text box's anchor references (relH="column" → cell content-left,
    /// relV="paragraph" → the anchoring paragraph's top) cannot be resolved by
    /// resolve_textbox_position (which only knows body geometry), so the cell
    /// render loop resolves the absolute top-left itself (S488 anchor model) and
    /// passes it as `resolved_origin`. When Some((ax, ay)), that IS the box
    /// top-left; when None, fall back to the body resolver.
    fn layout_text_box_at(&self, text_box: &TextBox, page: &Page, block_y_positions: &[f32], resolved_origin: Option<(f32, f32)>) -> Vec<LayoutElement> {
        let mut elements = Vec::new();

        // 1. Calculate absolute position
        let (abs_x, abs_y) = match resolved_origin {
            Some((ax, ay)) => (ax, ay),
            None => self.resolve_textbox_position(text_box, page, block_y_positions),
        };

        // 2. Background fill + border as a single BoxRect (supports corner radius)
        let has_fill = text_box.fill.is_some();
        let has_border = text_box.border;
        if has_fill || has_border {
            let fill_hex = text_box.fill.as_ref().map(|f| {
                if f.starts_with('#') { f.clone() } else { format!("#{}", f) }
            });
            let cr = text_box.corner_radius.unwrap_or(0.0);
            elements.push(LayoutElement::new(abs_x, abs_y, text_box.width, text_box.height, LayoutContent::BoxRect {
                    fill: fill_hex,
                    stroke_color: if has_border {
                        text_box.stroke_color.as_ref()
                            .map(|c| if c.starts_with('#') { c.clone() } else { format!("#{}", c) })
                            .or_else(|| Some("#000000".to_string()))
                    } else { None },
                    stroke_width: if has_border { text_box.stroke_width.unwrap_or(1.0) } else { 0.0 },
                    corner_radius: cr,
            }));
        }

        // 3. Clip region — all TextBox content is clipped to the box boundary
        elements.push(LayoutElement::new(abs_x, abs_y, text_box.width, text_box.height, LayoutContent::ClipStart));

        // 4. Content layout within text box
        // Word default inset: L/R = 7.2pt (0.1in = 91440 EMU), T/B = 3.6pt (0.05in = 45720 EMU)
        let inset_l = text_box.inset_left.unwrap_or(7.2);
        let inset_r = text_box.inset_right.unwrap_or(7.2);
        let inset_t = text_box.inset_top.unwrap_or(3.6);
        let inset_b = text_box.inset_bottom.unwrap_or(3.6);
        // roundRect corner pushes the inscribed text area inward by r·(1−cos 45°) ≈ 0.293r per side.
        let corner_inset = text_box.corner_radius
            .map(|r| r * (1.0 - std::f32::consts::FRAC_1_SQRT_2))
            .unwrap_or(0.0);
        let inner_x = abs_x + inset_l + corner_inset;
        let inner_width = (text_box.width - inset_l - inset_r - 2.0 * corner_inset).max(0.0);
        let inner_height = (text_box.height - inset_t - inset_b).max(0.0);
        // S481 (REVERTED, finding only): rendering vertOverflow="overflow" lines
        // (e.g. 2ea81a "＜＜記載例＞＞", dropped via avail=0) REGRESSED p2 −0.0009 —
        // the box's anchor Y is itself mis-positioned (S478/S469 float-anchor
        // family), so rendering the text lands it in the wrong place. The
        // overflow render needs the box POSITION fixed first. vert_overflow IS
        // parsed into the IR (correct data) for that future fix; not acted on here.
        // v-text-anchor: middle/bottom shifts content within textbox.
        // Initial cursor at top; for middle/bottom, compute content height first,
        // then offset all elements after layout.
        let v_anchor = text_box.v_text_anchor.as_deref().unwrap_or("t");
        let mut cursor = LayoutCursor::new(abs_y + inset_t);

        // We layout content inside the text box without page-breaking.
        // Use dummy page/elements vecs since we don't want page breaks inside text boxes.
        let mut dummy_pages: Vec<LayoutPage> = Vec::new();
        let mut dummy_elements: Vec<LayoutElement> = Vec::new();

        for block in &text_box.blocks {
            // Stop if we've exceeded the text box bounds
            if cursor.cursor_y > abs_y + text_box.height - inset_b {
                break;
            }

            match block {
                Block::Paragraph(para) => {
                    let clip_bottom = abs_y + text_box.height;
                    // Capture para start Y before layout_paragraph advances cursor_y.
                    // Used below to anchor inner-paragraph shapes at their declared offset.
                    let para_start_y = cursor.cursor_y;
                    let empty_fn_h_txbx = std::collections::HashMap::new();
                    let (para_elements, _) = self.layout_paragraph(
                        para,
                        inner_x,
                        &mut cursor,
                        inner_width,
                        inner_height,
                        abs_y + inset_t,
                        page,
                        &mut dummy_pages,
                        &mut dummy_elements,
                        // TextBox grid snap: enabled for "lines" grid, disabled for "linesAndChars"
                        if page.grid_char_pitch.is_some() { None } else { page.grid_line_pitch },
                        None, false, // no prev style/contextual tracking
                        true, // in_textbox: suppress CJK compression
                        0.0, None, None, None,
                        false, None,
                        0.0,
                        &empty_fn_h_txbx,
                    );
                    // Emit PresetShape elements for shapes attached to this inner
                    // paragraph. Without this, floating shapes (e.g. DML:line
                    // dividers, brackets) declared inside <w:txbxContent> never
                    // render. The shape anchor uses the inner paragraph's start Y
                    // plus the shape's own pos.y (relative offset from paragraph).
                    // Found via 3a4f9fbe1a83 baseline scan 2026-04-23.
                    for shape in &para.shapes {
                        if let Some(ref pos) = shape.position {
                            let sx = inner_x + pos.x;
                            let sy = para_start_y + pos.y;
                            elements.push(LayoutElement::new(
                                sx, sy, shape.width, shape.height,
                                LayoutContent::PresetShape {
                                    shape_type: shape.shape_type.clone(),
                                    stroke_color: shape.stroke_color.clone(),
                                    stroke_width: shape.stroke_width.unwrap_or(0.75), flip_h: shape.flip_h, flip_v: shape.flip_v, arrow_head: shape.arrow_head, arrow_tail: shape.arrow_tail,
                                },
                            ));
                        }
                    }
                    // Word behavior: TextBox overflow text is not rendered.
                    // Filter: (1) Y overflow, (2) in dark-filled TextBox, skip text with no explicit color.
                    // Word PDF omits runs without color attribute inside colored TextBoxes —
                    // these are overflow text that would be black-on-dark and shouldn't be visible.
                    // Only apply to dark fills (not white/light backgrounds where black text is normal).
                    let has_dark_fill = text_box.fill.as_ref().map_or(false, |f| {
                        let hex = f.trim_start_matches('#');
                        if hex.len() >= 6 {
                            let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(255);
                            let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(255);
                            let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(255);
                            (r as u16 + g as u16 + b as u16) < 600
                        } else {
                            false
                        }
                    });
                    // Line-count-aware overflow cutoff (2026-04-25):
                    // For tight-fit single-line textboxes (459f05 様式１: textbox height
                    // = inset+line_height exactly), the OLD `pe.y + pe.height > clip_bottom`
                    // filter dropped the only line because line slot extends past clip.
                    // Compute available lines = floor(inner_height / line_height) and
                    // drop only elements whose line index >= available (i.e., y >= line_avail_top).
                    // For multi-line overflow (2ea81a tbx 5: 3 lines, only 2 fit), this
                    // correctly drops the 3rd line while keeping the first 2.
                    // line_height is taken from the first text element in para_elements.
                    let inner_h = (text_box.height - inset_t - inset_b).max(0.0);
                    let line_h = para_elements.iter()
                        .find_map(|pe| match &pe.content {
                            LayoutContent::Text { .. } if pe.height > 0.5 => Some(pe.height),
                            _ => None,
                        })
                        .unwrap_or(0.0);
                    let line_cutoff_y = if line_h > 0.5 && inner_h > 0.0 {
                        let avail = (inner_h / line_h).floor();
                        abs_y + inset_t + avail * line_h
                    } else {
                        clip_bottom
                    };
                    let accept_and_fix_color = |pe: &mut LayoutElement| -> bool {
                        // Drop elements whose Y is past the last-line-allowed cutoff.
                        // For non-text elements (BoxRect inside textbox), use traditional
                        // bounds check.
                        match &pe.content {
                            LayoutContent::Text { .. } => {
                                if pe.y >= line_cutoff_y { return false; }
                            }
                            _ => {
                                if pe.y + pe.height > clip_bottom { return false; }
                            }
                        }
                        // Word omits runs that have no explicit color attribute when
                        // rendered inside a dark-filled shape: these are overflow/invisible
                        // glyphs that would otherwise be black-on-dark. Verified on 1ec1
                        // P2 heading box (4472C4 fill, runs 8-9 have no color attribute and
                        // are not visible in Word's rendering despite line layout including them).
                        if has_dark_fill {
                            if let LayoutContent::Text { ref color, .. } = pe.content {
                                if color.is_none() {
                                    return false;
                                }
                            }
                        }
                        true
                    };
                    for mut pe in para_elements {
                        if accept_and_fix_color(&mut pe) { elements.push(pe); }
                    }
                    for mut de in dummy_elements.drain(..) {
                        if accept_and_fix_color(&mut de) { elements.push(de); }
                    }
                }
                Block::Table(table) => {
                    // TextBox tables don't paginate — use large content_height
                    let mut tb_pages = Vec::new();
                    let mut tb_elems = Vec::new();
                    let table_elements = self.layout_table(
                        table,
                        inner_x,
                        &mut cursor,
                        inner_width,
                        None,
                        None,
                        None,
                        0.0, 99999.0, 0.0, 99999.0,
                        &mut tb_pages, &mut tb_elems,
                        None,
                        page,
                        false,
                    );
                    elements.extend(tb_elems);
                    elements.extend(table_elements);
                }
                Block::Image(img) => {
                    elements.push(LayoutElement::new(inner_x, cursor.visual_y, img.width.min(inner_width), img.height, LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                    }));
                    cursor.advance(img.height);
                }
                Block::UnsupportedElement(_) => {}
                Block::Math(_) => {
                    // Phase 2 stub: textbox-embedded OMML math not yet rendered.
                }
            }
        }

        // Use specified height (no autoFit by default in Word).
        // Only shrink if content is smaller AND autoFit is explicitly enabled.
        let actual_height = text_box.height;

        // v-text-anchor: shift content elements vertically for middle/bottom alignment.
        // "ctr" (DrawingML) or "middle" (VML) = vertically centered within textbox.
        // Use actual element bounding box for content height (cursor_y includes spacing
        // that overshoots the textbox, but actual rendered text may be smaller).
        let content_h = {
            let mut min_y = f32::MAX;
            let mut max_y = f32::MIN;
            for el in &elements {
                match &el.content {
                    LayoutContent::BoxRect { .. } | LayoutContent::ClipStart | LayoutContent::ClipEnd => {}
                    _ => {
                        if el.y < min_y { min_y = el.y; }
                        let bottom = el.y + el.height;
                        if bottom > max_y { max_y = bottom; }
                    }
                }
            }
            if min_y < max_y { max_y - min_y } else { cursor.cursor_y - (abs_y + inset_t) }
        };
        let v_shift = match v_anchor {
            "ctr" | "middle" | "middle-center" => ((inner_height - content_h) / 2.0).max(0.0),
            "b" | "bottom" | "bottom-center" => (inner_height - content_h).max(0.0),
            _ => 0.0, // "t" | "top" = default, no shift
        };
        if v_shift > 0.0 {
            // Shift all text/content elements (skip BoxRect and ClipStart at indices 0/1)
            for el in elements.iter_mut() {
                match &el.content {
                    LayoutContent::BoxRect { .. } | LayoutContent::ClipStart => {}
                    _ => { el.y += v_shift; }
                }
            }
        }

        // Patch background fill and clip elements with actual height
        for el in elements.iter_mut() {
            if el.x == abs_x && el.y == abs_y && el.height == text_box.height {
                el.height = actual_height;
            }
        }

        // End clip region
        elements.push(LayoutElement::new(abs_x, abs_y, text_box.width, actual_height, LayoutContent::ClipEnd));

        elements
    }

    #[allow(clippy::too_many_arguments)]
    #[allow(unused_assignments)]
    fn layout_paragraph(
        &self,
        para: &Paragraph,
        start_x: f32,
        cursor: &mut LayoutCursor,
        content_width: f32,
        content_height: f32,
        page_top: f32,
        page: &Page,
        pages: &mut Vec<LayoutPage>,
        current_elements: &mut Vec<LayoutElement>,
        grid_pitch: Option<f32>,
        prev_style_id: Option<&str>,
        prev_contextual_spacing: bool,
        #[allow(unused)] in_textbox: bool,
        prev_space_after: f32,
        body_para_index: Option<usize>,
        mut lm2_grid_cells: Option<&mut usize>,
        mut mult_cumul_raw: Option<&mut f32>,
        adjacent_to_empty_run: bool,
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
    ) -> (Vec<LayoutElement>, f32) {
        if let Some(v) = line_fn_refs_out.as_deref_mut() {
            if v.is_empty() { v.push(Vec::new()); }
        }
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
        let snap075 = |y: f32| -> f32 { (y / 0.75).round() * 0.75 };

        // Apply paragraph spacing (space_before).
        // Word uses max(prev_space_after, space_before) — spacing collapse.
        let space_before = if let (Some(bl), Some(pitch)) = (para.style.before_lines, grid_pitch) {
            // beforeLines is specified as percentage of linePitch (e.g., 50 = 0.5 lines).
            // COM-confirmed (2026-04-06): the value is exact (bl/100 * pitch), no grid snap.
            // beforeLines=50 at pitch=17.5 gives exactly 8.75pt, not 17.5pt.
            bl / 100.0 * pitch
        } else {
            para.style.space_before.unwrap_or(0.0)
        };

        // Spacing collapse: max(prev_sa, cur_sb) instead of prev_sa + cur_sb.
        // prev_space_after was NOT added to cursor_y by the caller.
        let collapsed_spacing = space_before.max(prev_space_after);

        // Contextual spacing: suppress spacing when EITHER paragraph has
        // contextualSpacing=true AND they share the same style (COM-confirmed).
        let mut effective_spacing = collapsed_spacing;
        if para.style.contextual_spacing || prev_contextual_spacing {
            if let (Some(cur_id), Some(prev_id)) = (para.style.style_id.as_deref(), prev_style_id) {
                if cur_id == prev_id {
                    effective_spacing = 0.0;
                }
            }
        }

        // Suppress space_before at the top of a page (page 2+).
        // COM-confirmed: page 1 preserves space_before (H1 sb=24 → y=96=72+24).
        // Page 2+ suppresses it.
        let is_page_2_plus = !pages.is_empty() || !current_elements.is_empty();
        if (cursor.cursor_y - page_top).abs() < 0.01 && is_page_2_plus {
            effective_spacing = 0.0;
        }

        cursor.advance(effective_spacing);

        // Debug: dump per-paragraph cursor_y for Class A FAIL root cause investigation.
        // Gated by env OXI_DUMP_CURSOR_Y. Day 33 part 7 (option B).
        if std::env::var("OXI_DUMP_CURSOR_Y").is_ok() {
            let pi_str = body_para_index.map(|v| v.to_string()).unwrap_or_else(|| "?".into());
            let n_runs = para.runs.len();
            let txt: String = para.runs.iter().flat_map(|r| r.text.chars()).take(20).collect();
            eprintln!(
                "[CY_DUMP] body_pi={} cursor_y={:.3} space_before={:.3} n_runs={} text={:?}",
                pi_str, cursor.cursor_y, effective_spacing, n_runs, txt
            );
        }

        // When both twip and *Chars values exist, twip is authoritative (pre-computed by Word).
        // Fall back to *Chars × 10.5pt only when twip value is absent.
        let indent_left = para.style.indent_left
            .or_else(|| para.style.indent_left_chars.map(|c| c / 100.0 * 10.5))
            .unwrap_or(0.0);
        let indent_right = para.style.indent_right
            .or_else(|| para.style.indent_right_chars.map(|c| c / 100.0 * 10.5))
            .unwrap_or(0.0);
        let first_line_indent_raw = para.style.indent_first_line
            .or_else(|| para.style.indent_first_line_chars.map(|c| c / 100.0 * 10.5))
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
        let mut first_line_indent = if list_consumes_hanging { 0.0 } else { first_line_indent_raw };

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
                para.runs.first().map(|r| &r.style).unwrap_or(&RunStyle::default()),
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
        let indent_left = if indent_absorbed_by_leading_ws { 0.0 } else { indent_left };
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
        let s342_no_snap_gate = std::env::var("OXI_S342_NO_SNAP_GATE").map(|v| v != "0" && v != "false").unwrap_or(true);
        let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT").map(|v| v != "0" && v != "false").unwrap_or(false);
        let snap_pass_through = s342_no_snap_gate || s344_fs_gate;
        let snap_gate_active = !snap_pass_through && !para.style.snap_to_grid;
        let effective_char_pitch = if in_textbox || snap_gate_active { None } else { page.grid_char_pitch };
        // 2026-05-05 Track A (Session 55+): COM-measured 8 paragraphs in b837
        // confirmed Word's wrap rule: available = content_w - indent_l - indent_r
        // for both charGrid and non-charGrid (full indent applied, no cell-based
        // tolerance). Combined with fn attribution fix below, b837 pagination
        // score improved 0.9524 → 0.9744 (Phase 1 gate). Other docs unchanged.
        let available_width = content_width - indent_left - indent_right;

        // Render list marker if present
        // S517 (2026-06-09): index of the emitted list-marker element so the
        // first-line loop can back-patch its text_y_off to share the body
        // baseline (the marker element is emitted here, before the line loop
        // computes text_y_off). Only set for NON-bullet markers (number markers
        // like ①/(1)); bullets keep their own marker_y_offset tuning untouched.
        let mut s517_marker_el_idx: Option<usize> = None;
        if let Some(ref marker) = para.style.list_marker {
            let default_style = RunStyle::default();
            let marker_style = para.runs.first().map(|r| &r.style).unwrap_or(&default_style);
            let marker_font_size = self.resolve_font_size(marker_style, &para.style);
            // Symbol font bullets (•/●) have large glyphs relative to em-square.
            // No font size adjustment needed — use the paragraph's font size directly.
            let marker_metrics = self.metrics_for(marker_style, &para.style);
            let marker_width: f32 = marker
                .chars()
                .map(|c| self.registry.char_width_pt_with_fallback(c, marker_font_size, marker_metrics))
                .sum();
            let list_indent = para.style.list_indent.unwrap_or(18.0);
            let marker_x = start_x + indent_left - list_indent;
            let line_height = self.line_height(marker_font_size, para.style.line_spacing, para.style.line_spacing_rule.as_deref(), marker_metrics, para.style.snap_to_grid, grid_pitch);

            // Determine marker text including suffix
            let suff = para.style.list_suff.as_deref().unwrap_or("tab");
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

            // Page break check for marker
            if cursor.cursor_y + line_height > page_top + content_height {
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                cursor.set(page_top);
            }

            // Bullet markers are scaled up (2x) so adjust Y to align with text center
            let marker_y_offset = if marker.contains('\u{2022}') || marker.contains('\u{25CF}') {
                -marker_font_size * 0.15  // shift up slightly
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
            } else {
                self.resolve_font_family_for_text(&marker_text, marker_style, &para.style)
                    .map(|s| s.to_string())
            };
            let marker_bold = self.resolve_bold(marker_style, &para.style);
            let marker_color = self.resolve_color(marker_style, &para.style).map(|s| s.to_string());
            let marker_base_y = if s467_vsnap { snap075(cursor.visual_y) } else { cursor.visual_y };
            // S517: remember this marker element's index so the first body line
            // can set its text_y_off to match (the marker shares the body
            // baseline — Word-confirmed dy=0 on b837 ①②③). Scoped to non-bullet
            // markers (marker_y_offset==0) so bullet placement is unchanged.
            if marker_y_offset == 0.0 {
                s517_marker_el_idx = Some(elements.len());
            }
            elements.push(LayoutElement::new(marker_x, marker_base_y + marker_y_offset, marker_width, line_height, LayoutContent::Text {
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
            }));
        }

        // Collect all text fragments with their styles, field types, and source indices
        let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = para
            .runs
            .iter()
            .enumerate()
            .map(|(i, r)| (r.text.as_str(), &r.style, r.field_type, i, 0usize))
            .collect();

        // Resolve font size for line breaking
        let default_style = RunStyle::default();
        let para_font_size = self.resolve_font_size(
            para.runs.first().map(|r| &r.style).unwrap_or(&default_style),
            &para.style,
        );

        // Round 7: pre-compute ruby paragraph-tail expansion once.
        // Greenfield-dormant: 0/177 baseline docs use w:ruby, so this is
        // 0.0 for all baseline paragraphs. Used at last-line cursor advance
        // and gates ruby-annotation emission below.
        let ruby_para_expansion_pt = ruby::paragraph_ruby_expansion_pt(&para.runs, para_font_size);

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
        let ruby_total_overhang_pt: f32 = para.runs.iter()
            .filter_map(|run| run.ruby.as_ref().map(|r| (run, r)))
            .map(|(run, ruby_ir)| {
                let base_pt = run.style.font_size.unwrap_or(para_font_size);
                let hps_pt = ruby_ir.hps_halfpt
                    .map(|h| h as f32 / 2.0)
                    .unwrap_or(base_pt / 2.0);
                let ruby_metrics = self.metrics_for_text(&ruby_ir.text, &run.style, &para.style);
                let base_metrics = self.metrics_for_text(&run.text, &run.style, &para.style);
                let ruby_w: f32 = ruby_ir.text.chars()
                    .map(|c| self.registry.char_width_pt_with_fallback(c, hps_pt, ruby_metrics))
                    .sum();
                let base_w: f32 = run.text.chars()
                    .map(|c| self.registry.char_width_pt_with_fallback(c, base_pt, base_metrics))
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
        let effective_first_indent = first_line_indent;
        // S342: mirror the snap_to_grid gate change for cw_ratio (see effective_char_pitch comment).
        let effective_cw_ratio = if in_textbox || snap_gate_active { None } else { page.grid_char_cw_ratio };
        let wrap_width = (available_width - ruby_total_overhang_pt).max(0.0);
        // S476: this is the MAIN BODY flow (s476_body=true) → S475/S476 yakumono
        // capacity may apply (the demand break). Aux/estimate/cell calls pass false.
        let para_has_lrpb = para.runs.iter().any(|r| r.has_last_rendered_page_break);
        let lines = self.break_into_lines(&fragments, wrap_width, effective_first_indent, &para.style, effective_char_pitch, effective_cw_ratio, page.doc_grid_lines_and_chars, true, matches!(para.alignment, Alignment::Justify | Alignment::Distribute), page.doc_grid_no_type, para_has_lrpb);

        // S168 Phase B-2 holistic bundle (b): per-line fn cumul delta.
        let committed_fn_delta_at_line: Vec<f32> = if !para_fn_heights.is_empty() {
            let mut out = Vec::with_capacity(lines.len());
            let mut cumulative = 0.0_f32;
            let mut seen: Vec<u32> = Vec::new();
            for line in &lines {
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
            }
            out
        } else { vec![0.0; lines.len()] };

        // Widow/orphan control: pre-compute line heights for lookahead
        let line_heights: Vec<f32> = lines.iter().map(|line| {
            self.line_height_for_line(line, &para.style, para_font_size, para.style.snap_to_grid, grid_pitch, page.doc_grid_no_type)
        }).collect();
        // Day 33 part 65 (2026-05-12): natural line heights (ascent+descent only)
        // for page-break threshold. Word allows grid-snap LEADING to extend
        // into bottom margin; only the text-occupying zone (ascent+descent)
        // must fit within content area. db9ca18 i=37 confirmed via COM:
        // line at y=758.25, grid line_h=18, line bottom=776.25 (5.25pt past
        // pgBot=771) — Word fits, while Oxi (using full line_h for break
        // check) rejected.
        let natural_line_heights: Vec<f32> = lines.iter().map(|line| {
            self.natural_line_height_for_line(line, &para.style, para_font_size)
        }).collect();
        // S576 (2026-06-15): glyph-ink line heights (typo_sum*fs ≈ em) for the
        // page-bottom break-fit. The natural_line_heights above are the SPACING
        // box (win*83/64 = 1.297*em for CJK), ~3.2pt larger than the real glyph
        // ink — that over-count rejected page-bottom lines Word fits (their grid
        // leading hangs into the margin). See break_threshold below.
        let ink_line_heights: Vec<f32> = lines.iter().map(|line| {
            self.ink_line_height_for_line(line, &para.style, para_font_size)
        }).collect();

        // COM-confirmed (2026-04-05, test_widow): Multiple spacing uses cumulative ceil
        // for intra-paragraph Y positions. Last line uses per-line ceil for paragraph gap.
        // COM-confirmed (2026-04-08, 683ffcab86e2): SINGLE spacing also benefits from
        // cumulative round in LM=0, but ONLY when raw > per-line round (i.e., cumulative
        // gives MORE advance than per-line), which preserves page-break decisions.
        // When raw < per-line round (e.g., Meiryo 10.5pt: raw=20.43 vs round=20.5),
        // cumulative would tighten content and shift page breaks (LOD_Handbook lost a
        // page in bdd9321 → reverted in cb35baa). Gating by sign keeps both gains.
        let is_multiple_spacing = match (para.style.line_spacing_rule.as_deref(), para.style.line_spacing) {
            (Some("exact"), _) | (Some("atLeast"), _) => false,
            (_, Some(f)) if (f - 1.0).abs() > 0.001 => true,
            _ => false,
        };
        let is_single_lm0 = !is_multiple_spacing && grid_pitch.is_none()
            && match (para.style.line_spacing_rule.as_deref(), para.style.line_spacing) {
                (Some("exact"), _) | (Some("atLeast"), _) => false,
                _ => true,
            };
        let use_cumulative_basis = is_multiple_spacing || is_single_lm0;
        let raw_spaced_tw: f32 = if use_cumulative_basis && !lines.is_empty() {
            let first_line = &lines[0];
            let base = {
                let mut ma: f32 = 0.0; let mut md: f32 = 0.0;
                let mut has_latin = false;
                if first_line.fragments.is_empty() {
                    // Match line_height_for_line_inner: empty paragraphs use
                    // pPr/rPr font + para_mark metrics (not doc_default).
                    let font_size = para.style.ppr_rpr.as_ref()
                        .and_then(|r| r.font_size)
                        .unwrap_or(para_font_size);
                    let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                    let m = self.metrics_for_para_mark(&rpr_ref, &para.style);
                    ma = m.word_ascent_pt(font_size); md = m.word_descent_pt(font_size);
                } else {
                    for frag in &first_line.fragments {
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let m = self.metrics_for_text(&frag.text, &frag.style, &para.style);
                        if m.word_ascent_pt(fs) > ma { ma = m.word_ascent_pt(fs); }
                        if m.word_descent_pt(fs) > md { md = m.word_descent_pt(fs); }
                        if !frag.text.chars().all(|c| kinsoku::is_cjk(c)) {
                            has_latin = true;
                        }
                    }
                    // COM-confirmed (2026-04-07): Latin text on a line causes Word to also
                    // consider the ASCII font's CJK 83/64 height for the base.
                    if has_latin {
                        if let Some(frag) = first_line.fragments.first() {
                            let fs = frag.style.font_size.unwrap_or(para_font_size);
                            let latin_m = self.metrics_for(&frag.style, &para.style);
                            if latin_m.is_cjk_83_64_font() {
                                let la = latin_m.word_ascent_pt(fs);
                                let ld = latin_m.word_descent_pt(fs);
                                if la > ma { ma = la; }
                                if ld > md { md = ld; }
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
                    // Empty paragraphs: compute no_grid from para mark font
                    // (matching line_height_for_line_inner's empty-para logic).
                    if first_line.fragments.is_empty() {
                        let font_size = para.style.ppr_rpr.as_ref()
                            .and_then(|r| r.font_size)
                            .unwrap_or(para_font_size);
                        let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        let m = self.metrics_for_para_mark(&rpr_ref, &para.style);
                        no_grid_max = m.word_line_height_no_grid(font_size);
                        no_grid_raw_max = (m.win_ascent + m.win_descent) * font_size;
                    }
                    for frag in &first_line.fragments {
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let m = self.metrics_for_text(&frag.text, &frag.style, &para.style);
                        let h = m.word_line_height_no_grid(fs);
                        if h > no_grid_max { no_grid_max = h; }
                        // Raw (un-floored) height for Multiple spacing cumulative base
                        let raw = (m.win_ascent + m.win_descent) * fs;
                        if raw > no_grid_raw_max { no_grid_raw_max = raw; }
                    }
                    if has_latin {
                        if let Some(frag) = first_line.fragments.first() {
                            let fs = frag.style.font_size.unwrap_or(para_font_size);
                            let latin_m = self.metrics_for(&frag.style, &para.style);
                            if latin_m.is_cjk_83_64_font() {
                                let h = latin_m.word_line_height_no_grid(fs);
                                if h > no_grid_max { no_grid_max = h; }
                                let raw = (latin_m.win_ascent + latin_m.win_descent) * fs;
                                if raw > no_grid_raw_max { no_grid_raw_max = raw; }
                            }
                        }
                    }
                    if is_multiple_spacing && no_grid_raw_max > 0.0 {
                        run_base.max(no_grid_raw_max)
                    } else {
                        run_base.max(no_grid_max)
                    }
                } else {
                    run_base
                }
            };
            let raw = base * para.style.line_spacing.unwrap_or(1.0) * 20.0;
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
                if para.style.snap_to_grid && !page.doc_grid_no_type && pitch > 0.0
                    && std::env::var("OXI_S584_DISABLE").is_err()
                {
                    raw.max(pitch * 20.0)
                } else { raw }
            } else { raw }
        } else { 0.0 };
        // LM2 (charGrid) and LM0 single spacing: carry cumul_line_idx across paragraphs.
        // COM-confirmed (2026-04-12, 0e7a p2): Word maintains cumulative line index
        // across paragraph boundaries for LM0 single spacing, producing continuous
        // 11.5/11.5/12.0 round pattern instead of resetting at each paragraph.
        // Multiple spacing (1.15x etc): reset per paragraph (cumul uses raw base).
        let carry_cumul = is_single_lm0 || grid_pitch.is_some();
        let mut cumul_line_idx: usize = if carry_cumul {
            lm2_grid_cells.as_deref().copied().unwrap_or(0)
        } else { 0 };

        for (line_idx, line) in lines.iter().enumerate() {
            let _first_style = line.fragments.first().map(|f| &f.style).unwrap_or(&default_style);
            let line_height = line_heights[line_idx];

            // Page break check with widow/orphan control
            // TextBox content: no page breaks, no widow/orphan. Overflow is clipped.
            let effective_lh = if is_multiple_spacing && raw_spaced_tw > 0.0 {
                let old_pos = mult_cumul_raw.as_deref().copied().unwrap_or(0.0);
                let new_pos = old_pos + raw_spaced_tw;
                let cn = (new_pos / 10.0).round() as i32 * 10;
                let cc = (old_pos / 10.0).round() as i32 * 10;
                (cn - cc) as f32 / 20.0
            } else { line_height };
            // Day 33 part 65 (2026-05-12): use natural_lh (ascent+descent) for
            // break threshold; the grid-snap LEADING (line_h − natural_lh) is
            // allowed to extend into bottom margin. Cursor still advances by
            // full line_h. COM-confirmed via db9ca18 i=37 (+5.25pt overflow
            // accepted by Word).
            let natural_lh = natural_line_heights.get(line_idx).copied().unwrap_or(effective_lh);
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
            let s562b_empty_full = std::env::var("OXI_S562B_DISABLE").is_err()
                && para.runs.iter().all(|r| r.text.is_empty());
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
                natural_lh
            } else {
                ink_line_heights.get(line_idx).copied().unwrap_or(natural_lh).min(natural_lh)
            };
            // S582 (2026-06-15) FALSIFIED the "S576 ink-leniency is ~1.75pt too
            // loose" hypothesis for ikujidetail's +1×2: an OXI_INK_MARGIN sweep
            // showed margin 0 (= ink=em, current) is OPTIMAL; +1.0/+1.5 unchanged,
            // +1.75 WORSE (+1×5), +2.0..box +1×12 (the S571b state). So the
            // page-bottom threshold is correct; the +1×2 are the doc-wide para-spill
            // break-POINT cascade (which line of a wrapped para lands at the bottom),
            // reset by the real pi=149/263 LRPBs — not a threshold calibration.
            let break_threshold = if s548b_exact_full || s562b_empty_full {
                effective_lh
            } else {
                ink_lh.min(effective_lh)
            };
            // R7.53: first-line lenient check using `first_line_extra_content_h`.
            // S168 Phase B-2 (c): per-line lenient.
            let line_lenient_extra = if !para_fn_heights.is_empty() {
                let committed = committed_fn_delta_at_line.get(line_idx).copied().unwrap_or(0.0);
                (first_line_extra_content_h - committed).max(0.0)
            } else if line_idx == 0 {
                first_line_extra_content_h
            } else {
                0.0
            };
            let effective_break_bottom = page_top + content_height + line_lenient_extra;
            let natural_needs_page_break = if in_textbox { false } else {
                cursor.cursor_y + break_threshold > effective_break_bottom
            };
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
                .unwrap_or(true);
            let s391_lrpb_break = if line_idx > 0 && !in_textbox && s391_on {
                let has_lrpb_here = line.fragments.iter().any(|f| {
                    f.char_offset == 0
                        && para.runs.get(f.run_index)
                            .map(|r| r.has_last_rendered_page_break)
                            .unwrap_or(false)
                });
                let s394_max = std::env::var("OXI_S394_LRPB_MAX").ok()
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
                // within 1 line of the bottom. Opt-out OXI_S581_DISABLE.
                let s581_stale = if std::env::var("OXI_S581_DISABLE").is_ok() {
                    false
                } else {
                    let over = cursor.cursor_y + break_threshold - effective_break_bottom;
                    over < -effective_lh
                };
                has_lrpb_here && page.total_lrpb_count <= s394_max && s563_full && !s581_stale
            } else {
                false
            };
            let needs_page_break = natural_needs_page_break || s391_lrpb_break;
            if std::env::var("OXI_DUMP_BREAK").is_ok()
                && (line_idx == 0 || (std::env::var("OXI_DUMP_BREAK_ALL").is_ok() && cursor.cursor_y > 700.0)) {
                let pi_str = body_para_index.map(|v| v.to_string()).unwrap_or_else(|| "?".into());
                let txt: String = para.runs.iter().flat_map(|r| r.text.chars()).take(15).collect();
                eprintln!(
                    "[BR_DUMP] pi={} line0 cursor_y={:.3} eff_lh={:.3} nat_lh={:.3} ink_lh={:.3} brk_thr={:.3} eff_bot={:.3} over={:.3} brk={} text={:?}",
                    pi_str, cursor.cursor_y, effective_lh, natural_lh, ink_lh, break_threshold,
                    effective_break_bottom, cursor.cursor_y + break_threshold - effective_break_bottom,
                    needs_page_break, txt
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
            let widow_effective = para.style.widow_control
                || (force_widow && lines.len() >= 5);
            let widow_orphan_break = if !in_textbox && widow_effective && lines.len() >= 2 {
                if line_idx == 0 && !needs_page_break {
                    // Orphan: check if the next line would overflow — that would leave
                    // only 1 line on this page. Push entire paragraph to next page.
                    // Orphan: check if the next line would overflow — that would leave
                    // only 1 line on this page. Push entire paragraph to next page.
                    let next_h = line_heights.get(1).copied().unwrap_or(0.0);
                    cursor.cursor_y + line_height + next_h > page_top + content_height
                        && !current_elements.is_empty()
                } else if line_idx == lines.len() - 2 && !needs_page_break {
                    // Widow: if the last line would overflow to the next page alone,
                    // break BEFORE this line so at least 2 lines go to the next page.
                    let next_h = line_heights.get(line_idx + 1).copied().unwrap_or(0.0);
                    cursor.cursor_y + line_height + next_h > page_top + content_height
                } else {
                    false
                }
            } else {
                false
            };
            if std::env::var("OXI_DUMP_WIDOW").is_ok() && line_idx == 0 {
                let txt: String = para.runs.iter().flat_map(|r| r.text.chars()).take(15).collect();
                eprintln!("[WIDOW] line0 lines={} wc={} cursor_y={:.2} lh={:.2} next_h={:.2} limit={:.2} curr_empty={} break={} text={:?}",
                    lines.len(), para.style.widow_control, cursor.cursor_y, line_height,
                    line_heights.get(1).copied().unwrap_or(0.0),
                    page_top + content_height, current_elements.is_empty(),
                    widow_orphan_break, txt);
            }

            if widow_orphan_break {
                // Push current page and move entire paragraph so far to next page
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                cursor.set(page_top);
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
                    v.push(carry);       // NEW page — earlier lines' refs move here
                }
            } else if needs_page_break {
                // Phantom-blank-page fix (2026-04-23): when an empty paragraph
                // with page_break_after overflows and would produce an empty
                // stub alone on a new page, followed by ANOTHER page break,
                // skip the stub entirely. Push the current page and return —
                // caller's page_break_after path is a no-op on empty elements,
                // so the next block renders on the fresh page directly.
                // d77a p.11 case: block 127 is just <w:br w:type="page"/>.
                // See project_d77a_phantom_page_11.md.
                if para.runs.is_empty() && para.style.page_break_after {
                    current_elements.extend(std::mem::take(&mut elements));
                    pages.push(LayoutPage {
                        width: page.size.width,
                        height: page.size.height,
                        elements: std::mem::take(current_elements),
                    });
                    cursor.set(page_top);
                    return (Vec::new(), 0.0);
                }
                // Mid-paragraph page break: keep already-laid-out lines on current page,
                // only the overflowing line (and subsequent) go to the next page.
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                cursor.set(page_top);
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
                if !s396_skip_cont_hl && line_idx > 0
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
            }

            // Step 0: record fn refs rendered on this line's final page
            // (after any pre-line page push above). A run with footnote_ref
            // maps its marker to this line if any fragment here references
            // that run_index.
            if let Some(v) = line_fn_refs_out.as_deref_mut() {
                let bucket = v.last_mut().unwrap();
                let mut seen: Vec<usize> = Vec::new();
                for f in &line.fragments {
                    if seen.contains(&f.run_index) { continue; }
                    seen.push(f.run_index);
                    if let Some(run) = para.runs.get(f.run_index) {
                        if let Some(id) = run.footnote_ref {
                            if !bucket.contains(&id) { bucket.push(id); }
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
                                if !bucket.contains(&id) { bucket.push(id); }
                                i += 1;
                            } else {
                                break;
                            }
                        }
                    }
                }
            }

            let extra_indent = if line_idx == 0 { first_line_indent } else { 0.0 };
            // COM-confirmed 2026-04-17 (measure_hanging_indent_v2.py): first-line
            // indent DOES shift line_x. Word places line 1 at margin+indent_left+
            // first_line_indent, continuation lines at margin+indent_left. Applies
            // to both positive firstLine (e.g. +21pt) and hanging (negative, e.g. -9pt).
            let line_x = start_x + indent_left + extra_indent;

            // Alignment offset
            let line_text_width: f32 = line.fragments.iter().map(|f| f.width).sum();
            let is_last_line = line_idx == lines.len() - 1;
            // For alignment/justify, use indent-adjusted width.
            // Justify/alignment uses indent-adjusted width. In charGrid mode,
            // break_into_lines may put more chars than fit in the indented area;
            // negative slack triggers punctuation compression to fit.
            let render_width = content_width - indent_left - indent_right;
            let align_offset = match para.alignment {
                Alignment::Left => 0.0,
                Alignment::Center => {
                    // Word GDI: integer pixel division at 96dpi for center alignment
                    let slack_tw = ((render_width - extra_indent - line_text_width) * 20.0).round() as i32;
                    let center_tw = slack_tw / 2; // integer division (truncate)
                    center_tw as f32 / 20.0
                },
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
                        if slack > 0.0 { slack / 2.0 } else { 0.0 }
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

            let is_soft_break_line = self.do_not_expand_shift_return
                && line.break_type == LineBreakType::SoftBreak;
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
                    line.fragments.iter().map(|f| {
                        let fs = f.style.font_size.unwrap_or(para_font_size);
                        f.text.chars().filter(|&c| crate::font::is_fullwidth(c) && fs < pitch)
                            .count() as f32 * (pitch - fs)
                    }).sum::<f32>()
                } else { 0.0 };
                let mut slack = render_width - extra_indent - line_text_width - grid_extra_on_line;

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
                        let cap = if !single { 0.0 } else { match c0 {
                            '、' | '，' => fs / 3.0,
                            '。' | '．' => fs / 2.0,
                            '」' | '』' | '】' | '〕' | '》' | '〉' | '｝' | '］' | '）' => fs / 2.0,
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
                        }};
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
                        let mut active: Vec<(usize, f32)> = comps.iter().map(|(fi, _, cap)| (*fi, *cap)).collect();
                        loop {
                            if active.is_empty() || needed <= 0.001 { break; }
                            let share = needed / active.len() as f32;
                            let capped: Vec<(usize, f32)> = active.iter().cloned()
                                .filter(|(_, cap)| *cap <= share).collect();
                            if capped.is_empty() {
                                for (fi, _) in &active { comp_amt[*fi] = share; }
                                break;
                            }
                            for (fi, cap) in &capped { comp_amt[*fi] = *cap; needed -= cap; }
                            active.retain(|(_, cap)| *cap > share);
                        }
                    }
                    // Set adjustments so each standalone 、。 renders at (natural − comp).
                    for (fi, fs, _) in &comps {
                        frag_width_adjustments[*fi] = (fs - comp_amt[*fi]) - line.fragments[*fi].width;
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
                                let is_opening_bracket = matches!(ch,
                                    '（' | '「' | '『' | '〔' | '【' | '《' | '〈' | '｛' | '［'
                                );
                                if is_opening_bracket {
                                    continue;
                                }
                                let fs = frag.style.font_size.unwrap_or(para_font_size);
                                let fm = self.metrics_for(&frag.style, &para.style);
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
                let post_phase1_ltw: f32 = line.fragments.iter().enumerate()
                    .map(|(i, f)| f.width + frag_width_adjustments[i])
                    .sum();
                slack = if phase1_compressed {
                    render_width - extra_indent - post_phase1_ltw
                } else {
                    render_width - extra_indent - post_phase1_ltw - grid_extra_on_line
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
                            if !matches!(ch, '、' | '。' | '，' | '．') { continue; }
                            let fs = frag.style.font_size.unwrap_or(para_font_size);
                            // 、 is CJK punct — use CJK metrics if available.
                            let fm_cjk = self.metrics_for_cjk(&frag.style, &para.style);
                            let fm = fm_cjk.unwrap_or_else(|| self.metrics_for(&frag.style, &para.style));
                            let natural = self.registry.char_width_pt_with_fallback(ch, fs, fm);
                            let current = frag.width + frag_width_adjustments[fi];
                            if current < natural * 0.95 {
                                compressed.push((fi, natural, current));
                            }
                        }
                    }
                    if !compressed.is_empty() {
                        let n = compressed.len() as f32;
                        let per_comp_cap = compressed.iter()
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
                    let space_count = line.fragments.iter()
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
                        let total_chars: usize = line.fragments.iter()
                            .map(|f| f.text.chars().count())
                            .sum();
                        let has_cjk = line.fragments.iter()
                            .any(|f| f.text.chars().any(|c| kinsoku::is_cjk(c)));
                        if has_cjk && total_chars > 1 {
                            let char_gap_count = total_chars - 1;
                            let per_char_gap = slack / char_gap_count as f32;
                            // Distribute: fragment-boundary gaps via frag_spacing_after,
                            // internal gaps via frag_width_adjustments (for layout width),
                            // AND set justify_char_spacing for renderer to apply letterSpacing.
                            for fi in 0..line.fragments.len() {
                                let frag_chars = line.fragments[fi].text.chars().count();
                                if frag_chars > 1 {
                                    frag_width_adjustments[fi] += per_char_gap * (frag_chars - 1) as f32;
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

            // Matches Word output: exact/atLeast line spacing places text at BOTTOM of line box.
            // Extra space goes above text (ascent increased, descent unchanged).
            // Session 76 Mech A fix: pass in_textbox so the function can distinguish
            // body/cell (top-align for exact) from shape (bottom-align).
            let text_y_off = self.text_y_offset_for_line(line, &para.style, para_font_size, line_height, grid_pitch, in_textbox);

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
                        }
                    }
                }
            }

            // Compute max ascent across all fragments for baseline alignment.
            // All fragments in a line share the same baseline (matches Word output).
            let line_max_ascent: f32 = if line.fragments.is_empty() {
                // COM-confirmed: empty lines use paragraph font (East Asian in CJK docs)
                self.metrics_for_para_mark(&RunStyle::default(), &para.style).word_ascent_pt(para_font_size)
            } else {
                line.fragments.iter().map(|f| {
                    let fs = f.style.font_size.unwrap_or(para_font_size);
                    self.metrics_for_text(&f.text, &f.style, &para.style).word_ascent_pt(fs)
                }).fold(0.0_f32, f32::max)
            };

            // R-10: track whether any fragment on this line came from a
            // revision-bearing source run; if so we emit one change-bar at
            // the line's left margin after the fragment loop finishes.
            // Paragraph-level revisions (ppr_change, paragraph_mark_revision)
            // count as revisions on every line of the paragraph — they're
            // detected once at paragraph entry rather than per-fragment.
            let mut line_has_revision = para.ppr_change.is_some()
                || para.paragraph_mark_revision.is_some();

            for (frag_idx, frag) in line.fragments.iter().enumerate() {
                let base_font_size = frag.style.font_size.unwrap_or(para_font_size);
                // Round 29: superscript/subscript rendering. Word default for
                // <w:vertAlign w:val="superscript"/> and "subscript":
                //   - font size = base * 0.583... (≈58%)
                //   - vertical offset = +/- (base_size * 0.333) from baseline
                //     (negative = up for superscript, positive = down for subscript)
                let (resolved_font_size, vert_offset) = match frag.style.vertical_align {
                    Some(VerticalAlign::Superscript) => {
                        let fs = base_font_size * 0.583;
                        // Raise the glyph: smaller font's baseline shifts up
                        // by ~1/3 of the original font size.
                        (fs, -(base_font_size * 0.333))
                    }
                    Some(VerticalAlign::Subscript) => {
                        let fs = base_font_size * 0.583;
                        (fs, base_font_size * 0.083)
                    }
                    _ => (base_font_size, 0.0),
                };
                let resolved_bold = self.resolve_bold(&frag.style, &para.style);
                let adjusted_width = frag.width + frag_width_adjustments[frag_idx];

                // Per-fragment baseline alignment: shift fragments with smaller ascent
                // so all share the same baseline (y + frag_ascent = cursor_y + text_y_off + line_max_ascent)
                let frag_metrics = self.metrics_for_text(&frag.text, &frag.style, &para.style);
                let frag_ascent = frag_metrics.word_ascent_pt(resolved_font_size);
                // COM-confirmed (2026-04-14, gen2_001): Word does NOT apply
                // per-fragment baseline adjustment. All fragments on the same line
                // share the same Y coordinate. GDI TextOutW aligns from font cell top.
                let baseline_adjust = 0.0;
                let _ = line_max_ascent;
                let _ = frag_ascent;

                // Session 75 Phase D (2026-05-17): y is LINE BOX TOP, renderer adds
                // text_y_off + baseline_adjust + vert_offset at draw time. See
                // memory/session71_y_convention_refactor_design.md.
                // S467: snap the emitted line top to the 0.75pt grid (Word's model).
                let emit_y = if s467_vsnap { snap075(cursor.visual_y) } else { cursor.visual_y };
                let mut el = LayoutElement::new(x, emit_y, adjusted_width, line_height, LayoutContent::Text {
                        text: frag.text.clone(),
                        font_size: resolved_font_size,
                        font_family: self.resolve_font_family_for_text(&frag.text, &frag.style, &para.style)
                            .map(|s| s.to_string()),
                        bold: resolved_bold,
                        italic: self.resolve_italic(&frag.style, &para.style),
                        underline: frag.style.underline,
                        underline_style: frag.style.underline_style.clone(),
                        strikethrough: frag.style.strikethrough,
                        double_strikethrough: frag.style.double_strikethrough,
                        color: self.resolve_color(&frag.style, &para.style).map(|s| s.to_string()),
                        highlight: frag.style.highlight.clone(),
                        field_type: frag.field_type,
                        character_spacing: if frag.style.fit_text.is_some() {
                            frag.style.character_spacing.unwrap_or(0.0) + justify_char_spacing
                        } else {
                            snap_character_spacing(frag.style.character_spacing.unwrap_or(0.0)) + justify_char_spacing
                        },
                        text_scale: frag.style.text_scale.unwrap_or(100.0),
                        is_vertical: false,
                });
                // Session 72 Phase A: populate text_y_off (y still includes it).
                el.text_y_off = text_y_off + baseline_adjust + vert_offset;
                if let Some(pi) = body_para_index {
                    el.paragraph_index = Some(pi);
                    el.run_index = Some(frag.run_index);
                    el.char_offset = Some(frag.char_offset);
                }
                // Round 7: capture base element x/y BEFORE push (move).
                // Used below to position the ruby annotation above the base.
                let base_el_x = el.x;
                let base_el_y = el.y;
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
                if frag.char_offset == 0 {
                    if let Some(run) = para.runs.get(frag.run_index) {
                        if let Some(ref ruby_ir) = run.ruby {
                            let base_pt = frag.style.font_size.unwrap_or(para_font_size);
                            let hps_pt = ruby_ir.hps_halfpt
                                .map(|h| h as f32 / 2.0)
                                .unwrap_or(base_pt / 2.0);
                            let hps_raise_pt = ruby_ir.hps_raise_halfpt
                                .map(|h| h as f32 / 2.0)
                                .unwrap_or(ruby::DEFAULT_HPS_RAISE_PT);
                            let ruby_text = ruby_ir.text.as_str();
                            let mut ruby_run_style = frag.style.clone();
                            ruby_run_style.font_size = Some(hps_pt);
                            let ruby_metrics = self.metrics_for_text(ruby_text, &ruby_run_style, &para.style);
                            // Round 7.6: precise per-char widths via GDI metrics.
                            // Replaces the previous `chars × font_size` CJK
                            // monospace approximation; matches non-CJK ruby
                            // and proportional fonts correctly.
                            let ruby_char_count = ruby_text.chars().count();
                            let ruby_w: f32 = ruby_text
                                .chars()
                                .map(|c| self.registry.char_width_pt_with_fallback(c, hps_pt, ruby_metrics))
                                .sum();
                            let base_metrics = self.metrics_for_text(run.text.as_str(), &frag.style, &para.style);
                            let base_w: f32 = run.text
                                .chars()
                                .map(|c| self.registry.char_width_pt_with_fallback(c, base_pt, base_metrics))
                                .sum();
                            // Round 7.5: rubyAlign positioning per ECMA-376 §17.3.3.26.
                            // ruby_position returns (x_offset_from_base, per_char_spacing).
                            let (ruby_x_offset, ruby_char_spacing) = ruby::ruby_position(
                                base_w, ruby_w, ruby_char_count, ruby_ir.align,
                            );
                            let ruby_x = base_el_x + ruby_x_offset;
                            let ruby_ascent = ruby_metrics.word_ascent_pt(hps_pt);
                            let frag_metrics = self.metrics_for_text(&frag.text, &frag.style, &para.style);
                            let base_ascent = frag_metrics.word_ascent_pt(base_pt);
                            let ruby_y = base_el_y + base_ascent - hps_raise_pt - ruby_ascent;
                            let ruby_color = self.resolve_color(&ruby_run_style, &para.style).map(|s| s.to_string());
                            let ruby_font_family = self.resolve_font_family_for_text(ruby_text, &ruby_run_style, &para.style)
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
                                },
                            );
                            if let Some(pi) = body_para_index {
                                ruby_el.paragraph_index = Some(pi);
                                ruby_el.run_index = Some(frag.run_index);
                                ruby_el.char_offset = Some(0);
                            }
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
                        },
                    );
                    // Session 72 Phase A: populate text_y_off (y still includes it).
                    el.text_y_off = text_y_off;
                    el.paragraph_index = Some(pi);
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
            let is_lm2_single = lm2_grid_cells.is_some()
                && page.grid_char_pitch.is_some()
                && grid_pitch.map_or(false, |p| p > 0.0)
                && para.style.snap_to_grid
                && match (para.style.line_spacing_rule.as_deref(), para.style.line_spacing) {
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
                let margin_tw = (page.margin.top * 20.0).round() as i32;
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
                    let ideal0 = if cursor.lm2_ideal_y > 0.0
                        && (cursor.lm2_ideal_y - cur_f).abs() < pitch * 0.5 {
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
                        if raw % 10 == 0 { raw } else { (raw / 10 + 1) * 10 }
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
                    let raw = cur_tw + cells * pitch_tw_i;
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
                cumul_line_idx += cells as usize;
                } // end S494 else (legacy cell-aligned / mid-cell branch)
            } else {
            // For single LM=0, gate by direction: only when raw advances MORE than rounded.
            let single_lm0_safe = if is_single_lm0 && raw_spaced_tw > 0.0 {
                let raw_pt = raw_spaced_tw / 20.0;
                let rounded_pt = (raw_pt * 2.0).round() / 2.0;
                raw_pt > rounded_pt
            } else { false };
            // LM=0 cumulative ROUND includes LAST line; LM≥1 cumulative CEIL excludes last.
            let use_cumulative = (is_multiple_spacing || single_lm0_safe) && raw_spaced_tw > 0.0
                && line_heights.iter().all(|&h| (h - line_heights[0]).abs() < 1.5)
                && (grid_pitch.is_none() || !is_last);
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
                if s467_vsnap && is_multiple_spacing {
                    // visual_y advances by the EXACT raw line height; cursor_y by the
                    // current rounded amount (page-break unchanged). Emit snaps visual_y.
                    cursor.advance_split((cn - cc) as f32 / 20.0, raw_spaced_tw / 20.0);
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
            if line_idx + 1 == lines.len() && ruby_para_expansion_pt > 0.0 {
                cursor.advance(ruby_para_expansion_pt);
            }
            // Only advance cumul index when cumulative round is active.
            // COM-confirmed (683f): paragraphs with non-uniform line heights
            // (use_cumulative=false) do NOT advance the cross-paragraph index.
            if use_cumulative {
                cumul_line_idx += 1;
            }
            } // end else (non-LM2 single)

            // Handle explicit page/column breaks after this line
            if line.break_type == LineBreakType::PageBreak || line.break_type == LineBreakType::ColumnBreak {
                // Day 33 part 59 (2026-05-12): the line that CARRIES the break_type
                // has its text already rendered into `elements` and should stay on
                // the CURRENT page (text BEFORE the `<w:br w:type="page"/>` belongs
                // to current page per OOXML semantics). Original code pushed
                // current_elements first then merged elements → pi=11 text ended up
                // on the NEW page. Fix: merge elements into the pushed page first.
                let mut page_elements = std::mem::take(current_elements);
                page_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: page_elements,
                });
                cursor.set(page_top);
            }
        }

        // COM-confirmed (2026-04-16, 683f p2 + minimal repro): content paragraphs
        // adjacent to a RUN of ≥2 consecutive empty paragraphs get +0.5pt extra advance.
        // Only applies to LM0 no-grid single spacing. Skip if paragraph caused page break.
        if adjacent_to_empty_run && is_single_lm0 && grid_pitch.is_none()
            && (cursor.cursor_y - page_top).abs() > 0.1 {
            cursor.advance(0.5);
        }

        let space_after = if let (Some(al), Some(pitch)) = (para.style.after_lines, grid_pitch) {
            // afterLines: exact value (al/100 * pitch), no grid snap needed.
            al / 100.0 * pitch
        } else {
            para.style.space_after.unwrap_or(0.0)
        };
        // NOTE: space_after is NOT added to cursor_y here.
        // It will be collapsed with the next paragraph's space_before via max(sa, sb).

        // Paragraph borders (e.g., Title style bottom border)
        if let Some(ref borders) = para.style.borders {
            let para_top = elements.first().map(|e| e.y).unwrap_or(start_x);
            let para_bottom = cursor.cursor_y;
            let border_x = start_x;
            let border_width = content_width;

            if let Some(ref bottom) = borders.bottom {
                let bw = bottom.width;
                let color = bottom.color.clone().unwrap_or_else(|| "000000".to_string());
                let border_y = para_bottom + bottom.space;
                elements.push(LayoutElement::new(border_x, border_y, border_width, bw.max(0.5), LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
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
                let pbdr_full = std::env::var("OXI_S467_PBDR_ENABLE").is_ok();
                cursor.set(border_y + if pbdr_full { bw } else { bw / 2.0 });
            }
            if let Some(ref top) = borders.top {
                let bw = top.width;
                let color = top.color.clone().unwrap_or_else(|| "000000".to_string());
                let border_y = para_top - top.space - bw;
                elements.push(LayoutElement::new(border_x, border_y, border_width, bw.max(0.5), LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
            // Between border (horizontal line between consecutive bordered paragraphs)
            if let Some(ref between) = borders.between {
                let bw = between.width;
                let color = between.color.clone().unwrap_or_else(|| "000000".to_string());
                let border_y = para_bottom + between.space;
                elements.push(LayoutElement::new(border_x, border_y, border_width, bw.max(0.5), LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
            // Left border
            if let Some(ref left) = borders.left {
                let bw = left.width;
                let color = left.color.clone().unwrap_or_else(|| "000000".to_string());
                let bx = border_x - left.space - bw;
                elements.push(LayoutElement::new(bx, para_top, bw.max(0.5), para_bottom - para_top, LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
            // Right border
            if let Some(ref right) = borders.right {
                let bw = right.width;
                let color = right.color.clone().unwrap_or_else(|| "000000".to_string());
                let bx = border_x + border_width + right.space;
                elements.push(LayoutElement::new(bx, para_top, bw.max(0.5), para_bottom - para_top, LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
        }

        if let Some(ref mut cells) = lm2_grid_cells {
            **cells = cumul_line_idx;
        }

        (elements, space_after)
    }

    #[allow(unused_assignments)]
    fn break_into_lines(
        &self,
        fragments: &[(&str, &RunStyle, Option<FieldType>, usize, usize)],
        available_width: f32,
        first_line_indent: f32,
        para_style: &ParagraphStyle,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
        lines_and_chars: bool,
        s476_body: bool,
        is_justified: bool,
        doc_grid_no_type: bool,
        para_has_lrpb: bool,
    ) -> Vec<Line> {
        // Helper: convert pt to twips for Word-GDI-compatible integer comparison
        let pt_to_tw = |pt: f32| -> i32 { (pt * 20.0).round() as i32 };
        let available_tw = pt_to_tw(available_width);

        // Day 33 part 19 (2026-05-10): paragraphs containing ONLY whitespace
        // (ASCII space, tab, U+3000 fullwidth space, etc.) render as a single
        // line in Word regardless of total natural width. COM-confirmed via
        // WS_10 / WS_50 / WS_100 / WS_142 / WS_300 minimal repros (all 5
        // produce identical 1-line break boundaries with BEFORE→AFTER advance
        // = 31pt = 1 line each). MIX_50_TEXT (50 spaces + text) DOES wrap
        // normally → rule is binary at paragraph level.
        // This is the safe inverse of commit 82de3fa (reverted 2026-05-03)
        // which used a per-character "trailing U+3000 immune" flag that
        // propagated to ALL U+3000s in a paragraph, regressing d77a mid-text
        // U+3000 indentation use. The all-whitespace gate is paragraph-binary
        // and never affects mixed-content paragraphs.
        let para_all_whitespace = fragments.iter().all(|(text, _, _, _, _)| {
            text.chars().all(|c| c.is_whitespace() || c == '\u{3000}')
        }) && fragments.iter().any(|(text, _, _, _, _)| !text.is_empty());

        let mut lines = Vec::new();
        let mut current_line = Line { fragments: vec![], ..Default::default() };
        let mut current_width = first_line_indent;
        // Integer twips accumulator for line break decisions.
        // Avoids f32 rounding drift that causes ±0.1pt error over 40+ characters.
        let mut current_width_tw: i32 = pt_to_tw(first_line_indent);
        // S475 (2026-06-01): capacity-adjusted break width = Σ(natural_adv −
        // max_yakumono_compress). Runs parallel to current_width_tw; only CONSULTED
        // when s475_break is ON (else default byte-identical). The break accepts a
        // char iff this accumulator (incl the char) ≤ available_tw — greedy first-fit
        // with punct-only demand compression folded into the fit width. See
        // session471 finding + workflow wtvi6fvix.
        let mut current_capw_tw: i32 = pt_to_tw(first_line_indent);
        let mut compress_used = false; // true after compression-based overflow absorption
        // S243 (2026-05-24): removed dead variable `current_grid_extra`
        // (assigned/incremented in 8 sites but never read).

        // Word buffer spans across fragment boundaries so that a single word
        // split across two runs (e.g. "te" in Run1 + "st" in Run2) is kept
        // together for line-break decisions.
        let mut word = String::new();
        let mut word_width: f32 = 0.0;
        let mut word_natural_width: f32 = 0.0; // 2-pass wrap: natural (pre-compression) width
        // S245 (2026-05-24): removed dead variable `word_grid_extra`
        // (assigned/incremented at 3 sites but never read after S243
        // removed `current_grid_extra`).
        let mut word_style: Option<RunStyle> = None;
        let mut word_field_type: Option<FieldType> = None;
        let mut word_run_index: usize = 0;
        let mut word_char_offset: usize = 0;

        // Helper: flush the accumulated word into current_line, breaking if needed.
        macro_rules! flush_word {
            ($style:expr) => {
                if !word.is_empty() {
                    let ws = word_style.take().unwrap_or_else(|| $style.clone());
                    let wft = word_field_type.take();
                    // COM-confirmed (2026-04-14): charGrid extra does NOT affect line
                    // break. Word wraps based on natural char widths (fontSize for
                    // fullwidth, smaller for halfwidth). Grid extra only affects
                    // character positioning within the line, not line break count.
                    let word_width_tw = pt_to_tw(word_width);
                    // Day 33 part 19: skip wrap break for all-whitespace paragraphs.
                    if current_width_tw + word_width_tw > available_tw && !current_line.fragments.is_empty()
                        && !para_all_whitespace {
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0; current_width_tw = 0; current_capw_tw = 0; compress_used = false;
                    }
                    current_line.fragments.push(LineFragment {
                        text: std::mem::take(&mut word),
                        width: word_width,
                        natural_width: word_natural_width,
                        style: ws,
                        tab_alignment: None,
                        tab_position: None,
                        field_type: wft,
                        run_index: word_run_index,
                        char_offset: word_char_offset,
                    });
                    current_width += word_width;
                    current_width_tw += word_width_tw;
                    current_capw_tw += word_width_tw; // S475: words have no punct capacity
                    word_width = 0.0;
                    word_natural_width = 0.0;
                }
            };
        }

        let n_fragments = fragments.len();
        // S547 (2026-06-12): w:kern resolved per paragraph (any fragment rPr
        // kern>0, else the paragraph/docDefaults default-run kern). Gates the
        // yakumono pair halving (break-time rule AND the S532 Stage-2 revert
        // protection). See the gate comment at yakumono_pair_enabled below.
        let para_kern_on = fragments.iter().any(|&(_, st, _, _, _)|
                st.kern.map_or(false, |k| k > 0.0))
            || para_style.default_run_style.as_ref()
                .and_then(|rs| rs.kern).map_or(false, |k| k > 0.0);
        let s547_kern_gate = std::env::var("OXI_S547_DISABLE").is_err();
        // S466 (2026-05-31, SHIPPED default-ON, opt-out OXI_S466_DISABLE):
        // hoisted once — see h8_trigger comment below. (a) compute positive grid
        // expansion for fs>=default, AND (b) apply it to char_width so the WRAP
        // (chars/line) matches Word, not just positioning. Pairs with the parser
        // raw_pitch change (ooxml.rs). Gate (drift-free OFF-vs-ON same binary):
        // charGrid family +0.0019, bottom-N floor up (tokumei p4/p5), only
        // tokumei p7 ×4 regress (above the floor); Phase-1 54/55 preserved.
        let s466_grid_expand = std::env::var("OXI_S466_DISABLE").is_err();
        for (frag_outer_idx, &(text, style, frag_field_type, frag_run_index, frag_char_start)) in fragments.iter().enumerate() {
            let font_size = self.resolve_font_size(style, para_style);
            let mut char_pos_in_run = frag_char_start;

            // fitText runs: skip GDI snap to preserve exact target width
            let cs = if style.fit_text.is_some() {
                style.character_spacing.unwrap_or(0.0)
            } else {
                snap_character_spacing(style.character_spacing.unwrap_or(0.0))
            };

            // Pre-resolve font metrics and GDI width maps for this fragment.
            // Avoids repeated font family resolution and HashMap lookups per character.
            let latin_metrics = self.metrics_for(style, para_style);
            let cjk_metrics = self.metrics_for_cjk(style, para_style);
            let latin_gdi_map = self.registry.get_gdi_char_widths(&latin_metrics.family, font_size);
            let cjk_gdi_map = cjk_metrics.map(|m| self.registry.get_gdi_char_widths(&m.family, font_size)).flatten();

            // Yakumono compression flags (約物詰め): COM-confirmed (2026-04-08,
            // refined 2026-04-18 bisect).
            // Trigger requires BOTH:
            //   - w:characterSpacingControl = "compressPunctuation" or "compressPunctuationAndJapaneseKana"
            //   - w:compat/w:compatSetting compatibilityMode >= 15 (Word 2013+)
            // See RESEARCH_LOG.md 2026-04-18 and pipeline_data/d77a_yakumono_bisect.json:
            //   - cSC alone: NO compression (minimal repro confirmed)
            //   - cSC + compat15: yakumono pair compression applies
            // Most modern docs use "doNotCompress" → no compression regardless of compat.
            //
            // 2026-04-21 update: COM evidence on 4 distinct compat=14+cP docs
            // (04b88, 7f272a, fded68, 34140) showed Word fits +1 to +3 more chars
            // on line 1 of yakumono+indent paragraphs vs Oxi. Minimal repro
            // (idx46_real with compat=14) confirmed Word=43 / Oxi=42. The
            // `compat>=15` gate excluded compat=14 docs that Word DOES compress.
            // Drop the compat gate — `compress_punctuation` alone matches
            // Word's behavior for both compat 14 and compat 15.
            //
            // 2026-04-27 attempted-but-reverted: V_CP + V_COMPAT15 8x8 matrix
            // measurement (180 fixtures) showed Word applies next-trigger
            // compression unconditionally on isolated 4-char paragraphs across
            // compat∈{14,15} × cSC∈{doNotCompress,compressPunctuation} ×
            // useFELayout∈{on,off} × kern∈{on,off}. Patched to
            // `yakumono_enabled = true` and ran pipeline.verify on 177 docs:
            // **18 page regressions, net -2.0184, bottom-5 floor 3.2645→2.9337
            // (-0.3308 catastrophic)**. Two docs collapsed:
            //   - 0e7af1ae8f21 pages 2-7,8,10: -0.18 to -0.29 each
            //   - 683ffcab86e2 pages 1-3: -0.04 to -0.25
            // 6 e3c5 pages improved (+0.01 to +0.03) but vastly outweighed.
            // **Implication**: COM 4-char isolated fixtures do not generalize
            // to multi-line real-world paragraphs. Word's actual gate involves
            // additional context (line position, surrounding chars, paragraph
            // structure?) NOT captured in the 8x8 grid. Reverted on 2026-04-27.
            // See RESEARCH_LOG 2026-04-27 falsified entry for full data.
            //
            // Day 34 part 23 (2026-05-13): COM measurement of e3c545 idx=29
            // (Meiryo 10.5pt, csControl=doNotCompress) showed Word DOES
            // compress the 、 of 、「 pair to 5.25pt (half). This contradicts
            // the OOXML doNotCompress flag — Word applies pair compression
            // when the CJK font has hwid (halfwidth-glyph) support.
            // Two-tier gate:
            //   - PAIR rule (close+open, ×0.5): compress_punctuation OR hwid
            //   - FULL rules (expand pair, standalone, line-start): compress_punctuation only
            //     (preserves existing behavior; 2026-04-27 unconditional broke MS Mincho docs)
            let cjk_font_has_hwid = cjk_metrics
                .map(|m| crate::font::font_supports_hwid(&m.family))
                .unwrap_or(false);
            // S547 (2026-06-12): the pair-halving gate is w:kern — NOT
            // compressPunctuation and NOT compat. 2×2×2 COM matrix
            // (_s547b_gate_matrix.py): kern=2 halves 、（/（「 even under
            // doNotCompress at any compat; kern absent never halves even with
            // compressPunctuation (the full 26×26 sweep at kern=0 had ZERO
            // non-natural pairs). S532's "unconditional" was measured on
            // kern=2 docs. kern is a RUN property; resolved per fragment:
            // run rPr → paragraph-style chain (default_run_style, basedOn
            // merge carries kern) → docDefaults (merge in ooxml.rs). The
    // pair scan below is per-fragment (chars_vec), so fragment
            // granularity is exact. Opt-out OXI_S547_DISABLE restores the
            // pre-S547 compress_punctuation gate.
            let frag_kern_on = style.kern
                .or_else(|| para_style.default_run_style.as_ref().and_then(|rs| rs.kern))
                .map_or(false, |k| k > 0.0);
            let yakumono_pair_enabled = if s547_kern_gate {
                frag_kern_on || cjk_font_has_hwid
            } else {
                self.compress_punctuation || cjk_font_has_hwid
            };
            let yakumono_enabled = self.compress_punctuation;
            // S472 (2026-06-01) DEMAND-DRIVEN yakumono refactor (user chose the big
            // refactor path). Word does NOT pre-compress standalone 、 at wrap time:
            // it uses NATURAL fullwidth (12pt) for the break decision, then compresses
            // 、 trailing space ONLY as much as a line's justify-slack demands
            // (COM: b837 p1 、 = 8.0-12.0pt variable per line; d77a divergent line
            // 、 = 11.2 = LIGHT compression, not the flat ×0.6667=8.0 Oxi pre-applies).
            // Oxi's flat pre-compression over-packs (fits 1 extra char/、-heavy line).
            // When enabled: (1) standalone 、，use natural width at break, (2) the
            // overflow-absorb budget becomes (count of standalone 、 on line)×(fs/3)
            // [max 4pt each at 12pt], (3) on absorb the line's 、 are retroactively
            // compressed by the absorbed overflow so the line fits exactly = matches
            // Word's demand-driven per-line compression. Default OFF (byte-identical)
            // for the canary. See session470 finding.
            // S473 (2026-06-01): break-flip-derived budget. Word's break-time punct
            // compression is DEMAND-driven up to a CAP of ~3.25pt/compressible
            // (fs×0.27, NOT fs/3=4.0), measured via repros/breakflip + d77a p1/p9.
            // s473_locomp implies the s472 upstream (leave 、 natural) + the s472
            // render water-fill, and additionally swaps the break budget to the
            // cap-based, no-0.95-exclusion model. Default OFF (byte-identical).
            // S474 (2026-06-01): pure-natural break diagnostic. Disables ALL
            // break-time standalone-punct compression (the ×0.6667 AND the
            // line-start narrow-yakumono reduction) and the demand-absorb, keeping
            // only the always-on pair compression. Renders the natural-greedy line
            // counts (Ng) = the count if Word broke at natural widths. Used to test
            // "Word breaks at natural, punct compression is render-only" and to
            // derive the fullness-gate rule (Ng vs Word count). Default OFF.
            // S492 (2026-06-03) — jc=left disentanglement (the R35 multi-session
            // refactor, user option A). Word does NOT apply yakumono break-time
            // compression to NON-justified paragraphs: jc=left/right/center break
            // at NATURAL widths + kinsoku (punct compression is justify-specific).
            // MEASURED decisively (tools/metrics/measure_jc_disentangle*.py):
            //   - jc=left synthetic repro = punct 12.0 natural, zero compression,
            //     at every punct density (10-50%); jc=both packs +1 via burasagari
            //     (hangable punct hangs past margin; mid-line punct stays 12.0) or
            //     light opener compression (「→11.25), NOT distributed K-compression.
            //   - Real docs 683f/0e7af/d77a (docGrid type=lines) OVER-PACK +1 on
            //     jc=left wrapping lines because Oxi's S475 capacity break is ungated
            //     on alignment; b837 (linesAndChars, grid-determined) is unaffected.
            // The render water-fill is ALREADY jc-gated (it lives inside the
            // should_justify block). Only the BREAK side leaks. When set + the para
            // is NOT justified, run the validated s474_natural pure-natural-greedy
            // path (disables standalone compression + absorb) and disable the S475
            // capacity break. Default OFF (byte-identical). Phase-1-sensitive (fewer
            // chars/line on jc=left → more lines → pagination shift) → env-gated +
            // full canary before ship.
            // S539 (2026-06-11): SHIPPED default-ON, scoped to NON-linesAndChars
            // (the former OXI_S492_JCNATURAL + OXI_S492_LINESONLY config).
            // The S492-era blocker was 3a4f p2-p5 SSIM regressions; S539 traced
            // them to the style-basedOn jc-inheritance bug (parser/styles.rs):
            // paragraphs Word justifies (Normal jc=both via pStyle chain) were
            // resolved jc=left by Oxi, so the natural break wrongly rewrapped
            // them. With jc resolution fixed, the full-corpus gate is clean:
            // SSIM 1 up (d77a p2 +0.0090) / 409 unchanged / 0 regress,
            // Phase-1 54/55 with 3a4f histogram identical to baseline.
            // linesAndChars (b837 family) stays EXCLUDED: full-scope round-1
            // gate showed b837 pagination cascade 0.9997->0.5775 (30 paras +1
            // page; p5-p7 SSIM -0.53) even though p1-p4 improved +0.039 — the
            // b837 jc=left grid-line break needs its own investigation before
            // widening (OXI_S492_JCNATURAL still forces the full scope).
            // Opt-out: OXI_S492_DISABLE restores capacity-break for all paras.
            let s492_full = std::env::var("OXI_S492_JCNATURAL").is_ok();
            // S572 (2026-06-14): extend the S568 LEGACY jc=left 約物 OIKOMI to
            // NO-TYPE docGrid docs (S568 covered only linesAndChars). ikujidetail
            // (compat=11, no-type docGrid linePitch=286, jc=left, compressPunctuation)
            // compresses a tight line's 約物 to fit — COM (_s572_charadv): on the
            // over-wrapping para i=199 Word renders the mid-line 、 at 9.0pt and the
            // line-end 。 at 5.25pt (half-width) where a slack line (i=5) keeps both
            // at 11.25 = DEMAND-DRIVEN oikomi, not general 約物詰め. Oxi broke at
            // NATURAL (natural_break_jc) → the trailing char over-wrapped 1→2 lines
            // (i=199/231/419) → +1×27. Same discriminator as S568 (compat<15 +
            // compressPunctuation + non-justified), just no-type instead of
            // linesAndChars. SCOPE: the ONLY compat<15 compressPunctuation no-type
            // docGrid doc in the corpus is ikujidetail (single-doc-scoped, like
            // S568). Opt-out OXI_S572_DISABLE.
            let s572_legacy_notype_oikomi = std::env::var("OXI_S572_DISABLE").is_err()
                && doc_grid_no_type && s476_body && !is_justified
                && self.compress_punctuation && self.compat_mode < 15;
            let natural_break_jc = std::env::var("OXI_S492_DISABLE").is_err()
                && !is_justified
                && (!lines_and_chars || s492_full)
                && !s572_legacy_notype_oikomi;
            let s474_natural = std::env::var("OXI_S474_NATURAL").is_ok() || natural_break_jc;
            // S589 (2026-06-16, opt-IN OXI_S589=1, default OFF = byte-identical):
            // LEGACY (compat<15) JUSTIFIED body paras break standalone 、。，． at
            // NATURAL width instead of the flat ×0.6667 pre-compress (mod.rs:6764)
            // that over-packs 、-heavy lines by ~1 char. ROOT of tokyoshugyo #2:
            // _tks_oidashi.py (char-stream-aligned Word-PDF vs Oxi) localized the
            // 賃金 chapter over-fit to 102 "Word-NAT / Oxi-½" mid-、 — Oxi compresses
            // 、 at break, Word breaks at natural (compressing only on justify-slack).
            // natural_break_jc/s557 cover only !is_justified / c15; legacy JUSTIFIED
            // (compat<15, type=lines) was uncovered → ×0.6667 fired. compat=11 docs:
            // tokyoshugyo. See [[char_budget_wall]], [[tokyoshugyo_wrap_not_cellheight]].
            let s589_legacy_just_natural = std::env::var("OXI_S589").ok().as_deref() == Some("1")
                && is_justified && self.compress_punctuation
                && self.compat_mode < 15 && !lines_and_chars;
            // S557 (2026-06-13, part of the OXI_S556_JUST15 opt-in scaffold):
            // c15-explicit JUSTIFIED paragraphs keep standalone 、。，．at
            // NATURAL width at break — Word defers ALL their compression to
            // the per-line pack decision (d77a para9-L6 ground truth: Word's
            // 38-char line = naturals overflowing 2.5 → packed −0.75×3 onto
            // 、）、; Oxi's legacy ×0.6667 pre-compress baked −4/punct into
            // width AND natural_width, blinding the pack tier's need (7.0 vs
            // Word's 14.5 for the 39th char) and its all-natural guard).
            let s557_natural_just15 = std::env::var("OXI_S556_JUST15").is_ok()
                && is_justified
                && self.compress_punctuation
                && self.compat_mode >= 15 && self.compat_mode_explicit
                && !lines_and_chars;
            // S475 capacity-budget break (env-gated, default OFF = byte-identical).
            // Greedy first-fit where each punct contributes break-compression CAPACITY
            // (pair-first 6.0 / solo 1.5, env-tunable; flat-K = equal). Bypasses the
            // ×0.6667 standalone pre-compress + the S472/S473 absorb, and routes render
            // through the s472_render water-fill so glyphs justify correctly.
            // SCOPED to NO-char-grid sections only (docGrid type=lines / none →
            // grid_char_pitch is None). docGrid type=linesAndChars (grid_char_pitch
            // Some, e.g. b837) is GRID-determined (fixed char count/line) — a
            // SEPARATE mechanism (charGrid charsLine); S475 yakumono capacity must
            // NOT apply there (it cascaded b837 7→9 pages). See session471 finding.
            // S475 SHIPPED default-ON (2026-06-01, opt-out OXI_S475_DISABLE). flatK
            // params (PAIR=SOLO=2.5) — the break-decision capacity; reproduces d77a
            // [39,38,40,41,…]-class packing on type=lines docs. Gate: Phase-1 54/55
            // (no PASS→FAIL, b837 7pg), SSIM net +0.0398, bottom-5 +0.0109 (d77a
            // +0.062, ed025c +0.097). c7b923 −0.036 = latin-mixed residual (lever B).
            // S476 (2026-06-02): extend the S475 yakumono capacity break to
            // linesAndChars docs' MAIN BODY (b837/b35/tokumei = corpus bottom-N).
            // Lever C count-cap was FALSIFIED (Oxi charsLine already = Word's; Word
            // does NOT cap; Oxi UNDER-packs linesAndChars). The real gap is the SAME
            // yakumono per-line packing as S475 — Word fits more per line by
            // compressing punct. linesAndChars needs a heavier cap (K≈3.0 vs lines'
            // 2.5; the char-grid context). COM-verified Phase-1-safe: the whole
            // linesAndChars family (b837/b35/1636/31420/6514/a1d6/87b29/29dc6e/1ec1)
            // keeps its baseline=Word page count at K=3.0. b837 +0.0455. aux/cell
            // calls (s476_body=false) stay excluded to avoid the 7→9 cell cascade.
            // S568 (2026-06-14): LEGACY (compat ≤14) linesAndChars compressPunctuation
            // docs apply jc=left 約物 OIKOMI (the s476 capacity break) that the
            // compat≥15 gate excludes. harassmanual (compat=11) orphans a trailing
            // char (く) that Word fits by compressing a mid-line 読点 、 to half-em
            // (COM _s568_p16_adv: 、 advance 6.0pt, every other char 12.0). The
            // discriminator is compat: modern (15) jc=left breaks at NATURAL widths
            // (S492/S539 measured + shipped), legacy demands oikomi (see the
            // compat_mode_explicit note at mod.rs:1512). The ONLY compat<15
            // linesAndChars compressPunctuation doc in the corpus is harassmanual
            // (compat=14 docs are type=lines, not linesAndChars), so this is a
            // single-doc-scoped change. Cap = full half-em (6.0). Opt-out OXI_S568_DISABLE.
            let s568_legacy_oikomi = std::env::var("OXI_S568_DISABLE").is_err()
                && lines_and_chars && s476_body
                && self.compress_punctuation && self.compat_mode < 15;
            let s476_grid = (std::env::var("OXI_S476_DISABLE").is_err()
                && lines_and_chars && s476_body
                && self.compress_punctuation && self.compat_mode >= 15)
                || s568_legacy_oikomi
                || s572_legacy_notype_oikomi;
            // S590 (2026-06-16, opt-IN OXI_S590=1, default OFF = byte-identical):
            // LEGACY (compat<15) JUSTIFIED body paras use the s475 CAPACITY break
            // (greedy + compress-to-fit only when overflow ≤ Σ約物-caps, cap≈2.5)
            // instead of the flat ×0.6667 pre-compress (−3.5pt, over-compresses) OR
            // S589 pure-natural (0, under-fits the real oikomi lines). DERIVED:
            // _tks_oidashi.py --absorb on the Word PDF — Word's per-約物 oikomi
            // compression caps at ~2.9pt (median 1.93), and only 14/219 full lines
            // compress (176 expand at natural 約物). So the capacity model with
            // cap≈2.5 (≈ Word max) reproduces Word's compress-14/expand-176 split,
            // unlike ×0.6667 (over) / S589-natural (under) / S543-fs/2 (way over).
            let s590_legacy_just_cap = std::env::var("OXI_S590").ok().as_deref() == Some("1")
                && is_justified && self.compress_punctuation
                && self.compat_mode < 15 && !lines_and_chars;
            let s475_break = ((std::env::var("OXI_S475_DISABLE").is_err()
                && self.compress_punctuation && self.compat_mode >= 15
                && !lines_and_chars)
                || s476_grid || s590_legacy_just_cap)
                && !natural_break_jc;  // S492: non-justified paras break at natural
            let s476_cap: f32 = std::env::var("OXI_S476_CAP").ok()
                .and_then(|v| v.parse().ok())
                .unwrap_or(if s568_legacy_oikomi || s572_legacy_notype_oikomi { 6.0 } else { 3.0 });
            // S558 (2026-06-13): s475_pair default 2.5 → 6.0. A CLOSING bracket
            // before another bracket collapses a full half-em at break (matching
            // the render pair-halving); the old 2.5 under-counted it, so
            // bracket-cluster justified lines (d77a para9 L3 ）」（) broke a char
            // early — an SSIM cascade. Comma/period-first pairs still trim lightly
            // (s475_max_compress split — see kinsoku.rs). Env-tunable.
            // S590 refinement (2026-06-16): per-TYPE caps — solo (、。，．) = 1.5
            // (the derived break demand), but bracket-PAIR clusters keep 6.0 (S558,
            // the lever-3 heavy cluster compression). Measured (S591 cells clamped,
            // break divergence): solo1.5/pair1.5=593 → solo1.5/pair6.0=542 (best).
            let s475_pair: f32 = if s476_grid { s476_cap } else {
                std::env::var("OXI_S475_PAIR").ok().and_then(|v| v.parse().ok())
                    .unwrap_or(6.0) };
            // S575 (2026-06-15): BODY oikomi — raise the solo 約物 cap to 3.0 for the
            // MAIN body flow (s476_body) so jc=both type=lines compat=15 bodies fit
            // Word's demand compression (ikujikaigo i=41/i=57: mid 、 renders 9.0 = −3.0;
            // +1×4 → PASS). GATED to paras WITHOUT a lastRenderedPageBreak: the higher cap
            // REDISTRIBUTES chars within a para's lines (line COUNT unchanged), which SHIFTS
            // which line a run's char_offset==0 lands on → the S391 per-line-LRPB respect
            // then fires on a DIFFERENT line → a spurious mid-para page break (d77a's
            // "イは、編集…" para, 1 LRPB: +18pt continuation cascade → cell over page 6/7 =
            // the 16-session d77a blocker, isolated via OXI_S391_PER_LINE_LRPB=0 + env
            // bisection). ikujikaigo i=41/i=57 have 0 LRPBs → safe to redistribute. The
            // break itself is solo-STABLE (count unchanged); only the LRPB attribution is
            // sensitive, so skip the oikomi when there's an LRPB to preserve. Opt-out
            // OXI_S575_DISABLE.
            // S590: derived break-time 約物 demand cap ≈ 1.5pt (sweep minimum,
            // _tks_oidashi: divergence 666@0 → 618@1.5 → 925@2.5). Word's BREAK
            // cap (~1.5) < its RENDER cap (~2.9) — break is conservative, render
            // (justify) compresses more. Env OXI_S475_SOLO overrides.
            let s475_solo_default = if s590_legacy_just_cap { 1.5 }
                else if s476_body && !para_has_lrpb
                && std::env::var("OXI_S575_DISABLE").is_err() { 3.0 } else { 2.5 };
            let s475_solo: f32 = if s476_grid { s476_cap } else {
                std::env::var("OXI_S475_SOLO").ok().and_then(|v| v.parse().ok()).unwrap_or(s475_solo_default) };
            let s473_locomp = std::env::var("OXI_S473_LOCOMP").is_ok();
            let s473_cap: f32 = std::env::var("OXI_S473_CAP").ok()
                .and_then(|v| v.parse().ok()).unwrap_or(3.25);
            // S473b (2026-06-01): per-type break caps. Render-advance COM showed
            // Word compresses brackets/。 ~4× more than 、 (）=remove 6.0 vs 、=remove
            // 1.5). A UNIFORM cap could not reconcile d77a (bracket-heavy lines, want
            // heavy) with b837 pi30 (、-only line, wants light). Per-type caps (env-
            // tunable for the sweep): comma/opening-bracket = light, period/closing-
            // bracket = heavy. Defaults from render data (1.5 / 6.0). Used only when
            // OXI_S473_ASYM is set (else the uniform s473_cap path runs).
            let s473_asym = std::env::var("OXI_S473_ASYM").is_ok();
            let s473_cc: f32 = std::env::var("OXI_S473_CC").ok()
                .and_then(|v| v.parse().ok()).unwrap_or(1.5);   // 、，
            let s473_cp: f32 = std::env::var("OXI_S473_CP").ok()
                .and_then(|v| v.parse().ok()).unwrap_or(6.0);   // 。．
            let s473_ccl: f32 = std::env::var("OXI_S473_CCL").ok()
                .and_then(|v| v.parse().ok()).unwrap_or(6.0);   // closing brackets
            let s473_cop: f32 = std::env::var("OXI_S473_COP").ok()
                .and_then(|v| v.parse().ok()).unwrap_or(1.5);   // opening brackets
            let s472_demand = std::env::var("OXI_S472_DEMAND").is_ok() || s473_locomp;
            let chars_vec: Vec<char> = text.chars().collect();
            // Yakumono pair compression for line break width calculation.
            // Rule 1 (close+open ×0.5) is gated by yakumono_pair_enabled
            // (compress_punctuation OR hwid font); Rules 2-4 below use
            // yakumono_enabled (compress_punctuation only).
            let yakumono_compressed: Vec<bool> = if yakumono_pair_enabled {
                let n = chars_vec.len();
                let mut v = vec![false; n];
                for i in 0..n {
                    let c = chars_vec[i];
                    if kinsoku::is_yakumono_closing(c) {
                        if i + 1 < n && kinsoku::is_yakumono_trigger(chars_vec[i + 1]) {
                            v[i] = true;
                        }
                    } else if kinsoku::is_yakumono_opening(c) {
                        if i > 0 && kinsoku::is_yakumono_trigger(chars_vec[i - 1]) && !v[i - 1] {
                            v[i] = true;
                        }
                    }
                }
                v
            } else {
                vec![false; chars_vec.len()]
            };

            for (char_index, ch) in chars_vec.iter().copied().enumerate() {
                let (char_metrics, gdi_map) = if kinsoku::is_cjk(ch) {
                    if let Some(cjk_m) = cjk_metrics {
                        (cjk_m, cjk_gdi_map)
                    } else {
                        (latin_metrics, latin_gdi_map)
                    }
                } else {
                    (latin_metrics, latin_gdi_map)
                };
                let mut char_width = self.registry.char_width_pt_with_gdi_map(ch, font_size, char_metrics, gdi_map);
                if let Some(scale) = style.text_scale {
                    if (scale - 100.0).abs() > 0.01 {
                        char_width *= scale / 100.0;
                    }
                }
                char_width += cs;
                // §17.15.1.7 balanceSingleByteDoubleByteWidth (Session 56 Finding 3,
                // COM-confirmed via V19/V25/V26/V27 minimal repros 2026-05-06):
                // when this compat flag is set, character_spacing is applied TWICE
                // for CJK fullwidth chars (effective_cs = 2 * cs). Apply the extra
                // cs here so per-char fragment advance reflects the doubled spacing.
                // Day 37 (2026-05-14): EXCLUDE fitText runs — resolve_fit_text_runs
                // already produces the FINAL effective cs (post-balance-doubling) so
                // adding here would over-pump by another factor.
                if self.balance_single_byte_double_byte_width
                    && crate::font::is_fullwidth(ch)
                    && !yakumono_compressed[char_index]
                    && style.fit_text.is_none()
                {
                    char_width += cs;
                }
                // 2-pass wrap: remember pre-yakumono width to compute yakumono savings.
                let pre_yakumono_width = char_width;
                // Physical yakumono compression (COM-confirmed b837 2026-04-16):
                //   Pair (both chars): 6pt (×0.5) — e.g., 。）→ 6+6pt
                //   Standalone 、。 between non-trigger CJK: 7pt (×0.583)
                //   Other brackets: use native font width (bracket shapes vary widely by
                //     context in Word — 6, 10.5, 11, 11.5, 12pt — no simple compression rule)
                // 2026-04-20: Opening brackets have visible glyph at right side of
                // cell (ABC A-offset = 7.5pt for 「, 11pt for （). Compressing advance
                // to 6pt would place next char at 6pt offset, overwriting the bracket
                // glyph at 7.5-11.25pt. Skip compression for these — keep fullwidth
                // advance so glyph fits within its cell. Closing brackets (A=0) are
                // unaffected and still compress fine.
                let is_opening_bracket = matches!(ch,
                    '（' | '「' | '『' | '〔' | '【' | '《' | '〈' | '｛' | '［'
                );
                if yakumono_compressed[char_index] && !is_opening_bracket {
                    char_width *= 0.5;
                } else if yakumono_enabled {
                    // S532 (2026-06-10): the former "expand pair" rule (a yakumono
                    // ADJACENT to a pair-compressed one also compresses ×0.5) is
                    // REMOVED — Word compresses ONLY the FIRST char of an adjacent
                    // pair; the second keeps its natural advance. Measured
                    // (_s532_pair_repro.py, MS Gothic 12pt, PDF per-char origins):
                    // 。」=6.00/12.00, ）」=6.00/12.00, 。「=6.00/12.00 — identical
                    // in centered, loose-justified and wrapping-justified lines.
                    // (The 2026-04-16 b837 "。）→6+6" COM note conflated the pair
                    // rule with justify-demand compression of the second char.)
                    let is_yakumono_any = matches!(ch,
                        '（' | '）' | '「' | '」' | '『' | '』' | '〔' | '〕' |
                        '【' | '】' | '《' | '》' | '〈' | '〉' | '｛' | '｝' |
                        '［' | '］' | '、' | '。' | '，' | '．'
                    );
                    if is_yakumono_any {
                        if matches!(ch, '、' | '。' | '，' | '．') {
                            // Standalone 、 。 between non-triggers: spec §4.7b round 5
                            // floor = fontSize × 2/3. Trying 0.667 instead of 0.583.
                            let prev_non_tr = char_index == 0
                                || !kinsoku::is_yakumono_trigger(chars_vec[char_index - 1]);
                            let next_non_tr = char_index + 1 >= chars_vec.len()
                                || !kinsoku::is_yakumono_trigger(chars_vec[char_index + 1]);
                            if prev_non_tr && next_non_tr {
                                // S472: ALL standalone 、，。．use NATURAL width at break
                                // (Word defers compression to justify-demand; COM: 、/。
                                // standalone = near-full, only compressed on line-slack).
                                // The demand-absorb below compresses any of them as a
                                // line's overflow requires.
                                if (s472_demand || s474_natural || s475_break || s557_natural_just15
                                        || s589_legacy_just_natural)
                                    && matches!(ch, '、' | '，' | '。' | '．') {
                                    // no compression at break; demand-absorb handles fit
                                    // (s474_natural: leave natural, no absorb either =
                                    // pure natural-greedy diagnostic; s557: c15 justified
                                    // keeps naturals for the pack tier)
                                } else {
                                    char_width *= 0.6667;
                                }
                            }
                        }
                    }
                }
                // Line-start yakumono demand-driven compression (COM-verified 2026-04-21
                // on d77a pi=24-27 + 3a4f pi=300 + 1ec1/e3c5 no-overflow):
                // Word compresses ・/、/。 at line start by ~2.5pt at 12pt when the
                // line would otherwise overflow. Apply the compression speculatively;
                // Stage 2 revert (below) undoes it on short lines where
                // natural_total_width ≤ available_width (loose-line rule).
                // Compression: font_size × 5/24 = 2.5pt at 12pt, 2.1875pt at 10.5pt.
                // Gated on compress_punctuation + compat_mode>=15 to match Word 2016+.
                if yakumono_enabled
                    && self.compat_mode >= 15
                    && !s474_natural
                    && matches!(ch, '・' | '、' | '。' | '，' | '．')
                    && current_line.fragments.is_empty()
                    && word.is_empty()
                {
                    let reduction = font_size * 5.0 / 24.0;
                    let floor = char_width * 0.5;
                    char_width = (char_width - reduction).max(floor);
                }
                // 2-pass wrap: compute yakumono savings (difference between pre-yakumono
                // and post-yakumono width). Natural = final_char_width + yakumono_saved.
                let yakumono_saved = (pre_yakumono_width - char_width).max(0.0);
                // §4.6.3 CJK-adjacent space widening — COM-confirmed 2026-04-08.
                // The Latin space (' ') is widened to ≈ font_size/2 (half-em) when:
                //   1. The run's <w:rFonts> has an EXPLICIT w:eastAsia attribute
                //      (theme fallback eastAsiaTheme="..." does NOT count), AND
                //   2. The space is adjacent to a CJK ideograph or kana on
                //      either side.
                // Verified via jfmb (no explicit eastAsia, space=natural ~3.5pt) vs
                // runtime-saved equivalent (explicit eastAsia, space=6.0pt at 12pt).
                if ch == ' ' && style.has_explicit_east_asia {
                    let prev_is_cjk = chars_vec.get(char_index.wrapping_sub(1))
                        .copied()
                        .map_or(false, kinsoku::is_cjk_ideograph_or_kana);
                    let next_is_cjk = chars_vec.get(char_index + 1)
                        .copied()
                        .map_or(false, kinsoku::is_cjk_ideograph_or_kana);
                    if prev_is_cjk || next_is_cjk {
                        char_width = font_size / 2.0;
                    }
                }
                let _ = char_index;
                // charGrid: ONLY full-width chars are padded to 1 grid cell.
                // §11.2.1 (Round 14, COM-confirmed): half-width Latin chars
                // (ASCII 0-9 / A-Z / etc.), CJK punctuation under yakumono
                // compression, and other halfwidth glyphs use their NATURAL
                // advance width — they are NOT snapped to the grid pitch.
                // Reference: b837808d0555 P13 L1 measurement showed
                //   '2'=6pt, '」'=6pt (yakumono), 成=15pt (12+autoSpaceDE),
                //   '7'=9pt (6+autoSpaceDE), ' '=6pt (TNR space natural).
                // Previous (buggy) behavior padded ALL chars, halving the
                // chars/line and causing 177-doc max-error of 0.5366 SSIM.
                // 2026-04-19 (revised): cw = fs + charSpace_pt (absolute, not scaled).
                // COM-measured b35 fs=9→8.3pt, fs=10.5→9.8pt: both = fs − 0.7pt.
                // Previous fs*ratio formula over-compressed at small fs in docs
                // where default_fs ≠ fs.
                // fit_text EXPAND mode (natural ≤ target, character_spacing>0): skip
                // charGrid padding so Word's fitText cs applies verbatim. Without this,
                // the negative char_grid_extra swallows the cs for CJK chars and breaks
                // uniform spread (b837 p1 meta block).
                // fit_text SCALE mode (natural > target, text_scale<100) keeps charGrid
                // padding — otherwise scaled CJK chars in table cells become narrower
                // than the grid pitch, shifting downstream content (3a4f regression).
                let fit_text_expand = style.fit_text.is_some()
                    && style.character_spacing.map_or(false, |cs| cs > 0.01);
                let char_grid_extra = if fit_text_expand {
                    0.0
                } else if let (Some(ratio), Some(pitch)) = (grid_char_cw_ratio, grid_char_pitch) {
                    if ratio > 0.0 && pitch > 0.0 && char_width > 0.0
                        && ch != ' ' && ch != '\t' && ch != '\n'
                        && crate::font::is_fullwidth(ch)
                        && !yakumono_compressed[char_index]
                    {
                        let default_fs = pitch / ratio;
                        let char_space_pt = pitch - default_fs;
                        // R7.59 (Day 36 part 3, 2026-05-13): hybrid grid-extra formula.
                        // charSpace>=0 (expansion): proportional. COM-verified d4d126
                        //   w_i=245 fs=10 default=10.5 cs=+0.575: Word renders ~10.547pt
                        //   advance (proportional, NOT the 10.5pt COM Information(WD_HPOS)
                        //   reports — that's the snapped logical width). Old linear
                        //   cw = 10+0.575 = 10.575pt over-expanded → 1-line→2-line wrap
                        //   regression on 35-char paragraphs.
                        // charSpace<0 (compression): linear. COM-verified b35 fs=9
                        //   cs=-0.66: Word=8.3pt (= 9-0.7).
                        // 10tw-snap variant tested 2026-05-13: slightly worse SSIM
                        //   (+2.1481 net vs +2.1859 net) because Word's INTERNAL
                        //   rendering uses raw proportional advance, not snapped.
                        // S141 H6 (2026-05-20): OXI_H6_GRID_GATE=1 gates expansion to
                        //   only fire when font_size >= default_fs. Word doesn't expand
                        //   small-font (sz < default) cell text to grid pitch even
                        //   though Oxi did via this formula. COM-confirmed: 法人等 sz=10
                        //   cell in a1d6/d4d126/de6e/6514f all render 1 line in Word
                        //   (33 chars × 10pt = 330pt natural fits 345pt cell) but Oxi
                        //   wraps to 2 (33 × 10.555 expansion = 348.3pt overflows).
                        let h6_gate_enabled = std::env::var("OXI_H6_GRID_GATE").is_ok();
                        let h7_gate_enabled = std::env::var("OXI_H7_GRID_GATE_LE").is_ok();
                        // S145 H8 (2026-05-21): per LibreOffice ww8par.cxx ImportDop,
                        // MS_WORD_COMP_GRID_METRICS is SET unconditionally for all Word
                        // imported docs. LibreOffice itrform2.cxx then SKIPS grid kern
                        // portions when MS_WORD_COMP_GRID_METRICS && !vertical. So MS
                        // Word actually NEVER applies grid char-pitch for horizontal text.
                        // OXI_H8_NO_GRID_KERN=1 skips entirely (no font_size check).
                        // S148 (2026-05-21) H8 refinement: only skip POSITIVE expansion
                        // (kern portions). Word DOES apply negative compression (e.g.
                        // b35 charSpace=-2714 → chars narrower than natural).
                        // S239 (2026-05-23): removed OXI_LEGACY_GRID_KERN
                        // legacy env-var fallback during hardening pass.
                        // S466 (2026-05-31, env-gated test): the H8 "Word never
                        // applies grid char-pitch for horizontal text" came from
                        // reading LibreOffice source (ww8par/itrform2). Direct Word
                        // COM on a charSpace=1453 (tokumei grid) BODY repro
                        // CONTRADICTS it: MS Mincho 10.5pt(=default) fits 44 chars/
                        // line in Word but Oxi (H8 skip => natural advance) fits 46.
                        // i.e. Word DOES expand fullwidth chars to the grid pitch
                        // when fs >= default_fs (S141 already COM-confirmed Word does
                        // NOT expand when fs < default, the cell case). So the correct
                        // skip is fs < default, not unconditional. Gated so the corpus
                        // (charGrid family tokumei/b35/b837) can be A/B re-gated;
                        // Phase-1-sensitive (re-wrap moves pagination).
                        let h8_trigger = char_space_pt > 0.0 && (!s466_grid_expand || font_size < default_fs);
                        let h7_trigger = h7_gate_enabled && char_space_pt > 0.0 && font_size <= default_fs;
                        let h6_trigger = h6_gate_enabled && char_space_pt > 0.0 && font_size < default_fs;
                        // S466: when the docGrid has NO charSpace (char_space_pt≈0), the
                        // grid is line-pitch-only and Word does NOT horizontally expand
                        // chars. Under raw_pitch (S466) such a doc yields char_space_pt=0
                        // (b837: charSpace absent, default 12pt), which would otherwise
                        // fall through to expected_w=fs and widen natural<fs chars,
                        // over-wrapping (7->9). Skip expansion for the no-charSpace case.
                        let s466_no_grid = s466_grid_expand && char_space_pt < 0.01;
                        // S344 (2026-05-27): when S344 fed grid values through despite
                        // snap_to_grid=false, gate compression to fs < default_fs only.
                        // (Effective only when paired with S342/S344 pass-through at
                        // mod.rs:4073/4246.)
                        let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT").map(|v| v != "0" && v != "false").unwrap_or(false);
                        let s344_skip = s344_fs_gate
                            && !para_style.snap_to_grid
                            && font_size >= default_fs;
                        if h6_trigger || h7_trigger || h8_trigger || s344_skip || s466_no_grid {
                            0.0
                        } else {
                            let expected_w = if char_space_pt >= 0.0 {
                                font_size * pitch / default_fs
                            } else {
                                font_size + char_space_pt
                            };
                            expected_w - char_width
                        }
                    } else { 0.0 }
                } else { 0.0 };
                // S239 (2026-05-23): removed OXI_LEGACY_GRID_KERN legacy
                // env-var-only branch (was `else if let Some(pitch) = grid_char_pitch`
                // gated entirely by env::var().is_ok()).
                // For negative extras, fold into char_width directly so fragment
                // widths (positioning) reflect the shrink. For positive extras,
                // keep the existing separate-accumulator model (padding for positioning).
                if char_grid_extra < 0.0 {
                    char_width += char_grid_extra;
                } else if s466_grid_expand && char_grid_extra > 0.0 {
                    // S466: fold POSITIVE grid expansion into char_width so the wrap
                    // (chars/line) reflects Word's grid-pitch advance. Default-OFF
                    // keeps the legacy separate-accumulator positioning behavior.
                    char_width += char_grid_extra;
                }

                if ch == ' ' || ch == '\t' || ch == '\n' || ch == '\x0C' || ch == '\x0B' {
                    // Whitespace: flush word, then handle the whitespace
                    flush_word!(style);

                    if ch == '\n' || ch == '\x0C' || ch == '\x0B' {
                        // Set break type on the current line before pushing
                        let break_type = match ch {
                            '\x0C' => LineBreakType::PageBreak,
                            '\x0B' => LineBreakType::ColumnBreak,
                            '\n' => LineBreakType::SoftBreak,
                            _ => LineBreakType::Normal,
                        };
                        current_line.break_type = break_type;
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0; current_width_tw = 0; current_capw_tw = 0; compress_used = false;
                    } else {
                        // Space or tab
                        if ch == '\t' {
                            // COM-confirmed: tab positions are absolute from left margin.
                            // current_width is relative to the indent start, so we add
                            // indent_left to get the absolute position from margin.
                            let indent_left = para_style.indent_left.unwrap_or(0.0);
                            let abs_pos = current_width + indent_left;
                            let (next_pos, tab_align) = if !para_style.tab_stops.is_empty() {
                                para_style.tab_stops.iter()
                                    .find(|ts| ts.position > abs_pos + 0.01)
                                    .map(|ts| (ts.position, ts.alignment))
                                    .unwrap_or_else(|| {
                                        let tab_stop = self.default_tab_stop;
                                        (((abs_pos / tab_stop).floor() + 1.0) * tab_stop, TabStopAlignment::Left)
                                    })
                            } else {
                                let tab_stop = self.default_tab_stop;
                                (((abs_pos / tab_stop).floor() + 1.0) * tab_stop, TabStopAlignment::Left)
                            };
                            // Convert absolute tab position back to relative width
                            let next_relative = next_pos - indent_left;
                            let w = (next_relative - current_width).max(char_width);
                            current_line.fragments.push(LineFragment {
                                text: TAB_STRING.to_owned(),
                                width: w,
                                natural_width: w,
                                style: style.clone(),
                                tab_alignment: Some(tab_align),
                                tab_position: Some(next_pos),
                                field_type: None,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            current_width += w;
                        } else {
                            // Regular space
                            current_line.fragments.push(LineFragment {
                                text: SPACE_STRING.to_owned(),
                                width: char_width,
                                natural_width: char_width,
                                style: style.clone(),
                                tab_alignment: None,
                                tab_position: None,
                                field_type: None,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            current_width += char_width; current_width_tw += pt_to_tw(char_width);
                            current_capw_tw += pt_to_tw(char_width); // S475: space, no punct capacity
                        }
                    }
                } else if is_break_after(ch) {
                    // Characters like '-', '/' that allow a line break AFTER them.
                    // Include them in the current word, flush, and allow a break.
                    if word_style.is_none() {
                        word_style = Some(style.clone());
                        word_field_type = frag_field_type;
                        word_run_index = frag_run_index;
                        word_char_offset = char_pos_in_run;
                    }
                    word.push(ch);
                    word_width += char_width;
                    word_natural_width += char_width + yakumono_saved;
                    flush_word!(style);
                } else if kinsoku::is_cjk(ch) {
                    // CJK characters always break at char boundaries (subject to kinsoku).
                    // ECMA-376 §17.3.1.40: wordWrap controls LATIN word-break only.
                    // V_JJ measurement (2026-05-02) confirmed: V_JJ2 (wordWrap=on) and
                    // V_JJ3 (wordWrap=off) produce identical CJK break points.
                    // Pre-2026-05-03: this branch was gated on `&& para_style.word_wrap`,
                    // causing CJK to accumulate as a single non-breakable word in
                    // wordWrap=off paragraphs (34 baseline docs / 108 instances).
                    // autoSpaceDE: add 2.5pt gap between Latin and CJK ideograph/kana.
                    // COM-confirmed (2026-04-07): only ideographs/kana trigger auto-space,
                    // not CJK punctuation (which gets no extra spacing from Latin).
                    // Session 95 (2026-05-18) split: autoSpaceDE gates ALPHABETIC
                    // boundaries, autoSpaceDN gates DIGIT boundaries. e3c545 has
                    // DE=on, DN=off → digits should NOT get the gap. Was previously
                    // gated on auto_space_de alone via is_ascii_alphanumeric().
                    let prev_alpha_local = !word.is_empty() && word.chars().last().map_or(false, |c| c.is_ascii_alphabetic());
                    let prev_digit_local = !word.is_empty() && word.chars().last().map_or(false, |c| c.is_ascii_digit());
                    let (prev_frag_alpha, prev_frag_digit) = if word.is_empty() {
                        let last_char = current_line.fragments.last().and_then(|f| f.text.chars().last());
                        (
                            last_char.map_or(false, |c| c.is_ascii_alphabetic()),
                            last_char.map_or(false, |c| c.is_ascii_digit()),
                        )
                    } else { (false, false) };
                    let prev_is_alpha = prev_alpha_local || prev_frag_alpha;
                    let prev_is_digit = prev_digit_local || prev_frag_digit;
                    flush_word!(style);
                    let cur_is_cjk = kinsoku::is_cjk_ideograph_or_kana(ch);
                    if cur_is_cjk
                        && ((prev_is_alpha && para_style.auto_space_de)
                            || (prev_is_digit && para_style.auto_space_dn))
                    {
                        // S546: gap = fs/4 true-space (old per-fontSize table = paint artifact).
                        let extra = s546_autospace_extra(font_size);
                        if let Some(last) = current_line.fragments.last_mut() {
                            last.width += extra;
                            last.natural_width += extra;
                        }
                        current_width += extra;
                        current_width_tw += pt_to_tw(extra);
                        current_capw_tw += pt_to_tw(extra); // S475: autoSpace, no punct capacity
                    }

                    let s475_capinc = if s475_break {
                        pt_to_tw(pre_yakumono_width
                            - kinsoku::s475_max_compress(ch, chars_vec.get(char_index + 1).copied(),
                                s475_pair, s475_solo, font_size))
                    } else { 0 };
                    let overflow_tw = if s475_break {
                        current_capw_tw + s475_capinc - available_tw
                    } else {
                        current_width_tw + pt_to_tw(char_width) - available_tw
                    };
                    // 82de3fa REVERTED 2026-05-03 (independently confirmed by
                    // Session 52 + Session 51 oxi-3 branch). The trailing-U+3000
                    // immune-from-wrap rule (originally added for ed025 p.1 +0.042)
                    // caused d77a p.10 -0.054, p.9 -0.037, p.8 -0.008 (net -0.099)
                    // because d77a has a paragraph with 142×U+3000 decorative run
                    // (verified by OOXML walk) that got marked immune. Mid-text
                    // U+3000s (used for indentation) were also affected — both
                    // per-fragment and per-paragraph trailing-run-length
                    // threshold gates failed to discriminate d77a's mid-text
                    // U+3000s from ed025's true trailing run, because the
                    // immune flag propagates to ALL U+3000s in a paragraph,
                    // not just the trailing run. Word's actual gate is likely
                    // line-fill aware (only elide when the line is already
                    // near-full).
                    // Net trade: +0.099 d77a recovery / -0.042 ed025 loss
                    // = +0.057 net on bottom-bucket. d77a min p.7=0.6268
                    // unchanged so bottom-5 floor is preserved (3.2377 →
                    // 3.2646, +0.0269 Path A strict positive). ed025 stays
                    // rank 18 (out of bottom-5).
                    // Future: re-attempt with line-fill-aware gate (e.g. only
                    // immune when the line is already at >95% of available_tw).
                    //
                    // R7.62 (Day 36 part 9, 2026-05-14): re-attempt the trailing-U+3000
                    // immunity rule with TWO conditions tightened to avoid the d77a
                    // regression of the previous 82de3fa attempt:
                    //   1. ch == U+3000 AND
                    //   2. ALL remaining chars (char_index+1..end) are also U+3000
                    //      (true trailing-U+3000 run — distinguishes ed025c wi=10's
                    //      5 trailing U+3000 from d77a's 142-char mid-text U+3000
                    //      decorative run where non-U+3000 chars follow) AND
                    //   3. current line is already at ≥95% of available_tw (near-full
                    //      — Word collapses trailing U+3000 only when the line has
                    //      legitimately filled with content first; this excludes
                    //      degenerate "empty + trailing" lines).
                    // ed025c wi=10: "32×U+3000 + 法人番号： + 5×U+3000". Char 38 (1st
                    // trailing U+3000 that overflows) has remaining 4 chars all
                    // U+3000 + line at 99.1% full → immune. Resolves the +16pt
                    // drift jump that cascades 149 paras +1, 69 paras +2/+3.
                    let trailing_u3000 = ch == '\u{3000}'
                        && chars_vec.iter().skip(char_index + 1).all(|&c| c == '\u{3000}');
                    let line_near_full = available_tw > 0
                        && (current_width_tw * 100) >= (available_tw * 95);
                    let is_immune_space = trailing_u3000 && line_near_full;
                    let line_compress_count = current_line.fragments.iter()
                        .flat_map(|f| f.text.chars())
                        .filter(|&c| kinsoku::is_cjk_compressible(c))
                        .count();
                    // Phase 2 pair-yakumono compression for compressPunctuation docs.
                    // COM-refined 2026-04-17: Word only absorbs overflow when a
                    // PAIR of adjacent yakumono is present that can actually
                    // compress. Previous `count * font_size * 0.5` formula
                    // overestimated available savings, letting Oxi fit 1-2 extra
                    // chars/line on long paragraphs. Restrict absorption to small
                    // overflow (≤ 10tw = 0.5pt) and only when we have evidence of
                    // pair-compressible yakumono on the line.
                    let has_pair = current_line.fragments.iter()
                        .flat_map(|f| f.text.chars())
                        .collect::<Vec<_>>()
                        .windows(2)
                        .any(|w| kinsoku::is_yakumono_trigger(w[0]) && kinsoku::is_yakumono_trigger(w[1]));
                    // 2026-04-21: allow absorb when line starts with narrow yakumono
                    // (・/、/。/，/．). COM-verified on d77a pi=24-27 — single-yakumono
                    // line-start needs -2.5pt compression to fit +1 char/line. Without
                    // this extension the compression applied above still leaves ~0.1pt
                    // residual tw overflow that breaks the line 1 char early.
                    let has_linestart_narrow_yakumono = current_line.fragments.first()
                        .and_then(|f| f.text.chars().next())
                        .map_or(false, |c| matches!(c, '・' | '、' | '。' | '，' | '．'));
                    // Threshold raised 10→50tw (2026-04-18) per
                    // project_wrap_overflow_analyzer_e3c545.md analysis:
                    // e3c545 idx=29 at +18tw triggers 20.5pt cascade; 4 of its
                    // 18 over-wraps cluster in 10-50tw and are gated by
                    // has_pair so d77a over-wraps (has_pair=false) remain
                    // unaffected.
                    // S472 demand-driven absorb: a line carrying standalone 、，(left
                    // at NATURAL width upstream when S472 is on) can absorb overflow up
                    // to (count of such 、)×(fontSize/3 ≈4pt each at 12pt) by compressing
                    // them — Word's per-line justify-demand compression. On absorb the
                    // 、 fragments already on the line are retroactively shrunk by the
                    // absorbed overflow so the line fits exactly (break count AND render
                    // both match Word). This replaces the over-eager flat pre-compress.
                    // S543 (2026-06-11): demand oikomi for the NON-justified natural
                    // path. Word compresses each compressible yakumono on the line by
                    // a LIGHT -0.75pt (at fs=10.5; scaled fs*0.75/10.5) to fit one
                    // more char when the natural overflow is within that budget;
                    // otherwise oidashi with zero compression (S492's zero-compression
                    // observation was the oidashi branch only). Repro-confirmed
                    // (tools/metrics/repro_s542_width.py, verbatim 7f272a ３． para:
                    // ．（、 all 9.75 mid-line, 45-char L1; short lines stay 10.50 =
                    // demand-gated). Opening brackets DO compress in this tier
                    // S545 (2026-06-11) THE GATE: the demand oikomi is a
                    // compatibilityMode ≤ 14 (Word 2010) layout behavior.
                    // Bidirectionally repro-confirmed: the isolated real ed025c
                    // para (compat 15, refuses) FIRES when flipped to 14; the
                    // synthetic (compat 14, fires) STOPS when flipped to 15.
                    // ABSENT compatSetting = legacy doc = Word lays out ≤14
                    // (d77a/34140b/04b88e/fded6 have none and oikomi in Word),
                    // but parse_compat_mode reports 15 for them → use the
                    // explicit flag. ed025c (explicit 15) is correctly excluded.
                    // Default ON (spec complete); opt-out OXI_S543_DISABLE.
                    let s543_oikomi = natural_break_jc
                        && std::env::var("OXI_S543_DISABLE").is_err()
                        && self.compress_punctuation
                        && (self.compat_mode <= 14 || !self.compat_mode_explicit);
                    // S556 scaffold (opt-IN OXI_S556_JUST15; default OFF =
                    // byte-identical): the c15 justified pack tier, re-applied
                    // for integration debugging. Slack-table rule from
                    // S551-S555 (T={1:6.5,2:4.15,3:3.3,4:2.65,5+:2.15}).
                    // OXI_S556_DEBUG=1 prints each candidate decision.
                    let s556_just15 = is_justified
                        && std::env::var("OXI_S556_JUST15").is_ok()
                        && self.compress_punctuation
                        && self.compat_mode >= 15 && self.compat_mode_explicit
                        && !lines_and_chars
                        // PLAIN pulls only: a line-start-prohibited overflow
                        // char is the KINSOKU path's case (S550 K matrix:
                        // c15 = oidashi, NO compression) — the pack matrices
                        // all pulled plain 国.
                        && !kinsoku::is_line_start_prohibited(ch)
                        && current_line.fragments.iter()
                            .all(|f| f.width >= f.natural_width - 0.001)
                        && {
                            let cs: Vec<char> = current_line.fragments.iter()
                                .flat_map(|f| f.text.chars()).collect();
                            !cs.windows(2).any(|w|
                                kinsoku::is_yakumono_trigger(w[0])
                                    && kinsoku::is_yakumono_trigger(w[1]))
                        };
                    let mut s472_absorb = false;
                    if (s472_demand || s543_oikomi || s556_just15) && overflow_tw > 0
                        && self.compress_punctuation
                        && (self.compat_mode >= 15 || s543_oikomi)
                    {
                        let nat = font_size;
                        // Cap-aware compressibles for the whole-line FIT budget:
                        // 、,，→fs/3 (8.0pt floor); 。．and CLOSING brackets→fs/2 (6.0pt
                        // floor; opening brackets never compress). Including closing
                        // brackets fixes bracket-heavy lines (d77a p1 「…規約」) that Word
                        // packs by compressing 」 but Oxi's 、-only budget under-packed.
                        let mut comps: Vec<(usize, f32)> = Vec::new(); // (fi, removable)
                        for (i, f) in current_line.fragments.iter().enumerate() {
                            if f.text.chars().count() != 1 { continue; }
                            let c = f.text.chars().next().unwrap_or(' ');
                            if s556_just15 && !s543_oikomi && !s472_demand {
                                // indices only; rem=0 keeps legacy budget at 0.
                                if !kinsoku::is_s473_compressible(c) { continue; }
                                comps.push((i, 0.0));
                            } else if s543_oikomi && !s472_demand {
                                // S543 light tier compressibles (、，。．+ opening AND
                                // closing brackets; ！？ never compress).
                                // S546 (2026-06-12): each punct can compress down to its
                                // HALVING floor (fs/2) — the margin-sweep fit boundaries
                                // (_s546e/_s546f: single 、 fits overflow 5.10 not 5.35;
                                // 3-punct line frees −1.5/−1.5/−2.25 at need 5.05)
                                // falsified the S543b flat 0.75 cap, which was a painted
                                // (1px) artifact of small demands. The LINE-TOTAL budget
                                // is fs/2 (see below). Pre-S546 (OXI_S546_DISABLE):
                                // flat 0.75/punct.
                                if !kinsoku::is_s473_compressible(c) { continue; }
                                let cap = if crate::font::s546_exact_halfwidth() {
                                    font_size / 2.0
                                } else { 0.75 };
                                let floor = font_size - cap;
                                let rem = (f.width - floor).max(0.0);
                                if rem > 0.001 { comps.push((i, rem)); }
                            } else if s473_locomp {
                                // S473: break budget = Σ cap over ALL compressibles
                                // (、。，．+ closing AND opening brackets — break-flip
                                // showed opening （ compresses at break too), cap =
                                // s473_cap (≈3.25pt = fs×0.27). NO 0.95 exclusion:
                                // removable = how much THIS fragment can still lose down
                                // to its floor (font_size − cap); = cap when at natural,
                                // tapering for already-compressed fragments. This is the
                                // remaining-capacity model that fixes d77a p1 under-pack
                                // (37→38) and b837 p5 (40→39 rows) without over-packing
                                // p9 (39 needs 3.6pt/、 > cap → still wraps).
                                if !kinsoku::is_s473_compressible(c) { continue; }
                                let cap_pt = if s473_asym {
                                    match c {
                                        '、' | '，' => s473_cc,
                                        '。' | '．' => s473_cp,
                                        _ if kinsoku::is_yakumono_opening(c) => s473_cop,
                                        _ => s473_ccl, // closing brackets
                                    }
                                } else { s473_cap };
                                let cap = cap_pt * (font_size / 12.0);
                                let floor = font_size - cap;
                                let rem = (f.width - floor).max(0.0);
                                if rem > 0.001 { comps.push((i, rem)); }
                            } else {
                                if f.width < nat * 0.95 { continue; }
                                let cap = match c {
                                    '、' | '，' => font_size / 3.0,
                                    '。' | '．' => font_size / 2.0,
                                    '」' | '』' | '】' | '〕' | '》' | '〉' | '｝' | '］' | '）' => font_size / 2.0,
                                    _ => continue,
                                };
                                comps.push((i, cap));
                            }
                        }
                        // S556: justified-c15 slack-table pack + quanta distribution.
                        if s556_just15 && !s543_oikomi && !s472_demand && !comps.is_empty() {
                            let n = comps.len();
                            let t_n = match n {
                                1 => 6.5f32,
                                2 => 4.15,
                                3 => 3.3,
                                4 => 2.65,
                                _ => 2.15,
                            };
                            let need = overflow_tw as f32 / 20.0;
                            let slack = char_width - need;
                            let fire = need > 0.0 && slack >= t_n;
                            if std::env::var("OXI_S556_DEBUG").is_ok() {
                                // S557 (2026-06-13): the width components expose that
                                // `overflow_tw` in the s475_break (justified c15)
                                // regime is a CAPACITY overflow (capw + capinc −
                                // avail), measured against s475's shallow 2.5pt/punct
                                // compression — NOT a natural overflow. wid−capw is
                                // the s475 compression already baked in. The d77a
                                // para9 "counterexample" is a CASCADE artifact (L3
                                // under-packs 39 vs Word 40 → all later windows shift
                                // +1; Word's に…あらか line is a different 38-char
                                // window than Oxi's drifted 拠…あら). NOT a pack-rule
                                // gap. See [[session557_*]].
                                let head: String = current_line.fragments.iter()
                                    .flat_map(|f| f.text.chars()).take(12).collect();
                                eprintln!("S556 cand ch={} need={:.2} slack={:.2} n={} t={:.2} fire={} | s475brk={} wid={:.2} capw={:.2} avail={:.2} capinc={:.2} cw={:.2} head={}",
                                    ch, need, slack, n, t_n, fire,
                                    s475_break, current_width_tw as f32/20.0, current_capw_tw as f32/20.0,
                                    available_tw as f32/20.0, s475_capinc as f32/20.0, char_width, head);
                            }
                            if fire {
                                let q = (need / 0.75).round() as i32;
                                if q >= 1 {
                                    let mut order: Vec<usize> =
                                        comps.iter().map(|(fi, _)| *fi).collect();
                                    order.sort_by_key(|fi| {
                                        let c = current_line.fragments[*fi].text.chars().next().unwrap_or(' ');
                                        let class = if matches!(c, '、' | '，' | '。' | '．') { 0usize } else { 1 };
                                        (class, *fi)
                                    });
                                    let floor = font_size / 2.0;
                                    let mut remaining = q;
                                    'rr: loop {
                                        let mut placed = false;
                                        for fi in order.iter() {
                                            if remaining == 0 { break 'rr; }
                                            if current_line.fragments[*fi].width - 0.75 >= floor - 0.001 {
                                                current_line.fragments[*fi].width -= 0.75;
                                                remaining -= 1;
                                                placed = true;
                                            }
                                        }
                                        if !placed { break; }
                                    }
                                    let saved = (q - remaining) as f32 * 0.75;
                                    current_width -= saved;
                                    current_width_tw = current_width_tw.saturating_sub(pt_to_tw(saved));
                                    current_capw_tw = current_capw_tw.saturating_sub(pt_to_tw(saved));
                                    s472_absorb = true;
                                }
                            }
                        }
                        // S546 (2026-06-12): for the S543 light tier the fit budget is
                        // LINE-TOTAL fs/2 — one halfwidth char worth — independent of
                        // punct count (margin-sweep boundaries: single 、 [5.10, 5.35],
                        // 4-punct line [5.10, 5.60], both bracketing 5.25 at fs=10.5;
                        // 0.75×count would cap the 4-punct line at 3.0). Capped by Σrem
                        // (puncts already compressed, e.g. S532 pairs, contribute less).
                        let budget_tw = if s543_oikomi && !s472_demand
                            && crate::font::s546_exact_halfwidth()
                        {
                            pt_to_tw((font_size / 2.0).min(comps.iter().map(|(_, c)| *c).sum()))
                        } else {
                            pt_to_tw(comps.iter().map(|(_, c)| *c).sum())
                        };
                        if !comps.is_empty() && overflow_tw <= budget_tw && s543_oikomi && !s472_demand {
                            // S545/S546 Word distribution rule: comma/period class
                            // (、，。．) before brackets, left-to-right within a class.
                            // S546 deep-demand refinement: the freed amount is assigned
                            // in 0.75pt QUANTA, round-robin across the ordered puncts,
                            // n_quanta = round(overflow/0.75) — reproduces BOTH the
                            // S545 min-count×0.75 observations (small demand: need 1.5
                            // /3 puncts → 2 quanta → 2 puncts −0.75, 1 natural) AND the
                            // deep-demand split (_s546d r=1358: need 5.05 → 7 quanta →
                            // 、−2.25 （−1.5 ）−1.5 = the COM-painted advances exactly).
                            // Pre-S546: full-rem greedy (cap 0.75 ⇒ identical behavior).
                            let mut order: Vec<(usize, f32)> = comps.clone();
                            order.sort_by_key(|(fi, _)| {
                                let c = current_line.fragments[*fi].text.chars().next().unwrap_or(' ');
                                let class = if matches!(c, '、' | '，' | '。' | '．') { 0usize } else { 1 };
                                (class, *fi)
                            });
                            let mut saved = 0.0f32;
                            if crate::font::s546_exact_halfwidth() {
                                let quantum = 0.75f32;
                                let mut n_quanta =
                                    ((overflow_tw as f32 / 20.0) / quantum).round() as i32;
                                let mut rem_cap: Vec<f32> =
                                    order.iter().map(|(_, r)| *r).collect();
                                'outer: loop {
                                    let mut assigned_any = false;
                                    for (oi, (fi, _)) in order.iter().enumerate() {
                                        if n_quanta <= 0 { break 'outer; }
                                        if rem_cap[oi] >= quantum - 0.001 {
                                            current_line.fragments[*fi].width -= quantum;
                                            rem_cap[oi] -= quantum;
                                            saved += quantum;
                                            n_quanta -= 1;
                                            assigned_any = true;
                                        }
                                    }
                                    if !assigned_any { break; }
                                }
                            } else {
                                let mut needed = (overflow_tw as f32) / 20.0;
                                for (fi, rem) in &order {
                                    if needed <= 0.001 { break; }
                                    current_line.fragments[*fi].width -= *rem;
                                    saved += *rem;
                                    needed -= *rem;
                                }
                            }
                            current_width -= saved;
                            current_width_tw = current_width_tw.saturating_sub(pt_to_tw(saved));
                            s472_absorb = true;
                        } else if !comps.is_empty() && overflow_tw <= budget_tw {
                            // water-fill the overflow across comps (cap-aware), reducing
                            // current_width so accumulation stays coherent.
                            let mut needed = (overflow_tw as f32) / 20.0;
                            let mut active = comps.clone();
                            let mut amt = vec![0.0f32; current_line.fragments.len()];
                            loop {
                                if active.is_empty() || needed <= 0.001 { break; }
                                let share = needed / active.len() as f32;
                                let capped: Vec<(usize, f32)> = active.iter().cloned()
                                    .filter(|(_, c)| *c <= share).collect();
                                if capped.is_empty() {
                                    for (fi, _) in &active { amt[*fi] = share; }
                                    break;
                                }
                                for (fi, c) in &capped { amt[*fi] = *c; needed -= c; }
                                active.retain(|(_, c)| *c > share);
                            }
                            let mut saved = 0.0f32;
                            for (i, _) in &comps {
                                current_line.fragments[*i].width -= amt[*i];
                                saved += amt[*i];
                            }
                            current_width -= saved;
                            current_width_tw = current_width_tw
                                .saturating_sub(pt_to_tw(saved));
                            s472_absorb = true;
                        }
                    }
                    let absorb = if s472_absorb { true }
                        else if !s474_natural && !s475_break && overflow_tw > 0 && overflow_tw <= 50
                        && self.compress_punctuation && self.compat_mode >= 15
                        && (has_pair || has_linestart_narrow_yakumono)
                    { true } else { false };
                    let _ = line_compress_count;
                    if absorb {
                        compress_used = true;
                    }
                    if overflow_tw > 0 && !absorb && !is_immune_space && !current_line.fragments.is_empty()
                        && !para_all_whitespace {
                        // Word CJK hybrid hang/oikomi rule — COM-confirmed 2026-04-08.
                        // See memory/hangable_oikomi_rule.md.
                        //
                        // Hang ch on current line (burasagari) only if:
                        //   1. ch is a hangable CJK punct (、。）」 etc.), AND
                        //   2. next char (if any) is NOT line-start-prohibited
                        //      (otherwise hanging would push a still-prohibited char to L2 head).
                        //
                        // S228 (2026-05-23) v2: block hang ONLY when the
                        // current line has already absorbed earlier overflow
                        // (compress_used=true) AND the hang would compound
                        // the cheat. c7b923 wi=43 line 3: char ん at +36tw
                        // overflow gets absorbed (has_pair from mid-line 。）),
                        // then 。 at +129tw hangs. Both cheats stack to
                        // produce 46 chars instead of Word's 33-34.
                        // Word likely refuses the second cheat: if a line
                        // already absorbed overflow, hanging a further
                        // sentence-terminator past the right margin is
                        // disallowed.
                        // Gate fires ONLY when:
                        //   - compress_used = true (line already cheated), AND
                        //   - ch is `。` or `．` (sentence terminator), AND
                        //   - last char of last fragment (= paragraph end)
                        // OXI_LEGACY_HANG_NO_S228_GATE=1 disables.
                        let next_ch = chars_vec.get(char_index + 1).copied();
                        let next_is_proh = next_ch.map_or(false, kinsoku::is_line_start_prohibited);
                        let legacy_s228 = std::env::var("OXI_LEGACY_HANG_NO_S228_GATE").is_ok();
                        let is_para_last_char = frag_outer_idx + 1 == n_fragments
                            && char_index + 1 == chars_vec.len();
                        let is_sentence_terminator = matches!(ch, '。' | '．');
                        // S472h: the S228 hang-block fires when a line already "cheated"
                        // (compress_used). Under the S472 demand model, a line legitimately
                        // compresses 、 by a small justify-demand amount (compress_used=true)
                        // yet should STILL let a trailing 。 hang (b837 para13 L4: る fits via
                        // 、-absorb, then 。 must hang to keep 38/line = Word). So exempt the
                        // S472 path from the S228 hang-block.
                        let s228_block_hang = !legacy_s228
                            && compress_used
                            && is_para_last_char
                            && is_sentence_terminator
                            && !s472_demand;
                        // S492: burasagari (ぶら下げ) is NOT cleanly justify-gated.
                        // The synthetic jc=left repro (国、×30 = 36, oidashi) does NOT
                        // hang, but e3c545 (doNotCompress, type=lines) DOES hang on
                        // its jc=left lines — disabling the hang there cascaded its
                        // pagination (0.9997->0.245, the sole S492 Phase-1 regression).
                        // Burasagari is a doc-level HangingPunctuation behaviour, not a
                        // justify effect; leave it ON for the non-justified path. (The
                        // synthetic comma jc=left then over-hangs +2 vs Word, an
                        // accepted edge case — real docs don't carry 50%-punct lines.
                        // Re-deriving the exact jc/HangingPunctuation gate is next-session
                        // work — see docs/spec/cjk_break_refactor_s492.md.)
                        // S506 (2026-06-08, opt-in OXI_S506_OIDASHI scaffold; default OFF =
                        // byte-identical) — the CORRECT gate STRUCTURE (compat≥15 oidashi tied
                        // to the F1 natural-break path), but still cascades pending the
                        // footnote-vs-body grid distinction. compat≥15 (Word 2013+) does
                        // OIDASHI not burasagari at line-end (S506 repro: compat 12/14 HANG,
                        // 15 OIDASHI; b837=15 / e3c545=14). TEST (OXI_S492_JCNATURAL +
                        // OXI_S506_OIDASHI): b837 STILL cascaded 7→8. ROOT: natural_break_jc
                        // fires for b837's BODY (fs12, linesAndChars) too — disabling its grid
                        // count — and scoping it out via OXI_S492_LINESONLY also kills it for
                        // the FOOTNOTE (fs11) which NEEDS oidashi. break_into_lines cannot tell
                        // the off-grid footnote (→ natural+oidashi) from the on-grid body (→
                        // grid count) within one linesAndChars doc. That distinction (does the
                        // para's font align to the docGrid char pitch?) is the IRREDUCIBLE core
                        // of the S492 Step-2 refactor — see docs/spec/cjk_break_refactor_s492.md
                        // §8 and session505_b837_kinsoku_oidashi. compat_mode>=15 correctly
                        // leaves compat-14 (e3c545) hanging.
                        // S506 (2026-06-08, opt-in OXI_S506_OIDASHI scaffold, default OFF =
                        // byte-identical) — compat≥15 (Word 2013+) does OIDASHI not burasagari
                        // at line-end (S506 repro: compat 12/14 HANG, 15 OIDASHI; b837=15 /
                        // e3c545=14). DEFINITIVE CONCLUSION (3 gate conditions tried —
                        // lines_and_chars / !s476_grid / !s476_body — ALL cascade b837 7→8):
                        // the cascade is NOT a gate-scope problem. The b837 footnote MUST grow
                        // 2→3 lines (oidashi, to match Word's char positions), but growing it
                        // overflows Oxi's layout where Word fits the SAME 3-line footnote in 7
                        // pages — b837 carries a SECOND, vertical, compensating error (Oxi ~1
                        // line of vertical space taller than Word; the wrong 2-line hang
                        // footnote was offsetting it). So the kinsoku oidashi and the vertical
                        // over-height MUST land TOGETHER — b837 is multiply-compensated, the
                        // S492 multi-session refactor. compat_mode>=15 correctly leaves
                        // compat-14 (e3c545) hanging. See docs/spec/cjk_break_refactor_s492.md
                        // §8 / session505_b837_kinsoku_oidashi.
                        // S507 (reverted): an oidashi gate for type=lines compat-15 docs
                        // (683f/0e7af/d77a) was a NO-OP — 0 glyphs changed on 683f. Their
                        // S492 §2 over-pack is the S475 CAPACITY break, not burasagari, so it
                        // is addressed by F1 (OXI_S492_JCNATURAL, disables S475) — a
                        // coverage-track fix that §6 found SSIM-swamped by their structural /
                        // weight-AA errors. The burasagari/oidashi gate (S506) is for the
                        // hang-overflow case (b837 footnote), which these docs don't hit.
                        let s506_oidashi = std::env::var("OXI_S506_OIDASHI").is_ok()
                            && !is_justified && self.compat_mode >= 15 && !s476_body;
                        // S548 (2026-06-12, default ON, opt-out OXI_S548_DISABLE):
                        // the S506-confirmed rule shipped with its correct scope.
                        // compat≥15 (Word 2013+) does OIDASHI, not burasagari, at
                        // line-end (S506 repro: compat 12/14 HANG, 15 OIDASHI).
                        // EXPLICIT compat only — absent compatSetting = legacy doc
                        // = Word 2010 layout = burasagari stays (d77a/34140b/
                        // 04b88e/fded6; same semantics as the S545 oikomi gate).
                        // BODY paragraphs included (S506's !s476_body excluded
                        // them = no-op for the 3a4f class: its kern=0 注釈 para
                        // 「…就業規則は、」hung the 、 at natural width 5.3pt PAST
                        // the right margin where Word compat-15 pushes は、 down
                        // → every 注釈 para one line short → the 5 delta=-1
                        // boundary paras = the Phase-1 sole FAIL).
                        // linesAndChars (b837) excluded: its footnote oidashi is
                        // coupled to a compensating vertical over-height (S506
                        // definitive conclusion) — needs the S492 §8 refactor.
                        // s476_body: MAIN BODY flow only — applying the oidashi to
                        // CELL/textbox paragraphs regressed ed025c p4 −0.0685 (its
                        // narrow fitText cells re-wrapped); Word's in-cell hang
                        // behavior is unmeasured — body is the COM-pinned scope.
                        let s548_oidashi = std::env::var("OXI_S548_DISABLE").is_err()
                            && !is_justified && s476_body
                            && self.compat_mode >= 15 && self.compat_mode_explicit
                            && !lines_and_chars;
                        let can_hang = kinsoku::is_hangable_punct(ch) && !next_is_proh
                            && !s228_block_hang && !s506_oidashi && !s548_oidashi;

                        if can_hang {
                            current_line.fragments.push(LineFragment {
                                text: char_to_string(ch),
                                width: char_width,
                                natural_width: char_width + yakumono_saved,
                                style: style.clone(),
                                tab_alignment: None,
                                tab_position: None,
                                field_type: frag_field_type,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            lines.push(std::mem::take(&mut current_line));
                            current_width = 0.0; current_width_tw = 0; current_capw_tw = 0; compress_used = false;
                            continue;
                        }

                        // Oikomi (押し下げ): pop fragments from end of current line until
                        // both conditions are satisfied:
                        //   - first char of next line is NOT line-start-prohibited
                        //   - last char of current line is NOT line-end-prohibited
                        let mut popped: Vec<LineFragment> = Vec::new();
                        loop {
                            let last_of_curr = current_line.fragments.last()
                                .and_then(|f| f.text.chars().last());
                            let next_first = if let Some(p) = popped.last() {
                                p.text.chars().next().unwrap_or(ch)
                            } else {
                                ch
                            };
                            let bad = kinsoku::is_line_start_prohibited(next_first)
                                || last_of_curr.map_or(false, kinsoku::is_line_end_prohibited);
                            if !bad || current_line.fragments.len() <= 1 { break; }
                            let f = current_line.fragments.pop().unwrap();
                            current_width -= f.width;
                            popped.push(f);
                        }
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0; current_width_tw = 0; current_capw_tw = 0; compress_used = false;
                        for f in popped.into_iter().rev() {
                            current_width += f.width;
                            current_width_tw += pt_to_tw(f.width);
                            // S475: re-added oikomi frag. Approximate its break capacity
                            // from its first char's natural − max_compress (edge path).
                            if s475_break {
                                let fc = f.text.chars().next().unwrap_or(' ');
                                let fnext = f.text.chars().nth(1);
                                current_capw_tw += pt_to_tw(f.natural_width
                                    - kinsoku::s475_max_compress(fc, fnext, s475_pair, s475_solo, font_size));
                            } else {
                                current_capw_tw += pt_to_tw(f.width);
                            }
                            current_line.fragments.push(f);
                        }
                    }

                    current_line.fragments.push(LineFragment {
                        text: char_to_string(ch),
                        width: char_width,
                        natural_width: char_width + yakumono_saved,
                        style: style.clone(),
                        tab_alignment: None,
                        tab_position: None,
                        field_type: frag_field_type,
                        run_index: frag_run_index,
                        char_offset: char_pos_in_run,
                    });
                    current_width += char_width;
                    current_width_tw += pt_to_tw(char_width);
                    current_capw_tw += if s475_break { s475_capinc } else { pt_to_tw(char_width) };
                } else {
                    // Regular word character — accumulate
                    // autoSpaceDE: add 2.5pt gap when transitioning from CJK ideograph/kana to Latin.
                    // COM-confirmed (2026-04-07): Word only adds auto-space between Latin and
                    // CJK ideographs/kana, NOT between Latin and CJK punctuation.
                    // Session 95 (2026-05-18) split DE (alpha) vs DN (digit).
                    let ch_is_alpha = ch.is_ascii_alphabetic();
                    let ch_is_digit = ch.is_ascii_digit();
                    let s95_de_fires = ch_is_alpha && para_style.auto_space_de;
                    let s95_dn_fires = ch_is_digit && para_style.auto_space_dn;
                    if word_style.is_none() && (s95_de_fires || s95_dn_fires) {
                        let prev_is_cjk_ideo = current_line.fragments.last().map_or(false, |f| {
                            f.text.chars().last().map_or(false, |c| kinsoku::is_cjk_ideograph_or_kana(c))
                        });
                        if prev_is_cjk_ideo {
                            // S546: gap = fs/4 true-space (old per-fontSize table = paint artifact).
                            let extra = s546_autospace_extra(font_size);
                            if let Some(last) = current_line.fragments.last_mut() {
                                last.width += extra;
                                last.natural_width += extra;
                            }
                            current_width += extra;
                            current_width_tw += pt_to_tw(extra);
                            current_capw_tw += pt_to_tw(extra); // S475: autoSpace, no punct capacity
                        }
                    }
                    if word_style.is_none() {
                        word_style = Some(style.clone());
                        word_field_type = frag_field_type;
                        word_run_index = frag_run_index;
                        word_char_offset = char_pos_in_run;
                    }
                    word.push(ch);
                    word_width += char_width;
                    word_natural_width += char_width + yakumono_saved;
                }
                char_pos_in_run += 1; // character index (not byte offset) for JS compatibility
            }
            // Do NOT flush word here — it may continue in the next fragment
        }

        // Flush any remaining word after all fragments
        if !word.is_empty() {
            let ws = word_style.take().unwrap_or_else(|| {
                fragments.last().map(|f| f.1.clone()).unwrap_or_default()
            });
            let wft = word_field_type.take();
            // Day 33 part 19: skip wrap break for all-whitespace paragraphs.
            if current_width_tw + pt_to_tw(word_width) > available_tw && !current_line.fragments.is_empty()
                && !para_all_whitespace {
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0; current_width_tw = 0; current_capw_tw = 0; compress_used = false;
            }
            current_line.fragments.push(LineFragment {
                text: word,
                width: word_width,
                natural_width: word_natural_width,
                style: ws,
                tab_alignment: None,
                tab_position: None,
                field_type: wft,
                run_index: word_run_index,
                char_offset: word_char_offset,
            });
            current_width += word_width;
        }

        // Flush last line
        if !current_line.fragments.is_empty() {
            lines.push(current_line);
        }

        // Ensure at least one empty line for empty paragraphs
        if lines.is_empty() {
            lines.push(Line { fragments: vec![], ..Default::default() });
        }

        // 2-pass wrap (Stage 1): compute per-line natural_total_width and
        // was_compressed flag. These are consumed by Stage 2+ for context-aware
        // yakumono handling (loose vs tight line).
        for line in &mut lines {
            let nat: f32 = line.fragments.iter().map(|f| f.natural_width).sum();
            let comp: f32 = line.fragments.iter().map(|f| f.width).sum();
            line.natural_total_width = nat;
            line.was_compressed = (nat - comp) > 0.5;
        }

        // 2-pass wrap (Stage 2): demand-scaled compression revert.
        //   - Full revert: natural fits within available → revert all compression
        //     (Word's loose-line rule: no compression when line has slack)
        //   - Partial revert: natural slightly exceeds available → keep just enough
        //     compression to make line fit, scale rest back toward natural
        //     (Word's demand-driven rule: compression amount matches actual overflow)
        //   - No revert: natural greatly exceeds available (demand ≥ total savings) →
        //     keep full compression
        // S532 (2026-06-10): PAIR-compressed yakumono (。」/）」 adjacency) is
        // EXCLUDED from the revert — Word compresses adjacent-pair punctuation
        // UNCONDITIONALLY (minimal repro _s532_pair_repro.py: 。 advance = 6.0pt
        // exactly, in centered, loose-justified AND wrapping-justified lines
        // alike; d77a title/body 。」 gate pixels agree). Only the demand-driven
        // compressions (standalone 、。 ×0.6667, line-start narrow yakumono)
        // revert on loose lines. The pair members are re-identified here by
        // mirroring the break-time pair rule over the line's char sequence
        // (a fragment is typically one CJK char). Fragments the break never
        // compressed have width==natural, so over-marking is a no-op.
        // opt-out OXI_S532_DISABLE.
        // S547: the pair revert-protection only applies where the pair rule
        // itself applies (w:kern docs). Without this, a kern-less doc whose
        // standalone 、 was pre-compressed by OTHER rules and happens to sit
        // before an opener would be wrongly protected from the loose-line
        // revert (Word keeps it natural — kern0 sweep had zero halved pairs).
        let s532_keep_pairs = std::env::var("OXI_S532_DISABLE").is_err()
            && (!s547_kern_gate || para_kern_on);
        for line in &mut lines {
            if !line.was_compressed { continue; }
            let pair_frag: Vec<bool> = if s532_keep_pairs {
                let line_chars: Vec<(usize, char)> = line.fragments.iter().enumerate()
                    .flat_map(|(i, f)| f.text.chars().map(move |c| (i, c)))
                    .collect();
                let n = line_chars.len();
                let mut v = vec![false; n];
                for k in 0..n {
                    let c = line_chars[k].1;
                    if kinsoku::is_yakumono_closing(c) {
                        if k + 1 < n && kinsoku::is_yakumono_trigger(line_chars[k + 1].1) {
                            v[k] = true;
                        }
                    } else if kinsoku::is_yakumono_opening(c) {
                        if k > 0 && kinsoku::is_yakumono_trigger(line_chars[k - 1].1) && !v[k - 1] {
                            v[k] = true;
                        }
                    }
                }
                let mut mask = vec![false; line.fragments.len()];
                for k in 0..n {
                    let (fi, c) = line_chars[k];
                    let is_opening = matches!(c,
                        '（' | '「' | '『' | '〔' | '【' | '《' | '〈' | '｛' | '［');
                    // Only the FIRST char of an adjacent pair compresses (S532
                    // measurement); the second keeps natural advance, so only
                    // v[k] members need revert protection.
                    if v[k] && !is_opening {
                        mask[fi] = true;
                    }
                }
                mask
            } else {
                vec![false; line.fragments.len()]
            };
            let savings: f32 = line.fragments.iter().enumerate()
                .filter(|(fi, _)| !pair_frag[*fi])
                .map(|(_, f)| (f.natural_width - f.width).max(0.0)).sum();
            if savings <= 0.5 { continue; }
            let demand = (line.natural_total_width - available_width).max(0.0);
            if demand <= 0.5 {
                // Full revert: loose line, no compression needed
                for (fi, f) in line.fragments.iter_mut().enumerate() {
                    if !pair_frag[fi] {
                        f.width = f.natural_width;
                    }
                }
                line.was_compressed = line.fragments.iter().enumerate()
                    .any(|(fi, f)| pair_frag[fi] && (f.natural_width - f.width) > 0.5);
            } else if demand < savings {
                // Partial revert: demand-scaled. Release (savings - demand) back to
                // compressed fragments proportionally, matching Word's per-line
                // demand-driven compression on line-start yakumono (d77a pi=24-27
                // COM: ・ compresses -0.5 to -2.5pt based on line overflow demand).
                let keep_ratio = demand / savings;
                for (fi, f) in line.fragments.iter_mut().enumerate() {
                    if pair_frag[fi] { continue; }
                    let f_saving = (f.natural_width - f.width).max(0.0);
                    if f_saving > 0.0 {
                        f.width = f.natural_width - f_saving * keep_ratio;
                    }
                }
            }
        }

        // Post-process: adjust tab fragment widths for Center/Right/Decimal alignment.
        // ECMA-376 §17.3.1.38: Center tabs center the following segment on the tab position,
        // Right tabs right-align, Decimal tabs align at the decimal point.
        for line in &mut lines {
            let frag_count = line.fragments.len();
            let mut i = 0;
            while i < frag_count {
                if let Some(align) = line.fragments[i].tab_alignment {
                    if align == TabStopAlignment::Left {
                        i += 1;
                        continue;
                    }
                    let _tab_pos = line.fragments[i].tab_position.unwrap_or(0.0);
                    // Measure the segment width after this tab until next tab or end of line
                    let mut segment_width: f32 = 0.0;
                    let mut decimal_offset: Option<f32> = None;
                    let mut j = i + 1;
                    while j < frag_count {
                        if line.fragments[j].tab_alignment.is_some() {
                            break;
                        }
                        if align == TabStopAlignment::Decimal && decimal_offset.is_none() {
                            // Find decimal point position within this fragment
                            let mut char_offset: f32 = 0.0;
                            let fs = line.fragments[j].style.font_size.unwrap_or(11.0);
                            let metrics = self.registry.default_metrics();
                            for ch in line.fragments[j].text.chars() {
                                if ch == '.' || ch == ',' {
                                    decimal_offset = Some(segment_width + char_offset);
                                    break;
                                }
                                char_offset += self.registry.char_width_pt_with_fallback(ch, fs, metrics);
                            }
                        }
                        segment_width += line.fragments[j].width;
                        j += 1;
                    }

                    // Calculate the desired tab width so the segment aligns correctly
                    // Current tab width advances cursor to tab_pos. We need to adjust it
                    // so the segment is positioned according to the alignment type.
                    let current_tab_width = line.fragments[i].width;
                    let adjustment = match align {
                        TabStopAlignment::Center => segment_width / 2.0,
                        TabStopAlignment::Right => segment_width,
                        TabStopAlignment::Decimal => decimal_offset.unwrap_or(segment_width),
                        TabStopAlignment::Left => 0.0,
                    };
                    // New tab width = original width - adjustment (shift left by adjustment)
                    let new_width = (current_tab_width - adjustment).max(0.0);
                    line.fragments[i].width = new_width;
                }
                i += 1;
            }
        }

        // S492 (2026-06-03) — paragraph-level DEMAND break optimizer (env OXI_S492_OPT,
        // default OFF = byte-identical). Replaces the char-greedy break for JUSTIFIED
        // linesAndChars paragraphs with a Knuth-Plass DP that minimizes per-line
        // underfull² with free residual compression (= fill each line to ~avail with
        // LIGHT compression, Word's demand behaviour). Validated 72% per-line match vs
        // Word on b837 (vs 58-62% for any per-line greedy; greedy+maxcomp over-packs at
        // 31%). Render unchanged (decides COUNTS only; render water-fill re-justifies).
        // Scope: justified + linesAndChars + all-Normal-break + no tab/field fragments;
        // re-derive scope before extending. See docs/spec/cjk_break_optimizer_design.md.
        let s492_opt = std::env::var("OXI_S492_OPT").is_ok();
        if s492_opt && is_justified && lines_and_chars && lines.len() > 1
            && lines.iter().all(|l| l.break_type == LineBreakType::Normal
                && l.fragments.iter().all(|f| f.tab_alignment.is_none() && f.field_type.is_none()))
        {
            let flat: Vec<LineFragment> =
                lines.iter().flat_map(|l| l.fragments.iter().cloned()).collect();
            let n = flat.len();
            if n > 1 {
                let avail_l0 = (available_width - first_line_indent).max(0.0);
                let avail_cont = available_width;
                let mut pn = vec![0.0f32; n + 1];
                let mut pc = vec![0.0f32; n + 1];
                for k in 0..n {
                    let mc = if flat[k].text.chars().count() == 1 {
                        let fs = self.resolve_font_size(&flat[k].style, para_style);
                        kinsoku::s492_max_compress(flat[k].text.chars().next().unwrap(), fs)
                    } else { 0.0 };
                    pn[k + 1] = pn[k] + flat[k].natural_width;
                    pc[k + 1] = pc[k] + mc;
                }
                // Cost weights (env-tunable during the canary). Defaults from the
                // Python fit; w_line breaks ties toward packing (the Python's implicit
                // tie order favoured packing; Rust needs it explicit, else lines
                // under-pack once slack hits 0).
                let w_slack: f32 = std::env::var("OXI_S492_WSLACK").ok()
                    .and_then(|v| v.parse().ok()).unwrap_or(1.0);
                let w_comp: f32 = std::env::var("OXI_S492_WCOMP").ok()
                    .and_then(|v| v.parse().ok()).unwrap_or(0.0);
                let w_line: f32 = std::env::var("OXI_S492_WLINE").ok()
                    .and_then(|v| v.parse().ok()).unwrap_or(0.0);
                let inf = f32::INFINITY;
                let mut best = vec![inf; n + 1];
                let mut prev = vec![0usize; n + 1];
                best[0] = 0.0;
                // Overflow tolerance (env-tunable): Word's linesAndChars grid fits a
                // partial trailing cell / hangs a punct, so a line may exceed avail by
                // up to ~half a cell even with little compression. Default 0.6; sweep.
                let tol: f32 = std::env::var("OXI_S492_TOL").ok()
                    .and_then(|v| v.parse().ok()).unwrap_or(0.6);
                for j in 1..=n {
                    // kinsoku: break after frag j-1 is invalid if the next frag (j) would
                    // start a line with a line-start-prohibited char, or frag j-1 ends
                    // with a line-end-prohibited char.
                    if j < n {
                        if let Some(c0) = flat[j].text.chars().next() {
                            if kinsoku::is_line_start_prohibited(c0) { continue; }
                        }
                    }
                    if let Some(cl) = flat[j - 1].text.chars().last() {
                        if kinsoku::is_line_end_prohibited(cl) { continue; }
                    }
                    for i in 0..j {
                        if !best[i].is_finite() { continue; }
                        let avail = if i == 0 { avail_l0 } else { avail_cont };
                        let natural = pn[j] - pn[i];
                        let comp = pc[j] - pc[i];
                        if natural - comp > avail + tol { continue; } // infeasible even compressed
                        let lc = if j == n {
                            0.0 // last line: ragged-right, free
                        } else {
                            let slack = (avail - natural).max(0.0);
                            let used = (natural - avail).max(0.0);
                            w_slack * slack * slack + w_comp * used * used
                        };
                        let t = best[i] + lc + w_line;
                        if t < best[j] { best[j] = t; prev[j] = i; }
                    }
                }
                if best[n].is_finite() {
                    let mut bounds = vec![n];
                    let mut j = n;
                    while j > 0 { j = prev[j]; bounds.push(j); }
                    bounds.reverse();
                    let mut new_lines: Vec<Line> = Vec::with_capacity(bounds.len());
                    let mut flat_iter = flat.into_iter();
                    let mut taken = 0usize;
                    for w in bounds.windows(2) {
                        let count = w[1] - w[0];
                        let mut frags: Vec<LineFragment> = Vec::with_capacity(count);
                        for _ in 0..count { frags.push(flat_iter.next().unwrap()); }
                        let _ = taken; taken += count;
                        let nat: f32 = frags.iter().map(|f| f.natural_width).sum();
                        let comp: f32 = frags.iter().map(|f| f.width).sum();
                        new_lines.push(Line {
                            fragments: frags,
                            break_type: LineBreakType::Normal,
                            natural_total_width: nat,
                            was_compressed: (nat - comp) > 0.5,
                        });
                    }
                    if !new_lines.is_empty() {
                        lines = new_lines;
                    }
                }
            }
        }

        lines
    }

    /// Calculate line height considering:
    /// 1. Font metrics (base single-line height)
    /// 2. Paragraph default font minimum (from style/docDefaults)
    /// 3. Line spacing multiplier (w:line/240, e.g. 1.15 for default)
    /// 4. Document grid snapping (linePitch)
    ///
    /// Word determines line height as the max of the run font's height
    /// and the paragraph's default font height (from the style/theme).
    /// Then applies the spacing multiplier and optionally snaps to grid.
    fn line_height(
        &self,
        font_size: f32,
        line_spacing: Option<f32>,
        line_spacing_rule: Option<&str>,
        metrics: &FontMetrics,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
    ) -> f32 {
        self.line_height_inner(font_size, line_spacing, line_spacing_rule, metrics, snap_to_grid, grid_pitch, false)
    }

    fn line_height_inner(
        &self,
        font_size: f32,
        line_spacing: Option<f32>,
        line_spacing_rule: Option<&str>,
        metrics: &FontMetrics,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
        in_table_cell: bool,
    ) -> f32 {
        // For Single/auto spacing (no explicit line_spacing or factor=1.0),
        // try COM-measured lookup table first (most accurate, includes GDI hinting)
        let is_single = match (line_spacing_rule, line_spacing) {
            (Some("exact"), _) | (Some("atLeast"), _) => false,
            (_, Some(f)) if (f - 1.0).abs() > 0.01 => false,
            _ => true,
        };

        if is_single {
            if in_table_cell {
                // Table cells: use GDI table if available, otherwise word_line_height_table_cell.
                // Grid snap is applied below via the normal path (compat_mode dependent).
                // Don't early-return here — fall through to GDI table + grid snap logic.
            } else {
                if let Some(lh) = self.registry.com_line_height(
                    &metrics.family, font_size,
                    if snap_to_grid { grid_pitch } else { None }
                ) {
                    return lh;
                }
            }
        }

        // Use GDI tmHeight table if available (most accurate).
        // Falls back to formula-based calculation.
        let ppem = (font_size * 96.0 / 72.0).round() as u32;
        let base = if let Some((h_px, _a_px, _d_px)) = self.registry.gdi_height(&metrics.family, ppem) {
            // GDI table stores tmHeight (MulDiv-based ascent + descent).
            // Body paragraphs with COM lookup use that (more accurate).
            // Table cells: tmHeight only (no tmExternalLeading) — COM-confirmed.
            let gdi_height_pt = h_px as f32 * 72.0 / 96.0;
            if metrics.is_cjk_83_64_font() {
                let raw = gdi_height_pt * 83.0 / 64.0;
                (raw * 8.0).floor() / 8.0
            } else {
                gdi_height_pt
            }
        } else if in_table_cell && self.adjust_line_height_in_table {
            metrics.word_line_height_standard(font_size)
        } else if in_table_cell {
            metrics.word_line_height_table_cell(font_size)
        } else {
            metrics.word_line_height(font_size, 96.0)
        };

        match (line_spacing_rule, line_spacing) {
            (Some("exact"), Some(val)) => val,
            (Some("atLeast"), Some(val)) => {
                // R55 (2026-05-17): `line=0 atLeast` uses NATURAL line height
                // without grid-snap. COM-confirmed via L1-L8 minimal repros:
                //   L1 (10.5pt line=0 atLeast): gap=13.5pt  (natural, no snap)
                //   L2 (14pt   line=0 atLeast): gap=18pt    (natural ≈18.75)
                //   L3 (10pt   line=0 atLeast): gap=12.75pt (natural)
                //   L4 (12pt   line=0 atLeast): gap=15.75pt (natural)
                //   L5 (10.5pt line=240 atLeast): gap=18pt  (snap to grid, then max)
                //   L8 (mixed  line=0 atLeast): each para uses its OWN natural
                // Only 2 baseline docs use line=0 atLeast: e201 + d1e8 (both
                // Phase 2 bottom-band with accumulating Y drift, Phase 1 PASS).
                if val == 0.0 {
                    return base;
                }
                // Non-zero val: original behavior (snap natural, then max with val)
                let snapped = if snap_to_grid {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            (((base + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch
                        } else { base }
                    } else { base }
                } else { base };
                snapped.max(val)
            }
            _ => {
                let spaced = match line_spacing {
                    Some(factor) => base * factor,
                    None => base,
                };
                // COM-confirmed (2026-04-03, gen2_023): grid snap is only applied when
                // lineSpacing is Single (factor=1.0) or unset. Multiple spacing (factor≠1.0)
                // does NOT get grid-snapped. MS Mincho 11pt line=276: gap=16.5pt (no snap),
                // NOT 18pt (with snap).
                let is_single = match line_spacing {
                    Some(f) => (f - 1.0).abs() < 0.001,
                    None => true,
                };
                // Bug B Day 28 (Phase β step 3): cell line snap is gated by the
                // settings.xml `<w:adjustLineHeightInTable/>` flag. COM-confirmed
                // via V70 minimal repro: cell paragraph dy = 18pt (snap) only when
                // the flag is present; absent = 13pt natural (= L7/L8 spec).
                // This unifies V1-V6 (no flag, no snap) and b5f706 real-doc
                // (with flag, snap to pitch) under one rule.
                let cell_snap_allowed = !in_table_cell || self.adjust_line_height_in_table;
                // S521 (2026-06-09) FALSIFIED + reverted: a controlled SINGLE-line AUTO-height
                // cell repro showed Word using natural cell line height (not the grid pitch), but
                // the cell-natural change regressed the corpus massively (33-ALH-doc gate net
                // −5.44, d77a pagination broke, b35 −0.12). The grid-snap is CORRECT for real
                // MULTI-line cells; the single-line repro does not generalize (Nth confirmation of
                // the cell row-height tombstone, cf S499). Do NOT re-attempt cell-natural-height.
                if snap_to_grid && is_single && cell_snap_allowed {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            let snapped = (((spaced + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch;
                            // S492y SHIP (2026-06-03, default ON, opt-out OXI_S492Y_DISABLE):
                            // snap the CELL line pitch to 96dpi device pixels (0.75pt). Word
                            // renders the cell line pitch at 96dpi px (17.5pt -> 23px -> 17.25pt;
                            // COM cell pitch measured 17.2-17.3, NOT the full grid 17.5); Oxi's
                            // 17.5 over-allocated ~0.25pt/line -> accumulated ~10px too LOW by
                            // page bottom (b35 p1 per-line vertical-align ceiling +0.157). GATE:
                            // Phase-1 54/55 preserved (0 PASS->FAIL); 40-affected-doc canary net
                            // +0.0456, bottom-3 sum +0.0361 (strictly up), b35 p1 +0.0357 /
                            // 29dc6e p3 +0.0089; 4 tiny regress all <0.001 (de6e32/15076df/
                            // 1636d28/a1d6e4 < the 0.005 threshold). Cell-only (non-table docs
                            // byte-identical). Screenshot ground-truth (not COM-logical) drove it.
                            if in_table_cell && std::env::var("OXI_S492Y_DISABLE").is_err() {
                                return ((snapped / 0.75).round() * 0.75).max(1.0);
                            }
                            return snapped;
                        }
                    }
                }
                // No grid snap: ceil to 10 twips (0.5pt) — Word internal line height.
                // COM-confirmed: 80/80 tests (5 fonts x 4 sizes x 4 spacings) all match.
                // Table cells use raw value (table row height has separate calculation).
                if !in_table_cell {
                    let tw = spaced * 20.0;
                    (tw / 10.0).ceil() * 10.0 / 20.0
                } else {
                    // S584 (2026-06-16): a TYPED docGrid CELL line is never shorter
                    // than 1 grid cell, even with a COMPRESSING auto multiplier
                    // (line<240). Word grid-snaps cell lines when snap_to_grid is
                    // set AND adjustLineHeightInTable is present; the single-spacing
                    // cell snap above (the `is_single` block) already does this, but
                    // the multiple-spacing fallthrough returned raw 0.85*natural. COM
                    // (tokyoshugyo パワハラ cell, line=204 0.85x, MS Mincho 10.5pt
                    // linePitch=360): Word renders 18.0pt (=1 cell, i=363-368 at 18pt
                    // advance), Oxi gave 11.5. Gated on `snap_to_grid` (NOT just
                    // cell_snap_allowed): a snap_to_grid=false cell (b35123
                    // linesAndChars + charSpace, fs=9 cells) uses its NATURAL height
                    // and must NOT be clamped — without this gate it regressed
                    // b35123/bd90b00 PASS→FAIL (the cell row-height tombstone). The
                    // is_single block uses the same `snap_to_grid && cell_snap_allowed`
                    // gate, so single+snap cells already snap there and never reach
                    // here; only multiple-spacing+snap cells do. mult>=1.25
                    // (spaced>pitch) is a no-op. Opt-out OXI_S584_DISABLE.
                    if snap_to_grid && cell_snap_allowed
                        && std::env::var("OXI_S584_DISABLE").is_err()
                    {
                        match grid_pitch {
                            Some(pitch) if pitch > 0.0 => spaced.max(pitch),
                            _ => spaced,
                        }
                    } else {
                        spaced
                    }
                }
            }
        }
    }

    /// Compute line height for a line with multiple runs using Word's algorithm:
    /// max(ascent across all runs) + max(descent across all runs).
    /// Uses EastAsia font metrics for CJK text (#2).
    fn line_height_for_line(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
        grid_no_type: bool,
    ) -> f32 {
        self.line_height_for_line_inner(line, para_style, para_font_size, snap_to_grid, grid_pitch, false, grid_no_type)
    }

    /// Returns ascent+descent only (no grid snap, no leading) for a line.
    /// Used for page-break threshold check (Day 33 part 65, 2026-05-12):
    /// COM-confirmed via db9ca18 i=37 — Word allows the LEADING portion of
    /// a grid-snapped line to extend into the bottom margin. Only the
    /// ascent+descent zone must fit within content area. Without this rule,
    /// Oxi rejects lines whose grid-pitch bottom exceeds pgBot by even a
    /// few points (db9ca18 +2pt), while Word fits them (db9ca18 i=37
    /// extends 5.25pt past pgBot in Word).
    fn natural_line_height_for_line(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
    ) -> f32 {
        let mut max_ascent: f32 = 0.0;
        let mut max_descent: f32 = 0.0;
        if line.fragments.is_empty() {
            let font_size = para_style.ppr_rpr.as_ref()
                .and_then(|r| r.font_size)
                .unwrap_or(para_font_size);
            let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
            let metrics = self.metrics_for_para_mark(&rpr_ref, para_style);
            max_ascent = metrics.word_ascent_pt(font_size);
            max_descent = metrics.word_descent_pt(font_size);
        } else {
            for frag in &line.fragments {
                let font_size = frag.style.font_size.unwrap_or(para_font_size);
                let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                let asc = metrics.word_ascent_pt(font_size);
                let des = metrics.word_descent_pt(font_size);
                if asc > max_ascent { max_ascent = asc; }
                if des > max_descent { max_descent = des; }
            }
        }
        max_ascent + max_descent
    }

    /// S576 (2026-06-15): glyph-INK height of a line (≈ em), for the
    /// page-bottom break-fit check. Word lets the line-spacing leading hang
    /// into the bottom margin and only requires the glyph ink to fit the
    /// content area. natural_line_height_for_line returns the SPACING box
    /// (win_sum*83/64 = 1.297*em for CJK); this returns typo_sum*fs (= em).
    /// See FontMetrics::glyph_ink_height_pt.
    fn ink_line_height_for_line(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
    ) -> f32 {
        let mut max_ink: f32 = 0.0;
        if line.fragments.is_empty() {
            let font_size = para_style.ppr_rpr.as_ref()
                .and_then(|r| r.font_size)
                .unwrap_or(para_font_size);
            let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
            let metrics = self.metrics_for_para_mark(&rpr_ref, para_style);
            max_ink = metrics.glyph_ink_height_pt(font_size);
        } else {
            for frag in &line.fragments {
                let font_size = frag.style.font_size.unwrap_or(para_font_size);
                let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                let ink = metrics.glyph_ink_height_pt(font_size);
                if ink > max_ink { max_ink = ink; }
            }
        }
        max_ink
    }

    fn line_height_for_line_inner(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
        in_table_cell: bool,
        grid_no_type: bool,
    ) -> f32 {
        let _default_style = RunStyle::default();

        let mut max_ascent: f32 = 0.0;
        let mut max_descent: f32 = 0.0;

        // adjustLineHeightInTable=true: use standard height without CJK 83/64
        let use_standard = in_table_cell && self.adjust_line_height_in_table;

        if line.fragments.is_empty() {
            // Empty paragraph: use pPr/rPr font size if available (direct paragraph property),
            // otherwise fall back to paragraph style's default run style.
            // COM-confirmed: 3a4f P1 empty, pPr/rPr/sz=48 (24pt) → uses Century 24pt height.
            let font_size = para_style.ppr_rpr.as_ref()
                .and_then(|r| r.font_size)
                .unwrap_or(para_font_size);
            let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
            let metrics = self.metrics_for_para_mark(&rpr_ref, para_style);
            if use_standard {
                let h = metrics.word_line_height_standard(font_size);
                max_ascent = h * metrics.win_ascent / (metrics.win_ascent + metrics.win_descent);
                max_descent = h - max_ascent;
            } else {
                max_ascent = metrics.word_ascent_pt(font_size);
                max_descent = metrics.word_descent_pt(font_size);
            }
        } else {
            for frag in &line.fragments {
                let font_size = frag.style.font_size.unwrap_or(para_font_size);
                let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                let (asc, des) = if use_standard {
                    let h = metrics.word_line_height_standard(font_size);
                    (h * metrics.win_ascent / (metrics.win_ascent + metrics.win_descent), h * metrics.win_descent / (metrics.win_ascent + metrics.win_descent))
                } else {
                    (metrics.word_ascent_pt(font_size), metrics.word_descent_pt(font_size))
                };
                if asc > max_ascent { max_ascent = asc; }
                if des > max_descent { max_descent = des; }
            }
        }

        let run_base = max_ascent + max_descent;

        // For LayoutMode=0 (no grid, grid_pitch=None), use direct font metrics formula.
        // COM-confirmed (2026-04-06): LayoutMode=0 uses floor(win_sum*fontSize*20/10)*10/20
        // without GDI pixel rounding. The ascent+descent formula uses pixel_round which
        // overshoots by 0.5pt (e.g. Calibri 11pt: 13.5 vs actual 13.0).
        let base = if grid_pitch.is_none() && !in_table_cell {
            // LayoutMode=0: use no-grid formula for each fragment.
            // Round 9 (2026-04-08): per-(font,size) lookup table from
            // `lm0_lineauto.json` overrides the formula when present
            // (COM-confirmed sweep across MS Mincho/Gothic, Yu Mincho/Gothic,
            // Calibri, TNR, Meiryo at sizes 7-25pt).
            // Round 18 (2026-04-28) — REVERT attempt at lookup removal:
            // Round 16 V18/V19 measurement showed LM0 values diverge from
            // formula by 1.5-9pt for the no-docGrid scenario. Removing the
            // lookup at this site (replacing with `fallback`) caused
            // pipeline.verify NG: test_line_heights p.4 (−0.0027) +
            // p.5 (−0.0012) regressed against +0.0045 gain at p.2. Net
            // ≈0, bottom-5 floor unchanged. LM0 lookup is load-bearing
            // for at least some Word style + lineSpacing combinations
            // (likely compensating for other Oxi quirks). Lookup retained
            // until the underlying compensation is identified.
            // Round 20 (2026-04-28) — TARGETED CJK-only skip attempt:
            // Made closure conditional `if metrics.is_cjk_83_64_font()`.
            // Verify result: 0 improved / 352 unchanged / 0 regressed
            // (all sub-threshold movements; 130 pages with diff < 0.001
            // each, total net +0.0039, bottom-5 floor +0.000018 = noise).
            // Path A gate technically passes (strict increase) but SSIM
            // value essentially zero. Reverted to keep code simple; LM0
            // lookup question requires per-entry verification (240 cells)
            // before further fix attempts.
            let lookup_no_grid = |family: &str, font_size: f32, fallback: f32| -> f32 {
                self.registry.lm0_lineauto_base(family, font_size).unwrap_or(fallback)
            };
            let mut no_grid_max: f32 = 0.0;
            if line.fragments.is_empty() {
                let font_size = para_style.ppr_rpr.as_ref()
                    .and_then(|r| r.font_size)
                    .unwrap_or(para_font_size);
                let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                let metrics = self.metrics_for_para_mark(&rpr_ref, para_style);
                let formula = metrics.word_line_height_no_grid(font_size);
                no_grid_max = lookup_no_grid(&metrics.family, font_size, formula);
            } else {
                let mut has_latin = false;
                for frag in &line.fragments {
                    let font_size = frag.style.font_size.unwrap_or(para_font_size);
                    let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                    let formula = metrics.word_line_height_no_grid(font_size);
                    let h = lookup_no_grid(&metrics.family, font_size, formula);
                    if h > no_grid_max { no_grid_max = h; }
                    // Track if this line has Latin text (non-CJK fragments)
                    if !frag.text.chars().all(|c| kinsoku::is_cjk(c)) {
                        has_latin = true;
                    }
                }
                // COM-confirmed (2026-04-07): when a line contains Latin text, Word also
                // considers the ASCII font's CJK 83/64 height for the line height basis.
                // This accounts for Japanese proportional fonts (游ゴシック, 游明朝, MS明朝)
                // having larger CJK 83/64 heights than the eastAsia font (ＭＳ 明朝).
                // Only applies when the ASCII font is a CJK 83/64 font.
                if has_latin {
                    // Use the first fragment's ASCII font (resolve via latin path)
                    if let Some(frag) = line.fragments.first() {
                        let font_size = frag.style.font_size.unwrap_or(para_font_size);
                        let latin_metrics = self.metrics_for(&frag.style, para_style);
                        if latin_metrics.is_cjk_83_64_font() {
                            let formula = latin_metrics.word_line_height_no_grid(font_size);
                            let h = lookup_no_grid(&latin_metrics.family, font_size, formula);
                            if h > no_grid_max { no_grid_max = h; }
                        }
                    }
                }
            }
            // Use max of no-grid height and run_base (for mixed-font baseline coverage)
            run_base.max(no_grid_max)
        } else {
            run_base
        };

        // Apply line spacing rule
        let line_spacing = para_style.line_spacing;
        let line_spacing_rule = para_style.line_spacing_rule.as_deref();
        match (line_spacing_rule, line_spacing) {
            (Some("exact"), Some(val)) => val,
            (Some("atLeast"), Some(val)) => {
                // R55 (2026-05-17): `line=0 atLeast` uses NATURAL line height
                // without grid-snap. COM-confirmed via L1-L8 minimal repros.
                // See line_height_inner counterpart for full reasoning.
                if val == 0.0 {
                    return base;
                }
                // Non-zero val: original behavior (snap natural, then max with val)
                let snapped = if para_style.snap_to_grid {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            if grid_no_type {
                                // S571: no-type docGrid uses device-snapped natural,
                                // not whole-cell ceil (see `_` arm below).
                                let dev = (base / 0.75).floor() * 0.75;
                                pitch.max(dev)
                            } else {
                            // S195: narrower grid-snap tolerance (see the `_` arm
                            // below for the full S195/S580b rationale — kept until the
                            // run-vs-TOC empty-para spec is re-derived).
                            let is_empty = line.fragments.iter()
                                .all(|f| f.text.is_empty());
                            let just_over_pitch = base > pitch && base <= pitch + 0.5;
                            let apply_tol = is_empty && just_over_pitch;
                            let tol = if apply_tol { 0.5 } else { 0.0 };
                            (((base - tol + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch
                            }
                        } else { base }
                    } else { base }
                } else { base };
                snapped.max(val)
            }
            _ => {
                let spaced = match line_spacing {
                    Some(factor) => base * factor,
                    None => base,
                };
                // COM-confirmed (2026-04-03): grid snap only for Single (factor=1.0) or unset.
                // Multiple spacing (factor≠1.0) does NOT get grid-snapped.
                let is_single_line = match line_spacing {
                    Some(f) => (f - 1.0).abs() < 0.001,
                    None => true,
                };
                if snap_to_grid && is_single_line {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            // S571 (2026-06-14): a NO-TYPE docGrid does NOT snap each
                            // line to whole grid cells. The linePitch is the DEFAULT
                            // line advance, but a line taller than the pitch (e.g. a
                            // 14pt heading in a 14.3pt grid) uses its NATURAL height
                            // device-snapped to 96dpi px, NOT ceil-to-2-cells.
                            // GOLD-STANDARD (ikujidetail PDF render-truth): 11pt body
                            // -> 14.28 (=pitch), 14pt heading -> 18.0
                            // (=floor(18.5/0.75)*0.75), NOT 28.6 (the ceil result).
                            // Rule = max(linePitch, floor(natural/0.75)*0.75). This
                            // realizes the long-dead doc_grid_no_type design intent
                            // (skip the CJK 83/64 whole-cell inflation). The COM
                            // Single-height table (grid288=16.5) was measured in a
                            // TYPED-grid context and does NOT match the no-type render.
                            if grid_no_type {
                                let dev = (spaced / 0.75).floor() * 0.75;
                                return pitch.max(dev);
                            }
                            // S195: narrower grid-snap tolerance — empty paragraph
                            // whose natural lh slightly exceeds pitch (e.g. 14pt MS
                            // Mincho 18.125pt at 18pt pitch) snaps to 1 cell (18pt),
                            // not 2 cells (36pt). This is the WESTERN-ascii empty-para
                            // case (model wi=173: ascii=Century, eastAsia=MS Mincho —
                            // Oxi's base comes from the eastAsia 83/64 font = 18.125,
                            // but Word measures the ¶ with the ASCII font = Century,
                            // which fits 1 cell). S583 (2026-06-16): the discriminator
                            // is the ASCII font's CJK-ness — a CJK-ascii empty (kojin's
                            // 様式 spacer: ascii=HGPｺﾞｼｯｸM) must NOT be snapped down; Word
                            // renders it at 2 cells (COM-confirmed gap=36.00). So the
                            // tolerance applies only when the ASCII font is Western.
                            // (S580b removed S195 entirely → regressed model because it
                            // lacked this ascii discriminator.) The tolerance window
                            // (base ≤ pitch+0.5) only ever fires at grid360 for 14pt, so
                            // non-360 grids (1ec1 g357, 2ea81a g323) are unaffected.
                            // Opt-out OXI_S583_DISABLE restores the old (model-only) gate.
                            // See [[empty_para_ascii_font]].
                            let is_empty = line.fragments.iter()
                                .all(|f| f.text.is_empty());
                            let just_over_pitch = spaced > pitch && spaced <= pitch + 0.5;
                            let ascii_is_cjk = is_empty
                                && std::env::var("OXI_S583_DISABLE").is_err()
                                && {
                                    let rpr_ref = para_style.ppr_rpr.as_ref()
                                        .cloned().unwrap_or_default();
                                    self.metrics_for(&rpr_ref, para_style).is_cjk_83_64_font()
                                };
                            let apply_tol = is_empty && just_over_pitch && !ascii_is_cjk;
                            let tol = if apply_tol { 0.5 } else { 0.0 };
                            return (((spaced - tol + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch;
                        }
                    }
                }
                // S584 (2026-06-16): a TYPED "lines" docGrid line is never
                // shorter than 1 grid cell, even with a COMPRESSING auto
                // multiplier (line<240). Word clamps the multiplied height UP
                // to the grid pitch. COM-confirmed (mult_grid repro, MS Mincho
                // 10.5pt linePitch=360): line=204 (0.85x) AND line=240 (1.0x)
                // both render 18.0pt (=1 cell); Oxi's un-snapped multiple path
                // gave 0.85*natural = 11.5pt. Scope: typed grid only
                // (!grid_no_type — a no-type grid uses device-snapped natural,
                // handled in the is_single branch above), BODY only
                // (!in_table_cell — cells have their own row-height machinery).
                // mult>=1.25 (spaced > pitch) is a no-op (the only corpus body
                // auto-mult cases are tokyoshugyo line=204 and the line=360
                // empty in 3a4f/model which already exceeds pitch). The exact
                // mult>=1.25 fractional-cell formula (9*mult+9 for 10.5pt) does
                // NOT generalize across font sizes (14pt is flat 29.25), so it
                // is intentionally NOT implemented — only the universal
                // "line >= 1 grid cell" floor. Opt-out OXI_S584_DISABLE.
                let spaced = if snap_to_grid && !grid_no_type && !in_table_cell
                    && std::env::var("OXI_S584_DISABLE").is_err()
                {
                    match grid_pitch {
                        Some(pitch) if pitch > 0.0 => spaced.max(pitch),
                        _ => spaced,
                    }
                } else {
                    spaced
                };
                // Round to 10 twips (0.5pt) — Word internal line height precision.
                // COM-confirmed (2026-04-07): LayoutMode=0 uses ROUND to 0.5pt.
                //   ＭＳ 明朝 10.5 LM=0: 13.625*1.15=15.66 → round 15.5pt (Word: 15.5pt)
                //   游明朝 10.5 LM=0:   17.5*1.15=20.125 → round 20.0pt (Word: 20.0pt)
                //   ＭＳ 明朝 14 LM=0:  18.125*1.15=20.84 → round 21.0pt (Word: 21.0pt)
                // LayoutMode≥1 uses CEIL.
                //   Meiryo 10.5: CJK 83/64=20.375pt → ceil 20.5pt for both.
                let tw = spaced * 20.0;
                if grid_pitch.is_none() {
                    (tw / 10.0).round() * 10.0 / 20.0
                } else {
                    (tw / 10.0).ceil() * 10.0 / 20.0
                }
            }
        }
    }

    /// Compute the vertical offset to apply to text within a line for exact/atLeast spacing.
    ///
    /// Word behavior depends on paragraph context (Session 76 Mech A fix, 2026-05-17):
    /// - **body / cell paragraphs**: top-align — text glyph at LINE BOX TOP, return 0.
    ///   COM-confirmed via 7 minimal repros in Session 70 (A1-A3, A7, B5-B6):
    ///   Word's exact rule on body paragraphs places glyph at top_margin = LINE BOX TOP.
    /// - **shape / textbox paragraphs**: bottom-align — text glyph at line box bottom,
    ///   return `(line_height - max_font_size).max(0.0)`. COM-confirmed on 1ec1 p1
    ///   Shape 4 (exact=22pt fontSize=14pt → 8pt offset).
    ///
    /// Returns the offset from line-box top to where text should start.
    #[allow(unused_assignments)]
    fn text_y_offset_for_line(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        line_height: f32,
        grid_pitch: Option<f32>,
        in_shape_context: bool,
    ) -> f32 {
        match (para_style.line_spacing_rule.as_deref(), para_style.line_spacing) {
            (Some("exact"), Some(_)) | (Some("atLeast"), Some(_)) => {
                // Session 76 Mech A fix: body/cell top-align, shape bottom-align.
                // Session 78 Mech A v2 refinement (2026-05-17): Word's actual
                // glyph offset for body exact = 0.5pt (NOT 0.0pt), per Session 70
                // COM measurements A1/A2/A3/A7 all showing Word offset = 0.50pt
                // = 10 twips. Returning 0.0 caused fded6 p.1 -0.10, 7f272a p.1
                // -0.06, 04b88e p.1 -0.06 SSIM regressions. See
                // memory/session78_mech_a_v2_05pt_offset.md.
                // Per spec §13.4 note: "GDI TextOutW character cell = fontSize".
                // offset = line_height - max_font_size = the empty space the exact box
                // leaves; Word places the text at the BOTTOM of the box (extra space above).
                // COM-confirmed on 1ec1 p1 Shape 4 (exact=22pt fontSize=14pt → 8pt offset).
                let mut max_font_size: f32 = 0.0;
                if line.fragments.is_empty() {
                    max_font_size = para_style.ppr_rpr.as_ref()
                        .and_then(|r| r.font_size)
                        .unwrap_or(para_font_size);
                } else {
                    for frag in &line.fragments {
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        if fs > max_font_size { max_font_size = fs; }
                    }
                }
                // S504 (2026-06-08, default-ON opt-out OXI_S504_DISABLE): for PURE-LATIN
                // exact lines the bottom-align subtrahend should be the font GLYPH CELL
                // (ascent+descent), not the point size. `line_height_pt` = max(win,hhea)/em
                // × fs ≈ 1.2×fs for Latin, so the point-size subtrahend over-reserves Latin
                // exact lines by ~0.2×fs (db9ca title sz28 line=420: +2.33pt too low; Word
                // offset 4.17 = line 21 − cell 16.83 = 1.20×14; S504 → +0.23). SCOPED to
                // lines whose every fragment is NON-CJK: a CJK font's line_height_pt can
                // exceed fs (win>1.0em) and shifting CJK exact lines REGRESSED 34140b
                // −0.0006 — so CJK / mixed lines keep the S495 point-size subtrahend
                // (unchanged). Floored at fs. See session504_s495_latin_exact_overshoot.
                let s504_latin_line = std::env::var("OXI_S504_DISABLE").is_err()
                    && !line.fragments.is_empty()
                    && line.fragments.iter().all(|frag|
                        !self.metrics_for_text(&frag.text, &frag.style, para_style).is_cjk_83_64_font());
                let max_font_cell: f32 = if !s504_latin_line {
                    max_font_size
                } else {
                    let mut c = 0.0_f32;
                    for frag in &line.fragments {
                        let fs = frag.style.font_size.unwrap_or(para_font_size);
                        let cell = self.metrics_for_text(&frag.text, &frag.style, para_style).line_height_pt(fs).max(fs);
                        if cell > c { c = cell; }
                    }
                    c
                };
                if !in_shape_context {
                    // S495 (2026-06-05): BODY/CELL exact also bottom-aligns when the box
                    // exceeds the font cell (text at bottom, extra space above) — render-truth
                    // confirmed: repro line15/font10 Oxi was -3.91 vs Word, line21/font14
                    // -5.27 (= b5f706 -3.91, db9ca -5.27 to the decimal). Word leaves
                    // (line - fontCell) above. Floored at the S78 0.5pt so the box~=font case
                    // (fded6/04b88e, line~=font) is unchanged. opt-out OXI_S495_EXACT_BOTTOM_DISABLE.
                    if std::env::var("OXI_S495_EXACT_BOTTOM_DISABLE").is_ok() {
                        return 0.5;
                    }
                    // NB: the ideal subtrahend is the font's GLYPH CELL (where the dwrite
                    // renderer actually places ascent+descent ≈ 1.06*fs for MS Mincho), not
                    // font_size — using font_size over-reserves by ~0.06*fs (the a47e box≈font
                    // -0.0015..-0.0026 residual). But Oxi's `line_height_pt` (raw win/hhea cell)
                    // = font_size for the CJK fonts (OS/2 winAscent+winDescent = 1.0em for MS
                    // Mincho), so (line - natural) is a NO-OP for CJK and cannot remove the over
                    // — it is a dwrite-vs-OS/2 cell gap, a renderer-metric issue, not a natural-
                    // height one. For CJK this stays at max_font_size (S504 leaves CJK/mixed
                    // lines untouched); for PURE-LATIN lines max_font_cell = the glyph cell
                    // (line_height_pt ≈ 1.2×fs) which removes the ~0.2×fs Latin overshoot.
                    return (line_height - max_font_cell).max(0.5);
                }
                // Shape context: text at bottom of line box (extra space above).
                (line_height - max_font_size).max(0.0)
            }
            _ => {
                // ECMA-376 §17.3.1.35 textAlignment: "baseline" / "top" place
                // glyph at line top (no centering offset above). COM behavior
                // confirmed 2026-04-24 on e3c545_LOD_Handbook: pPrDefault
                // textAlignment="baseline" → Word P1 at top_margin + 0pt, not
                // + (lh-fs)/2. Without this, all body paragraphs drift +5-6pt
                // below Word.
                //
                // R7.63 (Day 36 part 10, 2026-05-14): only suppress centering
                // when the baseline/top setting was INHERITED FROM pPrDefault
                // (document-wide). Per-paragraph override on a single paragraph
                // (ed025c wi=827: explicit `<w:textAlignment w:val="baseline"/>`
                // in pPr) should NOT disable centering — it's a glyph-alignment
                // hint within the line, not a line-positioning override. Without
                // this gate, ed025c wi=827 shifts up 4pt relative to wi=826,
                // collapsing the line gap from 18pt to 14pt and cascading 148
                // paras +1 / 69 paras +2 / 69 paras +3 in the downstream
                // pagination.
                if matches!(para_style.text_alignment.as_deref(), Some("baseline") | Some("top"))
                    && para_style.text_alignment_from_pprdefault
                {
                    // S459 (2026-05-31) ★ SHIP — the baseline/top textAlignment
                    // early-return places the glyph at line-box top (0 line-
                    // centering), which is correct for LINE positioning, but it
                    // ALSO skipped the S455-457 CJK 83/64 glyph-in-box correction.
                    // Pixel-measured on e3c545_LOD_Handbook p2 (Meiryo 10.5pt,
                    // pPrDefault textAlignment=baseline — the ONLY corpus doc on
                    // this path): Word glyphs sit ~3.2-3.4pt LOWER than Oxi's
                    // 0-offset placement, UNIFORMLY on every line (band-corr
                    // 0.94-0.99, ink-center +3.1..3.8pt). Same family as S454-457
                    // but in the baseline-alignment path. The magnitude (3.2 vs
                    // the LM0 1.75 constant) is larger because Meiryo's leading is
                    // much larger than MS Mincho 9pt — the leading-proportional
                    // dependence S458 noted but could not fit as a clean global
                    // law. Returning the LM0 centering `base` instead was tested
                    // and is a NO-OP (word_line_height_table_cell ≈ line_height
                    // for Meiryo → base≈0), so a constant is the model here.
                    // Gated to CJK 83/64 lines (non-CJK baseline lines stay at 0,
                    // unchanged). δ-sweep (DWrite gate renderer) peaks at 3.2:
                    // LOD mean 0.7901→0.8233 (+0.033), all 12 pages improve or
                    // hold (p1 +0.12, p2 +0.15, p3 +0.05), ZERO regress. ONLY LOD
                    // hits this branch corpus-wide (verified: it is the sole doc
                    // with pPrDefault textAlignment=baseline) → no other doc can
                    // regress. Render-only (text_y_off) → element.y / pagination
                    // unchanged → Phase-1 sentinel preserved by construction.
                    // Opt-out / override via OXI_S459_BASELINE_CJK_DY (0 disables).
                    let s459 = std::env::var("OXI_S459_BASELINE_CJK_DY")
                        .ok()
                        .and_then(|v| v.parse::<f32>().ok())
                        .unwrap_or(3.2);
                    if s459 != 0.0 {
                        let line_is_cjk_8364 = if !line.fragments.is_empty() {
                            line.fragments.iter().any(|f| {
                                self.metrics_for_text(&f.text, &f.style, para_style)
                                    .is_cjk_83_64_font()
                            })
                        } else {
                            let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                            self.metrics_for_para_mark(&rpr_ref, para_style).is_cjk_83_64_font()
                        };
                        if line_is_cjk_8364 {
                            return s459;
                        }
                    }
                    return 0.0;
                }
                // Grid-snapped lines: text is vertically centered within the grid cell.
                // COM-confirmed: P1 20pt in 35.7pt grid cell → 4.9pt offset above text.
                // Compute natural height and center within line_height.
                let mut max_ascent: f32 = 0.0;
                let mut max_descent: f32 = 0.0;
                if line.fragments.is_empty() {
                    let font_size = para_style.ppr_rpr.as_ref()
                        .and_then(|r| r.font_size)
                        .unwrap_or(para_font_size);
                    let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                    let metrics = self.metrics_for_para_mark(&rpr_ref, para_style);
                    max_ascent = metrics.word_ascent_pt(font_size);
                    max_descent = metrics.word_descent_pt(font_size);
                } else {
                    for frag in &line.fragments {
                        let font_size = frag.style.font_size.unwrap_or(para_font_size);
                        let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                        let asc = metrics.word_ascent_pt(font_size);
                        let des = metrics.word_descent_pt(font_size);
                        if asc > max_ascent { max_ascent = asc; }
                        if des > max_descent { max_descent = des; }
                    }
                }
                // Apply vertical centering offset in the line box.
                // - LM1/LM2 (has_grid): center GDI cell (= fontSize) within grid cell.
                // - LM0 (no grid): still center the GDI cell within line_height.
                //   Measured 2026-04-21 via `bugA_size_sweep.py` (27 font×size combos):
                //   Word places first-line glyph top at margin + ~(line_height - fontSize)/2,
                //   NOT at cursor_y. Previous "no offset for LM0" assumption was based on a
                //   single test doc (test_line_heights with 11pt) where the effect is ~0.3pt
                //   and visually masked; larger fonts (26pt) show a -3.84pt glyph-top shift
                //   that cascades through gen2_* series (46 docs).
                let has_grid = grid_pitch.map_or(false, |p| p > 0.0) && para_style.snap_to_grid;
                if has_grid {
                    // COM-confirmed (2026-04-04/16): text is vertically centered within
                    // its grid-cell allocation. For single-cell lines (fontSize ≤ pitch),
                    // line_height == pitch so offset = (pitch - natural)/2. For multi-cell
                    // lines where line is snapped to n*pitch (fontSize > pitch, e.g.
                    // 20pt title in 17.85pt pitch = 2 cells = 35.7pt), centering uses the
                    // full line_height: offset = (line_height - natural)/2, which
                    // reduces to the single-cell case when n=1. Round to 0.5pt.
                    //
                    // COM-measured 2026-04-16 (verify_lm2_multicell_firstline.py, 35 samples):
                    //   MS Mincho 14pt (2 cells=36, pitch=18): offset=8.80 ≈ (36-18.15)/2
                    //   MS Mincho 18pt (2 cells=36):          offset=6.30 ≈ (36-23.34)/2
                    //   MS Mincho 24pt (2 cells=36):          offset=2.30 ≈ (36-31.12)/2
                    //   Yu Mincho 20pt (2 cells=36):          offset=1.30 ≈ (36-33.4)/2
                    //   Meiryo 20pt (3 cells=54):             offset=7.30 ≈ (54-25.9)/2
                    // Centering offset for grid-snapped line: center the GDI character
                    // cell (= fontSize) within the allocated line_height per spec §13.4.
                    // For single-cell lines (line_height == pitch, fontSize ≤ pitch):
                    //   offset = (pitch - fontSize)/2
                    // For multi-cell lines (line_height = n*pitch when fontSize > pitch):
                    //   offset = (line_height - fontSize)/2
                    // Both reduce to (line_height - fontSize)/2 because line_height = pitch
                    // in the single-cell case.
                    // S166 (2026-05-21): centering uses font's natural line height,
                    // not raw font_size. Phase A measurement (Word COM ground truth for
                    // 49 baseline docs) showed Oxi's previous `(line_height - font_size)/2`
                    // formula gave consistent +1.25-2.25pt over-offset on 10+ clean docs.
                    // Word's actual formula uses the font's table-cell line height
                    // (MS Mincho 10.5pt = 14.5pt; Latin 10.5pt = ~12.5pt). Paired with
                    // visual-position IoU metric (element_iou_diff.py R166), full baseline:
                    // mean IoU 0.9254 → 0.9295 (+0.0041), pass 12 → 15, Phase 1 53/55 unchanged.
                    // 12 docs gain (incl. 6a39b1/8bc929 → 1.0), 6 small losses (worst -0.031).
                    // S238 (2026-05-23): removed OXI_LEGACY_TEXT_Y_FONT_SIZE
                    // legacy env-var fallback during hardening pass.
                    let centering_height = if !line.fragments.is_empty() {
                        line.fragments.iter()
                            .map(|f| {
                                let fs = f.style.font_size.unwrap_or(para_font_size);
                                let m = self.metrics_for_text(&f.text, &f.style, para_style);
                                m.word_line_height_table_cell(fs)
                            })
                            .fold(0.0_f32, f32::max)
                    } else {
                        let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        let m = self.metrics_for_para_mark(&rpr_ref, para_style);
                        m.word_line_height_table_cell(para_font_size)
                    };
                    let pitch = grid_pitch.unwrap_or(0.0);
                    if pitch > 0.0 {
                        let raw = (line_height - centering_height).max(0.0) / 2.0;
                        // S328 (2026-05-26) — env-gated FLOOR variant.
                        // Default formula `(raw*2 + 0.5).floor() / 2` is
                        // CEIL-half-up to 0.5pt grid: for raw=0.25 returns
                        // 0.5pt (over-applied by 0.5pt vs Word's measured
                        // 0.0pt for some font+size combos). 35/55 docs in
                        // Phase 2 baseline have exactly +0.5pt first-line
                        // bias matching this over-application. FLOOR
                        // variant `(raw*2).floor() / 2` would reduce these
                        // by 0.5pt. Default OFF preserves baseline exactly.
                        //
                        // S329 (2026-05-26) — env-gated ROUND variant.
                        // FLOOR was empirically falsified (d1e8ac8 -0.0071)
                        // because raw values in (0.25, 0.5) were over-
                        // corrected. ROUND-half-away-from-zero `(raw*2).round()/2`
                        // is conservative: matches CEIL for raw in [0.25, 0.5]
                        // (preserves d1e8ac8 tuning), differs only for
                        // raw < 0.25 (the strict over-application case).
                        let use_floor = std::env::var("OXI_S328_FLOOR_CENTER")
                            .map(|v| v != "0" && v != "false")
                            .unwrap_or(false);
                        let use_round = std::env::var("OXI_S329_ROUND_CENTER")
                            .map(|v| v != "0" && v != "false")
                            .unwrap_or(false);
                        let grid_base = if use_floor {
                            (raw * 2.0).floor() / 2.0
                        } else if use_round {
                            (raw * 2.0).round() / 2.0
                        } else {
                            (raw * 2.0 + 0.5).floor() / 2.0
                        };
                        // S457 (2026-05-30) ★ SHIP — CJK 83/64 glyph correction
                        // for the GRID branch (LM1/LM2 docGrid), default +2.5pt,
                        // opt-out/override OXI_S457_GRID_CJK_DY (set 0 to disable).
                        // Mirrors the S455/S456 LM0 fix: docGrid CJK body glyphs
                        // sit too high within the grid cell because the centering
                        // formula does not account for where Word places the
                        // 83/64 excess leading. d77a/b837/c7b923 (the #2/#3 bottom
                        // docs) show a uniform ~5px down-shift recovers corr
                        // 0.9-0.99 (d77a p1 corr 0.992). δ-sweep peaks at 2.5 for
                        // all three (larger than LM0's 1.75 because the grid cell/
                        // line_height is larger → more 83/64 excess to place).
                        //
                        // GATE (full 235-doc recompute, DWrite, δ=2.5): mean
                        // 0.8967→0.9058 (+0.0091), bottom-5 +0.158, bottom-10
                        // +0.233, <0.70 bucket 17→8 (halved!), ≥0.95 162→174;
                        // 79 improved (ALL hex: d77a p2 +0.20/p6 +0.18/p1 +0.14,
                        // b837 p6 +0.17/p2 +0.12, c7b923 p2 +0.13, …), only 2
                        // regress (ed025 p6 −0.018 — known Phase-1-sensitive
                        // cascade S305; 2ea81 p2 −0.001). b35 (the S430 opposite-
                        // direction CELL family) does NOT regress here (+0.003) —
                        // its −1.5 was the CELL glyph path, not this grid body
                        // path. Render-only → element.y/pagination unchanged →
                        // Phase-1 sentinel preserved. TODO: LM0 1.75 vs grid 2.5
                        // suggests a single line-height-proportional δ (the
                        // S455/S456 leading-proportional idea, now with two
                        // regimes to fit).
                        let s457_dy = if line.fragments.iter().any(|f| {
                            self.metrics_for_text(&f.text, &f.style, para_style)
                                .is_cjk_83_64_font()
                        }) {
                            std::env::var("OXI_S457_GRID_CJK_DY")
                                .ok()
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(2.5)
                        } else {
                            0.0
                        };
                        grid_base + s457_dy
                    } else {
                        0.0
                    }
                } else {
                    // LM0: same centering formula as LM1/LM2 single cell.
                    // S238 (2026-05-23): removed OXI_LEGACY_TEXT_Y_FONT_SIZE
                    // legacy env-var fallback during hardening pass.
                    let centering_height = if !line.fragments.is_empty() {
                        line.fragments.iter()
                            .map(|f| {
                                let fs = f.style.font_size.unwrap_or(para_font_size);
                                let m = self.metrics_for_text(&f.text, &f.style, para_style);
                                m.word_line_height_table_cell(fs)
                            })
                            .fold(0.0_f32, f32::max)
                    } else {
                        let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        let m = self.metrics_for_para_mark(&rpr_ref, para_style);
                        m.word_line_height_table_cell(para_font_size)
                    };
                    let raw = (line_height - centering_height).max(0.0) / 2.0;
                    // S328/S329 (2026-05-26) — see comments above (LM1/LM2 branch).
                    let use_floor = std::env::var("OXI_S328_FLOOR_CENTER")
                        .map(|v| v != "0" && v != "false")
                        .unwrap_or(false);
                    let use_round = std::env::var("OXI_S329_ROUND_CENTER")
                        .map(|v| v != "0" && v != "false")
                        .unwrap_or(false);
                    let base = if use_floor {
                        (raw * 2.0).floor() / 2.0
                    } else if use_round {
                        (raw * 2.0).round() / 2.0
                    } else {
                        (raw * 2.0 + 0.5).floor() / 2.0
                    };
                    // S454 (2026-05-30) [finding — env-gated OFF, default 0.0]
                    // — the LM0 (no-docGrid) body glyph is ~1.75pt too HIGH for
                    // the CONTRACT FAMILY (0e7af/683f, MS Mincho 9pt, pure
                    // single-column body) — the LM0 analogue of the S453
                    // cell-glyph offset. CORRECTS the 43-day-old
                    // project_0e7a_p2_drift memory: it is NOT a pagination drift
                    // — pagination AND line-box tops are CORRECT (COM: 第1条
                    // 69.5/69.75, para#61 93.0/93.0; element.y byte-identical
                    // across δ). Only the glyph-in-box vertical offset differs.
                    // SSIM δ-sweep (DWrite gate renderer) peaks ≈1.5 (0e7af) to
                    // 2.0 (683f); pixel best-shift +1.6..1.84pt; +0.13/page on
                    // 0e7af's 6 bottom pages.
                    //
                    // WHY NOT SHIPPED as a flat constant: the correction is the
                    // WRONG SHAPE. The default-font (10.5pt) test fixtures
                    // (test_widow/keepnext/line_height) peak SHARPLY at δ=0
                    // (already correctly placed) — any δ>0 regresses them
                    // (test_widow −0.12 at +1.5, −0.089 even at +1.0). Full
                    // 235-doc recompute at δ=1.5: mean 0.8862→0.8949 (+0.0086),
                    // bottom-5 +0.36, REAL-hex corpus clean (+22/−1) BUT 50
                    // SYNTHETIC fixtures regress. A flat constant is known-wrong
                    // for the default-font class → violates fidelity + CLAUDE.md
                    // anti-exception-stacking. The real fix is a per-font
                    // centering re-derivation (word_line_height_table_cell
                    // calibration for CJK-Mincho 9pt undershoots the half-
                    // leading) — deferred, multi-step COM work. Knob retained
                    // for that investigation. Render-only → Phase-1-safe.
                    let s454_dy = std::env::var("OXI_S454_LM0_GLYPH_DY")
                        .ok()
                        .and_then(|v| v.parse::<f32>().ok())
                        .unwrap_or(0.0);
                    // S455 (2026-05-30) ★ SHIP — LM0 body-glyph vertical
                    // correction scoped to CJK 83/64 fonts (MS Mincho/Gothic/
                    // Meiryo), default +1.5pt, opt-out/override via
                    // OXI_S455_CJK_GLYPH_DY (set 0 to disable). Roots the
                    // corpus #1 bottom doc 0e7af (contract, MS Mincho 9pt) and
                    // sister 683f: Oxi drew the LM0 body glyph ~1.75pt too HIGH
                    // within the (correctly-positioned, element.y byte-identical)
                    // line box — the LM0 analogue of the S453 cell-glyph offset.
                    //
                    // WHY CJK-SCOPED (the discriminator that made it shippable):
                    // the 83/64 multiplier inflates CJK line_height by
                    // ~0.297·tmHeight and Word places that extra leading ABOVE
                    // the glyph, but the centering formula above does not.
                    // Non-CJK fonts (Calibri/Cambria — the test_widow/keepnext/
                    // line_height fixtures) use a different line-height path and
                    // are ALREADY correct (peak SHARPLY at δ=0); an unscoped flat
                    // δ regressed all 50 of them (S454). Gating on the line's
                    // dominant font being CJK 83/64 is principled (tied to the
                    // 83/64 mechanism, NOT a per-font carve-out) and excludes
                    // them exactly (verified: test_* FLAT across δ).
                    //
                    // GATE (full 235-doc recompute, DWrite, δ=1.5): mean
                    // 0.8862→0.8953 (+0.0090), bottom-5 page sum +0.364,
                    // <0.70 bucket 25→18, ≥0.90 205→230, ≥0.99 42→45;
                    // 90 improved (22 real-hex: 0e7af +0.092, 683f +0.081,
                    // 9a8e8d +0.009, b837 +0.005, 1ec1 +0.003, …), 5 regress
                    // (all tiny CJK over-correction, worst −0.0059 gen_jp_report;
                    // 6295 order_09 −0.0029). Passes the Phase-3 gate (mean↑ AND
                    // bottom-N↑). Render-only (el.text_y_off) → element.y /
                    // pagination unchanged → Phase-1 sentinel provably preserved.
                    // TODO: 683f optimum ≈2.0; a leading-proportional δ
                    // (∝ the 83/64 excess) would fit larger-leading docs and
                    // shave the 5 small over-corrections.
                    let line_is_cjk_8364 = if !line.fragments.is_empty() {
                        line.fragments.iter().any(|f| {
                            self.metrics_for_text(&f.text, &f.style, para_style)
                                .is_cjk_83_64_font()
                        })
                    } else {
                        let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                        self.metrics_for_para_mark(&rpr_ref, para_style).is_cjk_83_64_font()
                    };
                    // S456 (2026-05-30) — magnitude refined 1.5 -> 1.75.
                    // 1.75 is the COM/pixel-measured offset (0e7af best-shift
                    // 1.6-1.84pt) AND 0e7af's own SSIM peak; per-doc optima
                    // (683f/9a8e8d/1ec1 want 2.0-2.25+, 0e7af 1.75) cluster
                    // above 1.5. Full 235-doc recompute 1.5 -> 1.75: mean
                    // 0.8953 -> 0.8967 (+0.0014), bottom-10 +0.0248, 63 up /
                    // 16 down. 2.0 was marginally higher mean (+0.0017) but
                    // doubled regressions (28) and lost bottom-10, so 1.75 is
                    // the principled peak. Per-doc optima don't fit a clean
                    // size-proportional law (683f and gen_jp_report both 10.5pt
                    // want 2.25 vs 1.0), so a constant is the right model.
                    let s455_dy = if line_is_cjk_8364 {
                        std::env::var("OXI_S455_CJK_GLYPH_DY")
                            .ok()
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(1.75)
                    } else {
                        0.0
                    };
                    base + s454_dy + s455_dy
                }
            }
        }
    }

    fn layout_table(
        &self,
        table: &Table,
        start_x: f32,
        cursor: &mut LayoutCursor,
        content_width: f32,
        grid_pitch: Option<f32>,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
        page_top: f32,
        content_height: f32,
        page_width: f32,
        page_height: f32,
        pages: &mut Vec<LayoutPage>,
        current_elements: &mut Vec<LayoutElement>,
        block_idx: Option<usize>,
        page: &Page,
        is_nested: bool,
    ) -> Vec<LayoutElement> {
        let mut elements = Vec::new();

        // Resolve column widths from grid_columns, cell widths, or equal split
        let col_widths = self.resolve_table_col_widths(table, content_width);
        let table_width: f32 = col_widths.iter().sum();

        // Table positioning: tblpPr horizontal or inline alignment
        let table_x = if let Some(ref pos) = table.style.position {
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
            let pad_l_default = table.style.default_cell_margins.as_ref().and_then(|m| m.left).unwrap_or(4.95);
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
            let border_offset = {
                let border_w = table.style.border_width.unwrap_or(0.5);
                match table.style.indent {
                    Some(v) if v.abs() < 0.01 && !table.style.explicit_borders => {
                        pad_l_default + border_w / 2.0
                    }
                    _ => 0.0,
                }
            };
            match table.style.alignment.as_deref() {
                Some("center") => start_x + (content_width - table_width) / 2.0,
                Some("right") => start_x + content_width - table_width,
                _ => start_x + table.style.indent.unwrap_or(0.0) - border_offset,
            }
        };

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
        let bug_a_enabled = if std::env::var("OXI_BUG_A_REVERT").is_ok() {
            false  // research toggle: always skip
        } else {
            // Default (S151): apply only for non-linesAndChars docs
            grid_char_pitch.is_none()
        };
        if bug_a_enabled && table.style.border {
            let top_bw = table.style.border_width.unwrap_or(0.4);
            cursor.advance(top_bw);
        }

        let num_rows = table.rows.len();
        let dump_table = std::env::var("OXI_DUMP_TABLE").is_ok();
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
                        p.runs.iter().any(|run| run.text.chars().any(kinsoku::is_cjk))
                    } else { false }
                })
            })
        });
        for (row_idx, row) in table.rows.iter().enumerate() {
            let mut row_height: f32 = 0.0;
            // Session 79c: visual_row_h = max cell content_h with emit-equivalent
            // line-height formula (grid-snapped when adjustLineHeightInTable). Used
            // ONLY for vAlign=center offset, NOT for row_height (page break logic
            // preserves the natural pre-pass to avoid 3a4f9f cascade — see
            // session79_adjust_lh_in_table_mixed_cell_valign_falsified.md).
            let mut visual_row_h: f32 = 0.0;
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
            let mut center_row_h: f32 = 0.0;
            let row_entry_cursor_y = cursor.cursor_y;

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
                && std::env::var("OXI_S361_TRHEIGHT_NO_LINE_SNAP").is_ok() {
                None
            } else {
                table_grid_pitch
            };

            // First pass: calculate row height
            let mut grid_idx = row.grid_before as usize;
            for cell in row.cells.iter() {
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
                let cell_w: f32 = col_widths[grid_idx..grid_idx + span].iter().sum();
                let _pad_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(default_pad_l);
                let _pad_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(default_pad_r);
                let mut pad_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(default_pad_t);
                let pad_b = cell.margins.as_ref().and_then(|m| m.bottom).unwrap_or(default_pad_b);
                // Round 30: implicit border padding (matches second pass)
                // S359 (2026-05-27): test confirmed Round 30 is load-bearing
                // (OXI_S359_NO_ROUND30=1 caused -0.0186 corpus regression).
                // S386 (2026-05-27): hypothesis "bug_a + Round30 double-count the
                // top border on row 0" FALSIFIED. OXI_S386_NO_DOUBLE_BORDER=1
                // (suppress Round30 on row 0 when bug_a fired) → corpus 0.9603
                // → 0.9521 (-0.0082, pass 18→17), and b5f706 barely moved
                // (0.9715→0.9707) because the iou_yrange_adj median absorbs
                // uniform per-table shifts. Round30 on row 0 is load-bearing.
                if pad_t == 0.0 && table.style.border {
                    pad_t = table.style.border_width.unwrap_or(0.4);
                }
                // COM-confirmed (2026-04-09, 10 minimal repros + 3 real docs):
                // Each row's height includes its BOTTOM-EDGE border:
                //   - Non-last rows: bottom edge = insideH width (0 if no insideH)
                //   - Last row: bottom edge = outer bottom border (0 if none)
                // Top/side borders do NOT add to row height.
                // OOXML default single border sz=4 = 0.5pt (4/8).
                let bw = table.style.border_width.unwrap_or(if table.style.border { 0.5 } else { 0.0 });
                let is_last = row_idx + 1 == num_rows;
                let _border_overhead = if is_last {
                    if table.style.border { bw } else { 0.0 }
                } else if table.style.has_inside_h {
                    bw
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
                let inner_w = cell_w.max(0.0);
                let mut cell_content_h = pad_t;
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

                // Session 131 (2026-05-20): vertical writing — cell height
                // along the page-y axis equals the sum of vertical-text lengths
                // (chars × font_size), not the wrapped-horizontal line count.
                // Gated by OXI_VERT_WRITING env var.
                let vert_writing_active = self.is_vert_writing_active(cell);
                for block in &cell.blocks {
                    match block {
                        Block::Paragraph(para) => {
                            let (para_h, para_h_visual, para_h_center) = if vert_writing_active {
                                let h = self.vert_para_height(para);
                                (h, h, h)
                            } else {
                                let p1 = self.estimate_para_height(para, inner_w, row_line_pitch, table.style.para_style.as_ref(), true, grid_char_pitch, grid_char_cw_ratio);
                                let p2 = self.estimate_para_height_emit(para, inner_w, row_line_pitch, table.style.para_style.as_ref(), true, grid_char_pitch, grid_char_cw_ratio);
                                // S503: render-line-height variant for centering floor
                                // (opt-in; default OFF avoids the extra estimate call).
                                let p3 = if s503_enable {
                                    self.estimate_para_height_emit_render(para, inner_w, row_line_pitch, table.style.para_style.as_ref(), true, grid_char_pitch, grid_char_cw_ratio)
                                } else { p2 };
                                (p1, p2, p3)
                            };
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
                            cell_content_h += para_h;
                            cell_content_h_visual += para_h_visual;
                            // S427: collapse this paragraph's space_before against
                            // the previous paragraph's space_after.
                            let (cur_sb, cur_sa) = self.cell_para_spacing(para, table.style.para_style.as_ref(), row_line_pitch);
                            if s427_collapse {
                                if let Some(psa) = prev_sa {
                                    let credit = psa.min(cur_sb);
                                    cell_content_h -= credit;
                                    cell_content_h_visual -= credit;
                                }
                            }
                            prev_sa = Some(cur_sa);
                        }
                        Block::Table(nested) => {
                            prev_sa = None;
                            // Estimate nested table height from rows
                            // COM-confirmed: nested table width = cell width - 2 × padding
                            let nested_w = (inner_w).max(0.0);
                            for nr in &nested.rows {
                                let mut nr_h = 0.0_f32;
                                for nc in &nr.cells {
                                    let mut nc_h = 0.0_f32;
                                    for nb in &nc.blocks {
                                        if let Block::Paragraph(np) = nb {
                                            nc_h += self.estimate_para_height(np, nested_w / 2.0, table_grid_pitch, nested.style.para_style.as_ref(), true, grid_char_pitch, grid_char_cw_ratio);
                                        }
                                    }
                                    nr_h = nr_h.max(nc_h);
                                }
                                if let Some(h) = nr.height {
                                    match nr.height_rule.as_deref() {
                                        Some("exact") => { nr_h = h; }
                                        Some("atLeast") => { nr_h = nr_h.max(h); }
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
                            cell_content_h += img.height;
                            cell_content_h_visual += img.height;
                        }
                        _ => {}
                    }
                }
                cell_content_h += pad_b;
                cell_content_h_visual += pad_b;
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
                if std::env::var("OXI_S375_HALF_INSIDEH").is_ok() && !is_last && table.style.has_inside_h {
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
                if std::env::var("OXI_S463_DISABLE").is_err()
                    && table.style.has_inside_h && table_is_latin
                    && !self.doc_body_has_cjk {
                    cell_content_h += _border_overhead * 0.375;
                    cell_content_h_visual += _border_overhead * 0.375;
                }

                if dump_table {
                    let ftext: String = cell.blocks.iter().filter_map(|b| {
                        if let Block::Paragraph(p) = b {
                            Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                        } else { None }
                    }).collect::<Vec<_>>().join("|");
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

            // Apply trHeight constraint.
            // 2026-04-09 (COM re-verified, 0e7a contract sample table 1):
            //   <w:trHeight w:val="830"/> with NO w:hRule attribute →
            //   Word reports HeightRule = atLeast (1) and renders the row
            //   at exactly val (41.5pt = 830tw), not at content height.
            //   Treat missing hRule as atLeast to match Word behavior.
            if let Some(h) = row.height {
                match row.height_rule.as_deref() {
                    Some("exact") => { row_height = h; }
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
                    _ => { row_height = row_height.max(h); }
                }
            }

            if row_height == 0.0 {
                let metrics = self.doc_default_metrics();
                row_height = self.line_height_inner(self.default_font_size, None, None, metrics, true, table_grid_pitch, true);
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
            let has_content = !elements.is_empty() || !current_elements.is_empty();
            let page_bottom = page_top + content_height;
            let row_overflows = cursor.cursor_y + row_height > page_bottom;
            // R7.47 (Day 34 part 16, 2026-05-13): row-level SOFT LRPB. When
            // ANY cell's FIRST paragraph carries `<w:lastRenderedPageBreak/>`
            // on its run[0], Word's saved render broke before this row.
            // Mirrors the body-paragraph LRPB SOFT rule at mod.rs:1888.
            // de6e / 29dc6e outliers (4 each) had LRPB-at-cell-start markers
            // that the table layout previously ignored.
            let row_has_lrpb_at_cell_start = row.cells.iter().any(|cell| {
                cell.blocks.first().map_or(false, |b| match b {
                    Block::Paragraph(p) => p.runs.first()
                        .map(|r| r.has_last_rendered_page_break).unwrap_or(false),
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
                && has_content
                && !row_overflows  // row would fit; LRPB hint says break anyway
                && consumed_row > lrpb_threshold;
            if std::env::var("OXI_DUMP_ROW_LRPB").is_ok() && row_has_lrpb_at_cell_start {
                let preview: String = row.cells.iter().filter_map(|c| {
                    c.blocks.first().and_then(|b| match b {
                        Block::Paragraph(p) => Some(p.runs.iter().flat_map(|r| r.text.chars()).take(20).collect::<String>()),
                        _ => None,
                    })
                }).next().unwrap_or_default();
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
                                if run.has_last_rendered_page_break
                                    && !(ci == 0 && is_first_para_in_cell && ri == 0)
                                {
                                    found = true;
                                    break;
                                }
                            }
                            if found { break; }
                        }
                    }
                    if found { break; }
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
            let is_single_row_single_cell = table.rows.len() == 1
                && table.rows.get(0).map_or(false, |r| r.cells.len() == 1);
            let widow_break_needed = row_idx == 0 && has_content && is_single_row_single_cell && {
                let free_space = page_bottom - cursor.cursor_y;
                let widow_threshold = if let Some(pitch) = table_grid_pitch {
                    pitch * 4.0
                } else { 58.0 };
                free_space > 0.0 && free_space < widow_threshold
            };

            // S533 (2026-06-10): a row carrying an inline IMAGE block is pushed
            // WHOLE to the next page instead of splitting (when it fits a fresh
            // page). Word treats the image's line as atomic — 3a4f's calendar
            // row (466pt cell: 321.75pt EMF + paragraphs, trHeight 7910 atLeast)
            // starts fresh on Word's p34; Oxi's element-level split stranded the
            // image across the boundary and the post-split cursor under-advanced,
            // overlapping the following content. Rows taller than a full page
            // still split (unavoidable).
            let row_has_image_block = row.cells.iter()
                .any(|c| c.blocks.iter().any(|b| matches!(b, Block::Image(_))));
            let image_atomic_push = row_has_image_block && row_height <= content_height;

            // needs_row_split: only when overflow + table allows split.
            // widow_break_needed overrides split — we want the whole table on next page.
            let needs_row_split = row_overflows && !row.cant_split && has_content
                && (is_single_cell_row || has_lrpb_mid_row)
                && !widow_break_needed
                && !image_atomic_push;

            if (row_overflows || lrpb_row_should_break || widow_break_needed) && has_content && !needs_row_split {
                // Push all accumulated elements (including previous rows) to current page
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });
                cursor.set(page_top);
            }

            // Second pass: render cells
            // Track actual content height per cell for row_height correction
            let is_exact_row = row.height_rule.as_deref() == Some("exact");
            let mut max_actual_cell_h: f32 = row_height;
            let elements_before_row = elements.len();
            // S500 (L1) FALSIFIED (2026-06-06): re-centering vAlign center/bottom cells against
            // the FINAL row height (fixing early cells centered before later/taller cells set
            // max_actual_cell_h) FIXED the synthetic repro vc_2cell_auto (-1.65->+0.10) but was
            // a NO-OP on the real corpus (net -0.0005; every bottom-N page +/-0.0004, 2ea81a
            // -0.0004) — the stale-height ordering doesn't manifest in real docs (centered cells
            // are single-cell rows or similar-height) and it does NOT fix d4d126's +3.3 (the
            // over-estimate direction). Reverted. See cellY_perdoc_scoped_design.md.
            // Apply gridBefore: skip leading grid columns
            let mut grid_idx: usize = row.grid_before as usize;
            let mut cell_x = table_x + col_widths[..grid_idx.min(col_widths.len())].iter().sum::<f32>();
            let _num_cells = row.cells.len();
            for (cell_idx, cell) in row.cells.iter().enumerate() {
                let span = cell.grid_span.max(1) as usize;
                // vMerge="continue" cells: skip content but still draw borders
                let is_vmerge_continue = cell.v_merge.as_deref() == Some("continue") || cell.v_merge.as_deref() == Some("");
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

                let pad_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(default_pad_l);
                let pad_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(default_pad_r);
                let mut pad_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(default_pad_t);
                let pad_b = cell.margins.as_ref().and_then(|m| m.bottom).unwrap_or(default_pad_b);

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
                let lead_absorb = self.compat_mode <= 14
                    && table.style.indent.map_or(false, |v| v > 0.1);
                if cell_start_grid == 0 && !is_nested && lead_absorb
                    && std::env::var("OXI_S496_TBLIND_DISABLE").is_err() {
                    cell_x -= pad_l;
                    cell_w += pad_l;
                }

                // Round 30 (2026-04-09): When cell top/bottom padding is 0 and
                // the table has borders, add the border width as implicit padding.
                // Word positions text below the top border line, not at the border.
                // COM-confirmed minimal repro: Table Grid with tcMar=0 all sides,
                // MS Mincho 12pt → text_y = topMargin + 0.5pt (= border width).
                // S359 (2026-05-27): test confirmed Round 30 is load-bearing
                // (OXI_S359_NO_ROUND30=1 caused -0.0186 corpus regression).
                // S386 (2026-05-27): double-border-count hypothesis FALSIFIED
                // (see height-calc site above; -0.0082 corpus regression).
                if pad_t == 0.0 && table.style.border {
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
                        elements.push(LayoutElement::new(cell_x, cursor.visual_y, cell_w, row_height, LayoutContent::CellShading {
                                color: color_hex,
                        }));
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

                // Layout blocks in document order (paragraphs and nested tables interleaved)
                let is_exact = row.height_rule.as_deref() == Some("exact");
                // R7.32: count Paragraph blocks within this cell so each cell
                // paragraph can be distinguished in the dump output.
                let mut cell_para_counter: usize = 0;
                // R7.73: track whether the immediately-previous cell paragraph
                // carried a `<w:lastRenderedPageBreak/>` on a non-run-0 run.
                // Reset to false at each cell start.
                let mut prev_cell_para_had_mid_lrpb: bool = false;
                // S427 (2026-05-29): track previous cell paragraph's space_after
                // for adjacent-paragraph spacing collapse (see pre-pass comment).
                let s427_collapse = std::env::var("OXI_S427_DISABLE").is_err();
                let mut prev_cell_sa: Option<f32> = None;
                if !is_vmerge_continue {
                // S428 (2026-05-29): index of the last cell block that carries
                // real content (a non-empty paragraph or a nested table). Used to
                // gate the empty-paragraph zero-glyph element emission below to
                // only INTERIOR empty paragraphs (those followed by content). A
                // trailing empty paragraph must NOT get an element, else it would
                // overflow a mid-cell page split alone and spawn a near-blank
                // continuation page (e3c545: a lone trailing empty cell paragraph
                // created a blank page 5, cascading every later page +1).
                let last_content_block_pos: Option<usize> = cell.blocks.iter().enumerate()
                    .filter(|(_, b)| match b {
                        Block::Paragraph(p) => p.runs.iter().any(|r| !r.text.is_empty()),
                        _ => true,
                    })
                    .map(|(i, _)| i)
                    .last();
                for (block_pos, block) in cell.blocks.iter().enumerate() {
                // S488: snapshot this block's content_h-relative top (aligns with
                // block_pos via enumerate; pushed before the exact-clip break so
                // blocks that fit are all recorded).
                debug_assert_eq!(cell_block_tops.len(), block_pos);
                cell_block_tops.push(content_h);
                // Clip content that overflows exact row height
                if is_exact && content_h + pad_t >= row_height {
                    break;
                }
                match block {
                Block::Table(nested) => {
                    // COM-confirmed: nested table width = outer cell width - 2 × padding
                    let nested_x = cell_x + pad_l;
                    let nested_content_w = (cell_w - pad_l - pad_r).max(0.0);
                    let mut nested_y = LayoutCursor::new(content_h);
                    let mut dummy_pages = Vec::new();
                    let mut dummy_elems = Vec::new();
                    let nested_elements = self.layout_table(
                        nested, nested_x, &mut nested_y, nested_content_w, table_grid_pitch,
                        grid_char_pitch,
                        grid_char_cw_ratio,
                        0.0, 99999.0, 0.0, 99999.0,
                        &mut dummy_pages, &mut dummy_elems,
                        block_idx,
                        page,
                        true,
                    );
                    for elem in nested_elements {
                        cell_elements.push(elem);
                    }
                    content_h = nested_y.cursor_y;
                    prev_cell_sa = None; // S427: nested table breaks paragraph adjacency
                }
                Block::Paragraph(para) => {
                let para = para;
                    // Session 131 (2026-05-20): vertical writing early-exit.
                    // For tbRlV cells, emit one Text element per paragraph at
                    // relative_y=0 (Word's COM Information(6) on a vert-cell
                    // paragraph returns the row-top y for all paragraphs in
                    // that cell). Cell content_h grows by vert_para_height so
                    // row-height calc reflects the vertical-text extent.
                    // The renderer (S132 GDI, S133 DWrite) is responsible for
                    // actual 90° CW rotation when emitting glyphs; this layout
                    // step only ensures positional correctness for pagination.
                    if self.is_vert_writing_active(cell) {
                        let vert_h = self.vert_para_height(para);
                        let first_run_style = para.runs.first()
                            .map(|r| r.style.clone())
                            .unwrap_or_default();
                        let first_run_fs = self.resolve_font_size(&first_run_style, &para.style);
                        let para_text: String = para.runs.iter()
                            .flat_map(|r| r.text.chars())
                            .collect();
                        let font_family = self.resolve_font_family_for_text(
                            &para_text, &first_run_style, &para.style,
                        ).map(|s| s.to_string());
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
                                underline_style: first_run_style.underline_style.clone(),
                                strikethrough: first_run_style.strikethrough,
                                double_strikethrough: first_run_style.double_strikethrough,
                                color: first_run_style.color.clone(),
                                highlight: first_run_style.highlight.clone(),
                                character_spacing: 0.0,
                                field_type: None,
                                text_scale: first_run_style.text_scale.unwrap_or(100.0),
                                // Session 132: flag for renderer rotation.
                                is_vertical: true,
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
                    // Apply table style pPr as fallback (ECMA-376: table style pPr < paragraph style < direct)
                    // Word resets line spacing to Single and space_after to 0 for table cell
                    // paragraphs that inherit from Normal style (no direct spacing in pPr).
                    // COM-measured: Normal outside table = ls=13.80(1.15x) sa=10,
                    //               Normal inside table = ls=12.00(Single) sa=0.
                    // COM-confirmed: Word resets docDefaults-only lineSpacing to Single in table cells,
                    // but keeps Normal style's lineSpacing. gen2_036: docDefaults line=276 → cell ls=12(Single).
                    // test_table_borders.docx: Normal style line=276 → cell ls=13.80(1.15x).
                    let effective_line_spacing = if para.style.line_spacing_from_doc_defaults {
                        None // Reset to single
                    } else {
                        para.style.line_spacing
                            .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.line_spacing))
                    };
                    let effective_line_rule = if para.style.line_spacing_from_doc_defaults {
                        None
                    } else {
                        para.style.line_spacing_rule.as_deref()
                            .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.line_spacing_rule.as_deref()))
                    };
                    let style_has_explicit_rule = effective_line_rule == Some("exact") || effective_line_rule == Some("atLeast");
                    let should_reset = !para.style.has_direct_spacing && !style_has_explicit_rule;
                    let tbl_has_ls = table.style.para_style.as_ref().and_then(|ps| ps.line_spacing).is_some();
                    let (effective_line_spacing, effective_line_rule) = if tbl_has_ls && !para.style.has_direct_spacing {
                        let tbl_ls = table.style.para_style.as_ref().and_then(|ps| ps.line_spacing);
                        let tbl_lr = table.style.para_style.as_ref().and_then(|ps| ps.line_spacing_rule.as_deref());
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
                    let effective_space_before = if should_reset {
                        // Day 33 part 17 (2026-05-10): Word suppresses spacing.before
                        // for the first paragraph in a cell. COM-confirmed via 8 repros
                        // (row1_attr_isolation): R1A_spacing_lineRule has spacing.before=4.35pt
                        // + lineRule=exact 12pt → Word renders row at 12.5pt (no spacing
                        // applied), Oxi was rendering at 16.85pt (+4.35pt over-pump).
                        // Same for R1A_all4. Affects 备考 cluster docs (d4d126/de6e/etc)
                        // where row 1 cell has style "ac"+spacing.before+lineRule=exact.
                        0.0
                    } else if let (Some(bl), Some(pitch)) = (para.style.before_lines, table_grid_pitch) {
                        bl / 100.0 * pitch
                    } else {
                        para.style.space_before
                            .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.space_before))
                            .unwrap_or(0.0)
                    };
                    let effective_space_after = if should_reset {
                        None
                    } else if let (Some(al), Some(pitch)) = (para.style.after_lines, table_grid_pitch) {
                        // Session 94 (2026-05-18) fix: afterLines was parsed into
                        // IR but not applied in cell rendering path. Body path at
                        // mod.rs:4554 already had this. Symmetric with before_lines
                        // handling at mod.rs:6497. TR33 (afterLines only, no after
                        // twip) measured Word pitch 13.50pt vs Oxi 12.00pt = +1.5pt
                        // gap closed by reading afterLines.
                        Some(al / 100.0 * pitch)
                    } else {
                        para.style.space_after
                            .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.space_after))
                    };
                    // S427: collapse this paragraph's space_before against the
                    // previous cell paragraph's space_after (max(sa,sb), not sum).
                    if s427_collapse {
                        if let Some(psa) = prev_cell_sa {
                            content_h -= psa.min(effective_space_before);
                        }
                    }
                    content_h += effective_space_before;
                    let para_content_start_h = content_h;
                    {
                        // Paragraph indentation within cell (relative to cell content area)
                        // COM-confirmed: *Chars multiplier = 10.5pt always
                        let p_indent_left = para.style.indent_left
                            .or_else(|| para.style.indent_left_chars.map(|c| c / 100.0 * 10.5))
                            .unwrap_or(0.0);
                        let p_indent_right = para.style.indent_right
                            .or_else(|| para.style.indent_right_chars.map(|c| c / 100.0 * 10.5))
                            .unwrap_or(0.0);
                        // When both firstLine (twip) and firstLineChars exist,
                        // twip value is authoritative (pre-computed by Word).
                        let p_first_line_indent_raw = para.style.indent_first_line
                            .or_else(|| para.style.indent_first_line_chars.map(|c| c / 100.0 * 10.5))
                            .unwrap_or(0.0);
                        // COM-confirmed (2026-04-25): numbered list + hanging + suff=tab/default
                        // => marker consumes hanging, text starts at `left`. See body path.
                        let p_list_consumes_hanging = para.style.list_marker.is_some()
                            && p_first_line_indent_raw < 0.0
                            && matches!(para.style.list_suff.as_deref(), None | Some("tab"));
                        let p_first_line_indent = if p_list_consumes_hanging { 0.0 } else { p_first_line_indent_raw };
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
                        let s301_layout_fixed = std::env::var("OXI_S301_DISABLE").is_err()
                            && table.style.layout.as_deref() == Some("fixed")
                            && (pad_l + pad_r) > 0.0
                            && row.cells.len() == 2
                            && cell_w <= content_width
                            && !para.style.word_wrap;  // tight discriminator: wordWrap=0 only
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
                        let s412_cellmar_subtract = std::env::var("OXI_S412_DISABLE").is_err()
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
                        let s531_singlecell_cellmar = std::env::var("OXI_S531_DISABLE").is_err()
                            && row.cells.len() == 1
                            && table.style.default_cell_margins.is_some()
                            && !table.style.has_explicit_cellmar
                            && !matches!(para.alignment, Alignment::Right | Alignment::Center)
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
                        let s559_tcw_gt = cell.width.map_or(false, |tcw| tcw - cell_w >= 8.0);
                        let s559_base = !s559_disabled
                            && row.cells.len() == 1
                            && !table.style.has_explicit_cellmar
                            && !matches!(para.alignment, Alignment::Right | Alignment::Center)
                            && cell_w <= content_width;
                        let s559_justified =
                            matches!(para.alignment, Alignment::Justify | Alignment::Distribute);
                        let s559_cellmar = s559_base && match s559_mode.as_deref() {
                            Some("all") => true,
                            Some("tcwgt") => s559_tcw_gt,
                            Some("justl0") => s559_tcw_gt && s559_justified && p_indent_left.abs() < 1.0,
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
                            .ok().and_then(|v| v.parse().ok()).unwrap_or(11.0);
                        let s585_tblw_auto = table.style.width_type.as_deref() == Some("auto");
                        let s585_over = cell_w - content_width;
                        let s585_cellmar = std::env::var("OXI_S585_DISABLE").is_err()
                            && row.cells.len() == 1
                            && !table.style.has_explicit_cellmar
                            && !matches!(para.alignment, Alignment::Right | Alignment::Center)
                            && cell_w > content_width
                            && (s585_over < 5.0 || (s585_tblw_auto && s585_over < s585_auto_over));
                        // S594 (2026-06-17, opt-IN OXI_S594=1): narrow the S585b cell
                        // wrap_base by ONE EXTRA cellMar. S585c: Oxi's cell right border is
                        // at content+2×cellMar vs Word's +1× → Oxi's cell_w over-computes by
                        // ~1 cellMar, so wrap_base=cell_w−2×pad=417.95 is still ~5pt wider
                        // than Word's true content ≈413 (=cell_w−3×cellMar). The 66 residual
                        // CELL over-fit roots (S7m) are this width over. Subtract one more
                        // cellMar to reach Word's content (single-cell S585b tables only;
                        // 3a4f/model 第N条 are in BODY → unaffected).
                        let s594_extra = if std::env::var("OXI_S594").ok().as_deref() == Some("1") { pad_l } else { 0.0 };
                        let wrap_base = if s585_cellmar {
                            (cell_w - pad_l - pad_r - s594_extra).max(0.0)
                        } else if cell_hang_inner || s301_layout_fixed || s412_cellmar_subtract || s531_singlecell_cellmar || s559_cellmar {
                            (cell_w - pad_l - pad_r).max(0.0)
                        } else {
                            cell_w
                        };
                        let wrap_w = (wrap_base - p_indent_left - p_indent_right).max(0.0);
                        let mut first_line_wrap_w = if p_first_line_indent < 0.0 {
                            (wrap_base - (p_indent_left + p_first_line_indent).max(0.0) - p_indent_right).max(0.0)
                        } else {
                            (wrap_w - p_first_line_indent).max(0.0)
                        };
                        if std::env::var("OXI_DUMP_CELLX").is_ok() {
                            let preview: String = para.runs.iter().flat_map(|r| r.text.chars()).take(8).collect();
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
                        let list_marker_info: Option<(String, f32, f32)> = para.style.list_marker.as_ref().map(|marker| {
                            let marker_style = para.runs.first().map(|r| &r.style).cloned().unwrap_or_default();
                            let marker_fs = self.resolve_font_size(&marker_style, &para.style);
                            let marker_metrics = self.metrics_for(&marker_style, &para.style);
                            let marker_width: f32 = marker.chars()
                                .map(|c| self.registry.char_width_pt_with_fallback(c, marker_fs, marker_metrics))
                                .sum();
                            (marker.clone(), marker_fs, marker_width)
                        });
                        // S592 (2026-06-17, opt-IN OXI_S592=1): a SPACE-suffix numbered CELL
                        // paragraph (the 賃金 regulation 第N条 boxes) renders its number INLINE at
                        // the cell-left indent (NOT outdented), body flowing AFTER it. Word does
                        // not outdent a space-suffix number (no tab). Oxi outdented the marker
                        // (marker_x −list_indent) AND did not reserve marker width on line-1 →
                        // body over-fit + overlapped the marker. FIX: reserve marker width on
                        // line-1 (wrap), place the marker at cell-left, start the body after it.
                        // Cross-doc PDF: tokyoshugyo 第４条 x97.6, 3a4f 第２条 x95.7 = cell-left.
                        let s592_cell_space = std::env::var("OXI_S592").ok().as_deref() == Some("1")
                            && matches!(para.style.list_suff.as_deref(), Some("space"))
                            && para.style.list_indent.unwrap_or(0.0) > 0.5;
                        let s592_marker_reserve = if s592_cell_space {
                            list_marker_info.as_ref().map(|(_, fs, w)| w + fs * 0.25).unwrap_or(0.0)
                        } else { 0.0 };
                        if s592_cell_space {
                            first_line_wrap_w = (first_line_wrap_w - s592_marker_reserve).max(0.0);
                        }

                        // Collect runs into lines with greedy wrapping
                        // Tuple: (text, font_size, width, bold, italic, underline, underline_style, strikethrough, font_family, color, highlight, character_spacing, text_scale)
                        let mut lines: Vec<Vec<(String, f32, f32, bool, bool, bool, Option<String>, bool, Option<String>, Option<String>, Option<String>, f32, f32)>> = Vec::new();
                        let mut current_line: Vec<(String, f32, f32, bool, bool, bool, Option<String>, bool, Option<String>, Option<String>, Option<String>, f32, f32)> = Vec::new();
                        let mut line_x: f32 = 0.0;
                        // Session 118 jc=both refactor — gated by OXI_JCBOTH_REFACTOR env var.
                        // When enabled, calls compute_compression from jc_both_compress module
                        // for wrap-decision lookahead. Track per-char CharContext alongside
                        // string buf so we can pass to compute_compression.
                        // S166 (2026-05-21): default ON. Baseline: mean IoU 0.9301 → 0.9303,
                        // 15076df 0.8799 → 0.8850 (+0.005), no other doc moved.
                        // S238 (2026-05-23): removed OXI_LEGACY_NO_JCBOTH_REFACTOR
                        // legacy env-var fallback during hardening pass.
                        let jc_gate_active = true
                            && matches!(para.alignment, Alignment::Justify | Alignment::Distribute)
                            && self.balance_single_byte_double_byte_width
                            && self.compress_punctuation;
                        // S497b FALSIFIED (2026-06-05): extending the compute_compression wrap
                        // lookahead to left-aligned compressPunctuation cells (to model Word's
                        // end-of-line yakumono oikomi at wrap for non-justified paras) was a NO-OP
                        // on the whole tokumei family + ed025c/3a4f (dwrite ΔTOTAL +0.0000). The
                        // tokumei cells do NOT wrap early at yakumono boundaries — S497 (the
                        // line-start-prohibited hang) already covered the one real case (15076df).
                        // The remaining tokumei gap is cumulative sub-pixel precision / weight, not
                        // fixable wrap-precision. Reverted; not gated.
                        let mut current_line_chars: Vec<crate::layout::jc_both_compress::CharContext> = Vec::new();
                        let mut is_first_line = true;
                        // R7.51 (2026-05-13): autoSpaceDE state for CJK↔Latin transitions.
                        // Tracks the last emitted character across runs/buffers so we can
                        // detect transitions and add Word's 2.5pt (10.5pt font) gap. The
                        // body renderer (break_into_lines) applies this; this cell-renderer
                        // path historically did not, causing d77a58 w_i=47 wrap mismatch
                        // (5 lines Oxi vs 6 lines Word).
                        let mut prev_char_emitted: Option<char> = None;
                        // S443: the widened oikomi must fire ONLY on tab-bearing
                        // (list-marker) paragraphs. 3a4f has hanging-indent paras
                        // but ZERO hanging+tab paras; gating oikomi on hanging
                        // alone made it fire on 3a4f's tab-less hanging cells and
                        // cratered it (909 paras +1). Restricting to para_has_tab
                        // excludes 3a4f entirely while keeping d77a's カ/タ list items.
                        let para_has_tab = para.runs.iter().any(|r| r.text.contains('\t'));
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
                        // root). ★HELD default-OFF: page-44 alone exposes sub-bug #2 (the 賃金
                        // chapter is a real ~1-page over-fit, Word p46-64=19pg vs Oxi 18pg; Oxi
                        // fits more chars/line in its cells = the char-budget OIDASHI side, same
                        // wall as #1, NOT separately pinnable — per-element localization blocked by
                        // the doc's repeated-phrase text + table-dense fitz). Ship BOTH when #2/the
                        // wall lands (cf. the S506 OIDASHI scaffold pattern). See
                        // [[tokyoshugyo_wrap_not_cellheight]], [[char_budget_wall]].
                        let s586_para_chars: Vec<char> = para.runs.iter().flat_map(|r| r.text.chars()).collect();
                        let s586_orphan = std::env::var("OXI_S586").ok().as_deref() == Some("1")
                            && self.compat_mode < 15 && self.compress_punctuation;
                        let s586_cap: f32 = std::env::var("OXI_S586_CAP").ok()
                            .and_then(|v| v.parse().ok()).unwrap_or(3.5);
                        let mut s586_run_offset = 0usize;

                        for run in &para.runs {
                            let font_size = self.resolve_font_size(&run.style, &para.style);
                            let bold = self.resolve_bold(&run.style, &para.style);
                            let font_family = self.resolve_font_family_for_text(&run.text, &run.style, &para.style)
                                .map(|s| s.to_string());

                            // Split text character by character for wrapping
                            let cs = if run.style.fit_text.is_some() {
                                run.style.character_spacing.unwrap_or(0.0)
                            } else {
                                snap_character_spacing(run.style.character_spacing.unwrap_or(0.0))
                            };
                            let mut buf = String::new();
                            let mut buf_w: f32 = 0.0;
                            // S118: per-char context for jc_both_compress integration.
                            let mut buf_chars: Vec<crate::layout::jc_both_compress::CharContext> = Vec::new();
                            let s586_run_chars: Vec<char> = run.text.chars().collect();
                            for (s586_ci, ch) in s586_run_chars.iter().copied().enumerate() {
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
                                if ch == '\n' || ch == '\x0B' || ch == '\x0C' {
                                    if !buf.is_empty() {
                                        current_line.push((
                                            buf.clone(), font_size, buf_w, bold,
                                            run.style.italic, run.style.underline,
                                            run.style.underline_style.clone(),
                                            run.style.strikethrough,
                                            font_family.clone(),
                                            run.style.color.clone(),
                                            run.style.highlight.clone(),
                                            cs, run.style.text_scale.unwrap_or(100.0),
                                        ));
                                        buf.clear();
                                        buf_w = 0.0;
                                        current_line_chars.extend(buf_chars.drain(..));
                                    }
                                    lines.push(std::mem::take(&mut current_line));
                                    line_x = 0.0;
                                    current_line_chars.clear();
                                    is_first_line = false;
                                    prev_char_emitted = None;
                                    continue;
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
                                if ch == '\t' && std::env::var("OXI_S443_DISABLE").is_err()
                                    && p_first_line_indent < 0.0 {
                                    let indent_off = if is_first_line {
                                        (p_indent_left + p_first_line_indent).max(0.0)
                                    } else {
                                        p_indent_left
                                    };
                                    let abs_pos = line_x + buf_w + indent_off;
                                    let next_pos = if !para.style.tab_stops.is_empty() {
                                        para.style.tab_stops.iter()
                                            .find(|ts| ts.position > abs_pos + 0.01)
                                            .map(|ts| ts.position)
                                            .unwrap_or_else(|| ((abs_pos / self.default_tab_stop).floor() + 1.0) * self.default_tab_stop)
                                    } else {
                                        ((abs_pos / self.default_tab_stop).floor() + 1.0) * self.default_tab_stop
                                    };
                                    let tab_w = (next_pos - abs_pos).max(0.0);
                                    buf.push(ch);
                                    buf_w += tab_w;
                                    buf_chars.push(crate::layout::jc_both_compress::CharContext {
                                        ch, natural_advance: tab_w, font_size,
                                    });
                                    prev_char_emitted = Some(ch);
                                    continue;
                                }
                                let cm = self.metrics_for_char(ch, &run.style, &para.style);
                                let mut cw = self.registry.char_width_pt_with_fallback(ch, font_size, cm);
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
        let s342_no_snap_gate = std::env::var("OXI_S342_NO_SNAP_GATE").map(|v| v != "0" && v != "false").unwrap_or(true);
                                let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT").map(|v| v != "0" && v != "false").unwrap_or(false);
                                let snap_ok = s342_no_snap_gate || s344_fs_gate || para.style.snap_to_grid;
                                if run.style.fit_text.is_none() && snap_ok {
                                    if let (Some(ratio), Some(pitch)) = (grid_char_cw_ratio, grid_char_pitch) {
                                        if ratio > 0.0 && pitch > 0.0 && cw > 0.0
                                            && crate::font::is_fullwidth(ch)
                                        {
                                            let default_fs = pitch / ratio;
                                            let char_space_pt = pitch - default_fs;
                                            // R7.59 hybrid (see break_into_lines comment).
                                            // S141 H6: skip expansion when font_size < default_fs
                                            let h6_skip = std::env::var("OXI_H6_GRID_GATE").is_ok()
                                                && char_space_pt > 0.0 && font_size < default_fs;
                                            let h7_skip = std::env::var("OXI_H7_GRID_GATE_LE").is_ok()
                                                && char_space_pt > 0.0 && font_size <= default_fs;
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
                                            let h8_skip = char_space_pt > 0.0;
                                            // S344: when snap_to_grid=false and S344 enabled,
                                            // skip compression unless fs < default_fs.
                                            let s344_skip = s344_fs_gate
                                                && !para.style.snap_to_grid
                                                && font_size >= default_fs;
                                            if !(h6_skip || h7_skip || h8_skip || s344_skip) {
                                                cw = if char_space_pt >= 0.0 {
                                                    font_size * pitch / default_fs
                                                } else {
                                                    font_size + char_space_pt
                                                };
                                            }
                                        }
                                    }
                                }
                                if let Some(scale) = run.style.text_scale {
                                    if (scale - 100.0).abs() > 0.01 {
                                        cw *= scale / 100.0;
                                    }
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
                                let balance_extra_cs = if self.balance_single_byte_double_byte_width
                                    && crate::font::is_fullwidth(ch)
                                    && run.style.fit_text.is_none()
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
                                    let prev_cjk_ideo = prev_char_emitted.map_or(false, kinsoku::is_cjk_ideograph_or_kana);
                                    let prev_alpha = prev_char_emitted.map_or(false, |c| c.is_ascii_alphabetic());
                                    let prev_digit = prev_char_emitted.map_or(false, |c| c.is_ascii_digit());
                                    let cur_cjk_ideo = kinsoku::is_cjk_ideograph_or_kana(ch);
                                    let cur_alpha = ch.is_ascii_alphabetic();
                                    let cur_digit = ch.is_ascii_digit();
                                    let de_boundary = (prev_cjk_ideo && cur_alpha) || (prev_alpha && cur_cjk_ideo);
                                    let dn_boundary = (prev_cjk_ideo && cur_digit) || (prev_digit && cur_cjk_ideo);
                                    if (de_boundary && para.style.auto_space_de)
                                        || (dn_boundary && para.style.auto_space_dn) {
                                        s546_autospace_extra(font_size)
                                    } else { 0.0 }
                                };
                                let cw = cw + auto_space_extra;
                                let effective_wrap = if is_first_line { first_line_wrap_w } else { wrap_w };
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
                                let would_overflow_natural = line_x + buf_w + cw > effective_wrap;
                                let run_has_neg_cs = cs <= -0.1;
                                // S586 orphan + small-overflow + 約物→opener oikomi (scaffold).
                                let s586_overflow_fixed = if s586_orphan && would_overflow_natural
                                    && !kinsoku::is_cjk_compressible(ch) {
                                    let gpos = s586_run_offset + s586_ci;
                                    let chars_to_end = s586_para_chars.len().saturating_sub(gpos);
                                    let overflow = (line_x + buf_w + cw) - effective_wrap;
                                    if chars_to_end <= 2 && overflow <= s586_cap {
                                        let seq: Vec<char> = current_line_chars.iter()
                                            .chain(buf_chars.iter()).map(|c| c.ch)
                                            .chain(std::iter::once(ch)).collect();
                                        // collapse a 約物 immediately before an opening bracket
                                        // to ~3.0pt (full 、「 inter-space vanish), natural − 3.0.
                                        let collapse: f32 = seq.iter().enumerate()
                                            .filter(|(i, &c)| matches!(c, '、' | '。' | '，' | '．')
                                                && seq.get(i + 1).map_or(false, |&n| kinsoku::is_yakumono_opening(n)))
                                            .map(|_| (font_size - 3.0).max(0.0))
                                            .sum();
                                        collapse > 0.0
                                            && (line_x + buf_w + cw - collapse) <= effective_wrap
                                    } else { false }
                                } else { false };
                                let would_overflow = if s586_overflow_fixed {
                                    false
                                } else if jc_gate_active && run_has_neg_cs && would_overflow_natural {
                                    let ch_ctx = crate::layout::jc_both_compress::CharContext {
                                        ch,
                                        natural_advance: cw,
                                        font_size,
                                    };
                                    let mut trial: Vec<crate::layout::jc_both_compress::CharContext> =
                                        Vec::with_capacity(current_line_chars.len() + buf_chars.len() + 1);
                                    trial.extend(current_line_chars.iter().cloned());
                                    trial.extend(buf_chars.iter().cloned());
                                    trial.push(ch_ctx);
                                    let r = crate::layout::jc_both_compress::compute_compression(
                                        &trial, effective_wrap, true,
                                    );
                                    !r.fits
                                } else {
                                    would_overflow_natural
                                };
                                if !is_space && would_overflow && !(current_line.is_empty() && buf.is_empty()) {
                                    // Kinsoku: line-start-prohibited chars (）。、etc.) stay on current line
                                    if kinsoku::is_line_start_prohibited(ch) {
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
                                        if std::env::var("OXI_S421_DISABLE").is_err()
                                            && (s412_cellmar_subtract
                                                || (std::env::var("OXI_S443_DISABLE").is_err()
                                                    && p_first_line_indent < 0.0
                                                    && para_has_tab)) {
                                            let ch_ctx = crate::layout::jc_both_compress::CharContext { ch, natural_advance: cw, font_size };
                                            let mut carry: Vec<crate::layout::jc_both_compress::CharContext> = vec![ch_ctx];
                                            loop {
                                                let head = carry[0].ch;
                                                let tail = buf.chars().last();
                                                let need = kinsoku::is_line_start_prohibited(head)
                                                    || tail.map_or(false, kinsoku::is_line_end_prohibited);
                                                let remaining_on_line = buf.chars().count()
                                                    + current_line.iter().map(|f| f.0.chars().count()).sum::<usize>();
                                                if !need || remaining_on_line <= 1 { break; }
                                                if buf.pop().is_some() {
                                                    if let Some(pc) = buf_chars.pop() {
                                                        buf_w -= pc.natural_advance;
                                                        carry.insert(0, pc);
                                                    }
                                                } else if std::env::var("OXI_S443_DISABLE").is_err() {
                                                    // S443: when buf is empty, pop the companion
                                                    // from current_line (already-flushed chars).
                                                    // d77a J's overflowing 「。」 has its companion
                                                    // 「す」 in current_line, not buf — the S421
                                                    // buf-only oikomi could not reach it and
                                                    // force-fit instead. Pop from current_line_chars
                                                    // (per-char ctx) + trim the matching glyph off
                                                    // the last fragment's text/width.
                                                    if let Some(pc) = current_line_chars.pop() {
                                                        if let Some(frag) = current_line.last_mut() {
                                                            if frag.0.pop().is_some() {
                                                                frag.2 -= pc.natural_advance;
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
                                                    current_line.push((buf.clone(), font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family.clone(), run.style.color.clone(), run.style.highlight.clone(), cs, run.style.text_scale.unwrap_or(100.0)));
                                                    buf.clear();
                                                    buf_w = 0.0;
                                                    current_line_chars.extend(buf_chars.drain(..));
                                                }
                                                lines.push(std::mem::take(&mut current_line));
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
                                            current_line.push((buf.clone(), font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family.clone(), run.style.color.clone(), run.style.highlight.clone(), cs, run.style.text_scale.unwrap_or(100.0)));
                                            buf.clear();
                                            buf_w = 0.0;
                                            current_line_chars.extend(buf_chars.drain(..));
                                        }
                                        lines.push(std::mem::take(&mut current_line));
                                        line_x = 0.0;
                                        current_line_chars.clear();
                                        is_first_line = false;
                                        continue;
                                    }
                                    // Flush buffer to current line, then wrap
                                    if !buf.is_empty() {
                                        current_line.push((buf.clone(), font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family.clone(), run.style.color.clone(), run.style.highlight.clone(), cs, run.style.text_scale.unwrap_or(100.0)));
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
                                    && buf.is_empty() && current_line.is_empty()
                                    && !lines.is_empty()
                                    && kinsoku::is_line_start_prohibited(ch)
                                {
                                    if let Some(last) = lines.last_mut() {
                                        last.push((char_to_string(ch), font_size, cw, bold,
                                            run.style.italic, run.style.underline,
                                            run.style.underline_style.clone(), run.style.strikethrough,
                                            font_family.clone(), run.style.color.clone(),
                                            run.style.highlight.clone(), cs,
                                            run.style.text_scale.unwrap_or(100.0)));
                                        prev_char_emitted = Some(ch);
                                        continue;
                                    }
                                }
                                buf.push(ch);
                                buf_w += cw;
                                buf_chars.push(crate::layout::jc_both_compress::CharContext {
                                    ch, natural_advance: cw, font_size,
                                });
                                prev_char_emitted = Some(ch);
                            }
                            if !buf.is_empty() {
                                current_line.push((buf, font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family, run.style.color.clone(), run.style.highlight.clone(), cs, run.style.text_scale.unwrap_or(100.0)));
                                line_x += buf_w;
                                current_line_chars.extend(buf_chars.drain(..));
                            }
                            s586_run_offset += s586_run_chars.len();
                        }
                        if !current_line.is_empty() {
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
                            let pprrpr_fs = para.style.ppr_rpr.as_ref().and_then(|r| r.font_size);
                            let empty_lh = if let Some(empty_fs) = pprrpr_fs {
                                let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                let empty_metrics = self.metrics_for_para_mark(&rpr_ref, &para.style);
                                self.line_height_inner(empty_fs, effective_line_spacing, effective_line_rule, empty_metrics, para.style.snap_to_grid, row_line_pitch, true)
                            } else {
                                let metrics = self.doc_default_metrics();
                                self.line_height_inner(self.default_font_size, effective_line_spacing, effective_line_rule, metrics, para.style.snap_to_grid, row_line_pitch, true)
                            };
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
                            let is_interior_empty = last_content_block_pos
                                .map_or(false, |last| block_pos < last);
                            if is_interior_empty && std::env::var("OXI_S428_DISABLE").is_err() {
                                let mut empty_el = LayoutElement::new(
                                    cell_x + pad_l,
                                    content_h,
                                    0.0,
                                    empty_lh,
                                    LayoutContent::Text {
                                        text: String::new(),
                                        font_size: pprrpr_fs.unwrap_or(self.default_font_size),
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

                        let total_lines = lines.len();
                        for (line_idx, line) in lines.iter().enumerate() {
                            // Clip content that overflows exact row height
                            if is_exact && content_h + pad_t >= row_height {
                                break;
                            }
                            // Line height = max of all runs in line (in_table_cell=true: no default font minimum)
                            let lh: f32 = line.iter().map(|(_text, fs, _, _, _, _, _, _, font_family, _, _, _, _)| {
                                let metrics = match font_family.as_deref() {
                                    Some(ff) => self.registry.get(ff),
                                    None => self.registry.default_metrics(),
                                };
                                self.line_height_inner(*fs, effective_line_spacing, effective_line_rule, metrics, para.style.snap_to_grid, row_line_pitch, true)
                            }).fold(0.0_f32, f32::max);

                            // Paragraph indentation: first line uses indent_left + first_line_indent
                            let line_indent = p_indent_left + if line_idx == 0 { p_first_line_indent } else { 0.0 };

                            // Calculate line total width for alignment
                            let line_total_w: f32 = line.iter().map(|(_, _, tw, _, _, _, _, _, _, _, _, _, _)| tw).sum();
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
                            let effective_wrap = if line_idx == 0 { first_line_wrap_w } else { wrap_w };

                            // Justify: non-last lines for jc=both, all lines for distribute
                            let is_last_line = line_idx == total_lines - 1;
                            let should_justify = (para.alignment == Alignment::Justify && !is_last_line)
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
                            let pad_adjust = if cell_hang_inner || s301_layout_fixed || s412_cellmar_subtract
                                || std::env::var("OXI_S493J_DISABLE").is_ok() {
                                0.0
                            } else {
                                pad_l + pad_r
                            };
                            let align_avail = (effective_wrap - pad_adjust).max(0.0);
                            let align_offset = if should_justify {
                                0.0
                            } else {
                                match para.alignment {
                                    Alignment::Center => (align_avail - line_total_w).max(0.0) / 2.0,
                                    Alignment::Right => (align_avail - line_total_w).max(0.0),
                                    _ => 0.0,
                                }
                            };

                            // Justify: CJK punctuation compression + space/gap distribution
                            let mut frag_width_adj: Vec<f32> = vec![0.0; line.len()];
                            let mut frag_spacing: Vec<f32> = vec![0.0; line.len()];
                            let mut justify_char_spacing: f32 = 0.0;
                            // 2026-04-19: allow single-fragment justify for CJK content.
                            // Word distributes chars within a single CJK run for jc=both
                            // non-last lines (b35 "組織的管" row: 4 chars spread across cell).
                            if should_justify && !line.is_empty() {
                                let mut slack = effective_wrap - line_total_w;

                                // Phase 1: CJK punctuation compression (only when overflowing)
                                if slack < 0.0 {
                                    for (fi, (text, fs, _, _, _, _, _, _, _, _, _, _, _)) in line.iter().enumerate() {
                                        for ch in text.chars() {
                                            if kinsoku::is_cjk_compressible(ch) {
                                                let fm = self.registry.default_metrics();
                                                let char_w = self.registry.char_width_pt_with_fallback(ch, *fs, fm);
                                                let savings = char_w * 0.5;
                                                frag_width_adj[fi] -= savings;
                                                slack += savings;
                                            }
                                        }
                                    }
                                }

                                // Phase 2: Distribute slack at word spaces, then CJK gaps
                                if slack > 0.0 {
                                    let space_count = line.iter()
                                        .enumerate()
                                        .filter(|(i, (text, _, _, _, _, _, _, _, _, _, _, _, _))| *i < line.len() - 1 && text.trim().is_empty())
                                        .count();
                                    if space_count > 0 {
                                        let per_space = slack / space_count as f32;
                                        for (fi, (text, _, _, _, _, _, _, _, _, _, _, _, _)) in line.iter().enumerate() {
                                            if fi < line.len() - 1 && text.trim().is_empty() {
                                                frag_spacing[fi] += per_space;
                                            }
                                        }
                                    } else {
                                        // No word spaces: distribute between ALL CJK character gaps
                                        // Only activate when line is noticeably short (>10% slack);
                                        // COM-confirmed 2026-04-19: for b35 "法令の理解" row with
                                        // 4% slack Word does NOT distribute, showing natural widths.
                                        let has_cjk = line.iter().any(|(text, _, _, _, _, _, _, _, _, _, _, _, _)| text.chars().any(|c| kinsoku::is_cjk(c)));
                                        let slack_ratio = if effective_wrap > 0.0 { slack / effective_wrap } else { 0.0 };
                                        if has_cjk && slack_ratio > 0.10 {
                                            let total_chars: usize = line.iter()
                                                .map(|(text, _, _, _, _, _, _, _, _, _, _, _, _)| text.chars().count())
                                                .sum();
                                            if total_chars > 1 {
                                                let per_char_gap = slack / (total_chars - 1) as f32;
                                                for (fi, (text, _, _, _, _, _, _, _, _, _, _, _, _)) in line.iter().enumerate() {
                                                    let n = text.chars().count();
                                                    if n > 1 {
                                                        frag_width_adj[fi] += per_char_gap * (n - 1) as f32;
                                                    }
                                                    if fi < line.len() - 1 && n > 0 {
                                                        frag_spacing[fi] += per_char_gap;
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
                            let cell_max_fs: f32 = line.iter()
                                .map(|(_, fs, _, _, _, _, _, _, _, _, _, _, _)| *fs)
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
                                line.iter()
                                    .map(|(text, fs, _, bold, italic, _underline, _us, _strikethrough, font_family, _color, _hl, _cs, _ts)| {
                                        let mut rs = RunStyle::default();
                                        rs.font_size = Some(*fs);
                                        rs.bold = *bold;
                                        rs.italic = *italic;
                                        if let Some(ff) = font_family { rs.font_family = Some(ff.clone()); }
                                        let m = self.metrics_for_text(text, &rs, &para.style);
                                        m.word_line_height_table_cell(*fs)
                                    })
                                    .fold(0.0_f32, f32::max)
                            };
                            let cell_text_y_off = match (effective_line_rule, effective_line_spacing) {
                                (Some("exact"), Some(_)) | (Some("atLeast"), Some(_)) => {
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
                                        "center" => ((lh - cell_centering_height).max(0.0) / 2.0 * 2.0 + 0.5).floor() / 2.0,
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
                                    let raw = (lh - cell_centering_height).max(0.0) / 2.0;
                                    // S360 (2026-05-27): cell centering uses same CEIL-half-up
                                    // rounding as body (S328). The +1.0pt table dy cluster
                                    // (S357: 400 paragraphs / 16 docs) may stem from this CEIL
                                    // over-application. Env-gated FLOOR / ROUND variants to test.
                                    let use_floor = std::env::var("OXI_S360_CELL_FLOOR")
                                        .map(|v| v != "0" && v != "false").unwrap_or(false);
                                    let use_round = std::env::var("OXI_S360_CELL_ROUND")
                                        .map(|v| v != "0" && v != "false").unwrap_or(false);
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
                            let cell_glyph_dy = std::env::var("OXI_S453_CELL_GLYPH_DY")
                                .ok()
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(1.5);
                            let cell_text_y_off = cell_text_y_off + cell_glyph_dy;
                            // S592: line-1 body starts AFTER the inline marker (no overlap).
                            let mut rx = if s592_cell_space && line_idx == 0 { s592_marker_reserve } else { 0.0_f32 };
                            // Emit list marker on the first line of the paragraph.
                            if line_idx == 0 {
                                if let Some((ref mk_text, mk_fs, mk_w)) = list_marker_info {
                                    let list_indent = para.style.list_indent.unwrap_or(18.0);
                                    let marker_style = para.runs.first().map(|r| &r.style).cloned().unwrap_or_default();
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
                                            font_family: self.resolve_font_family_for_text(mk_text, &marker_style, &para.style).map(|s| s.to_string()),
                                            bold: self.resolve_bold(&marker_style, &para.style),
                                            italic: marker_style.italic,
                                            underline: marker_style.underline,
                                            underline_style: marker_style.underline_style.clone(),
                                            strikethrough: marker_style.strikethrough,
                                            double_strikethrough: marker_style.double_strikethrough,
                                            color: self.resolve_color(&marker_style, &para.style).map(|s| s.to_string()),
                                            highlight: marker_style.highlight.clone(),
                                            character_spacing: 0.0,
                                            field_type: None,
                                            text_scale: 100.0,
                                            is_vertical: false,
                                        },
                                    );
                                    // Session 72 Phase A: populate text_y_off.
                                    marker_el.text_y_off = cell_text_y_off;
                                    cell_elements.push(marker_el);
                                }
                            }
                            for (frag_idx, (text, fs, tw, bold, italic, underline, underline_style, strikethrough, font_family, color, highlight, cs, ts)) in line.iter().enumerate() {
                                let adj_w = *tw + frag_width_adj[frag_idx];
                                // 2026-04-19: Inject charSpace delta into GDI cs so TextOutW
                                // renders at layout-correct advance (prevents glyph overlap
                                // between fragments when pitch<natural).
                                let grid_cs_adj = if let (Some(ratio), Some(pitch)) = (grid_char_cw_ratio, grid_char_pitch) {
                                    if ratio > 0.0 && pitch > 0.0 {
                                        let default_fs = pitch / ratio;
                                        let char_space_pt = pitch - default_fs;
                                        // Only apply to fullwidth CJK content (halfwidth chars render naturally)
                                        if text.chars().any(|c| crate::font::is_fullwidth(c)) {
                                            char_space_pt
                                        } else { 0.0 }
                                    } else { 0.0 }
                                } else { 0.0 };
                                // Session 75 Phase D: y is LINE BOX TOP; renderer adds cell_text_y_off.
                                let mut cell_el = LayoutElement::new(cell_x + pad_l + line_indent + align_offset + rx, content_h, adj_w, lh, LayoutContent::Text {
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
                                        character_spacing: *cs + justify_char_spacing + grid_cs_adj,
                                        field_type: None,
                                        text_scale: *ts,
                                        is_vertical: false,
                                });
                                // Session 72 Phase A: populate text_y_off (y still includes it).
                                cell_el.text_y_off = cell_text_y_off;
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
                                if line_idx == 0 && frag_idx == 0 && cell_para_counter > 0 {
                                    let para_has_lrpb_on_run0 = para.runs.first()
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
                                    if para_has_lrpb_on_run0 || prev_cell_para_had_mid_lrpb {
                                        cell_el.is_paragraph_start_with_lrpb = true;
                                    }
                                }
                                cell_elements.push(cell_el);
                                rx += adj_w + frag_spacing[frag_idx];
                            }
                            content_h += lh;
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
                        prev_cell_para_had_mid_lrpb = para.runs.iter().enumerate().any(|(i, r)|
                            i > 0 && r.has_last_rendered_page_break);

                        // Render shapes attached to this paragraph (e.g. bracketPair)
                        // pos.y = offset from paragraph start (Word COM confirmed)
                        for shape in &para.shapes {
                            if let Some(ref pos) = shape.position {
                                cell_elements.push(LayoutElement::new(cell_x + pad_l + pos.x, para_content_start_h + pos.y, shape.width, shape.height, LayoutContent::PresetShape {
                                        shape_type: shape.shape_type.clone(),
                                        stroke_color: shape.stroke_color.clone(),
                                        stroke_width: shape.stroke_width.unwrap_or(0.5),
                                        flip_h: shape.flip_h, flip_v: shape.flip_v, arrow_head: shape.arrow_head, arrow_tail: shape.arrow_tail,
                                }));
                            }
                        }
                    }
                }
                Block::Image(img) => {
                    // S533 (2026-06-10): place inline images inside table cells.
                    // The parser forwards cell-paragraph inline images as sibling
                    // Block::Image (S331, default ON as of S533); without this arm
                    // the image occupied no height and emitted no element, so an
                    // image-bearing cell collapsed to its text height (3a4f p34:
                    // the 321.75pt year-calendar EMF cell rendered ~28pt, pulling
                    // ~7 paragraphs up a page = the Phase-1 sole FAIL).
                    cell_elements.push(LayoutElement::new(
                        cell_x + pad_l, content_h, img.width, img.height,
                        LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                        }));
                    content_h += img.height;
                    prev_cell_sa = None;
                }
                _ => {}
                } // match block
                } // for block
                } // if !is_vmerge_continue

                // Track actual cell height for row_height correction.
                // content_h is the sum of per-paragraph layout heights; elements may
                // be positioned with a text_y_offset (vertical centering inside the
                // line box), which makes their bottom extend past content_h. Using
                // max(content_h, elem_bottom) double-counts this offset and inflates
                // row height. Trust content_h as the authoritative sum.
                let is_vmerge_restart = cell.v_merge.as_deref() == Some("restart");
                if !is_vmerge_continue && !is_exact_row && !is_vmerge_restart {
                    let actual = pad_t + content_h + pad_b;
                    if actual > max_actual_cell_h {
                        max_actual_cell_h = actual;
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
                if is_vmerge_restart_for_valign
                    && cell.v_align.is_some()
                {
                    let target_grid = cell_start_grid;
                    let mut span_count = 1usize;
                    for next_ri in (row_idx + 1)..num_rows {
                        let next_row = &table.rows[next_ri];
                        let mut next_grid = next_row.grid_before as usize;
                        let mut continues = false;
                        for next_cell in &next_row.cells {
                            let next_span = next_cell.grid_span.max(1) as usize;
                            if next_grid == target_grid {
                                if matches!(next_cell.v_merge.as_deref(),
                                            Some("continue") | Some("")) {
                                    continues = true;
                                }
                                break;
                            }
                            next_grid += next_span;
                        }
                        if continues { span_count += 1; } else { break; }
                    }
                    if span_count > 1 {
                        let mut span_h: f32 = 0.0;
                        let mut all_have_h = true;
                        for ri in row_idx..(row_idx + span_count) {
                            if let Some(h) = table.rows[ri].height {
                                span_h += h;
                            } else { all_have_h = false; break; }
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
                                        if rule == Some("exact") { continue; }
                                        let nat = self.estimate_table_row_natural_h(
                                            r, &col_widths,
                                            default_pad_l, default_pad_r,
                                            default_pad_t, default_pad_b,
                                            table, table_grid_pitch,
                                            grid_char_pitch, grid_char_cw_ratio,
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
                                            r, &col_widths,
                                            default_pad_l, default_pad_r,
                                            default_pad_t, default_pad_b,
                                            table, table_grid_pitch,
                                            grid_char_pitch, grid_char_cw_ratio,
                                        );
                                        match (r.height, r.height_rule.as_deref()) {
                                            (Some(h), Some("exact")) => h,
                                            (Some(h), _) => nat.max(h),
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
                                        r, &col_widths,
                                        default_pad_l, default_pad_r,
                                        default_pad_t, default_pad_b,
                                        table, table_grid_pitch,
                                        grid_char_pitch, grid_char_cw_ratio,
                                    );
                                    match (r.height, r.height_rule.as_deref()) {
                                        (Some(h), Some("exact")) => h,
                                        (Some(h), _) => nat.max(h),
                                        (None, _) => nat,
                                    }
                                };
                                relaxed_span_h += eff_h;
                            }
                            effective_row_h = effective_row_h.max(relaxed_span_h);
                        }
                    }
                }
                let v_offset = match cell.v_align.as_deref() {
                    Some("center") => ((effective_row_h - pad_t - pad_b - content_h) / 2.0).max(0.0),
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
                let is_vmerge_restart = cell.v_merge.as_deref() == Some("restart");
                for mut elem in cell_elements {
                    elem.y += dy;
                    // Also update y-coords inside TableBorder content (nested tables)
                    if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = elem.content {
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
                        && elem.y > page_bottom + 0.5
                        && elem.cell_paragraph_index.map_or(false, |cpi| cpi > 0)
                    {
                        elem.vmerge_restart_overflow_to_next_page = true;
                    }
                    elements.push(elem);
                }

                // Draw cell borders if table has borders OR cell has its own borders
                let has_cell_borders = cell.borders.as_ref().map_or(false, |b| {
                    b.top.is_some() || b.bottom.is_some() || b.left.is_some() || b.right.is_some()
                });
                if table.style.border || has_cell_borders {
                    let bx = cell_x;
                    let by = cursor.visual_y;

                    // Resolve border color, width and S480 style from cell borders,
                    // falling back to table style.
                    let resolve_border = |side: Option<&BorderDef>| -> (Option<String>, f32, Option<String>) {
                        if let Some(b) = side {
                            // S482: explicit w:val="nil"/"none" cell edge SUPPRESSES
                            // the border (do NOT fall through to the table border).
                            if b.style == "none" {
                                return (None, 0.0, None);
                            }
                            let c = b.color.as_ref().map(|c| {
                                if c.starts_with('#') { c.clone() } else { format!("#{}", c) }
                            });
                            (c, b.width, Some(b.style.clone()))
                        } else if table.style.border {
                            // Table-level borders: use table style color, default to black
                            let c = Some(table.style.border_color.as_ref()
                                .map(|c| if c.starts_with('#') { c.clone() } else { format!("#{}", c) })
                                .unwrap_or_else(|| "#000000".to_string()));
                            (c, table.style.border_width.unwrap_or(0.4), table.style.border_style.clone())
                        } else {
                            (None, 0.4, None)
                        }
                    };

                    let cell_borders = cell.borders.as_ref();
                    let (top_color, top_width, top_style) = resolve_border(cell_borders.and_then(|b| b.top.as_ref()));
                    let (bot_color, bot_width, bot_style) = resolve_border(cell_borders.and_then(|b| b.bottom.as_ref()));
                    let (left_color, left_width, left_style) = resolve_border(cell_borders.and_then(|b| b.left.as_ref()));
                    let (right_color, right_width, right_style) = resolve_border(cell_borders.and_then(|b| b.right.as_ref()));

                    // When cells have their own borders (tcBorders), draw each side per cell.
                    // When using table-level borders, use collapsed model to avoid double-drawing.
                    let use_collapsed = table.style.border && !has_cell_borders;

                    // Top — skip for vMerge continue cells (internal to merged range)
                    if !is_vmerge_continue && top_color.is_some() && (!use_collapsed || row_idx == 0) {
                        elements.push(LayoutElement::new(bx, by, cell_w, 0.0, LayoutContent::TableBorder {
                                x1: bx, y1: by, x2: bx + cell_w, y2: by,
                                color: top_color, width: top_width, style: top_style,
                        }));
                    }
                    // Bottom — skip for vMerge continue cells unless next row is not continue
                    let next_is_continue = if row_idx + 1 < num_rows {
                        table.rows[row_idx + 1].cells.get(cell_idx)
                            .map_or(false, |nc| nc.v_merge.as_deref() == Some("continue") || nc.v_merge.as_deref() == Some(""))
                    } else {
                        false
                    };
                    if bot_color.is_some() && !next_is_continue {
                        elements.push(LayoutElement::new(bx, by + row_height, cell_w, 0.0, LayoutContent::TableBorder {
                                x1: bx, y1: by + row_height, x2: bx + cell_w, y2: by + row_height,
                                color: bot_color, width: bot_width, style: bot_style,
                        }));
                    }
                    // Left
                    if left_color.is_some() && (!use_collapsed || cell_idx == 0) {
                        elements.push(LayoutElement::new(bx, by, 0.0, row_height, LayoutContent::TableBorder {
                                x1: bx, y1: by, x2: bx, y2: by + row_height,
                                color: left_color, width: left_width, style: left_style,
                        }));
                    }
                    // Right
                    if right_color.is_some() {
                        elements.push(LayoutElement::new(bx + cell_w, by, 0.0, row_height, LayoutContent::TableBorder {
                                x1: bx + cell_w, y1: by, x2: bx + cell_w, y2: by + row_height,
                                color: right_color, width: right_width, style: right_style,
                        }));
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
                if !cell.cell_text_boxes.is_empty()
                    && std::env::var("OXI_S487_ENABLE").is_ok()
                {
                    let cell_content_left = cell_x + pad_l;
                    for tb in &cell.cell_text_boxes {
                        let (px, py) = tb.position.as_ref()
                            .map(|p| (p.x, p.y)).unwrap_or((0.0, 0.0));
                        let h_rel = tb.position.as_ref()
                            .and_then(|p| p.h_relative.as_deref());
                        let v_rel = tb.position.as_ref()
                            .and_then(|p| p.v_relative.as_deref());
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
                        let tb_elems = self.layout_text_box_at(
                            tb, page, &[], Some((abs_x, abs_y)));
                        elements.extend(tb_elems);
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
                            if (*y1 - old_bottom).abs() < 0.5 { *y1 = new_bottom; }
                            if (*y2 - old_bottom).abs() < 0.5 { *y2 = new_bottom; }
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
            let row_bottom = cursor.cursor_y + row_height;
            if row_bottom > page_bottom + 0.5 && !row.cant_split {
                // R7.56 (Day 34 part 25, 2026-05-13): respect mid-cell LRPB markers.
                // If any element in this row carries `is_paragraph_start_with_lrpb`,
                // force the split before the FIRST such element above page_bottom
                // (i.e., pull split_y back to that element's y so it goes to next page).
                // e3c545 cpi=81 LRPB at y=765.62: without this pull-back, the element
                // bottom (777.24) fits split_y=785.2 (page_bottom) and stays on current
                // page; with pull-back, split_y becomes 765.62 → element goes overflow.
                let row_elements = elements.split_off(elements_before_row);
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
                let lrpb_split_y = row_elements.iter()
                    .find(|e| e.is_paragraph_start_with_lrpb && e.y > cursor.cursor_y + 0.5)
                    .map(|e| e.y)
                    .unwrap_or(f32::INFINITY);
                let split_y = if lrpb_split_y.is_finite() && lrpb_split_y < page_bottom {
                    lrpb_split_y
                } else {
                    page_bottom
                };
                // Partition elements: those fitting on current page vs overflow
                let mut current_page_elems: Vec<LayoutElement> = Vec::new();
                let mut next_page_elems: Vec<LayoutElement> = Vec::new();

                for elem in row_elements {
                    let _elem_top = elem.y;
                    match &elem.content {
                        LayoutContent::TableBorder { y1, y2, x1, x2, ref color, width, ref style } => {
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
                                    if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = e.content {
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
                                        elem.x, elem.y, elem.width, split_y - vy_top,
                                        LayoutContent::TableBorder {
                                            x1: *x1, y1: vy_top, x2: *x2, y2: split_y,
                                            color: color.clone(), width: *width, style: style.clone(),
                                        },
                                    ));
                                }
                                // Next page portion
                                if vy_bot > split_y {
                                    let shift = split_y - page_top;
                                    let new_y1 = page_top;
                                    let new_y2 = vy_bot - shift;
                                    next_page_elems.push(LayoutElement::new(
                                        elem.x, new_y1, elem.width, new_y2 - new_y1,
                                        LayoutContent::TableBorder {
                                            x1: *x1, y1: new_y1, x2: *x2, y2: new_y2,
                                            color: color.clone(), width: *width, style: style.clone(),
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
                                    elem.x, elem.y, elem.width, clip_h,
                                    LayoutContent::CellShading { color: color.clone() },
                                ));
                            }
                            if shade_bottom > split_y {
                                let shift = split_y - page_top;
                                let new_y = (elem.y - shift).max(page_top);
                                let new_h = shade_bottom - shift - new_y;
                                next_page_elems.push(LayoutElement::new(
                                    elem.x, new_y, elem.width, new_h.max(0.0),
                                    LayoutContent::CellShading { color: color.clone() },
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
                            let elem_bottom = elem.y + elem.height;
                            // S402 (2026-05-28): tested OXI_S402_TIGHTEN to make
                            // ed025 cell-3 page 13 spill 1 × × × to p14. Even
                            // 3pt of tightening CATASTROPHICALLY regressed ed025
                            // from 0.9986 (1 misplaced) to 0.8025 (140 misplaced)
                            // because this split decision fires across ALL rows;
                            // many correctly-fitting cells got pushed down a page.
                            // The bug needs per-cell positioning correction
                            // upstream (Oxi cell cursor 4pt above Word at p13
                            // continuation start), not split-threshold tweaking.
                            if elem_bottom <= split_y + 0.1 {
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
                let min_overflow_text_y = next_page_elems.iter()
                    .filter(|e| matches!(&e.content,
                        LayoutContent::Text { text, .. } if !s570 || !text.trim().is_empty()))
                    .map(|e| e.y)
                    .fold(f32::INFINITY, f32::min);
                if s570 && min_overflow_text_y.is_finite() {
                    next_page_elems.retain(|e| !matches!(&e.content,
                        LayoutContent::Text { text, .. }
                            if text.trim().is_empty() && e.y < min_overflow_text_y - 0.1));
                }
                if min_overflow_text_y.is_finite() {
                    let original_shift = split_y - page_top;
                    let correct_shift = (min_overflow_text_y + original_shift) - page_top;
                    let adjust = correct_shift - original_shift;
                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        let first_after = min_overflow_text_y - adjust;
                        eprintln!("[REANCHOR] page_top={:.2} split_y={:.2} min_overflow_text_y={:.2} orig_shift={:.2} correct_shift={:.2} adjust={:.2} -> first_overflow_line_y={:.2}",
                            page_top, split_y, min_overflow_text_y, original_shift, correct_shift, adjust, first_after);
                    }
                    for e in next_page_elems.iter_mut() {
                        if matches!(e.content, LayoutContent::Text { .. }) {
                            e.y -= adjust;
                        }
                    }
                }

                // Step 3 (2026-04-23): On continuation page, re-close the box
                // to match actual overflow content. The shifted row_bottom lands
                // at a position that doesn't reflect the continuation line's
                // actual bottom (Oxi's row_height undersizes by one overflow line).
                // Word draws (a) top horizontal border at page_top AND (b) bottom
                // border at continuation content bottom. COM-verified on d77a p.7:
                // top=71.04 (=page_top), bottom=89.28 (=line_top 71 + line_height 18).
                {
                    let max_cont_text_bottom = next_page_elems.iter()
                        .filter_map(|e| match &e.content {
                            LayoutContent::Text { .. } => Some(e.y + e.height),
                            _ => None,
                        })
                        .fold(f32::NEG_INFINITY, f32::max);

                    // Find a horizontal bottom border in next_page_elems (the
                    // shifted row_bottom from the split row).
                    let bot_border_idx = next_page_elems.iter().position(|e| {
                        matches!(&e.content,
                            LayoutContent::TableBorder { y1, y2, .. }
                                if (*y1 - *y2).abs() < 0.1)
                    });

                    // Template for new top border (from any vertical border).
                    let vertical_template = next_page_elems.iter().find_map(|e| match &e.content {
                        LayoutContent::TableBorder { y1, y2, x1, x2, color, width, style }
                            if (*y1 - *y2).abs() >= 0.1 => {
                            Some((*x1, *x2, color.clone(), *width, style.clone()))
                        }
                        _ => None,
                    });

                    if let (Some(bi), Some((_, _, color, vw, vstyle))) =
                        (bot_border_idx, vertical_template.clone()) {
                        if max_cont_text_bottom.is_finite() {
                            // Only apply when border is above content bottom
                            // (the broken-box-top-of-page case).
                            let cur_border_y = match &next_page_elems[bi].content {
                                LayoutContent::TableBorder { y1, .. } => *y1,
                                _ => f32::INFINITY,
                            };
                            if cur_border_y < max_cont_text_bottom - 0.5 {
                                // Move bottom horizontal border down to content bottom.
                                if let LayoutContent::TableBorder { y1, y2, .. } =
                                    &mut next_page_elems[bi].content {
                                    *y1 = max_cont_text_bottom;
                                    *y2 = max_cont_text_bottom;
                                }
                                next_page_elems[bi].y = max_cont_text_bottom;

                                // Collect vertical border x1/x2 range, extend y2.
                                let mut min_vx = f32::INFINITY;
                                let mut max_vx = f32::NEG_INFINITY;
                                for e in next_page_elems.iter_mut() {
                                    if let LayoutContent::TableBorder { y1, y2, x1, x2, .. } =
                                        &mut e.content {
                                        if (*y1 - *y2).abs() >= 0.1 {
                                            if *y2 < max_cont_text_bottom {
                                                *y2 = max_cont_text_bottom;
                                                e.height = *y2 - *y1;
                                            }
                                            if *x1 < min_vx { min_vx = *x1; }
                                            if *x2 < min_vx { min_vx = *x2; }
                                            if *x1 > max_vx { max_vx = *x1; }
                                            if *x2 > max_vx { max_vx = *x2; }
                                        }
                                    }
                                }

                                // Add top horizontal border at page_top.
                                if min_vx.is_finite() && max_vx.is_finite() {
                                    next_page_elems.push(LayoutElement::new(
                                        min_vx, page_top, max_vx - min_vx, 0.0,
                                        LayoutContent::TableBorder {
                                            x1: min_vx, y1: page_top,
                                            x2: max_vx, y2: page_top,
                                            color, width: vw, style: vstyle,
                                        },
                                    ));
                                }
                            }
                        }
                    }
                }

                // Step 2 v9 (2026-04-23): close_y = last_line.y + natural_height,
                // gated by `close_y <= split_y - 10pt` (skip if content packs too
                // close to page_bottom). COM-derived from 11 minimal repros
                // C1-C11 (MS Mincho/Gothic/Meiryo × fs 10.5/12/14 × pitch 15/18).
                // See project_z_step2_v9_* memos and ECMA-376 §17.4.33.
                {
                    let last_text_y = current_page_elems.iter()
                        .filter_map(|e| match &e.content {
                            LayoutContent::Text { .. } => Some(e.y),
                            _ => None,
                        })
                        .fold(f32::NEG_INFINITY, f32::max);

                    if last_text_y.is_finite() {
                        let line_eps = 2.0;
                        let mut max_nat: f32 = 0.0;
                        for e in current_page_elems.iter() {
                            if let LayoutContent::Text { font_size, font_family, .. } = &e.content {
                                if (e.y - last_text_y).abs() < line_eps {
                                    let mut rpr = crate::ir::RunStyle::default();
                                    rpr.font_family = font_family.clone();
                                    rpr.font_size = Some(*font_size);
                                    let para_style = crate::ir::ParagraphStyle::default();
                                    let metrics = self.metrics_for_text("", &rpr, &para_style);
                                    let h = metrics.word_ascent_pt(*font_size)
                                          + metrics.word_descent_pt(*font_size);
                                    if h > max_nat { max_nat = h; }
                                }
                            }
                        }
                        let close_y = last_text_y + max_nat;
                        if close_y <= split_y - 10.0 {
                            let template = next_page_elems.iter().find_map(|e| match &e.content {
                                LayoutContent::TableBorder { y1, y2, x1, x2, color, width, style }
                                    if (*y1 - *y2).abs() < 0.1 => {
                                    Some((*x1, *x2, color.clone(), *width, style.clone()))
                                }
                                _ => None,
                            });
                            if let Some((bx1, bx2, color, bw, bstyle)) = template {
                                for e in current_page_elems.iter_mut() {
                                    if let LayoutContent::TableBorder { y1, y2, .. } = &mut e.content {
                                        if (*y1 - *y2).abs() >= 0.1 {
                                            if *y2 > close_y { *y2 = close_y; }
                                            if *y1 > close_y { *y1 = close_y; }
                                        }
                                    }
                                }
                                current_page_elems.push(LayoutElement::new(
                                    bx1, close_y, bx2 - bx1, 0.0,
                                    LayoutContent::TableBorder {
                                        x1: bx1, y1: close_y, x2: bx2, y2: close_y,
                                        color, width: bw, style: bstyle,
                                    },
                                ));
                            }
                        }
                    }
                }

                // Push current page elements
                if std::env::var("OXI_DBG_SPLIT").is_ok() {
                    let cur_txt = current_page_elems.iter().filter(|e| matches!(&e.content, LayoutContent::Text { .. })).count();
                    let nxt_txt = next_page_elems.iter().filter(|e| matches!(&e.content, LayoutContent::Text { .. })).count();
                    let nxt_maxy = next_page_elems.iter().map(|e| match &e.content { LayoutContent::TableBorder { y2, .. } => *y2, _ => e.y + e.height }).fold(0.0_f32, f32::max);
                    eprintln!("[SPLIT] split_y={:.1} ptop={:.1} pbot={:.1} pages_pre={} | cur_elems={} (txt {}) | next_elems={} (txt {}) next_maxy={:.1}",
                        split_y, page_top, page_bottom, pages.len(), current_page_elems.len(), cur_txt, next_page_elems.len(), nxt_txt, nxt_maxy);
                }
                elements.extend(current_page_elems);
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });

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
                loop {
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
                    let max_y = remaining.iter().map(|e| {
                        match &e.content {
                            LayoutContent::TableBorder { y2, .. } => *y2,
                            LayoutContent::PresetShape { .. } => e.y, // ignore height
                            _ => e.y + e.height,
                        }
                    }).fold(0.0_f32, f32::max);

                    if max_y <= page_bottom + 0.5 {
                        // Everything fits on this page
                        break;
                    }

                    // Need another split at page_bottom (or earlier if a mid-cell
                    // LRPB marker exists at y < page_bottom).
                    // R7.56 (Day 34 part 25): pull split back to first LRPB-marked
                    // element above page_top. Same logic as the first-split path.
                    let lrpb_next_split = remaining.iter()
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
                    let next_split = if lrpb_next_split.is_finite()
                        && lrpb_next_split < page_bottom
                        && s565_half_full {
                        lrpb_next_split
                    } else {
                        page_bottom
                    };
                    let mut this_page: Vec<LayoutElement> = Vec::new();
                    let mut overflow: Vec<LayoutElement> = Vec::new();

                    for elem in remaining {
                        let _elem_top = elem.y;
                        match &elem.content {
                            LayoutContent::TableBorder { y1, y2, x1, x2, ref color, width, ref style } => {
                                if (y1 - y2).abs() < 0.1 {
                                    if *y1 <= next_split + 0.5 {
                                        this_page.push(elem);
                                    } else {
                                        let shift = next_split - page_top;
                                        let mut e = elem;
                                        e.y -= shift;
                                        if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = e.content {
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
                                            elem.x, elem.y, elem.width, next_split - vy_top,
                                            LayoutContent::TableBorder {
                                                x1: *x1, y1: vy_top, x2: *x2, y2: next_split,
                                                color: color.clone(), width: *width, style: style.clone(),
                                            },
                                        ));
                                    }
                                    if vy_bot > next_split {
                                        let shift = next_split - page_top;
                                        let new_y1 = page_top;
                                        let new_y2 = vy_bot - shift;
                                        overflow.push(LayoutElement::new(
                                            elem.x, new_y1, elem.width, new_y2 - new_y1,
                                            LayoutContent::TableBorder {
                                                x1: *x1, y1: new_y1, x2: *x2, y2: new_y2,
                                                color: color.clone(), width: *width, style: style.clone(),
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
                                        elem.x, elem.y, elem.width, clip_h,
                                        LayoutContent::CellShading { color: color.clone() },
                                    ));
                                }
                                if shade_bottom > next_split {
                                    let shift = next_split - page_top;
                                    let new_y = (elem.y - shift).max(page_top);
                                    let new_h = shade_bottom - shift - new_y;
                                    overflow.push(LayoutElement::new(
                                        elem.x, new_y, elem.width, new_h.max(0.0),
                                        LayoutContent::CellShading { color: color.clone() },
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
                                if elem_bottom <= next_split + 0.1 {
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

                    if std::env::var("OXI_DBG_SPLIT").is_ok() {
                        let tp_txt = this_page.iter().filter(|e| matches!(&e.content, LayoutContent::Text { .. })).count();
                        let ov_txt = overflow.iter().filter(|e| matches!(&e.content, LayoutContent::Text { .. })).count();
                        eprintln!("[SPLIT-LOOP] next_split={:.1} -> pushed this_page txt={} | overflow txt={}", next_split, tp_txt, ov_txt);
                    }
                    pages.push(LayoutPage {
                        width: page_width,
                        height: page_height,
                        elements: this_page,
                    });
                    remaining = overflow;
                }

                if std::env::var("OXI_DBG_SPLIT").is_ok() {
                    let rem_txt = remaining.iter().filter(|e| matches!(&e.content, LayoutContent::Text { .. })).count();
                    let rem_maxy = remaining.iter().map(|e| match &e.content { LayoutContent::TableBorder { y2, .. } => *y2, _ => e.y + e.height }).fold(0.0_f32, f32::max);
                    eprintln!("[SPLIT-END] pages_now={} | remaining(=p_cont) elems={} txt={} maxy={:.1}", pages.len(), remaining.len(), rem_txt, rem_maxy);
                }
                elements = remaining;
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
                let fix_disabled = std::env::var("OXI_PATTERN_A_DISABLE").is_ok();
                if !fix_disabled && allow_fix {
                    let last_cont_top = elements.iter()
                        .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                        .map(|e| e.y)
                        .fold(f32::NEG_INFINITY, f32::max);
                    let trailing_empty_count = row.cells.iter()
                        .map(|cell| {
                            cell.blocks.iter().rev()
                                .take_while(|b| matches!(b, Block::Paragraph(p)
                                    if p.runs.iter().all(|r| r.text.is_empty())))
                                .count()
                        })
                        .max()
                        .unwrap_or(0);
                    if last_cont_top.is_finite() {
                        if let Some(lh) = table_grid_pitch {
                            cursor.set(last_cont_top + lh * (1.0 + trailing_empty_count as f32));
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
                            let last_cont_h: f32 = elements.iter()
                                .filter(|e| matches!(&e.content, LayoutContent::Text { .. }))
                                .filter(|e| (e.y - last_cont_top).abs() < 0.5)
                                .map(|e| e.height)
                                .fold(0.0_f32, f32::max);
                            if last_cont_h > 0.0 {
                                cursor.set(
                                    last_cont_top
                                        + last_cont_h
                                            * (1.0 + trailing_empty_count as f32),
                                );
                            } else {
                                let overflow_on_next = row_bottom - split_y;
                                let pages_used =
                                    ((overflow_on_next) / content_height).floor() as usize;
                                cursor.set(
                                    page_top
                                        + overflow_on_next
                                        - (pages_used as f32 * content_height),
                                );
                            }
                        }
                    } else {
                        // No text on final page (border-only): geometric fallback.
                        let overflow_on_next = row_bottom - split_y;
                        let pages_used = ((overflow_on_next) / content_height).floor() as usize;
                        cursor.set(page_top + overflow_on_next - (pages_used as f32 * content_height));
                    }
                } else {
                    let overflow_on_next = row_bottom - split_y;
                    let pages_used = ((overflow_on_next) / content_height).floor() as usize;
                    cursor.set(page_top + overflow_on_next - (pages_used as f32 * content_height));
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
                if apply_plus_half {
                    cursor.advance_split(row_height, row_height + 0.5);
                } else {
                    cursor.advance(row_height);
                }
            }
        }

        elements
    }

    /// Resolve column widths for a table.
    /// Priority: grid_columns > cell widths > equal split.
    fn resolve_table_col_widths(&self, table: &Table, content_width: f32) -> Vec<f32> {
        // 1. Use grid_columns if available
        // When nested table overflows parent cell, Word keeps earlier columns
        // at their specified width and shrinks only the last column to fit.
        if !table.grid_columns.is_empty() {
            let total: f32 = table.grid_columns.iter().sum();
            let indent = table.style.indent.unwrap_or(0.0);
            let available = content_width - indent;
            // Floating tables (tblpPr) are not constrained by content_width
            let is_floating = table.style.position.is_some();
            // Day 33 part 69 R7.24 (2026-05-12): fixed-layout tables keep
            // their declared widths even if they exceed content area. Word
            // does NOT shrink fixed-layout cells. Only autofit/auto layout
            // shrinks. a47e6 table 1: tblLayout=fixed, grid 503pt > content
            // 481.9pt — previously Oxi shrunk col 1 (396.9→375.8) costing
            // 21.1pt of wrap width, causing fullwidth+年月日 paragraph to
            // overflow by 0.55pt and Oxi wrap to 2 lines instead of 1
            // (Word renders 1 line). +25pt cumulative row 0 over-pump.
            let is_fixed = table.style.layout.as_deref() == Some("fixed");
            if !is_floating && !is_fixed && total > available && table.grid_columns.len() > 1 {
                let mut cols = table.grid_columns.clone();
                let prefix_sum: f32 = cols[..cols.len() - 1].iter().sum();
                let last = (available - prefix_sum).max(0.0);
                *cols.last_mut().unwrap() = last;
                return cols;
            }
            return table.grid_columns.clone();
        }

        // 2. Use cell widths from first row
        if let Some(first_row) = table.rows.first() {
            let cell_widths: Vec<f32> = first_row.cells.iter()
                .filter_map(|c| c.width)
                .collect();
            if cell_widths.len() == first_row.cells.len() && !cell_widths.is_empty() {
                return cell_widths;
            }
        }

        // 3. Use table style width
        if let Some(tw) = table.style.width {
            let num_cols = table.rows.first().map_or(1, |r| r.cells.len().max(1));
            return vec![tw / num_cols as f32; num_cols];
        }

        // 4. Equal split fallback
        let num_cols = table.rows.first().map_or(1, |r| r.cells.len().max(1));
        vec![content_width / num_cols as f32; num_cols]
    }

    /// Estimate paragraph height for table cell height calculation.
    /// Count lines in a cell paragraph using the same wrap logic as the cell renderer
    /// (mod.rs:4480-4566). Char-by-char, kinsoku line-start prohibited, fullwidth
    /// grid pitch via grid_char_cw_ratio. No yakumono compression, no 2-pass wrap.
    ///
    /// Fix C: estimate_para_height previously used break_into_lines (with yakumono
    /// compression) which fit more chars per line than the cell renderer actually
    /// produces. estimate under-counted → render overflowed → page cascade.
    #[allow(unused_assignments)]
    fn count_cell_lines(
        &self,
        para: &Paragraph,
        wrap_w: f32,
        first_line_wrap_w: f32,
        // S443: raw indents (pt) so the tab-stop estimate matches the render
        // path; the line-count drives row height / element Y placement, so the
        // estimate MUST account for tab advancement too or the cascade won't move.
        indent_l_pt: f32,
        first_indent_pt: f32,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
    ) -> usize {
        if para.runs.is_empty() {
            return 1;
        }
        let mut lines: usize = 0;
        let mut line_x: f32 = 0.0;
        let mut buf_w: f32 = 0.0;
        let mut buf_nonempty = false;
        let mut line_nonempty = false;
        let mut is_first_line = true;
        // R7.51: mirror autoSpaceDE applied in cell renderer (mod.rs:6181 path).
        // Without this the line-count estimate under-counts vs render, re-introducing
        // the Fix C estimate<render mismatch this function exists to prevent.
        let mut prev_char_emitted: Option<char> = None;
        // Session 123 (2026-05-20): mirror the S118 jc=both wrap-decision lookahead
        // from the cell renderer (mod.rs:6722-6727 and 6864-6895). Without this
        // mirror, count_cell_lines predicts more lines than the renderer emits
        // when OXI_JCBOTH_REFACTOR=1 packs an extra char on a line via per-char
        // compression — re-introducing the very estimate<render mismatch this
        // function exists to prevent (Fix C). Identical gate conditions: env var
        // + jc∈{Justify,Distribute} + balanceSBDB + compressPunctuation, and
        // per-run cs ≤ -0.1pt (= ≤ -2tw, S122 threshold).
        // S166: default ON (see comment near mod.rs:6995).
        // S238 (2026-05-23): removed OXI_LEGACY_NO_JCBOTH_REFACTOR
        // legacy env-var fallback during hardening pass.
        let jc_gate_active = matches!(para.alignment, Alignment::Justify | Alignment::Distribute)
            && self.balance_single_byte_double_byte_width
            && self.compress_punctuation;
        let mut current_line_chars: Vec<crate::layout::jc_both_compress::CharContext> = Vec::new();

        for run in &para.runs {
            let font_size = self.resolve_font_size(&run.style, &para.style);
            let cs = if run.style.fit_text.is_some() {
                run.style.character_spacing.unwrap_or(0.0)
            } else {
                snap_character_spacing(run.style.character_spacing.unwrap_or(0.0))
            };
            // S123: per-char trial-line for jc=both wrap-decision lookahead.
            let mut buf_chars: Vec<crate::layout::jc_both_compress::CharContext> = Vec::new();
            for ch in run.text.chars() {
                // Session 109 (2026-05-19): mirror the cell renderer's soft-line-
                // break handling so the line-count estimate matches what the
                // renderer actually produces. Otherwise estimate < render →
                // overflow cascade (the very failure mode this function exists
                // to prevent, per the "Fix C" doc-comment above).
                if ch == '\n' || ch == '\x0B' || ch == '\x0C' {
                    lines += 1;
                    line_x = 0.0;
                    buf_w = 0.0;
                    buf_nonempty = false;
                    line_nonempty = false;
                    is_first_line = false;
                    prev_char_emitted = None;
                    current_line_chars.clear();
                    buf_chars.clear();
                    continue;
                }
                // S443 (2026-05-30, env-gated): TAB-STOP advancement in the cell
                // line-count estimate (mirrors the render-loop fix). Must match the
                // render so estimate==render (Fix C invariant). See render path.
                if ch == '\t' && std::env::var("OXI_S443_DISABLE").is_err()
                    && first_indent_pt < 0.0 {
                    let indent_off = if is_first_line {
                        (indent_l_pt + first_indent_pt).max(0.0)
                    } else {
                        indent_l_pt
                    };
                    let abs_pos = line_x + buf_w + indent_off;
                    let next_pos = if !para.style.tab_stops.is_empty() {
                        para.style.tab_stops.iter()
                            .find(|ts| ts.position > abs_pos + 0.01)
                            .map(|ts| ts.position)
                            .unwrap_or_else(|| ((abs_pos / self.default_tab_stop).floor() + 1.0) * self.default_tab_stop)
                    } else {
                        ((abs_pos / self.default_tab_stop).floor() + 1.0) * self.default_tab_stop
                    };
                    buf_w += (next_pos - abs_pos).max(0.0);
                    buf_nonempty = true;
                    line_nonempty = true;
                    prev_char_emitted = Some(ch);
                    continue;
                }
                let cm = self.metrics_for_char(ch, &run.style, &para.style);
                let mut cw = self.registry.char_width_pt_with_fallback(ch, font_size, cm);
                // S342 (2026-05-27): see effective_char_pitch at line 4073 for
                // OXI_S342_NO_SNAP_GATE gate-drop rationale.
                // S342 SHIP (2026-05-27): default ON. Drops snap_to_grid gate from
        // char-grid (horizontal compression) per OOXML §17.3.1.32. Env-var
        // preserved as opt-OUT.
        let s342_no_snap_gate = std::env::var("OXI_S342_NO_SNAP_GATE").map(|v| v != "0" && v != "false").unwrap_or(true);
                // S344 (2026-05-27): refine S342 — when snap_to_grid=false,
                // only apply char-grid compression for fs < default_fs.
                // S343 per-element diff (b35123) showed S342 over-applied
                // compression to fs == default_fs paragraphs (i=9, i=12 etc.),
                // causing -0.0126 cascade. Real Word behavior:
                //   - fs <  default_fs + snap_to_grid=false → compress (i=89)
                //   - fs == default_fs + snap_to_grid=false → no compress (i=9)
                let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT").map(|v| v != "0" && v != "false").unwrap_or(false);
                let snap_ok = s342_no_snap_gate || s344_fs_gate || para.style.snap_to_grid;
                if run.style.fit_text.is_none() && snap_ok {
                    if let (Some(ratio), Some(pitch)) = (grid_char_cw_ratio, grid_char_pitch) {
                        if ratio > 0.0 && pitch > 0.0 && cw > 0.0
                            && crate::font::is_fullwidth(ch)
                        {
                            let default_fs = pitch / ratio;
                            let char_space_pt = pitch - default_fs;
                            // R7.59 hybrid (see break_into_lines comment).
                            // S141 H6: skip expansion when font_size < default_fs
                            let h6_skip = std::env::var("OXI_H6_GRID_GATE").is_ok()
                                && char_space_pt > 0.0 && font_size < default_fs;
                            let h7_skip = std::env::var("OXI_H7_GRID_GATE_LE").is_ok()
                                && char_space_pt > 0.0 && font_size <= default_fs;
                            // S158 (2026-05-21): added missing H8 site — was the only
                            // cell-related site without H8 gating. V800y bisection
                            // traced a1d6 +14pt residual drift to this code path.
                            // S239 (2026-05-23): removed OXI_LEGACY_GRID_KERN.
                            let h8_skip = char_space_pt > 0.0;
                            // S344: when snap_to_grid=false and S344 enabled,
                            // skip compression unless fs < default_fs.
                            let s344_skip = s344_fs_gate
                                && !para.style.snap_to_grid
                                && font_size >= default_fs;
                            if !(h6_skip || h7_skip || h8_skip || s344_skip) {
                                cw = if char_space_pt >= 0.0 {
                                    font_size * pitch / default_fs
                                } else {
                                    font_size + char_space_pt
                                };
                            }
                        }
                    }
                }
                if let Some(scale) = run.style.text_scale {
                    if (scale - 100.0).abs() > 0.01 {
                        cw *= scale / 100.0;
                    }
                }
                // Session 56 Finding 3: balanceSingleByteDoubleByteWidth doubles
                // cs for CJK fullwidth chars (count_cell_lines for height estimation).
                // Day 37 (2026-05-14): EXCLUDE fitText runs (cs already accounts for
                // balance doubling via resolve_fit_text_runs). Mirrors cell renderer fix.
                let balance_extra_cs = if self.balance_single_byte_double_byte_width
                    && crate::font::is_fullwidth(ch)
                    && run.style.fit_text.is_none()
                {
                    cs
                } else {
                    0.0
                };
                let cw = cw + cs + balance_extra_cs;
                // R7.51: autoSpaceDE mirror — see cell renderer note.
                // Session 95 (2026-05-18): split DE (alpha) vs DN (digit).
                let auto_space_extra = {
                    let prev_cjk_ideo = prev_char_emitted.map_or(false, kinsoku::is_cjk_ideograph_or_kana);
                    let prev_alpha = prev_char_emitted.map_or(false, |c| c.is_ascii_alphabetic());
                    let prev_digit = prev_char_emitted.map_or(false, |c| c.is_ascii_digit());
                    let cur_cjk_ideo = kinsoku::is_cjk_ideograph_or_kana(ch);
                    let cur_alpha = ch.is_ascii_alphabetic();
                    let cur_digit = ch.is_ascii_digit();
                    let de_boundary = (prev_cjk_ideo && cur_alpha) || (prev_alpha && cur_cjk_ideo);
                    let dn_boundary = (prev_cjk_ideo && cur_digit) || (prev_digit && cur_cjk_ideo);
                    if (de_boundary && para.style.auto_space_de)
                        || (dn_boundary && para.style.auto_space_dn) {
                        s546_autospace_extra(font_size)
                    } else { 0.0 }
                };
                let cw = cw + auto_space_extra;
                let effective_wrap = if is_first_line { first_line_wrap_w } else { wrap_w };
                let is_space = ch == ' ' || ch == '\u{3000}';
                // S123: mirror cell renderer's compute_compression lookahead.
                // When gate active + this run has neg cs + natural would overflow,
                // ask whether per-char compression would let trial line fit. If yes,
                // do NOT wrap — matches the actual render decision.
                let would_overflow_natural = line_x + buf_w + cw > effective_wrap;
                let run_has_neg_cs = cs <= -0.1;
                let would_overflow = if jc_gate_active && run_has_neg_cs && would_overflow_natural {
                    let ch_ctx = crate::layout::jc_both_compress::CharContext {
                        ch, natural_advance: cw, font_size,
                    };
                    let mut trial: Vec<crate::layout::jc_both_compress::CharContext> =
                        Vec::with_capacity(current_line_chars.len() + buf_chars.len() + 1);
                    trial.extend(current_line_chars.iter().cloned());
                    trial.extend(buf_chars.iter().cloned());
                    trial.push(ch_ctx);
                    let r = crate::layout::jc_both_compress::compute_compression(
                        &trial, effective_wrap, true,
                    );
                    !r.fits
                } else {
                    would_overflow_natural
                };
                if !is_space && would_overflow && !(!line_nonempty && !buf_nonempty) {
                    if kinsoku::is_line_start_prohibited(ch) {
                        buf_w += cw;
                        buf_nonempty = true;
                        buf_chars.push(crate::layout::jc_both_compress::CharContext {
                            ch, natural_advance: cw, font_size,
                        });
                        line_x += buf_w;
                        line_nonempty = true;
                        current_line_chars.extend(buf_chars.drain(..));
                        lines += 1;
                        line_x = 0.0;
                        buf_w = 0.0;
                        buf_nonempty = false;
                        line_nonempty = false;
                        is_first_line = false;
                        current_line_chars.clear();
                        continue;
                    }
                    if buf_nonempty {
                        line_x += buf_w;
                        line_nonempty = true;
                        buf_w = 0.0;
                        buf_nonempty = false;
                        current_line_chars.extend(buf_chars.drain(..));
                    }
                    lines += 1;
                    line_x = 0.0;
                    line_nonempty = false;
                    is_first_line = false;
                    current_line_chars.clear();
                }
                buf_w += cw;
                buf_nonempty = true;
                buf_chars.push(crate::layout::jc_both_compress::CharContext {
                    ch, natural_advance: cw, font_size,
                });
                prev_char_emitted = Some(ch);
            }
            if buf_nonempty {
                line_x += buf_w;
                line_nonempty = true;
                buf_w = 0.0;
                buf_nonempty = false;
                current_line_chars.extend(buf_chars.drain(..));
            }
        }
        if line_nonempty {
            lines += 1;
        }
        lines.max(1)
    }

    /// Session 131 (2026-05-20): vertical-writing helpers.
    ///
    /// `is_vert_writing_active(cell)` returns true when `OXI_VERT_WRITING=1`
    /// env var is set AND the cell has `text_direction == "tbRlV"`. We only
    /// support tbRlV in this implementation (chars rotated 90° CW). The
    /// tbRl variant (chars upright) is not implemented; none of the 4
    /// affected baseline docs use it.
    ///
    /// `vert_para_height(para)` returns the natural vertical extent of the
    /// paragraph along the writing direction: sum(n_chars × font_size)
    /// per run. COM-measured against 2ea81a tbl=1 row=8: 14 chars × 8pt
    /// = 112pt, Word cell.Height = 113.15pt (≈1pt inter-paragraph gap).
    fn is_vert_writing_active(&self, cell: &TableCell) -> bool {
        // S166 (2026-05-21): vertical writing default ON. Stable across 70+
        // sessions; S236 (2026-05-23) removed OXI_LEGACY_NO_VERT_WRITING
        // legacy env-var fallback during hardening pass.
        cell.text_direction.as_deref() == Some("tbRlV")
    }

    fn vert_para_height(&self, para: &Paragraph) -> f32 {
        let mut h = 0.0_f32;
        for run in &para.runs {
            let fs = self.resolve_font_size(&run.style, &para.style);
            let n_chars = run.text.chars().count() as f32;
            h += n_chars * fs;
        }
        // Empty paragraph still occupies one line-height slot (matches
        // Word's behavior of preserving an empty vertical paragraph mark).
        if h <= 0.0 {
            let default_fs = self.resolve_font_size(&RunStyle::default(), &para.style);
            return default_fs;
        }
        h
    }

    fn estimate_para_height(
        &self,
        para: &Paragraph,
        available_width: f32,
        grid_pitch: Option<f32>,
        table_para_style: Option<&ParagraphStyle>,
        in_cell: bool,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
    ) -> f32 {
        self.estimate_para_height_inner(para, available_width, grid_pitch, table_para_style,
            in_cell, grid_char_pitch, grid_char_cw_ratio, false, false)
    }

    /// S218 (2026-05-23) / S222 (2026-05-23): emit-equivalent row height
    /// for a single table row, matching what the emit pass renders.
    /// Mirrors the inlined logic at the top of the table row-loop
    /// (lines 6473-6604) but uses `estimate_para_height_emit`
    /// (force_grid_snap=true) instead of `estimate_para_height` so the
    /// returned height accounts for the cell line-height grid snap that
    /// the emit pass applies. S218 originally used the non-snap variant,
    /// which underestimated row height by ~14pt for typical 2-line cells
    /// (b5f706 p2 row 1: pre-pass 21.75 vs emit 36.5). S220 attempted
    /// this fix but was blocked by derive_oxi_heights metric noise,
    /// resolved in S221; S222 re-applies. Caller layers trHeight semantic
    /// and zero-fallback.
    fn estimate_table_row_natural_h(
        &self,
        row: &TableRow,
        col_widths: &[f32],
        default_pad_l: f32,
        default_pad_r: f32,
        default_pad_t: f32,
        default_pad_b: f32,
        table: &Table,
        table_grid_pitch: Option<f32>,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
    ) -> f32 {
        let _ = (default_pad_l, default_pad_r); // unused; inner_w uses cell_w directly
        let mut row_height: f32 = 0.0;
        let mut grid_idx = row.grid_before as usize;
        for cell in row.cells.iter() {
            let span = cell.grid_span.max(1) as usize;
            if cell.v_merge.as_deref() == Some("continue")
                || cell.v_merge.as_deref() == Some("")
                || cell.v_merge.as_deref() == Some("restart")
            {
                grid_idx += span;
                continue;
            }
            if grid_idx + span > col_widths.len() {
                grid_idx += span;
                continue;
            }
            let cell_w: f32 = col_widths[grid_idx..grid_idx + span].iter().sum();
            let mut pad_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(default_pad_t);
            let pad_b = cell.margins.as_ref().and_then(|m| m.bottom).unwrap_or(default_pad_b);
            if pad_t == 0.0 && table.style.border {
                pad_t = table.style.border_width.unwrap_or(0.4);
            }
            let inner_w = cell_w.max(0.0);
            let mut cell_content_h = pad_t;
            let vert_writing_active = self.is_vert_writing_active(cell);
            // S427 (2026-05-29): adjacent-paragraph spacing collapse (see pre-pass
            // comment at the cell_content_h loop). Keeps this row-fit estimate
            // consistent with the row-height pre-pass and content placement.
            let s427_collapse = std::env::var("OXI_S427_DISABLE").is_err();
            let mut prev_sa: Option<f32> = None;
            for block in &cell.blocks {
                match block {
                    Block::Paragraph(para) => {
                        let para_h = if vert_writing_active {
                            self.vert_para_height(para)
                        } else {
                            self.estimate_para_height_emit(para, inner_w, table_grid_pitch,
                                table.style.para_style.as_ref(), true,
                                grid_char_pitch, grid_char_cw_ratio)
                        };
                        // S239 (2026-05-23): removed OXI_LEGACY_SB_SUPPRESS and
                        // OXI_SB_NO_SUPPRESS legacy env-var fallbacks (dead code
                        // since LEGACY var default false). S151 default ON.
                        cell_content_h += para_h;
                        let (cur_sb, cur_sa) = self.cell_para_spacing(para, table.style.para_style.as_ref(), table_grid_pitch);
                        if s427_collapse {
                            if let Some(psa) = prev_sa {
                                cell_content_h -= psa.min(cur_sb);
                            }
                        }
                        prev_sa = Some(cur_sa);
                    }
                    Block::Table(nested) => {
                        prev_sa = None;
                        let nested_w = inner_w.max(0.0);
                        for nr in &nested.rows {
                            let mut nr_h = 0.0_f32;
                            for nc in &nr.cells {
                                let mut nc_h = 0.0_f32;
                                for nb in &nc.blocks {
                                    if let Block::Paragraph(np) = nb {
                                        nc_h += self.estimate_para_height_emit(np, nested_w / 2.0,
                                            table_grid_pitch, nested.style.para_style.as_ref(),
                                            true, grid_char_pitch, grid_char_cw_ratio);
                                    }
                                }
                                nr_h = nr_h.max(nc_h);
                            }
                            if let Some(h) = nr.height {
                                match nr.height_rule.as_deref() {
                                    Some("exact") => { nr_h = h; }
                                    Some("atLeast") => { nr_h = nr_h.max(h); }
                                    _ => {}
                                }
                            }
                            cell_content_h += nr_h;
                        }
                    }
                    Block::Image(img) => {
                        // S533: inline image in a cell contributes its height
                        // (mirrors the pre-pass arm at ~8532 and the placement
                        // arm; without it the row-height estimate ignored the
                        // image and the row collapsed to text height).
                        cell_content_h += img.height;
                        prev_sa = None;
                    }
                    _ => {}
                }
            }
            cell_content_h += pad_b;
            row_height = row_height.max(cell_content_h);
            grid_idx += span;
        }
        row_height
    }

    /// Session 79c (2026-05-17): variant that matches the actual cell emit
    /// line-height formula (mod.rs:6711-6717) when adjustLineHeightInTable
    /// triggers cell grid snap. Used to compute `visual_row_h` (max content
    /// height as actually emitted) for vAlign=center offset only — does NOT
    /// affect `row_height` (page break logic) so 3a4f9f cascade is avoided.
    /// b5f706 row 1: 23 cells, mixed 1-/2-paragraph counts with vAlign=center.
    /// Word centers 1-para cells at row middle. Oxi pre-pass uses no-grid lh
    /// (10.625pt for 9pt MS Gothic) so row_height=21.75pt undercounts emit
    /// (36pt grid-snapped). Without this fix, 1-para cells render +1.6pt
    /// instead of Word's +9pt.
    fn estimate_para_height_emit(
        &self,
        para: &Paragraph,
        available_width: f32,
        grid_pitch: Option<f32>,
        table_para_style: Option<&ParagraphStyle>,
        in_cell: bool,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
    ) -> f32 {
        self.estimate_para_height_inner(para, available_width, grid_pitch, table_para_style,
            in_cell, grid_char_pitch, grid_char_cw_ratio, true, false)
    }

    /// S503: same as estimate_para_height_emit but the cell line-height uses the
    /// GDI render height (line_height_inner) so the result equals the ACTUAL emitted
    /// content height. Centering-only (vAlign=center floor), NOT pagination.
    fn estimate_para_height_emit_render(
        &self,
        para: &Paragraph,
        available_width: f32,
        grid_pitch: Option<f32>,
        table_para_style: Option<&ParagraphStyle>,
        in_cell: bool,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
    ) -> f32 {
        self.estimate_para_height_inner(para, available_width, grid_pitch, table_para_style,
            in_cell, grid_char_pitch, grid_char_cw_ratio, true, true)
    }

    fn estimate_para_height_inner(
        &self,
        para: &Paragraph,
        available_width: f32,
        grid_pitch: Option<f32>,
        table_para_style: Option<&ParagraphStyle>,
        in_cell: bool,
        grid_char_pitch: Option<f32>,
        grid_char_cw_ratio: Option<f32>,
        force_grid_snap: bool,
        // S503 (2026-06-08): when true, the single/auto cell line-height uses
        // line_height_inner (the GDI render height, ~13.5) instead of
        // word_line_height_table_cell (~12.625), so the result equals the ACTUAL
        // emitted cell content height. Used ONLY for a centering-only row-height
        // floor (vAlign=center v_offset) — NOT for row_height/pagination. Decouples
        // centering from the pagination estimate (S499 regressed e3c545 by changing
        // the SHARED estimate; this is additive and pagination-safe).
        use_render_lh: bool,
    ) -> f32 {
        let mut height = 0.0;
        // Table cells snap to grid in default Word mode
        let _snap = para.style.snap_to_grid;
        // COM-confirmed (2026-03-31): table cells inherit Normal style's lineSpacing.
        // Only override with table style if it explicitly defines lineSpacing.
        let raw_ls = para.style.line_spacing
            .or_else(|| table_para_style.and_then(|ps| ps.line_spacing));
        let raw_lr = para.style.line_spacing_rule.as_deref()
            .or_else(|| table_para_style.and_then(|ps| ps.line_spacing_rule.as_deref()));
        let style_has_explicit_rule = raw_lr == Some("exact") || raw_lr == Some("atLeast");
        let should_reset = !para.style.has_direct_spacing && !style_has_explicit_rule;
        let tbl_has_ls = table_para_style.and_then(|ps| ps.line_spacing).is_some();
        let (eff_ls, eff_lr): (Option<f32>, Option<&str>) = if tbl_has_ls && !para.style.has_direct_spacing {
            let tbl_ls = table_para_style.and_then(|ps| ps.line_spacing);
            let tbl_lr = table_para_style.and_then(|ps| ps.line_spacing_rule.as_deref());
            (tbl_ls, tbl_lr)
        } else {
            (raw_ls, raw_lr)
        };

        // estimate_para_height is called for table cell content.
        // COM-confirmed: table cells use no-grid line height (grid snap disabled inside cells).
        // Use COM table with grid_pitch=None to get no_grid value.
        if para.runs.is_empty() {
            // Use pPr/rPr font for empty paragraph height
            let empty_fs = para.style.ppr_rpr.as_ref()
                .and_then(|r| r.font_size)
                .unwrap_or(self.resolve_font_size(&RunStyle::default(), &para.style));
            let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
            let metrics = self.metrics_for_para_mark(&rpr_ref, &para.style);
            let is_single_empty = eff_lr.is_none() || eff_lr == Some("auto");
            // Session 79c: force_grid_snap=true switches the formula to match
            // the actual cell emit when adjustLineHeightInTable triggers grid
            // snap (line_height_inner cell_snap_allowed path). Used only for
            // visual_row_h (vAlign=center offset), NOT for row_height.
            let snap_in_cell = force_grid_snap
                && self.adjust_line_height_in_table
                && para.style.snap_to_grid
                && grid_pitch.is_some();
            let h_added = if is_single_empty {
                if snap_in_cell {
                    self.line_height_inner(empty_fs, eff_ls, eff_lr, metrics, true, grid_pitch, true)
                } else if use_render_lh {
                    // S503: actual GDI render line-height (matches emit), centering-only.
                    self.line_height_inner(empty_fs, eff_ls, eff_lr, metrics, para.style.snap_to_grid, grid_pitch, true)
                } else {
                    metrics.word_line_height_table_cell(empty_fs)
                }
            } else {
                self.line_height_inner(empty_fs, eff_ls, eff_lr, metrics, false, None, true)
            };
            if std::env::var("OXI_DUMP_TABLE").is_ok() {
                let pprrpr_fs = para.style.ppr_rpr.as_ref().and_then(|r| r.font_size);
                eprintln!(
                    "[TBL_DUMP]   estimate_empty_para empty_fs={} ppr_rpr_fs={:?} is_single={} h_added={:.3} eff_ls={:?} eff_lr={:?}",
                    empty_fs, pprrpr_fs, is_single_empty, h_added, eff_ls, eff_lr
                );
            }
            height += h_added;
        } else {
            let _para_font_size = self.resolve_font_size(
                para.runs.first().map(|r| &r.style).unwrap_or(&RunStyle::default()),
                &para.style,
            );
            // Use break_into_lines for accurate line count (handles kinsoku, word break, etc.)
            // Twip values are authoritative when present; *Chars × 10.5 is fallback.
            let indent_l = para.style.indent_left
                .or_else(|| para.style.indent_left_chars.map(|c| c / 100.0 * 10.5))
                .unwrap_or(0.0);
            let indent_r = para.style.indent_right
                .or_else(|| para.style.indent_right_chars.map(|c| c / 100.0 * 10.5))
                .unwrap_or(0.0);
            let first_indent_raw = para.style.indent_first_line
                .or_else(|| para.style.indent_first_line_chars.map(|c| c / 100.0 * 10.5))
                .unwrap_or(0.0);
            // COM-confirmed (2026-04-25): numbered list + hanging + suff=tab/default
            // => marker consumes hanging, text starts at `left`. See body path.
            let est_list_consumes_hanging = para.style.list_marker.is_some()
                && first_indent_raw < 0.0
                && matches!(para.style.list_suff.as_deref(), None | Some("tab"));
            let first_indent = if est_list_consumes_hanging { 0.0 } else { first_indent_raw };
            let effective_width = (available_width - indent_l - indent_r).max(1.0);

            // Fix C (2026-04-22): in cell context, count lines using the same
            // char-by-char wrap as the cell renderer (mod.rs:4480+). break_into_lines
            // applies yakumono compression that the cell renderer does not, so it
            // under-counts lines → estimate < render → overflow cascade.
            let line_count = if in_cell {
                // first line wrap width: same indent math as render (mod.rs:4461)
                let first_line_wrap_w = if first_indent < 0.0 {
                    (available_width - (indent_l + first_indent).max(0.0) - indent_r).max(0.0)
                } else {
                    (effective_width - first_indent).max(0.0)
                };
                // S348 (2026-05-27): decouple cell-paragraph HEIGHT from
                // visible WRAP. Per S347 analysis: Word uses natural (uncompressed)
                // line count for cell height allocation, but compressed line count
                // for visible char positions. Oxi current (with S342) uses compressed
                // for both → cell shorter than Word → cascade regression on b35123.
                //
                // OXI_S348_NATURAL_HEIGHT=1: force natural line count for height
                // by passing None grid params. The visible wrap (cell renderer at
                // mod.rs:7800+) continues to use grid params via its own call path,
                // unaffected by this gate.
                //
                // Only meaningful when paired with S342 (otherwise grid params are
                // already None when snap_to_grid=false at line 4073 effective_char_pitch).
                // S348 SHIP (2026-05-27): default ON. Decouples cell-paragraph
                // height from visible wrap. Word uses natural (uncompressed) line
                // count for height even when grid compresses visible wrap. Env-var
                // preserved as opt-OUT.
                let s348_natural_height = std::env::var("OXI_S348_NATURAL_HEIGHT").map(|v| v != "0" && v != "false").unwrap_or(true);
                let (gcp_for_count, gcr_for_count) = if s348_natural_height && !para.style.snap_to_grid {
                    (None, None)
                } else {
                    (grid_char_pitch, grid_char_cw_ratio)
                };
                self.count_cell_lines(para, effective_width, first_line_wrap_w, indent_l, first_indent, gcp_for_count, gcr_for_count)
            } else {
                let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = para.runs.iter().enumerate()
                    .map(|(ri, run)| (run.text.as_str(), &run.style, None, ri, 0))
                    .collect();
                let lines = self.break_into_lines(&fragments, effective_width, first_indent, &para.style, None, None, true, false, matches!(para.alignment, Alignment::Justify | Alignment::Distribute), false, false);
                lines.len().max(1)
            };

            // Session 79c: same snap_in_cell condition as empty-para branch.
            let snap_in_cell = force_grid_snap
                && self.adjust_line_height_in_table
                && para.style.snap_to_grid
                && grid_pitch.is_some();
            let mut max_line_height: f32 = 0.0;
            for run in &para.runs {
                let font_size = self.resolve_font_size(&run.style, &para.style);
                let metrics = self.metrics_for_text(&run.text, &run.style, &para.style);
                let is_single_run = match (eff_lr, eff_ls) {
                    (Some("exact"), _) | (Some("atLeast"), _) => false,
                    (_, Some(f)) if (f - 1.0).abs() > 0.01 => false,
                    _ => true,
                };
                // S499 FALSIFIED (2026-06-06): routing the is_single_run estimate through
                // line_height_inner (GDI render line-height, ~13.5) instead of
                // word_line_height_table_cell (12.625) FIXED the vc_2cell_auto repro
                // (−1.65→+0.10) but corpus dwrite SSIM was net −1.20: e3c545 −0.0974,
                // b35 −0.0158, and it did NOT fix the target d4d126 (+0.0000). The
                // estimate's word_line_height_table_cell is corpus-correct; the render/repro
                // is the outlier. C2 row-height tombstone (S349/361/445) re-confirmed —
                // do NOT change the cell row-height estimate corpus-wide. Reverted.
                let lh = if is_single_run {
                    if snap_in_cell {
                        self.line_height_inner(font_size, eff_ls, eff_lr, metrics, true, grid_pitch, true)
                    } else if use_render_lh {
                        // S503: actual GDI render line-height (matches emit), centering-only.
                        self.line_height_inner(font_size, eff_ls, eff_lr, metrics, para.style.snap_to_grid, grid_pitch, true)
                    } else {
                        metrics.word_line_height_table_cell(font_size)
                    }
                } else {
                    self.line_height_inner(font_size, eff_ls, eff_lr, metrics, false, None, true)
                };
                if lh > max_line_height { max_line_height = lh; }
            }
            if std::env::var("OXI_DUMP_TABLE").is_ok() {
                let first_run_fs = self.resolve_font_size(
                    para.runs.first().map(|r| &r.style).unwrap_or(&RunStyle::default()),
                    &para.style,
                );
                let preview: String = para.runs.iter().flat_map(|r| r.text.chars()).take(8).collect();
                eprintln!(
                    "[TBL_DUMP]   estimate_runs_para first_run_fs={} max_lh={:.3} lines={} h_added={:.3} text={:?}",
                    first_run_fs, max_line_height, line_count, max_line_height * line_count as f32, preview
                );
            }
            height += max_line_height * line_count as f32;

            // Ruby paragraph-tail expansion (§18 spec, V3/V6/V9 measurement).
            // Greenfield: 0/177 baseline docs use w:ruby, so this branch is
            // dormant on the baseline and adds 0.0pt to height. For ruby-
            // bearing paragraphs (V1-V10 fixtures), it adds the calibrated
            // expansion to make pagination match Word's larger paragraph box.
            let para_default_pt = self.resolve_font_size(&RunStyle::default(), &para.style);
            let ruby_exp = ruby::paragraph_ruby_expansion_pt(&para.runs, para_default_pt);
            height += ruby_exp;
        }

        if should_reset {
            // Word resets inherited Normal-style spacing to 0 in table cells
            // but preserves style-defined exact/atLeast spacing
        } else {
            height += if let (Some(bl), Some(pitch)) = (para.style.before_lines, grid_pitch) {
                bl / 100.0 * pitch
            } else {
                para.style.space_before
                    .or_else(|| table_para_style.and_then(|ps| ps.space_before))
                    .unwrap_or(0.0)
            };
            height += para.style.space_after
                .or_else(|| table_para_style.and_then(|ps| ps.space_after))
                .unwrap_or(0.0);
        }
        height
    }

    /// S427 (2026-05-29): effective (space_before, space_after) for a table-cell
    /// paragraph, mirroring the layout-pass logic at mod.rs:7736-7765. Used to
    /// compute the adjacent-paragraph spacing-collapse credit. When Word resets
    /// inherited Normal-style spacing in a cell (`should_reset`), both are 0.
    /// before_lines/after_lines convert via the cell grid pitch (hundredths of a
    /// line); otherwise the twip space_before/space_after applies.
    fn cell_para_spacing(
        &self,
        para: &Paragraph,
        table_para_style: Option<&ParagraphStyle>,
        grid_pitch: Option<f32>,
    ) -> (f32, f32) {
        let raw_lr = para.style.line_spacing_rule.as_deref()
            .or_else(|| table_para_style.and_then(|ps| ps.line_spacing_rule.as_deref()));
        let style_has_explicit_rule = raw_lr == Some("exact") || raw_lr == Some("atLeast");
        let should_reset = !para.style.has_direct_spacing && !style_has_explicit_rule;
        if should_reset {
            return (0.0, 0.0);
        }
        let sb = if let (Some(bl), Some(pitch)) = (para.style.before_lines, grid_pitch) {
            bl / 100.0 * pitch
        } else {
            para.style.space_before
                .or_else(|| table_para_style.and_then(|ps| ps.space_before))
                .unwrap_or(0.0)
        };
        let sa = if let (Some(al), Some(pitch)) = (para.style.after_lines, grid_pitch) {
            al / 100.0 * pitch
        } else {
            para.style.space_after
                .or_else(|| table_para_style.and_then(|ps| ps.space_after))
                .unwrap_or(0.0)
        };
        (sb, sa)
    }
}

#[derive(Default)]
struct Line {
    fragments: Vec<LineFragment>,
    /// What kind of break follows this line (normal line break, page break, or column break)
    break_type: LineBreakType,
    /// 2-pass wrap: sum of fragment.natural_width (un-compressed). Used to determine
    /// if line was "tight" (needed compression to fit) or "loose" (fits naturally).
    natural_total_width: f32,
    /// 2-pass wrap: true if yakumono compression was actually applied to any fragment.
    /// Used by render stage to decide 「 leading-gap direction (-3pt tight / +6pt loose).
    was_compressed: bool,
}

impl Default for LineBreakType {
    fn default() -> Self {
        Self::Normal
    }
}

#[derive(Clone)]
struct LineFragment {
    text: String,
    width: f32,
    /// 2-pass wrap: width BEFORE yakumono compression (natural advance).
    /// If no compression was applied, natural_width == width.
    natural_width: f32,
    style: RunStyle,
    /// For tab fragments: the alignment type of the tab stop they target.
    /// None for non-tab fragments.
    tab_alignment: Option<TabStopAlignment>,
    /// For tab fragments: the absolute position (from left margin) of the tab stop.
    tab_position: Option<f32>,
    /// Field type for dynamic content (PAGE, NUMPAGES)
    field_type: Option<FieldType>,
    /// Source run index within the paragraph (for editing support)
    run_index: usize,
    /// Source character byte offset within the run (for editing support)
    char_offset: usize,
}

/// Marker for page/column break after a line
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LineBreakType {
    Normal,
    SoftBreak,   // \n (Shift+Enter) — explicit line break within paragraph
    PageBreak,   // \x0C
    ColumnBreak, // \x0B
}

#[cfg(test)]
mod tests {
    #[allow(unused_imports)]
    use super::*;

    /// S99/S100 LayoutCursor invariants (Phase A1).
    /// Future Phase B will allow cursor_y and visual_y to diverge.
    #[test]
    fn layout_cursor_new_initializes_both_tracks() {
        let c = LayoutCursor::new(72.0);
        assert_eq!(c.cursor_y, 72.0);
        assert_eq!(c.visual_y, 72.0);
    }

    #[test]
    fn layout_cursor_advance_mirrors_both_tracks() {
        let mut c = LayoutCursor::new(10.0);
        c.advance(5.5);
        assert_eq!(c.cursor_y, 15.5);
        assert_eq!(c.visual_y, 15.5);
        c.advance(-2.0);
        assert_eq!(c.cursor_y, 13.5);
        assert_eq!(c.visual_y, 13.5);
    }

    #[test]
    fn layout_cursor_advance_split_diverges_tracks() {
        // Phase B divergence: cursor_y and visual_y advance by different amounts.
        let mut c = LayoutCursor::new(10.0);
        c.advance_split(15.0, 18.0);
        assert_eq!(c.cursor_y, 25.0);
        assert_eq!(c.visual_y, 28.0);
        // Visual delta accumulates
        c.advance_split(5.0, 7.0);
        assert_eq!(c.cursor_y, 30.0);
        assert_eq!(c.visual_y, 35.0);
    }

    #[test]
    fn layout_cursor_set_resets_both_tracks() {
        let mut c = LayoutCursor::new(50.0);
        c.advance(25.0);
        // Page boundary reset
        c.set(80.0);
        assert_eq!(c.cursor_y, 80.0);
        assert_eq!(c.visual_y, 80.0);
    }

    /// R-05d: balloons that don't overlap stay at their natural anchor.
    #[test]
    fn stack_balloon_ys_no_overlap_keeps_anchors() {
        let mut positions = vec![(100.0, 30.0), (200.0, 40.0), (300.0, 20.0)];
        stack_balloon_ys(&mut positions, 6.0);
        assert_eq!(positions, vec![(100.0, 30.0), (200.0, 40.0), (300.0, 20.0)]);
    }

    /// R-05d: when a balloon would overlap the previous one's bottom + gap,
    /// it gets pushed down. First balloon never moves.
    #[test]
    fn stack_balloon_ys_pushes_overlapping_balloons_down() {
        // (100, 30) ends at 130. (140, 40) starts at 140 which is above 130+6=136 floor.
        // 140 >= 136 so no push needed.
        let mut positions = vec![(100.0, 30.0), (140.0, 40.0)];
        stack_balloon_ys(&mut positions, 6.0);
        assert_eq!(positions, vec![(100.0, 30.0), (140.0, 40.0)]);

        // (100, 30) ends at 130. (135, 40) starts at 135 < 136 floor → push to 136.
        let mut positions = vec![(100.0, 30.0), (135.0, 40.0)];
        stack_balloon_ys(&mut positions, 6.0);
        assert_eq!(positions, vec![(100.0, 30.0), (136.0, 40.0)]);

        // Cascade: 3 balloons all anchored at the same Y get fanned out.
        let mut positions = vec![(100.0, 20.0), (100.0, 30.0), (100.0, 25.0)];
        stack_balloon_ys(&mut positions, 6.0);
        // 1st stays at 100 (height 20 → bottom 120).
        // 2nd: floor = 120 + 6 = 126; 100 < 126 → push to 126 (height 30 → bottom 156).
        // 3rd: floor = 156 + 6 = 162; 100 < 162 → push to 162.
        assert_eq!(positions, vec![(100.0, 20.0), (126.0, 30.0), (162.0, 25.0)]);
    }

    /// R-05d: degenerate inputs (empty / single) are no-ops.
    #[test]
    fn stack_balloon_ys_handles_degenerate_inputs() {
        let mut empty: Vec<(f32, f32)> = Vec::new();
        stack_balloon_ys(&mut empty, 6.0);
        assert!(empty.is_empty());

        let mut single = vec![(50.0, 100.0)];
        stack_balloon_ys(&mut single, 6.0);
        assert_eq!(single, vec![(50.0, 100.0)]);
    }

    #[test]
    #[ignore]
    fn bench_layout_multi() {
        // Benchmark multiple documents to find the pattern
        let docs_dir = "../../tools/golden-test/documents/docx";
        let mut results = Vec::new();
        if let Ok(entries) = std::fs::read_dir(docs_dir) {
            for entry in entries.flatten() {
                let path = entry.path();
                if path.extension().map_or(false, |e| e == "docx") {
                    if let Ok(data) = std::fs::read(&path) {
                        if let Ok(doc) = crate::parse_docx(&data) {
                            let engine = LayoutEngine::for_document(&doc);
                            let _ = engine.layout(&doc); // warmup
                            let start = std::time::Instant::now();
                            let r = engine.layout(&doc);
                            let ms = start.elapsed().as_micros() as f64 / 1000.0;
                            let elems: usize = r.pages.iter().map(|p| p.elements.len()).sum();
                            results.push((path.file_name().unwrap().to_string_lossy().to_string(), ms, r.pages.len(), elems));
                        }
                    }
                }
            }
        }
        results.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
        println!("\nTop 10 slowest:");
        for (name, ms, pages, elems) in results.iter().take(10) {
            println!("  {:.1}ms  {}p {}el  {}", ms, pages, elems, name);
        }
        let total: f64 = results.iter().map(|r| r.1).sum();
        println!("Total: {:.0}ms for {} docs", total, results.len());
    }

    #[test]
    #[ignore]
    fn bench_layout_1ec_detail() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/1ec1091177b1_006.docx")
            .expect("read docx");
        let doc = crate::parse_docx(&data).expect("parse");
        let engine = LayoutEngine::for_document(&doc);
        let _ = engine.layout(&doc); // warmup

        // Measure engine creation + layout
        let n = 20;
        let start = std::time::Instant::now();
        let mut result = None;
        for _ in 0..n {
            let eng = LayoutEngine::for_document(&doc);
            result = Some(eng.layout(&doc));
        }
        let full_ms = start.elapsed().as_micros() as f64 / 1000.0 / n as f64;

        // Measure layout only (engine reused)
        let start = std::time::Instant::now();
        for _ in 0..n {
            result = Some(engine.layout(&doc));
        }
        let layout_ms = start.elapsed().as_micros() as f64 / 1000.0 / n as f64;

        // Measure engine creation only
        let start = std::time::Instant::now();
        for _ in 0..n {
            let _eng = LayoutEngine::for_document(&doc);
        }
        let engine_ms = start.elapsed().as_micros() as f64 / 1000.0 / n as f64;

        let r = result.unwrap();
        let total_elems: usize = r.pages.iter().map(|p| p.elements.len()).sum();
        println!("Engine: {:.1}ms, Layout: {:.1}ms, Full: {:.1}ms, Pages: {}, Elements: {}", engine_ms, layout_ms, full_ms, r.pages.len(), total_elems);
    }

    #[test]
    #[ignore]
    fn bench_layout_1ec() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/1ec1091177b1_006.docx")
            .expect("read docx");
        // Profile: parse vs layout
        let parse_start = std::time::Instant::now();
        let doc = crate::parse_docx(&data).expect("parse");
        let parse_ms = parse_start.elapsed().as_millis();
        println!("Parse: {}ms", parse_ms);

        // Count blocks, runs, characters
        let mut total_chars = 0usize;
        let mut total_runs = 0usize;
        let mut total_blocks = 0usize;
        let mut total_table_cells = 0usize;
        for page in &doc.pages {
            total_blocks += page.blocks.len();
            for block in &page.blocks {
                match block {
                    crate::ir::Block::Paragraph(p) => {
                        total_runs += p.runs.len();
                        for r in &p.runs { total_chars += r.text.len(); }
                    }
                    crate::ir::Block::Table(t) => {
                        for row in &t.rows {
                            for cell in &row.cells {
                                total_table_cells += 1;
                                for b in &cell.blocks {
                                    if let crate::ir::Block::Paragraph(p) = b {
                                        total_runs += p.runs.len();
                                        for r in &p.runs { total_chars += r.text.len(); }
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            // TextBox content
            for tb in &page.text_boxes {
                for b in &tb.blocks {
                    if let crate::ir::Block::Paragraph(p) = b {
                        total_runs += p.runs.len();
                        for r in &p.runs { total_chars += r.text.len(); }
                    }
                }
            }
        }
        println!("Doc: {} blocks, {} runs, {} chars, {} table_cells", total_blocks, total_runs, total_chars, total_table_cells);

        // Warmup
        let engine = LayoutEngine::for_document(&doc);
        let _ = engine.layout(&doc);
        // Measure
        let n = 10;
        let start = std::time::Instant::now();
        for _ in 0..n {
            let engine = LayoutEngine::for_document(&doc);
            let _ = engine.layout(&doc);
        }
        let elapsed = start.elapsed();
        println!("Layout: {:.1}ms avg ({} runs, {:.0}ms total)",
            elapsed.as_millis() as f64 / n as f64, n, elapsed.as_millis());
    }

    #[test]
    #[ignore] // debug only
    fn debug_1ec_y_positions() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/1ec1091177b1_006.docx")
            .expect("read docx");
        let doc = crate::parse_docx(&data).expect("parse");

        // Print table structure in detail
        for page in &doc.pages {
            for (bi, block) in page.blocks.iter().enumerate() {
                if let crate::ir::Block::Table(t) = block {
                    println!("B{}: Table {}rows", bi, t.rows.len());
                    for (ri, row) in t.rows.iter().enumerate() {
                        let hr = row.height_rule.as_deref().unwrap_or("auto");
                        let hv = row.height.unwrap_or(0.0);
                        println!("  Row{}: h_spec={:.1} rule={} cells={}", ri, hv, hr, row.cells.len());
                        for (ci, cell) in row.cells.iter().enumerate() {
                            let pad = &cell.margins;
                            let pad_t = pad.as_ref().and_then(|m| m.top).unwrap_or(-1.0);
                            let pad_b = pad.as_ref().and_then(|m| m.bottom).unwrap_or(-1.0);
                            println!("    Cell{}: paras={} pad_t={:.1} pad_b={:.1} vmerge={:?}",
                                ci, cell.blocks.len(), pad_t, pad_b, cell.v_merge);
                            for (pi, blk) in cell.blocks.iter().enumerate() {
                                if let crate::ir::Block::Paragraph(p) = blk {
                                    let text: String = p.runs.iter().flat_map(|r| r.text.chars()).take(30).collect();
                                    let snap = p.style.snap_to_grid;
                                    let ls = p.style.line_spacing;
                                    let lr = p.style.line_spacing_rule.as_deref().unwrap_or("?");
                                    let sa = p.style.space_after.unwrap_or(0.0);
                                    let sb = p.style.space_before.unwrap_or(0.0);
                                    let font = p.runs.first().map(|r| r.style.font_family.as_deref().unwrap_or("?")).unwrap_or("?");
                                    let fsz = p.runs.first().map(|r| r.style.font_size.unwrap_or(0.0)).unwrap_or(0.0);
                                    println!("      P{}: snap={} ls={:?} lr={} sa={:.1} sb={:.1} font={}@{:.1} \"{}\"",
                                        pi, snap, ls, lr, sa, sb, font, fsz, text);
                                }
                            }
                        }
                    }
                }
            }
        }

        let engine = LayoutEngine::for_document(&doc);
        let result = engine.layout(&doc);
        println!("\nPages: {}", result.pages.len());
        // Show all elements with their Y positions grouped by row
        for (pi, lpage) in result.pages.iter().enumerate() {
            println!("--- Page {} ---", pi);
            let mut prev_y: f32 = -1.0;
            for el in &lpage.elements {
                match &el.content {
                    LayoutContent::Text { ref text, font_size, .. } => {
                        if (el.y - prev_y).abs() > 0.1 {
                            let snippet: String = text.chars().take(25).collect();
                            println!("  TEXT y={:.1} h={:.1} fs={:.1} \"{}\"", el.y, el.height, font_size, snippet);
                            prev_y = el.y;
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    #[test]
    #[ignore]
    fn debug_fded_positions() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/gen2_052_Privacy_Policy.docx")
            .expect("read docx");
        let doc = crate::parse_docx(&data).expect("parse");

        let engine = LayoutEngine::for_document(&doc);
        let result = engine.layout(&doc);
        println!("Pages: {}", result.pages.len());

        // Show elements in table area, grouped by Y
        let mut prev_y: f32 = -100.0;
        for el in &result.pages[0].elements {
            match &el.content {
                LayoutContent::Text { ref text, .. } => {
                    if (el.y - prev_y).abs() > 0.5 {
                        let s: String = text.chars().take(20).collect();
                        println!("  y={:.1} x={:.1} w={:.1} h={:.1} \"{}\"", el.y, el.x, el.width, el.height, s);
                        prev_y = el.y;
                    }
                }
                _ => {}
            }
        }
    }

    #[test]
    #[ignore]
    fn debug_db9c_line_breaks() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")
            .expect("read docx");
        let doc = crate::parse_docx(&data).expect("parse");

        let engine = LayoutEngine::for_document(&doc);
        let result = engine.layout(&doc);
        println!("\nPages: {}", result.pages.len());

        for (pi, lpage) in result.pages.iter().enumerate() {
            println!("--- Page {} ---", pi + 1);
            let mut prev_y: f32 = -100.0;
            let mut line_text = String::new();
            let mut line_x = 0.0_f32;
            for el in &lpage.elements {
                match &el.content {
                    LayoutContent::Text { ref text, .. } => {
                        if (el.y - prev_y).abs() > 0.5 {
                            if !line_text.is_empty() {
                                let chars: usize = line_text.chars().count();
                                let snippet: String = line_text.chars().take(120).collect();
                                println!("  y={:.1} x={:.1} [{}c] \"{}\"", prev_y, line_x, chars, snippet);
                            }
                            line_text = text.clone();
                            line_x = el.x;
                            prev_y = el.y;
                        } else {
                            line_text.push_str(text);
                        }
                    }
                    _ => {}
                }
            }
            if !line_text.is_empty() {
                let chars: usize = line_text.chars().count();
                let snippet: String = line_text.chars().take(80).collect();
                println!("  y={:.1} x={:.1} [{}c] \"{}\"", prev_y, line_x, chars, snippet);
            }
        }
    }
}
