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
}

impl LayoutCursor {
    pub fn new(y: f32) -> Self {
        Self { cursor_y: y, visual_y: y }
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
            compress_punctuation: false,
            do_not_expand_shift_return: false,
            balance_single_byte_double_byte_width: false,
            balloon_column_width: 0.0,
            show_comments: true,
            show_revisions: ShowRevisions::All,
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
            for elem in &mut page.elements {
                if let LayoutContent::Text { text, field_type: Some(ft), font_size, .. } = &mut elem.content {
                    let new_text = match ft {
                        FieldType::Page => format!("{}", page_idx + 1),
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
                let lines = self.break_into_lines(&fragments, 1e6, 0.0, para_style, None, None);
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
        let num_columns = page.columns.as_ref().map(|c| c.num.max(1) as usize).unwrap_or(1);
        let mut col_x_positions: Vec<f32> = Vec::with_capacity(num_columns);
        let mut col_widths: Vec<f32> = Vec::with_capacity(num_columns);

        if num_columns > 1 {
            if let Some(ref cols) = page.columns {
                if !cols.columns.is_empty() {
                    // Unequal width columns: use explicit definitions
                    let mut x = page.margin.left;
                    for col_def in &cols.columns {
                        col_x_positions.push(x);
                        col_widths.push(col_def.width);
                        x += col_def.width + col_def.space.unwrap_or(0.0);
                    }
                } else {
                    // Equal width columns
                    let spacing = cols.space.unwrap_or(36.0); // default 36pt
                    let col_w = (total_content_width - spacing * (num_columns - 1) as f32) / num_columns as f32;
                    let mut x = page.margin.left;
                    for _ in 0..num_columns {
                        col_x_positions.push(x);
                        col_widths.push(col_w);
                        x += col_w + spacing;
                    }
                }
            }
        }
        if col_x_positions.is_empty() {
            col_x_positions.push(page.margin.left);
            col_widths.push(total_content_width);
        }

        let mut current_column: usize = 0;
        let mut start_x = col_x_positions[0];
        let mut content_width = col_widths[0];

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
            block_y_positions.push(cursor.cursor_y);
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
                                    stroke_width: shape.stroke_width.unwrap_or(0.75),
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
                    if cursor.cursor_y + img.height > start_y + content_height {
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
                    cursor.advance(img.height);
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
                    let x = page.margin.left;
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
        for text_box in &page.text_boxes {
            let target_page = block_page_indices
                .get(text_box.anchor_block_index)
                .copied()
                .unwrap_or(0);
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

                        // Draw the footnote separator line: ~1/3 of content width
                        // (Word default), 1pt thick, black, anchored at the left margin.
                        // Word renders this as a 1px line at the top of the footnote area.
                        let sep_w = hdr_width * 0.33;
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
                                stroke_width: shape.stroke_width.unwrap_or(0.75),
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
        let abs_y = if abs_y + img.height > page.size.height {
            (page.size.height - img.height).max(0.0)
        } else {
            abs_y
        };

        (abs_x, abs_y)
    }

    /// Layout a single text box: background, borders, and inner content.
    fn layout_text_box(&self, text_box: &TextBox, page: &Page, block_y_positions: &[f32]) -> Vec<LayoutElement> {
        let mut elements = Vec::new();

        // 1. Calculate absolute position
        let (abs_x, abs_y) = self.resolve_textbox_position(text_box, page, block_y_positions);

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
                                    stroke_width: shape.stroke_width.unwrap_or(0.75),
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
            let marker_font_family = self
                .resolve_font_family_for_text(&marker_text, marker_style, &para.style)
                .map(|s| s.to_string());
            let marker_bold = self.resolve_bold(marker_style, &para.style);
            let marker_color = self.resolve_color(marker_style, &para.style).map(|s| s.to_string());
            elements.push(LayoutElement::new(marker_x, cursor.visual_y + marker_y_offset, marker_width, line_height, LayoutContent::Text {
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
        let lines = self.break_into_lines(&fragments, wrap_width, effective_first_indent, &para.style, effective_char_pitch, effective_cw_ratio);

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
            self.line_height_for_line(line, &para.style, para_font_size, para.style.snap_to_grid, grid_pitch)
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
            base * para.style.line_spacing.unwrap_or(1.0) * 20.0
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
            let break_threshold = natural_lh.min(effective_lh);
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
                has_lrpb_here && page.total_lrpb_count <= s394_max
            } else {
                false
            };
            let needs_page_break = natural_needs_page_break || s391_lrpb_break;
            if std::env::var("OXI_DUMP_BREAK").is_ok() && line_idx == 0 {
                let pi_str = body_para_index.map(|v| v.to_string()).unwrap_or_else(|| "?".into());
                let txt: String = para.runs.iter().flat_map(|r| r.text.chars()).take(15).collect();
                eprintln!(
                    "[BR_DUMP] pi={} line0 cursor_y={:.3} eff_lh={:.3} line_h={:.3} sum={:.3} pg_top={:.3} pg_bot={:.3} brk={} text={:?}",
                    pi_str, cursor.cursor_y, effective_lh, line_height,
                    cursor.cursor_y + effective_lh, page_top, page_top + content_height,
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

                // Phase 1: CJK punctuation compression (full-width -> half-width)
                // Only compress when the line overflows (slack < 0).
                // Matches Word output: TextBox content does NOT use punctuation compression.
                // 2026-04-20 fix: Skip chars whose fragment.width is ALREADY smaller than
                // natural (indicates break_into_lines already compressed them — applying
                // Phase 1 again would DOUBLE-compress, crushing 「」 to w=0pt).
                if slack < 0.0 && !in_textbox {
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
                if slack > 0.5 && !in_textbox {
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
                let mut el = LayoutElement::new(x, cursor.visual_y, adjusted_width, line_height, LayoutContent::Text {
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
                    let cn = (new_pos / 10.0).ceil() as i32 * 10;
                    let cc = (old_pos / 10.0).ceil() as i32 * 10;
                    (cn, cc)
                } else if is_multiple_spacing {
                    // COM-confirmed (2026-04-14, mixed font repro): Multiple spacing
                    // uses cumulative raw position model with ROUND. Each paragraph
                    // adds its raw_tw to a shared running total.
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
                cursor.advance((cn - cc) as f32 / 20.0);
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
                // Advance cursor to border midpoint (COM-confirmed: space + bw/2).
                // gen2_036 Title 26pt: lineH=34 + space(4) + bw/2(0.5) = 38.5 = Word.
                cursor.set(border_y + bw / 2.0);
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
                        current_width = 0.0; current_width_tw = 0; compress_used = false;
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
                    word_width = 0.0;
                    word_natural_width = 0.0;
                }
            };
        }

        let n_fragments = fragments.len();
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
            let yakumono_pair_enabled = self.compress_punctuation || cjk_font_has_hwid;
            let yakumono_enabled = self.compress_punctuation;
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
                    // Expand pair rule: when yakumono_compressed[neighbor] = true AND
                    // this char is also yakumono, both compress.
                    let is_yakumono_any = matches!(ch,
                        '（' | '）' | '「' | '」' | '『' | '』' | '〔' | '〕' |
                        '【' | '】' | '《' | '》' | '〈' | '〉' | '｛' | '｝' |
                        '［' | '］' | '、' | '。' | '，' | '．'
                    );
                    if is_yakumono_any {
                        let prev_compressed = char_index > 0
                            && yakumono_compressed[char_index - 1];
                        let next_compressed = char_index + 1 < chars_vec.len()
                            && yakumono_compressed[char_index + 1];
                        if (prev_compressed || next_compressed) && !is_opening_bracket {
                            // Adjacent to pair-compressed yakumono: also compress
                            // (except opening brackets — see explanation above)
                            char_width *= 0.5;
                        } else if matches!(ch, '、' | '。' | '，' | '．') {
                            // Standalone 、 。 between non-triggers: spec §4.7b round 5
                            // floor = fontSize × 2/3. Trying 0.667 instead of 0.583.
                            let prev_non_tr = char_index == 0
                                || !kinsoku::is_yakumono_trigger(chars_vec[char_index - 1]);
                            let next_non_tr = char_index + 1 >= chars_vec.len()
                                || !kinsoku::is_yakumono_trigger(chars_vec[char_index + 1]);
                            if prev_non_tr && next_non_tr {
                                char_width *= 0.6667;
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
                        let h8_trigger = char_space_pt > 0.0;
                        let h7_trigger = h7_gate_enabled && char_space_pt > 0.0 && font_size <= default_fs;
                        let h6_trigger = h6_gate_enabled && char_space_pt > 0.0 && font_size < default_fs;
                        // S344 (2026-05-27): when S344 fed grid values through despite
                        // snap_to_grid=false, gate compression to fs < default_fs only.
                        // (Effective only when paired with S342/S344 pass-through at
                        // mod.rs:4073/4246.)
                        let s344_fs_gate = std::env::var("OXI_S344_FS_LT_DEFAULT").map(|v| v != "0" && v != "false").unwrap_or(false);
                        let s344_skip = s344_fs_gate
                            && !para_style.snap_to_grid
                            && font_size >= default_fs;
                        if h6_trigger || h7_trigger || h8_trigger || s344_skip {
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
                        current_width = 0.0; current_width_tw = 0; compress_used = false;
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
                        // §4.6.2 autoSpaceDE per-fontSize formula — COM-confirmed 2026-04-08.
                        //   9-10.5pt → 2.5pt, 11-12pt → 3.0pt, 14pt → 3.5pt,
                        //   16pt → 4.0pt, 18pt → 4.5pt (8 sizes verified, both directions).
                        let extra = ((font_size / 2.0) + 0.5).floor() * 0.5;
                        if let Some(last) = current_line.fragments.last_mut() {
                            last.width += extra;
                            last.natural_width += extra;
                        }
                        current_width += extra;
                        current_width_tw += pt_to_tw(extra);
                    }

                    let overflow_tw = current_width_tw + pt_to_tw(char_width) - available_tw;
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
                    let absorb = if overflow_tw > 0 && overflow_tw <= 50
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
                        let s228_block_hang = !legacy_s228
                            && compress_used
                            && is_para_last_char
                            && is_sentence_terminator;
                        let can_hang = kinsoku::is_hangable_punct(ch) && !next_is_proh
                            && !s228_block_hang;

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
                            current_width = 0.0; current_width_tw = 0; compress_used = false;
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
                        current_width = 0.0; current_width_tw = 0; compress_used = false;
                        for f in popped.into_iter().rev() {
                            current_width += f.width;
                            current_width_tw += pt_to_tw(f.width);
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
                            // §4.6.2 autoSpaceDE per-fontSize formula — COM-confirmed 2026-04-08.
                            //   9-10.5pt → 2.5pt, 11-12pt → 3.0pt, 14pt → 3.5pt,
                            //   16pt → 4.0pt, 18pt → 4.5pt (8 sizes verified, both directions).
                            let extra = ((font_size / 2.0) + 0.5).floor() * 0.5;
                            if let Some(last) = current_line.fragments.last_mut() {
                                last.width += extra;
                                last.natural_width += extra;
                            }
                            current_width += extra;
                            current_width_tw += pt_to_tw(extra);
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
                current_width = 0.0; current_width_tw = 0; compress_used = false;
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
        for line in &mut lines {
            if !line.was_compressed { continue; }
            let savings: f32 = line.fragments.iter()
                .map(|f| (f.natural_width - f.width).max(0.0)).sum();
            if savings <= 0.5 { continue; }
            let demand = (line.natural_total_width - available_width).max(0.0);
            if demand <= 0.5 {
                // Full revert: loose line, no compression needed
                for f in &mut line.fragments {
                    f.width = f.natural_width;
                }
                line.was_compressed = false;
            } else if demand < savings {
                // Partial revert: demand-scaled. Release (savings - demand) back to
                // compressed fragments proportionally, matching Word's per-line
                // demand-driven compression on line-start yakumono (d77a pi=24-27
                // COM: ・ compresses -0.5 to -2.5pt based on line overflow demand).
                let keep_ratio = demand / savings;
                for f in &mut line.fragments {
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
                if snap_to_grid && is_single && cell_snap_allowed {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            return (((spaced + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch;
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
                    spaced
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
    ) -> f32 {
        self.line_height_for_line_inner(line, para_style, para_font_size, snap_to_grid, grid_pitch, false)
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

    fn line_height_for_line_inner(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
        in_table_cell: bool,
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
                            // S195: narrower grid-snap tolerance — empty paragraph
                            // whose natural lh slightly exceeds pitch (e.g. 14pt MS
                            // Mincho 18.126pt at 18pt pitch). Without this it
                            // over-snaps to 2 cells (36pt). 3a4f pi=132 pattern.
                            // Gate: empty paragraph + base just over pitch (≤ pitch + 0.5).
                            // S238 (2026-05-23): removed OXI_LEGACY_NO_GRID_SNAP_TOL
                            // legacy env-var fallback during hardening pass.
                            let is_empty = line.fragments.iter()
                                .all(|f| f.text.is_empty());
                            let just_over_pitch = base > pitch && base <= pitch + 0.5;
                            let apply_tol = is_empty && just_over_pitch;
                            let tol = if apply_tol { 0.5 } else { 0.0 };
                            (((base - tol + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch
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
                            // S195: narrower grid-snap tolerance (see comment above)
                            let is_empty = line.fragments.iter()
                                .all(|f| f.text.is_empty());
                            let just_over_pitch = spaced > pitch && spaced <= pitch + 0.5;
                            let apply_tol = is_empty && just_over_pitch;
                            let tol = if apply_tol { 0.5 } else { 0.0 };
                            return (((spaced - tol + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch;
                        }
                    }
                }
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
                if !in_shape_context {
                    return 0.5;
                }
                // Shape context: text at bottom of line box (extra space above).
                // Per spec §13.4 note: "GDI TextOutW character cell = fontSize".
                // offset = line_height - max_font_size. COM-confirmed on 1ec1 p1
                // Shape 4 exact=22pt fontSize=14pt → 8pt offset (not 3.85pt).
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
                        if use_floor {
                            (raw * 2.0).floor() / 2.0
                        } else if use_round {
                            (raw * 2.0).round() / 2.0
                        } else {
                            (raw * 2.0 + 0.5).floor() / 2.0
                        }
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
                    if use_floor {
                        (raw * 2.0).floor() / 2.0
                    } else if use_round {
                        (raw * 2.0).round() / 2.0
                    } else {
                        (raw * 2.0 + 0.5).floor() / 2.0
                    }
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
            let border_w = table.style.border_width.unwrap_or(0.5);
            // 2026-04-19: Apply margin-padding-border offset ONLY when indent is
            // EXPLICITLY set to 0 (Some(0.0)). When absent (None), use plain margin.
            // b35 measurement: Word table at margin+5.6 (= cell text start), border at
            // margin. Prior formula subtracted 5.2pt from margin → 5.2pt left-shift bug.
            // 2026-05-03: Additionally gate on `!explicit_borders`. 683ff has
            // explicit `<w:tblBorders>` in tblPr — Word renders its border AT
            // the margin (no padding subtraction). Style-only-bordered tables
            // (e.g. b35/gen2_052) need the offset; explicit-borders tables don't.
            let border_offset = match table.style.indent {
                Some(v) if v.abs() < 0.01 && !table.style.explicit_borders => {
                    pad_l_default + border_w / 2.0
                }
                _ => 0.0,
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
        for (row_idx, row) in table.rows.iter().enumerate() {
            let mut row_height: f32 = 0.0;
            // Session 79c: visual_row_h = max cell content_h with emit-equivalent
            // line-height formula (grid-snapped when adjustLineHeightInTable). Used
            // ONLY for vAlign=center offset, NOT for row_height (page break logic
            // preserves the natural pre-pass to avoid 3a4f9f cascade — see
            // session79_adjust_lh_in_table_mixed_cell_valign_falsified.md).
            let mut visual_row_h: f32 = 0.0;
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
                // Word allows text to extend into cell margins for wrapping purposes
                let inner_w = cell_w.max(0.0);
                let mut cell_content_h = pad_t;
                // Session 79c: parallel emit-equivalent content_h for visual_row_h
                let mut cell_content_h_visual = pad_t;

                // Session 131 (2026-05-20): vertical writing — cell height
                // along the page-y axis equals the sum of vertical-text lengths
                // (chars × font_size), not the wrapped-horizontal line count.
                // Gated by OXI_VERT_WRITING env var.
                let vert_writing_active = self.is_vert_writing_active(cell);
                for block in &cell.blocks {
                    match block {
                        Block::Paragraph(para) => {
                            let (para_h, para_h_visual) = if vert_writing_active {
                                let h = self.vert_para_height(para);
                                (h, h)
                            } else {
                                let p1 = self.estimate_para_height(para, inner_w, row_line_pitch, table.style.para_style.as_ref(), true, grid_char_pitch, grid_char_cw_ratio);
                                let p2 = self.estimate_para_height_emit(para, inner_w, row_line_pitch, table.style.para_style.as_ref(), true, grid_char_pitch, grid_char_cw_ratio);
                                (p1, p2)
                            };
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
                        }
                        Block::Table(nested) => {
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

                row_height = row_height.max(cell_content_h);
                visual_row_h = visual_row_h.max(cell_content_h_visual);
                grid_idx += span;
            }

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

            // needs_row_split: only when overflow + table allows split.
            // widow_break_needed overrides split — we want the whole table on next page.
            let needs_row_split = row_overflows && !row.cant_split && has_content
                && (is_single_cell_row || has_lrpb_mid_row)
                && !widow_break_needed;

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
                let cell_w: f32 = col_widths[cell_start_grid..cell_end_grid].iter().sum();

                let pad_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(default_pad_l);
                let pad_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(default_pad_r);
                let mut pad_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(default_pad_t);
                let pad_b = cell.margins.as_ref().and_then(|m| m.bottom).unwrap_or(default_pad_b);

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

                // Layout blocks in document order (paragraphs and nested tables interleaved)
                let is_exact = row.height_rule.as_deref() == Some("exact");
                // R7.32: count Paragraph blocks within this cell so each cell
                // paragraph can be distinguished in the dump output.
                let mut cell_para_counter: usize = 0;
                // R7.73: track whether the immediately-previous cell paragraph
                // carried a `<w:lastRenderedPageBreak/>` on a non-run-0 run.
                // Reset to false at each cell start.
                let mut prev_cell_para_had_mid_lrpb: bool = false;
                if !is_vmerge_continue {
                for block in &cell.blocks {
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
                    );
                    for elem in nested_elements {
                        cell_elements.push(elem);
                    }
                    content_h = nested_y.cursor_y;
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
                        // S405 (2026-05-28) tested OXI_S405_INNER_INDENT
                        // (subtract padding from wrap_base when p_first_line_indent>0)
                        // for ed025 T16 cell 2 pi=4 "(× × ×)" wrap. Env-gate fires
                        // correctly (cell_w=90.40, pad=4.95+4.95, p_fli=10.5pt;
                        // wrap_base went 90.40→80.50, first_line_wrap_w 79.9→70pt)
                        // but pi=4 STILL renders as 1 line h=18 (ed025 score
                        // unchanged 0.9986). Root cause is character-width: Oxi
                        // treats × (U+00D7) as HALFWIDTH (5.25pt at 10.5pt MS
                        // Mincho) since 0x00D7 is in Latin-1 Supplement range,
                        // NOT in is_fullwidth()'s 0x2200-0x22FF (Math Operators).
                        // Oxi text width: （(10.5)+×(5.25)+　(10.5)+×(5.25)+　(10.5)
                        // +×(5.25)+）(10.5) = 57.75pt → fits even in 70pt wrap.
                        // Word likely treats × as FULLWIDTH in CJK font context
                        // (text=73.5pt → wraps). Global change to is_fullwidth()
                        // would affect many docs that use × as math operator.
                        // Reverted; needs S406+ careful per-font/per-context handling.
                        let wrap_base = if cell_hang_inner || s301_layout_fixed {
                            (cell_w - pad_l - pad_r).max(0.0)
                        } else {
                            cell_w
                        };
                        let wrap_w = (wrap_base - p_indent_left - p_indent_right).max(0.0);
                        let first_line_wrap_w = if p_first_line_indent < 0.0 {
                            (wrap_base - (p_indent_left + p_first_line_indent).max(0.0) - p_indent_right).max(0.0)
                        } else {
                            (wrap_w - p_first_line_indent).max(0.0)
                        };

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
                        let mut current_line_chars: Vec<crate::layout::jc_both_compress::CharContext> = Vec::new();
                        let mut is_first_line = true;
                        // R7.51 (2026-05-13): autoSpaceDE state for CJK↔Latin transitions.
                        // Tracks the last emitted character across runs/buffers so we can
                        // detect transitions and add Word's 2.5pt (10.5pt font) gap. The
                        // body renderer (break_into_lines) applies this; this cell-renderer
                        // path historically did not, causing d77a58 w_i=47 wrap mismatch
                        // (5 lines Oxi vs 6 lines Word).
                        let mut prev_char_emitted: Option<char> = None;

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
                            for ch in run.text.chars() {
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
                                        ((font_size / 2.0) + 0.5).floor() * 0.5
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
                                let would_overflow = if jc_gate_active && run_has_neg_cs && would_overflow_natural {
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
                                        // Add to buffer and break AFTER this char
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
                            if let Some(empty_fs) = pprrpr_fs {
                                let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
                                let empty_metrics = self.metrics_for_para_mark(&rpr_ref, &para.style);
                                content_h += self.line_height_inner(empty_fs, effective_line_spacing, effective_line_rule, empty_metrics, para.style.snap_to_grid, row_line_pitch, true);
                            } else {
                                let metrics = self.doc_default_metrics();
                                content_h += self.line_height_inner(self.default_font_size, effective_line_spacing, effective_line_rule, metrics, para.style.snap_to_grid, row_line_pitch, true);
                            }
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
                            let effective_wrap = if line_idx == 0 { first_line_wrap_w } else { wrap_w };

                            // Justify: non-last lines for jc=both, all lines for distribute
                            let is_last_line = line_idx == total_lines - 1;
                            let should_justify = (para.alignment == Alignment::Justify && !is_last_line)
                                || para.alignment == Alignment::Distribute;

                            // Alignment within cell content area (cell_w - padding).
                            // 2026-04-19: effective_wrap already = cell_w - pad_l - pad_r (v9).
                            // Previous code subtracted padding AGAIN → center off by ~5pt.
                            // Fix: use effective_wrap directly as content area.
                            let align_avail = effective_wrap.max(0.0);
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
                                    0.25
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
                            let mut rx = 0.0_f32;
                            // Emit list marker on the first line of the paragraph.
                            if line_idx == 0 {
                                if let Some((ref mk_text, mk_fs, mk_w)) = list_marker_info {
                                    let list_indent = para.style.list_indent.unwrap_or(18.0);
                                    let marker_style = para.runs.first().map(|r| &r.style).cloned().unwrap_or_default();
                                    // Session 75 Phase D: y is LINE BOX TOP; renderer adds cell_text_y_off.
                                    let mut marker_el = LayoutElement::new(
                                        cell_x + pad_l + line_indent - list_indent,
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
                                }));
                            }
                        }
                    }
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

                    // Resolve border color and width from cell borders, falling back to table style
                    let resolve_border = |side: Option<&BorderDef>| -> (Option<String>, f32) {
                        if let Some(b) = side {
                            let c = b.color.as_ref().map(|c| {
                                if c.starts_with('#') { c.clone() } else { format!("#{}", c) }
                            });
                            (c, b.width)
                        } else if table.style.border {
                            // Table-level borders: use table style color, default to black
                            let c = Some(table.style.border_color.as_ref()
                                .map(|c| if c.starts_with('#') { c.clone() } else { format!("#{}", c) })
                                .unwrap_or_else(|| "#000000".to_string()));
                            (c, table.style.border_width.unwrap_or(0.4))
                        } else {
                            (None, 0.4)
                        }
                    };

                    let cell_borders = cell.borders.as_ref();
                    let (top_color, top_width) = resolve_border(cell_borders.and_then(|b| b.top.as_ref()));
                    let (bot_color, bot_width) = resolve_border(cell_borders.and_then(|b| b.bottom.as_ref()));
                    let (left_color, left_width) = resolve_border(cell_borders.and_then(|b| b.left.as_ref()));
                    let (right_color, right_width) = resolve_border(cell_borders.and_then(|b| b.right.as_ref()));

                    // When cells have their own borders (tcBorders), draw each side per cell.
                    // When using table-level borders, use collapsed model to avoid double-drawing.
                    let use_collapsed = table.style.border && !has_cell_borders;

                    // Top — skip for vMerge continue cells (internal to merged range)
                    if !is_vmerge_continue && top_color.is_some() && (!use_collapsed || row_idx == 0) {
                        elements.push(LayoutElement::new(bx, by, cell_w, 0.0, LayoutContent::TableBorder {
                                x1: bx, y1: by, x2: bx + cell_w, y2: by,
                                color: top_color, width: top_width,
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
                                color: bot_color, width: bot_width,
                        }));
                    }
                    // Left
                    if left_color.is_some() && (!use_collapsed || cell_idx == 0) {
                        elements.push(LayoutElement::new(bx, by, 0.0, row_height, LayoutContent::TableBorder {
                                x1: bx, y1: by, x2: bx, y2: by + row_height,
                                color: left_color, width: left_width,
                        }));
                    }
                    // Right
                    if right_color.is_some() {
                        elements.push(LayoutElement::new(bx + cell_w, by, 0.0, row_height, LayoutContent::TableBorder {
                                x1: bx + cell_w, y1: by, x2: bx + cell_w, y2: by + row_height,
                                color: right_color, width: right_width,
                        }));
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
                        LayoutContent::TableBorder { y1, y2, x1, x2, ref color, width } => {
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
                                            color: color.clone(), width: *width,
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
                                            color: color.clone(), width: *width,
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
                let min_overflow_text_y = next_page_elems.iter()
                    .filter(|e| matches!(e.content, LayoutContent::Text { .. }))
                    .map(|e| e.y)
                    .fold(f32::INFINITY, f32::min);
                if min_overflow_text_y.is_finite() {
                    let original_shift = split_y - page_top;
                    let correct_shift = (min_overflow_text_y + original_shift) - page_top;
                    let adjust = correct_shift - original_shift;
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
                        LayoutContent::TableBorder { y1, y2, x1, x2, color, width }
                            if (*y1 - *y2).abs() >= 0.1 => {
                            Some((*x1, *x2, color.clone(), *width))
                        }
                        _ => None,
                    });

                    if let (Some(bi), Some((_, _, color, vw))) =
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
                                            color, width: vw,
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
                                LayoutContent::TableBorder { y1, y2, x1, x2, color, width }
                                    if (*y1 - *y2).abs() < 0.1 => {
                                    Some((*x1, *x2, color.clone(), *width))
                                }
                                _ => None,
                            });
                            if let Some((bx1, bx2, color, bw)) = template {
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
                                        color, width: bw,
                                    },
                                ));
                            }
                        }
                    }
                }

                // Push current page elements
                elements.extend(current_page_elems);
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });

                // Handle multi-page overflow: if next_page_elems still overflow,
                // keep splitting into additional pages.
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
                    let next_split = if lrpb_next_split.is_finite() && lrpb_next_split < page_bottom {
                        lrpb_next_split
                    } else {
                        page_bottom
                    };
                    let mut this_page: Vec<LayoutElement> = Vec::new();
                    let mut overflow: Vec<LayoutElement> = Vec::new();

                    for elem in remaining {
                        let _elem_top = elem.y;
                        match &elem.content {
                            LayoutContent::TableBorder { y1, y2, x1, x2, ref color, width } => {
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
                                                color: color.clone(), width: *width,
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
                                                color: color.clone(), width: *width,
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

                    pages.push(LayoutPage {
                        width: page_width,
                        height: page_height,
                        elements: this_page,
                    });
                    remaining = overflow;
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
                        ((font_size / 2.0) + 0.5).floor() * 0.5
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
            in_cell, grid_char_pitch, grid_char_cw_ratio, false)
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
                    }
                    Block::Table(nested) => {
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
            in_cell, grid_char_pitch, grid_char_cw_ratio, true)
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
                self.count_cell_lines(para, effective_width, first_line_wrap_w, gcp_for_count, gcr_for_count)
            } else {
                let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = para.runs.iter().enumerate()
                    .map(|(ri, run)| (run.text.as_str(), &run.style, None, ri, 0))
                    .collect();
                let lines = self.break_into_lines(&fragments, effective_width, first_indent, &para.style, None, None);
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
                let lh = if is_single_run {
                    if snap_in_cell {
                        self.line_height_inner(font_size, eff_ls, eff_lr, metrics, true, grid_pitch, true)
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
