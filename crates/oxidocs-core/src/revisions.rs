//! S-03 — IR-level accept / reject commands for tracked changes.
//!
//! These mutate the `Document` in place, baking the accepted (or rejected)
//! state into the IR. They are the editor-side counterparts to the
//! `LayoutEngine::with_show_revisions(...)` view modes (S-02), which only
//! filter at render time without touching the underlying IR.
//!
//! Behaviour mirrors Word's "Accept" / "Reject" review actions:
//!
//! - **Accept**: insertion runs (`<w:ins>`) become permanent body text;
//!   deletion runs (`<w:del>`) are removed. `moveFrom` is removed,
//!   `moveTo` becomes permanent body text. The run's `tracked_change`
//!   field is cleared on the survivors.
//! - **Reject**: inverse — deletions become permanent (i.e. survive),
//!   insertions are removed.
//!
//! See `docs/spec/comments_tracked_changes/attack_matrix.md` row S-03.

use crate::ir::{Block, Document, Run};

/// Accept every tracked change in the document. Equivalent to running the
/// layout pipeline with `ShowRevisions::Final`, then persisting the result.
pub fn accept_all(doc: &mut Document) {
    apply_review(doc, ReviewMode::AcceptAll, None);
}

/// Reject every tracked change in the document. Mirror of `accept_all`.
pub fn reject_all(doc: &mut Document) {
    apply_review(doc, ReviewMode::RejectAll, None);
}

/// Accept a single tracked change identified by `pair_id` (the `w:id`
/// attribute on the `<w:ins>` / `<w:del>` / `<w:moveFrom>` / `<w:moveTo>`
/// wrapper). Other revisions are left untouched.
pub fn accept_revision(doc: &mut Document, id: &str) {
    apply_review(doc, ReviewMode::AcceptOne, Some(id));
}

/// Reject a single tracked change identified by `pair_id`.
pub fn reject_revision(doc: &mut Document, id: &str) {
    apply_review(doc, ReviewMode::RejectOne, Some(id));
}

#[derive(Clone, Copy, PartialEq)]
enum ReviewMode {
    AcceptAll,
    RejectAll,
    AcceptOne,
    RejectOne,
}

fn apply_review(doc: &mut Document, mode: ReviewMode, target_id: Option<&str>) {
    fn visit(blocks: &mut Vec<Block>, mode: ReviewMode, target_id: Option<&str>) {
        for block in blocks.iter_mut() {
            match block {
                Block::Paragraph(p) => {
                    p.runs.retain_mut(|run| handle_run(run, mode, target_id));
                }
                Block::Table(t) => {
                    for row in &mut t.rows {
                        for cell in &mut row.cells {
                            visit(&mut cell.blocks, mode, target_id);
                        }
                    }
                }
                Block::Image(_) | Block::UnsupportedElement(_) => {}
            }
        }
    }
    // Cover body, headers, footers, footnotes, endnotes, and textbox
    // contents — anywhere a tracked-change run may live.
    for page in &mut doc.pages {
        visit(&mut page.blocks, mode, target_id);
        visit(&mut page.header, mode, target_id);
        visit(&mut page.footer, mode, target_id);
        for footnote in &mut page.footnotes {
            visit(&mut footnote.blocks, mode, target_id);
        }
        for endnote in &mut page.endnotes {
            visit(&mut endnote.blocks, mode, target_id);
        }
        for tb in &mut page.text_boxes {
            visit(&mut tb.blocks, mode, target_id);
        }
    }
}

/// Returns `false` to drop the run, `true` to keep it. Mutates the run's
/// `tracked_change` and parser-applied styling in place when keeping.
fn handle_run(run: &mut Run, mode: ReviewMode, target_id: Option<&str>) -> bool {
    let Some(tc) = run.tracked_change.as_ref() else {
        return true;
    };

    // For per-id modes, leave non-matching revisions untouched (return true,
    // don't clear tracked_change).
    if let Some(id) = target_id {
        if tc.pair_id.as_deref() != Some(id) {
            return true;
        }
    }

    // Accept = keep insertions, drop deletions.
    // Reject = drop insertions, keep deletions.
    let is_insertion =
        matches!(tc.change_type.as_str(), "insert" | "moveTo");
    let is_deletion =
        matches!(tc.change_type.as_str(), "delete" | "moveFrom");
    let accept = matches!(mode, ReviewMode::AcceptAll | ReviewMode::AcceptOne);

    let drop = (accept && is_deletion) || (!accept && is_insertion);
    if drop {
        return false;
    }

    // Survivor — strip `tracked_change` and the parser's pre-applied
    // tracked-change styling (underline + FF0000 on insertions, strikethrough
    // + FF0000 on deletions). After acceptance/rejection the run reads as
    // plain body text.
    let kind = tc.change_type.clone();
    run.tracked_change = None;
    if matches!(kind.as_str(), "insert" | "moveTo") {
        run.style.underline = false;
        run.style.underline_style = None;
        if run.style.color.as_deref() == Some("FF0000") {
            run.style.color = None;
        }
    } else if matches!(kind.as_str(), "delete" | "moveFrom") {
        run.style.strikethrough = false;
        if run.style.color.as_deref() == Some("FF0000") {
            run.style.color = None;
        }
    }
    true
}
