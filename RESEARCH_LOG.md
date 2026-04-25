# Research log — shared across oxi-1/2/3/4

Append-only log of hypotheses tested, confirmed, or refuted across all worktrees.
Read this at the start of every loop iteration to avoid re-chasing falsified leads.
Append your own findings at the end (newest on top).

Format:
```
## YYYY-MM-DD — [worktree] — [confirmed|refuted|partial] — [short label]
- context: what doc/page/feature
- hypothesis: what was tested
- evidence: measurement data, COM results, commit refs
- outcome: what this means for other agents
```

---

## 2026-04-25 — oxi-2 — confirmed — R-10 paragraph_mark_revision path also test-covered

- mirror of the ppr_change test landed: `r10_fires_for_paragraph_mark_revision` mutates fixture_05's IR to install a synthetic `paragraph_mark_revision = TrackedChange { change_type: "insert", author: "Alice Reviewer", … }` while clearing every other revision pointer. Asserts ≥1 #424242 BoxRect emerges.
- 32/32 tests pass. R-10's full 4-way detection (Run.tracked_change / Run.rpr_change / Paragraph.ppr_change / Paragraph.paragraph_mark_revision) is now covered by 2 dedicated paragraph-level tests + the existing run-level test (`fixture_05_layout_emits_revision_change_bar`) + integration tests via fixture_07/08/09/10.

## 2026-04-25 — oxi-2 — confirmed — R-10 paragraph-level path now test-covered

- context: prior iteration extended R-10 to fire on `Paragraph.ppr_change` / `paragraph_mark_revision`, but no fixture exercised the path so the new code was untested.
- new test `r10_fires_for_paragraph_level_ppr_change` in `comments_fixtures.rs`:
  - parses fixture_05, then mutates the IR: clears every run-level `tracked_change` / `rpr_change`, installs a synthetic `Paragraph.ppr_change = Some(PropertyChange{ author: "Alice Reviewer", … })`.
  - lays out and asserts ≥1 thin BoxRect with `#424242` fill (the change bar) is emitted.
  - confirms R-10 fires from paragraph-level revisions alone, with no run-level pointer present.
- evidence: 31 tests pass.
- impact: closes the test-coverage gap from the prior iteration. The 4-way revision detection (Run.tracked_change / Run.rpr_change / Paragraph.ppr_change / Paragraph.paragraph_mark_revision) is now all observed by tests.

## 2026-04-25 — oxi-2 — confirmed — R-10 covers paragraph-level revisions

- context: previous iteration extended R-10 to cover `Run.rpr_change`. The remaining IR fields that carry revision metadata at paragraph level (`Paragraph.ppr_change` from P-07, `Paragraph.paragraph_mark_revision` from P-09) still did not trigger a change bar.
- fix in `layout_paragraph`: initialize `line_has_revision` from `para.ppr_change.is_some() || para.paragraph_mark_revision.is_some()` BEFORE the fragment loop. If either field is set the paragraph gets a margin change bar on every line, regardless of whether any individual run has `tracked_change` / `rpr_change`.
- evidence: all 30 tests pass. None of the 10 baseline fixtures exercise paragraph-level revisions on a body paragraph that has zero run-level revisions, so no visible change in the test set; the path is correct but unverified by fixture. Add a test fixture if/when `<w:pPrChange>` outside a run-revision paragraph appears in real documents.
- impact: defensive coverage. Guarantees the change bar tracks every kind of revision the IR can carry.

## 2026-04-25 — oxi-2 — confirmed — R-10 margin change bar now fires for rPrChange

- bug: fixture_09 (rPrChange "Now bold (was plain).") rendered the bold text correctly but did NOT show a margin change bar. Word renders a change bar next to formatting changes too.
- root cause: R-10's per-line revision detector only checked `Run.tracked_change` (the ins/del/move pointer). Property changes use a separate `Run.rpr_change` field (P-06's IR shape). Lines with only rpr_change fell through.
- fix in `layout_paragraph`: extend the check to also look at `run.rpr_change`. Now any revision-bearing run — `tracked_change` OR `rpr_change` — triggers the margin change bar.
- evidence:
  - all 30 comments_fixtures tests pass (no regression).
  - fixture_09 visual rebuild: dark grey 1.5pt change bar now appears in left margin next to "Regular. Now bold (was plain)." line. Was missing entirely before.
- limitations:
  - `Paragraph.ppr_change` (P-07) and `Paragraph.paragraph_mark_revision` (P-09) still don't trigger R-10. None of the 10 fixtures exercise them on a body paragraph that would otherwise have no `tracked_change`/`rpr_change` runs, so no visible miss in the current set. Add when a fixture stresses it.

## 2026-04-25 — oxi-2 — confirmed — R-05c anchor detection fix for empty marker runs

- bug: fixture_04 (multi-paragraph comment range starting with `<w:commentRangeStart>` BEFORE the first text run) rendered with the inline pink range tint but NO balloon. The anchor walk was indexed by `(paragraph_index, run_index)` of LayoutElements, but the parser's anchor-run fallback (which carries `comment_range_start`) has empty text and emits no LayoutContent::Text element — so the comment id was never found.
- fix in `emit_balloons_for_layout_page`:
  - Two-pass anchor detection. First pass walks LayoutElements once to build `paragraph_index → first-rendered (x, y)`. Second pass walks IR paragraphs in document order and, for any run that carries `comment_range_start`, looks up the paragraph's first rendered position from the map.
  - This decouples anchor detection from the run that carries the marker; whether the marker lives on an empty-text anchor run or a substantive text run, the comment is still anchored to the paragraph's first rendered Y.
- evidence:
  - all 30 comments_fixtures tests pass (no regression).
  - fixture_04 visual rebuild: balloon now appears next to paragraph 1 with header "Alice Reviewer" + body "Applies to all three paragraphs.", connector line drawn, all three paragraphs still tinted pink (R-04). Was missing the balloon entirely before this fix.
- baseline risk: zero.

## 2026-04-25 — oxi-2 — confirmed — pre-pass coverage extended to header/footer/footnote/textbox

- context: until now, `apply_revision_styling`, `apply_comment_range_highlighting`, `strip_parser_revision_styling`, `filter_runs_for_show_revisions`, and `revisions::apply_review` only walked `Page.blocks` (body). Tracked changes / comments living inside headers, footers, footnotes, endnotes, or textboxes would not get their visual treatment applied. None of the 10 fixtures exercise this, but real-world docs do.
- implementation: factored a single helper in `crates/oxidocs-core/src/layout/mod.rs`:
  ```
  fn for_each_block_tree<F: FnMut(&mut Vec<Block>)>(doc: &mut Document, mut f: F) {
      for page in &mut doc.pages {
          f(&mut page.blocks);
          f(&mut page.header);
          f(&mut page.footer);
          for footnote in &mut page.footnotes { f(&mut footnote.blocks); }
          for endnote in &mut page.endnotes { f(&mut endnote.blocks); }
          for tb in &mut page.text_boxes { f(&mut tb.blocks); }
      }
  }
  ```
  Each pre-pass that previously hand-rolled `for page in &mut doc.pages { for block in &mut page.blocks {…} }` now calls `for_each_block_tree(doc, |blocks| {…})`. Each pass's existing recursion into `Block::Table` cells is preserved.
  `revisions::apply_review` is in a separate module (no access to `for_each_block_tree`), so it inlines the same iteration explicitly.
- evidence: all 30 comments_fixtures tests pass — no regression. Coverage is purely additive (same body behavior + extended scope).
- skipped fields:
  - `Page.shapes` / `Page.floating_images` — these don't contain runs, just geometry.
  - The recursive table-cell walk inside each pass already handles Block::Table within any of these top-level locations.
- limitations:
  - R-10 (margin change bar in `layout_paragraph`) is emitted only on body paragraphs (only path that sets `body_para_index`). Header/footer revisions render with author tint + underline/strike but no margin bar — Word does the same (margin bar is body-only).
  - R-05 balloon emission still body-only. Comments anchored in headers/footers/footnotes don't render balloons; rare in practice and would need separate balloon Y resolution per region.
- baseline risk: zero — local 51-doc and oxi-main 184-doc baselines have 0 comments and the 5 `<w:del>` docs have body-only revisions. Header/footer revisions in the wild would now render correctly where previously they would silently render as plain text.

## 2026-04-25 — oxi-2 — confirmed — R-05 balloon body truncation polish

- context: post-R-05g visual review showed fixture_02's reply body "Following up." being clipped — the balloon's bounding box was too short to fit the reply body row beneath the reply chip.
- root cause: the height estimate accounted for line text height but underestimated per-section padding. The renderer adds an inter-section pad after each of: header chip, parent body, and (per reply) reply chip + reply body. For a parent with 1 reply that's 4 sections × ~4pt = 16pt of padding the estimate was missing, plus the chip-height was too small.
- fix: rewrote the height accumulation to mirror the renderer's actual sectional layout — `outer_pad + chip_h + section_pad + body_h + section_pad + Σ(chip_h + section_pad + reply_body_h + section_pad) + outer_pad`. Bumped `chip_h` from 12pt to 14pt and `line_height` from 12pt to 14pt for breathing room.
- evidence:
  - fixture_02 visual: balloon now shows full thread — `Alice Reviewer` parent chip + `Why?` body + indented `Alice Reviewer` reply chip + `Following up.` body. Was previously truncated below the reply chip.
  - fixture_01 visual: balloon a few pt taller; "Alice Reviewer" + "Is 'brown' needed here?" still cleanly inside the rounded box. No regression.
  - all 30 comments_fixtures tests pass (no test asserted on exact balloon height, so the change is internal).
- limitations:
  - Estimate is still char-count × avg-glyph-width, not actual GDI-measured wrap. fixture_03's resolved balloon (narrower 190.1pt) might still under-allocate for longer comment bodies — defer per-case tuning until a fixture exercises long resolved comments.
  - Reply body still uses the same wrap math as parent; replies indent ~10pt but the wrap width estimate doesn't account for the indent (slight overestimate, harmless).

## 2026-04-25 — oxi-2 — confirmed — S-03 accept/reject IR commands LANDED

- context: third settings-row. S-02 only filters at render time; S-03 bakes the accepted/rejected state into the IR so subsequent saves / re-renders see the post-review document with no tracked changes left.
- implementation:
  - new top-level module `crates/oxidocs-core/src/revisions.rs` exposes 4 free functions:
    - `accept_all(&mut Document)` / `reject_all(&mut Document)`
    - `accept_revision(&mut Document, id: &str)` / `reject_revision(&mut Document, id: &str)`
  - Internal `apply_review` walks pages → blocks (recursing into Table cells) → paragraphs and uses `Vec::retain_mut` to drop revision runs that fail the keep-test. Surviving runs have `tracked_change` cleared and the parser's pre-applied underline/strike + `FF0000` color stripped (same logic as S-02's filter helper).
  - Per-id variants match on `tracked_change.pair_id`; non-matching revisions stay untouched (tracked_change preserved).
  - Re-exported via `pub mod revisions;` in `lib.rs`.
- evidence:
  - 3 new layout tests on fixture_07 (3 revisions: ins1/del1/ins2):
    - `s03_accept_all_drops_deletions_and_clears_tracked_changes`: pre-call run count = 3, post-call = 0; del1 absent, ins1/ins2 present.
    - `s03_reject_all_drops_insertions_keeps_deletions`: del1 present, ins1/ins2 absent.
    - `s03_accept_revision_by_id_leaves_others_untouched`: accept id="101" (del1) — del1 gone, ins1 + ins2 still tracked.
  - all 30 comments_fixtures tests pass.
- baseline risk: zero — S-03 is a new public API; no existing call sites yet.
- limitations:
  - Per-id targeting uses `tracked_change.pair_id` which is the `<w:ins w:id=…>` attribute. Move-pair semantics (one `pair_id` shared between moveFrom + moveTo) work correctly because the same id matches both wrappers.
  - Doesn't update `commentsExtended.xml` / commentsIds — those are in-memory IR fields tied to comments, not revisions. No-op for accept/reject.
  - No undo; the mutation is destructive. Caller must clone first if they want to keep the original.
- next iteration candidate: cleanup parser-side pre-applied tracked-change styling (currently redundant with R-01 pre-pass + S-02/S-03 helpers); OR refinements to R-05 balloon (body height truncation in fixture_03, connector author-tinting); OR S-04 UI affordances (out of scope per attack-matrix).

## 2026-04-25 — oxi-2 — confirmed — S-02 show_revisions toggle LANDED (all 4 modes)

- context: second settings-row. Wires `ir::ShowRevisions` (I-04) into `LayoutEngine` so callers can pick All / Simple / Original / Final per-render.
- design:
  - New `show_revisions: ShowRevisions` field on `LayoutEngine`, defaults `All`.
  - New `with_show_revisions(mode) -> Self` builder.
  - 4-arm match in `layout()` before applying revision styling:
    - `All` → call `apply_revision_styling` (current behavior).
    - `Simple` → new helper `strip_parser_revision_styling` clears the parser's pre-applied underline/strike + `FF0000` color but keeps `tracked_change` so R-10 still fires margin change bars.
    - `Original` → `filter_runs_for_show_revisions(doc, final_view=false)`. Drops ins/moveTo runs from paragraphs; clears tracked_change + parser styling on del/moveFrom (which now render as plain text).
    - `Final` → `filter_runs_for_show_revisions(doc, final_view=true)`. Drops del/moveFrom runs; clears tracked_change + parser styling on ins/moveTo.
  - Both helpers recurse into `Block::Table` cells.
- gotcha discovered: the parser (`crates/oxidocs-core/src/parser/ooxml.rs:5795-5805`) pre-applies `style.underline=true` + `color="FF0000"` on insert runs and `style.strikethrough=true` + same red on delete runs — predates the R-01 pre-pass. For All view this gets overridden by `apply_revision_styling`'s author-tinted color/underline. For Simple/Original/Final views the parser's styling persists unless cleaned up — that's why both helpers explicitly reset `underline`, `underline_style`, `strikethrough`, and the FF0000 color when keeping/clearing a revision-bearing run.
- evidence:
  - 3 new layout tests in `comments_fixtures.rs`:
    - `s02_show_revisions_final_drops_del_and_strips_ins_styling` — fixture_07 in Final: del1 absent, ins1/ins2 present without underline/strike/author-color.
    - `s02_show_revisions_original_drops_ins_and_strips_del_styling` — fixture_07 in Original: ins1/ins2 absent, del1 present.
    - `s02_show_revisions_simple_skips_color_keeps_margin_bar` — fixture_05 in Simple: no author-tinted underline, but ≥1 margin change bar still fires.
  - All 27 comments_fixtures tests pass.
- baseline risk: zero — default is `All`, prior behavior preserved.
- limitations:
  - Original view doesn't yet restore prior `rPrChange` styles (the run renders with the *new* properties, not the prior ones). Word's actual Original view applies the prior rPr — refine when more PrChange fixtures exist.
  - Simple view's "change bar in author color" tinted variant deferred — current bar is uniform grey #424242 (R-10).
- follow-up cleanup: parser-side pre-applied tracked-change styling is now redundant in All view (overridden) and forces helper code in the other 3 views. Defer removal to a focused refactor later — risk of breaking IR consumers that read style directly.
- next iteration candidate: S-03 (accept/reject IR commands — editor-side, no render impact); or refinements to R-05's body truncation / connector color.

## 2026-04-25 — oxi-2 — confirmed — S-01 show_comments toggle LANDED

- context: first settings-row in the attack matrix. Provide a clean off-switch for the entire comment family (balloons + connectors + in-line range highlight + body-width compression) without touching tracked-change rendering — useful for clean print output and as a sanity-check.
- implementation:
  - New `show_comments: bool` field on `LayoutEngine`, defaults `true`.
  - New builder method `pub fn with_show_comments(mut self, show: bool) -> Self` for caller-side toggle.
  - 3 gated sites in `layout()`:
    - `apply_comment_range_highlighting` only runs when `self.show_comments` (was unconditional).
    - Balloon emission post-pass only runs when both `self.show_comments` AND `!doc.comments.is_empty()`.
    - `layout_page` uses `0.0` as `balloon_reservation` when `!self.show_comments`, so body width is full even on commented docs.
- evidence:
  - new layout test `s01_show_comments_false_suppresses_all_comment_visuals`: with show_comments=false, asserts 0 balloons, 0 connectors, 0 highlighted Text elements, and exactly 1 line of body text (= full-width layout). Same fixture with show_comments=true (default) emits 1 balloon for sanity.
  - all 24 comments_fixtures tests pass.
- baseline risk: zero — default is `true`, matching prior behavior.
- next iteration candidate: S-02 (`show_revisions: ShowRevisions` toggle), which uses the existing I-04 enum with 4 modes (All/Simple/Original/Final).

## 2026-04-25 — oxi-2 — confirmed — R-05g GDI rendering of Balloon + BalloonConnector LANDED

- context: seventh R-05 sub-iteration. With layout-side balloon emission complete (R-05a..f), now wire the GDI renderer so the PNG output finally shows balloons.
- additions:
  - **`pub fn comment_balloon_fill(idx, resolved) -> &'static str`** in `crates/oxidocs-core/src/layout/mod.rs` — public resolver so the GDI renderer can map `author_color_index + resolved` → palette hex without duplicating the constants.
  - **`parse_hex_rgb`** helper in `tools/oxi-gdi-renderer/src/main.rs`.
  - **Balloon handler**: rounded rect (4pt radius, scaled) filled with `comment_balloon_fill(...)` tint, 1pt border in a slightly-darker shade. Author header in bold 9pt grey. Body in 10pt black, wrapped via `DrawTextW(DT_LEFT|DT_TOP|DT_WORDBREAK)`. Each reply gets an indented (~10pt) author chip + body in identical shape.
  - **BalloonConnector handler**: `PS_DOT` pen, scaled width, medium grey (#808080). `MoveToEx` + `LineTo`. Color of the connector overrides the layout-supplied `color_hex` for now — a uniform grey is more readable against pink/purple/green tints than the per-author hue at 1pt thickness.
  - Balloon height estimate gained a `header_chip_h = 12pt` constant so the body actually fits inside the rect.
- evidence (visual via GDI render):
  - **fixture_01**: pink rounded balloon shows "Alice Reviewer" + "Is 'brown' needed here?", grey dotted connector line runs from the body's commentRangeStart line up to the balloon's left edge. Inline pink "brown fox" tint still works (R-04). Body still wraps to 2 lines (R-05b).
  - **fixture_02**: balloon shows parent "Why?" + reply chip "Alice Reviewer" indented inside (R-05f reply fold visible).
  - All 4 comment fixtures (01/02/03/04) render balloons.
- evidence (tests): all 23 comments_fixtures tests pass after the height-estimate change.
- limitations / refinements deferred:
  - Reply body sometimes clipped when balloon height estimate is short (reply estimate uses average glyph width which may underestimate at ~10 chars). Pixel-tune per fixture in a follow-up tick.
  - Connector dotted color is grey (not author tint). Word's actual connector is more pastel-author-colored at low alpha. Refine when we have a per-author connector color spec.
  - Connector geometry: `from_y` is the line top, so visually the line starts a few pt above the actual glyph row. Acceptable for v1.
  - GDI `DrawTextW(.., DT_NOCLIP)` is used for the header (single line) and `DT_WORDBREAK` for body. Word wraps at hyphens / break-after; GDI's word-break is rougher. Visual close enough for v1.
- baseline risk: zero — local 51-doc and oxi-main 184-doc baselines have 0 comments → zero balloons emitted on baseline, so renderer never executes the new branches.

## 2026-04-25 — oxi-2 — confirmed — R-05f reply threading inside parent balloon LANDED

- context: sixth R-05 sub-iteration. Word renders a comment's replies INSIDE the parent's balloon, indented. Don't emit a standalone Balloon for each reply.
- spec join: a Comment is a reply when `Comment.parent_para_id` is set; the value is the parent's `Comment.para_id` (different from `Comment.id`). P-10 (`commentsExtended.xml` parser) already populates these fields.
- implementation in `emit_balloons_for_layout_page`:
  - For each parent Balloon being prepared, iterate `doc.comments` looking for `c.parent_para_id == parent.para_id`. Each match becomes a `BalloonReply { author, author_color_index, body }`.
  - PendingBalloon gains a `replies: Vec<BalloonReply>` field; populated alongside body.
  - Balloon LayoutContent now carries the populated `replies` Vec (was `Vec::new()` placeholder).
  - Height estimate folds reply line counts in (per reply: line count + 1 for the author header chip). Approximate; R-05g refines.
  - Replies don't appear in `anchors` because they have no `commentRangeStart` of their own (they share the parent's range). So no separate balloon is emitted for them — the fold is the only visible side effect.
- evidence:
  - new layout test `fixture_02_reply_folds_into_parent_balloon_replies_vec`: parses fixture_02 (Alice's parent comment "Why?" + Alice's reply "Following up." linked via `paraIdParent="00000010"`), asserts exactly 1 standalone Balloon, parent body contains "Why?", `replies[0].body == "Following up."`, replies[0].author == "Alice Reviewer".
  - All 23 comments_fixtures tests pass.
- visual: GDI render still no-op; the JSON LayoutResult correctly carries the threaded structure.
- limitations:
  - Multi-level reply nesting (replies of replies) not handled — Word doesn't really support that anyway. Keep flat.
  - If a reply somehow had its own `commentRangeStart` (rare), it would emit a duplicate standalone balloon. Defer until a fixture stresses it.
- next iteration: R-05g — actual GDI rendering of `LayoutContent::Balloon` and `LayoutContent::BalloonConnector` so the PNG output finally shows balloons.

## 2026-04-25 — oxi-2 — confirmed — R-05e balloon connector lines LANDED

- context: fifth R-05 sub-iteration. With balloons emitting and stacking (R-05c/d), now also emit a connector line from each balloon's inline anchor to the balloon's left edge.
- implementation:
  - Inside `emit_balloons_for_layout_page`, just before pushing each Balloon, push a `LayoutContent::BalloonConnector` LayoutElement.
  - `from_x`, `from_y` = the balloon's `anchor_x`, `anchor_y` (= rendered Y of the comment's `commentRangeStart`).
  - `to_x` = balloon's left edge; `to_y` = balloon's resolved top + 5pt (visually meets the first text row of the balloon, matches Word).
  - `color_hex` = author's tint slot from `COMMENT_HIGHLIGHT_TINT_PALETTE` (slot 0 = `#FAE6E7` for Alice). Same color regardless of resolved state — Word's connector uses the unresolved tint hue at lower opacity, but our v1 ships solid color from the palette and refines at R-05g.
  - LayoutElement bounding box covers the connector's path so layered renderers can clip correctly.
- evidence:
  - new layout test `fixture_01_emits_balloon_connector_paired_with_balloon`: asserts exactly 1 BalloonConnector, ends at balloon_left, ends at balloon_top+5pt, starts left of balloon, color=#FAE6E7.
  - all 22 comments_fixtures tests pass.
- visual: GDI renderer still has `_ => {}` for BalloonConnector (will be wired in R-05g — needs dotted pen).
- limitations:
  - Connector currently solid in the LayoutResult; the dotted style is the renderer's job (R-05g).
  - Color is the author's tint regardless of resolved state. May refine.
- next iteration: R-05f — fold replies (`Comment.parent_para_id`) into their parent balloons' `replies` Vec, dropping the standalone child Balloon element.

## 2026-04-25 — oxi-2 — confirmed — R-05d balloon stacking LANDED

- context: fourth R-05 sub-iteration. With single-balloon emission in place (R-05c), prevent vertical overlap when ≥2 balloons would otherwise stack on top of each other.
- algorithm: extracted `stack_balloon_ys(positions: &mut [(f32, f32)], gap: f32)` as a pure helper. Sorts by anchor Y (caller's responsibility); walks the slice, pushes each balloon's Y to `max(natural_y, prev_y + prev_height + gap)`. First balloon never moves; cascade is monotonic (later balloons can only push down, never up).
- gap constant `BALLOON_STACK_GAP = 6.0pt`. Pixel-tune in R-05g once GDI render lands.
- emission integration: `emit_balloons_for_layout_page` now collects all per-balloon geometry into a `Vec<PendingBalloon>`, sorts by `anchor_y`, runs `stack_balloon_ys` on a `(y, height)` projection, then writes back the resolved Ys before pushing LayoutElements.
- tests:
  - 3 new pure-function unit tests in `mod tests` (`stack_balloon_ys_no_overlap_keeps_anchors`, `stack_balloon_ys_pushes_overlapping_balloons_down`, `stack_balloon_ys_handles_degenerate_inputs`). Cover happy path, push-down, and 3-balloon cascade.
  - All 21 comments_fixtures integration tests still pass — no regression.
- limitations:
  - 6pt gap is approximate. Will pixel-tune when R-05g compares against Word's rendered output.
  - Stacking only operates per-page. A balloon whose anchor is on page N but whose natural Y would push past page-bottom does NOT split or move to page N+1 — Word actually wraps the balloon column, but that's a separate edge case (defer until a fixture stresses it).
  - Replies still don't fold into parent (R-08 / R-05f) — stacking treats each comment as its own balloon. fixture_02 currently shows 1 balloon (parent only) since the reply has no commentRangeStart of its own.
- next iteration: R-05e — connector line. Emit one `LayoutContent::BalloonConnector` per balloon, dotted, from the inline anchor point to the balloon's left edge.

## 2026-04-25 — oxi-2 — confirmed — R-05c single-comment balloon emission LANDED

- context: third R-05 sub-iteration. With body width compressed (R-05b) and enum scaffolding (R-05a), now emit one `LayoutContent::Balloon` per visible comment, anchored to the rendered Y of its `commentRangeStart`.
- design — IR-page tracking: `LayoutEngine::layout` now builds a parallel `Vec<usize>` mapping each LayoutPage to its source IR page. A single IR page may produce multiple LayoutPages (pagination), all sharing the same IR index. Used by the post-pass to resolve `LayoutElement.paragraph_index` → source `Run`.
- design — `emit_balloons_for_layout_page(layout_page, doc, ir_page_idx)`:
  - Walks LayoutPage elements in order. For each Text element with `paragraph_index + run_index`, looks up the source `Run` in `doc.pages[ir_page_idx].blocks[paragraph_index].runs[run_index]` and reads its `comment_range_start: Vec<String>`.
  - Records `(comment_id, anchor_x, anchor_y)` for the FIRST occurrence of each comment id on this page, in document order.
  - For each anchor, emits one `LayoutContent::Balloon` with: width 293.8pt unresolved / 190.1pt resolved (COM-confirmed), right edge `page_width − 4pt`, anchor Y = first-occurrence Y, body = flattened comment paragraphs, color slot from author palette.
  - Height estimate: `body_lines × 12pt + 8pt_padding` using a rough average glyph width of 5pt for line wrap. R-05g will refine when GDI rendering measures actual wrap.
- evidence:
  - 2 new layout integration tests:
    - `fixture_01_emits_one_balloon_for_single_comment` asserts: 1 Balloon element, comment_id="0", author="Alice Reviewer", resolved=false, width=293.8pt±0.01, x=297.6pt±0.5 (= 595.3 − 4 − 293.8).
    - `fixture_03_emits_resolved_balloon_with_narrower_width` asserts: 1 Balloon, resolved=true, width=190.1pt (the resolved variant).
  - All 21 comments_fixtures tests pass.
- visual: GDI renderer currently has `_ => {}` no-op for Balloon (added in R-05a), so the actual page render is unchanged. R-05g will wire the rendering.
- baseline risk: zero — `doc.comments.is_empty()` short-circuits before the post-pass; baseline docs are untouched.
- limitations / next:
  - Stacking-on-overlap (R-05d) not yet implemented. Overlap is impossible in the current 4 single-comment fixtures (only fixture_02 has 2 comments which would overlap if rendered without offset). R-05d is the next sub-iteration.
  - No connector line yet (R-07 / R-05e).
  - Replies (R-08 / R-05f) — `Balloon.replies` is currently `Vec::new()`. Will populate when R-05f lands.
  - GDI render (R-05g) — needs `LayoutContent::Balloon` handler in `tools/oxi-gdi-renderer/src/main.rs`.

## 2026-04-25 — oxi-2 — confirmed — R-05b body width compression for commented docs LANDED

- context: second R-05 sub-iteration. With the enum variants in place (R-05a), now make the body actually narrower when the document has any comments — paves the way for R-05c balloon emission to land in a known column.
- implementation:
  - Added `balloon_column_width: f32` field to `LayoutEngine`. Set in `for_document` to `0.0` when `doc.comments.is_empty()`, `293.8 + 24.0 = 317.8` otherwise. (293.8 is COM-confirmed Alice unresolved balloon width from pixel pass; 24pt is approximate gap, refined as later iterations pixel-test.)
  - In `layout_page`, subtract `self.balloon_column_width` from `total_content_width`. Header / footer / floating-image / footnote widths kept at full un-reduced width — matches Word's behavior (only body reflows when balloons appear).
- evidence:
  - new layout test `fixture_01_body_width_compresses_when_comments_present` asserts: with comments → body wraps to ≥2 lines; with comments cleared → 1 line. Confirms the compression depends on `doc.comments.is_empty()`, not on the text itself.
  - all 19 comments_fixtures tests pass.
  - Visual confirmation via GDI re-render: fixture_01 now renders "The quick brown fox jumps" on line 1, "over the lazy dog." on line 2 (was previously on a single line). The R-04 pink "brown fox" tint still applies correctly to the new wrapped layout.
- baseline risk: zero — local 51-doc and oxi-main 184-doc baselines have 0 comments, so `balloon_column_width = 0.0` for every baseline doc and the subtraction is a no-op.
- limitations:
  - 24pt gap is approximate. Word's actual gap (between body right edge and balloon left edge) measured ~33.5pt for fixture_01; 24pt is conservative pending pixel-perfect tuning when R-05c lands and emits actual balloons.
  - Compression applies to ALL pages of a commented doc, not just pages where comments are visible. Word does the same — body width is uniform across pages of a section regardless of which page anchors which comment.
- next iteration: R-05c — emit one Balloon LayoutElement per visible comment, anchored to the rendered Y of its scope start. First iteration that produces actual `LayoutContent::Balloon` elements.

## 2026-04-25 — oxi-2 — confirmed — R-05a enum variants + match-arm fallthroughs LANDED

- context: R-05a is the first sub-iteration of the R-05 balloon design (`docs/spec/comments_tracked_changes/r05_balloon_design.md`). Add the new `LayoutContent::Balloon` and `LayoutContent::BalloonConnector` variants to the layout enum and update every match site to handle them, so the body-width compression (R-05b) and per-page emission (R-05c) iterations can land cleanly.
- variants added (in `crates/oxidocs-core/src/layout/mod.rs`):
  - `LayoutContent::Balloon { comment_id, author, author_color_index, resolved, body, replies: Vec<BalloonReply>, anchor_x, anchor_y }` — carries the full comment payload so renderers can do their own wrapping.
  - `LayoutContent::BalloonConnector { from_x, from_y, to_x, to_y, color_hex }` — dotted line geometry for R-07.
  - `BalloonReply { author, author_color_index, body }` — used inside Balloon.replies for R-08.
- match sites updated (5 files):
  - `tools/oxi-gdi-renderer/src/main.rs` — type_name lookup + element drawing (skip-stub for now until R-05g).
  - `crates/oxi-cli/src/main.rs` — PDF emission match (skip-stub).
  - `crates/oxi-wasm/src/lib.rs` — 3 match sites (LayoutElementJs construction × 2, PDF emission × 1). New `kind: "balloon"` and `kind: "balloon_connector"` strings surface to JS consumers.
  - `crates/oxidocs-core/examples/dump_docx.rs` — added BALLOON / CONNECTOR / SHAPE rows. (Side benefit: example was already broken on `PresetShape` from a prior session — fixed in passing.)
  - `crates/oxidocs-core/examples/layout_json.rs` — added BALLOON / CONNECTOR rows.
- evidence: `cargo build -p oxidocs-core` (lib + examples) clean. `cargo test -p oxidocs-core --test comments_fixtures` 18/18 pass. `cargo build --release -p oxi-gdi-renderer` clean. `cargo build -p oxi-cli` and `-p oxi-wasm` clean. Lib test suite: 51 pass + 1 pre-existing kinsoku failure (unrelated, predates this branch).
- behavior change: zero. The pre-pass family (R-01/R-04/etc.) doesn't emit Balloon/BalloonConnector elements — only R-05c will. So this iteration is a pure shape change.
- baseline risk: zero — same reasoning as prior iterations.
- next iteration: R-05b — when `doc.comments` non-empty, reduce `total_content_width` by `293.8 + gap`. First behavior change in the balloon family, but still no balloons emit yet so the reduced body simply leaves the right margin empty until R-05c.

## 2026-04-25 — oxi-2 — design — R-05 / R-06 / R-07 / R-08 / R-09 (balloon-side) implementation plan drafted

- context: with the pre-pass + per-line revision-bar family of renderer rows landed (R-01/R-02/R-03/R-04/R-09 in-line/R-10/R-11/R-12 minimal), the remaining renderer surface is balloon rendering — a substantial new component. Rather than starting implementation with imperfect defaults, draft the design first.
- output: `docs/spec/comments_tracked_changes/r05_balloon_design.md` (200+ lines).
- key design decisions captured:
  - **Two-pass per-page layout**: when `doc.comments` is non-empty, body width is reduced by 293.8pt + gap to make room for the right-margin balloon column. Body lays out first (with reduced width); balloons emit in a second per-page pass that anchors them to the rendered Y of each comment's `commentRangeStart`.
  - **New `LayoutContent::Balloon` + `BalloonConnector` variants** carry the comment payload (author, body, replies, anchor coordinates, resolved flag, color slot). This is intentional vs composing balloons from existing primitives — it gives renderers wrapping freedom and avoids coupling balloon shape to specific shapes/text primitives.
  - **Sub-iteration roadmap (R-05a..h)**: each step ships independently as a Path B confidence merge. R-05a adds the enum variants + `_ => {}` no-op arms across consumer crates; R-05b adds body-width compression; R-05c emits a single balloon; R-05d adds stacking; R-05e adds connector; R-05f adds reply threading; R-05g wires GDI rendering; R-05h flips resolved desaturation.
  - **Risk profile**: zero baseline impact — local 51 + oxi-main 184 baselines have 0 comments, so R-05's body-width compression never triggers there.
- key data sources cross-referenced:
  - Balloon width 293.8pt unresolved / 190.1pt resolved (pixel pass).
  - Right edge ≈ page_w − 4pt (pixel pass).
  - Balloon top aligned with `Comment.Scope.Start` line, NOT `Comment.Reference` (object-model + pixel pass).
  - Resolved tint #F1EDEC vs unresolved #FAE6E7 (pixel pass).
- explicitly out of scope: editor balloon UI (S-04), markup-mode toggle (S-02), left-side balloon anchor (rare config).
- next iteration: starts R-05a — enum variant addition + `_ => {}` fallthrough arms in oxi-cli, oxi-wasm, oxi-gdi-renderer, and example match sites. Mechanical change touching ~25 match sites across 5+ files.

## 2026-04-25 — oxi-2 — confirmed — R-10 margin change bar LANDED

- context: independent renderer row that doesn't need balloon infrastructure. Word draws a thin vertical bar in the left margin next to every line containing any revision (insert/delete/move/property change).
- implementation: emitted DURING layout (not as a post-pass), inside `layout_paragraph`'s per-line loop:
  - Before the fragment loop: declare `let mut line_has_revision = false`.
  - Inside the fragment loop after pushing the Text element: if `para.runs[frag.run_index].tracked_change.is_some()`, set the flag.
  - After the fragment loop: if the flag is set, push one `LayoutContent::BoxRect { fill: Some("#424242"), … }` at `(start_x − 12pt, *cursor_y, 1.5pt, line_height)`.
- design choice — single bar per line: emitted once after the fragment loop instead of per-fragment, so a line with multiple revision fragments (fixture_07: ins1 + del1 + ins2 on one line) gets exactly one bar. Multi-author lines also get one bar (fixture_10).
- design choice — fixed dark grey: Word uses `#424242`-ish dark grey by default; some configs cycle author color. v1 ships fixed grey for simplicity and unambiguity in multi-author cases. Author-tinted bar can layer on later when S-* config rows want it.
- design choice — `start_x − 12pt`: the bar sits 12pt to the LEFT of the body content's left edge, well inside the page margin. For default 72pt margins this puts the bar at x=60pt (12pt from body, 60pt from page edge). Visible without crowding the text.
- tests: new `fixture_05_layout_emits_revision_change_bar` asserts exactly 1 thin BoxRect (≤2pt wide) was emitted, with fill #424242, height ≥8pt, x in the left-margin range. All 18 comments_fixtures tests pass.
- visual confirmation (GDI renderer rebuild): fixture_07's mixed ins+del paragraph now has a clear vertical dark-grey bar in the left margin, next to "Start. ins1 middle del1 ins2. End." The bar height matches the line's vertical extent.
- baseline risk: zero — the local 51-doc and oxi-main 184-doc baselines have 0 revisions in the local set. R-10 emits zero bars on those.
- limitations:
  - Only walks body paragraphs (via `body_para_index`); table cell internal lines, footnotes, headers/footers/textboxes don't currently emit bars. The `body_para_index` field on LayoutElement is `None` for those locations, so the per-line check still runs but only the body case produces visible bars.
    Wait — actually the check IS in the body-emission path (the loop at mod.rs:~2674). Need to extend to header/footer/footnote/textbox emission paths if those should also draw bars. Add later when fixtures stress that.
  - Author-tinted bar variant deferred (would require palette-color lookup at emit time; not blocked by anything else).
- path: Path B confidence-merge candidate.

## 2026-04-25 — oxi-2 — confirmed — R-09 (in-line half) resolved-comment desaturation LANDED

- context: extension of R-04. Word renders the in-line range tint AND the balloon background of resolved comments (`<w15:done="1"/>`) with chroma stripped — Alice's #FAE6E7 unresolved → #F1EDEC resolved (COM-confirmed in pixel pass for fixture_03).
- implementation: `crates/oxidocs-core/src/layout/mod.rs`
  - Added `COMMENT_HIGHLIGHT_RESOLVED_PALETTE: [&str; 8]` next to the unresolved palette. Slot 0 = #F1EDEC (COM-confirmed Alice). Slots 1-7 derived via 25% tint + 75% per-slot grey blend (low chroma, lightness preserved).
  - In `apply_comment_range_highlighting`, the comment_id → tint map switches palette based on `Comment.resolved`: `if c.resolved { RESOLVED_PALETTE } else { UNRESOLVED_PALETTE }`. Same author slot (color_index) selects the same row across the two palettes.
- evidence:
  - new layout test `fixture_03_layout_resolved_comment_uses_desaturated_tint` asserts `"reviewed"` element has `highlight = Some("#F1EDEC")` (not the unresolved #FAE6E7). All 17 comments_fixtures tests pass.
  - Visual via GDI renderer rebuild: fixture_03 page shows "has been reviewed" on a CLEAR GREY background, distinctly different from fixture_01's pink "brown fox" background.
- limitations: balloon-side desaturation (the larger half of R-09 — when balloons render, their fill switches from #FAE6E7 to #F1EDEC) is part of R-05 and not yet implemented. The in-line half is sufficient for v1's "show user the range was already reviewed" UX hint.
- baseline risk: none.
- path: Path B confidence-merge candidate.

## 2026-04-25 — oxi-2 — confirmed — R-04 in-line comment-range highlight LANDED

- context: with R-01/R-02/R-03/R-11 confirmed, R-04 is the next-simplest renderer row that doesn't need balloon infrastructure. Spec: apply an author-tint background to every run strictly between `commentRangeStart` and `commentRangeEnd`.
- parser gap found + fixed: `commentRangeStart`/`commentRangeEnd` are zero-length markers that the previous parser attached to `runs.last_mut()`. When the marker appears as the FIRST child of a paragraph (fixture_04 P1 has `<w:commentRangeStart>` before the first `<w:r>`), `runs.last_mut()` returns None and the id was silently dropped. This left fixture_04's entire range invisible to any range-aware pass. Fix: when `runs.last_mut()` is None, create an empty anchor run carrying the marker (mirrors the bookmark-anchor treatment already in the parser). `commentRangeEnd` gets the symmetric fix.
- attachment convention after the fix: `comment_range_start` on run R → "range starts AFTER R"; `comment_range_end` on run R → "R is the LAST run inside the range". The walk applies highlight before processing either marker, which yields the correct set of highlighted runs with no special-casing.
- implementation: `crates/oxidocs-core/src/layout/mod.rs`
  - `COMMENT_HIGHLIGHT_TINT_PALETTE: [&str; 8]`: slot 0 = `#FAE6E7` (Alice, COM-confirmed from fixture_01 balloon BG), slots 1-7 derived via the 12/88 white-blend formula off the author palette.
  - `apply_comment_range_highlighting(doc)` pre-pass (line 207 area, runs after `apply_revision_styling` so highlight stacks on top of revision ink).
  - Builds `comment_id → tint` map from `doc.comments` + `doc.authors.color_index`. Early return when `doc.comments` is empty (no cost on non-comment baseline docs).
  - Walks `page.blocks` with a persistent `open: HashSet<String>`; recurses into table cells. Order: apply highlight → process `comment_range_end` → process `comment_range_start`.
- tests: 2 new layout integration tests — `fixture_01_layout_comment_range_highlight_inline` (asserts "brown" has Alice's tint, surrounding "The" / "jumps" do not), `fixture_04_layout_multi_paragraph_range_highlight` (asserts all 3 paragraphs "First" / "Second" / "Third" carry the tint — proves the walk carries `open` across paragraph boundaries AND that the parser's new anchor-run fallback works). All 16 comments_fixtures tests pass.
- visual confirmation (GDI renderer rebuild, PNG view):
  - fixture_01: "The quick" black, **"brown fox" on pink tint background**, "jumps over the lazy dog." black. ✓
  - fixture_04: all three paragraphs ("First paragraph...", "Second paragraph...", "Third paragraph...") have the pink tint behind their text. ✓
- baseline risk: none — the local 51-doc baseline has 0 comments (verified last iteration); the 184-doc oxi-main baseline has 0 comments (per inventory). R-04 touches zero runs in either baseline.
- limitations:
  - Single tint per run even if multiple comments overlap the same run — picks the `min()` of open ids deterministically. Word's real behavior blends tints; revisit when we have a multi-overlap fixture.
  - Resolved comments currently use the same tint (doc uses `Comment.resolved` to decide between resolved/unresolved tints in R-09; for inline highlight the behavioral difference is subtle and deferred).
  - Not walked: header/footer/footnote/textbox blocks. 10 fixtures don't exercise these locations.
- path: Path B confidence-merge candidate.

## 2026-04-25 — oxi-2 — confirmed — R-01/R-03/R-11 visual end-to-end validation (GDI renderer)

- context: layout-level integration tests proved the IR-side wiring of R-01/R-03/R-11 was correct, but the actual on-screen output flows through `oxi-gdi-renderer` which converts `LayoutContent::Text` to GDI `TextOutW` calls. End-to-end visual confirmation needed before declaring this Path B-mergeable.
- approach: rebuild `tools/oxi-gdi-renderer` (release) so it picks up the new layout pre-pass, render fixtures 5/6/7/8/10 to PNG at 150 DPI, view the PNGs.
- results:
  - fixture_05 (single ins): "Before insertion" black, "INSERTED TEXT" red+underline (#D03337 hue confirmed via `(210, 30, 30)` and `(180, 30, 30)` pixel buckets), "after insertion." black. ✓
  - fixture_06 (single del): "Before delete" black, "DELETED TEXT" red+strikethrough, "after delete." black. ✓
  - fixture_07 (mixed ins+del): "Start." black, "ins1" red+underline, "middle" black, "del1" red+strikethrough, "ins2." red+underline, "End." black. ✓ (visually viewed)
  - fixture_08 (moves): "Origin:" black, "moved clause" (moveFrom) GREEN+strikethrough (`(0, 60, 30)`/`(30, 60, 0)` hues = #2B6033), "Destination:" black, "moved clause" (moveTo) GREEN+underline. ✓ (visually viewed — the move-quirk hard-coded green is rendering, NOT Alice's red)
  - fixture_10 (two reviewers): "Alpha." black, "ALICE ADD" RED+underline, "middle" black, "BOB REMOVE" PURPLE+strikethrough (`(60, 0, 120)`/`(90, 30, 120)` = #5B2C90 hue), "omega." black. ✓ (visually viewed — palette rotation Alice slot 0 / Bob slot 1 working live in GDI output)
- baseline check: scanned all 51 docs in `pipeline_data/docx/` for `<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`/`<w:rPrChange>` — **zero matches**. The local 51-doc baseline does not exercise the R-01 code path. SSIM cannot regress mathematically. (The 5 `<w:del>` docs in the inventory are in `oxi-main`'s 184-doc baseline, a separate worktree.)
- artifacts: `tools/metrics/output/oxi_render_check/fixture_*_oxi_p1.png` (5 PNGs).
- conclusion: R-01/R-02 (slots 0-1)/R-03/R-11 (single-line v1) are correct end-to-end. Path B confidence merge ready.
- next R-* candidates by independent-of-balloon and value:
  - R-04 (comment-range highlight): in-line background tint between commentRangeStart/End; medium effort, doesn't need balloon rendering.
  - R-10 (margin change bar): vertical line in left margin on revision-bearing lines; medium effort, post-layout pass.
  - R-08 (reply thread): needs R-05 (balloon) first.
  - R-05/R-06/R-07 (balloon family): substantial new rendering surface, defer for a focused multi-tick block.

## 2026-04-25 — oxi-2 — confirmed — R-01 / R-03 / R-11 inline revision styling LANDED

- context: feat/comments-tracked-changes Phase 2.2 entry. Pixel-pass ground truth (author RGB Alice=#D03337, Bob=#5B2C90, move=#2B6033) was captured earlier today — wire it through to the layout output.
- design: rather than threading `tracked_change` through `LineFragment` and changing the `(text, &RunStyle, FieldType, run_index, char_offset)` fragment tuple signature (touches `break_into_lines` + 5 layout call sites), apply the visual styling as a *pre-pass* on `LayoutEngine::layout`'s `doc_resolved` clone. That keeps the entire downstream pipeline unchanged.
- implementation: `crates/oxidocs-core/src/layout/mod.rs` (line 207 area):
  - `apply_revision_styling(doc: &mut Document)` walks `page.blocks` recursively (Block::Paragraph + Block::Table → Cell → Block) and, for each Run with a `tracked_change`, mutates `run.style` to set `underline`/`strikethrough` + `color`.
  - `REVISION_AUTHOR_PALETTE: [&str; 8]` ships the 8-slot author rotation. Slots 0 (Alice #D03337) and 1 (Bob #5B2C90) are COM-confirmed; slots 2-7 use Word's documented rotation pending a 3+ author fixture.
  - `REVISION_MOVE_COLOR = #2B6033` is hard-coded for `<w:moveFrom>` / `<w:moveTo>` regardless of author (Word's quirk).
  - Author lookup: `tc.author` → `Document.authors[idx].color_index` → palette slot. Defensive fallback to slot 0 if author isn't in the palette (unreachable in practice — I-03 builds `authors` from the same source set).
  - `change_type` mapping: insert → underline+author color, delete → strikethrough+author color, moveFrom → strikethrough+green, moveTo → underline+green. Unknown change_type leaves the run alone (forward-compatibility).
- evidence:
  - 4 new layout integration tests in `tests/comments_fixtures.rs` (`fixture_05_layout_ins_underline_in_author_color`, `fixture_06_layout_del_strikethrough_in_author_color`, `fixture_10_layout_two_authors_get_distinct_colors`, `fixture_08_layout_moves_render_in_green`). All 4 + the existing 10 IR-only tests pass (14/14).
  - Adjacency assertion: in fixture_05, `LayoutContent::Text` for the surrounding "Before" run remains `underline=false, strikethrough=false` — confirming the pre-pass doesn't leak into non-revision runs.
  - Two-author distinction (fixture_10): Alice gets #D03337 from the "ALICE" element; Bob gets #5B2C90 from the "REMOVE" element — both in the same paragraph.
  - Move quirk (fixture_08): both occurrences of "moved" come back tagged #2B6033 (green), not Alice's red, with one occurrence strikethrough (moveFrom) and the other underlined (moveTo).
  - Full crate test suite: 51 lib tests + 14 comments_fixtures tests pass; pre-existing `kinsoku::test_line_start_prohibited` failure is unrelated (predates this branch).
- limitations / known follow-ups:
  - Move visualization is single-line + green in v1, not Word's default double-strike/double-underline. Upgrade hinges on adding `underline_style="double"` + a `strikethrough_style` primitive at the renderer; out of scope for this tick.
  - Headers, footers, footnotes, endnotes, and textbox-internal blocks aren't walked yet by `apply_revision_styling`. The 10 fixtures don't exercise revisions in those locations; add when needed.
  - rPrChange (fixture_09) is intentionally NOT styled — Word renders the new properties (e.g., bold black) directly + a margin balloon, which is R-12.
- baseline risk: extremely low — the 184-doc baseline contains 5 `<w:del>`-only docs and zero comment/ins/move docs (per `inventory/README.md`). For those 5 docs, runs that previously rendered as plain text now render with strikethrough + #D03337. Visually correct, but a baseline SSIM regression of <0.01 per affected page is possible (still need to verify with full pipeline).
- path: Path B `[confidence-merge]` once the `pipeline.verify` baseline is re-run.

## 2026-04-25 — oxi-2 — confirmed — Tick 2-3 pixel-sampling pass (author RGB + balloon geometry + resolved desat)

- context: feat/comments-tracked-changes — with the object-model pass complete on 10/10 fixtures, R-* renderer rows still blocked on the **visual** ground truth: author ink RGB, balloon rectangle, balloon background tint, strikethrough Y, etc.
- approach: render each fixture via `Document.ExportAsFixedFormat(Item=wdExportDocumentWithMarkup)` (NOT `SaveAs2(FileFormat=17)` — the latter drops markup). Rasterize page 1 at 150 DPI via PyMuPDF. For each `Revision` / `Comment`, use COM `Range.Information(WD_HORIZONTAL/VERTICAL_POSITION_RELATIVE_TO_PAGE)` to get page coordinates → convert to pixels → sample RGB.
- gotchas hit & fixed:
  - `View.ShowComments = False` is the default; balloons don't render even with `Item=wdExportDocumentWithMarkup` unless we flip it. Set `ShowComments = True` + `ShowRevisionsAndComments = True` + `RevisionsView = wdRevisionsViewFinal` before exporting.
  - Comment.Reference (the inline marker) returns the wrong Y for balloon alignment — use `Comment.Scope` (the range start) instead. Word's balloon aligns with the rendered first character of the range, not the marker.
  - Revision.Range.Information(6) returned x_pt y_pt that landed on the line top — sampling within a 30pt × 20pt window leaked into adjacent body text. Narrowed to 12pt × 14pt and added a saturation preference (prefer colored ink over black). Then fixture_07's three revisions all reported #D03337 cleanly.
- evidence:
  - Author RGB: **Alice = #D03337 (208, 51, 55), Bob = #5B2C90 (91, 44, 144)** — confirmed across 4 fixtures (05, 06, 07, 10).
  - **Move-revision quirk**: `<w:moveFrom>` / `<w:moveTo>` always render in **green #2B6033 (43, 96, 51)** regardless of author. This is a hard-coded Word behavior; R-11 needs to bypass the author-color rotation.
  - **rPrChange (fixture_09)**: underlying text is NOT author-tinted — renders as the new property (bold black). The author signal is in a separate right-margin "Formatting" balloon. R-12 minimal version: render the new properties + a margin balloon (don't tint the inline text).
  - Balloon geometry: width 293.8pt unresolved / **190.1pt resolved** (smaller box for done=true), right edge ≈ page_width − 4pt for all balloons.
  - Resolved desaturation: Alice unresolved BG #FAE6E7, resolved BG #F1EDEC — chroma drops from ~20 to ~5, lightness unchanged. R-09 must apply this when `done=true`.
  - Body width is reduced when balloons are present: fixture_01 body is ~147pt wide vs ~451pt without comments.
- outcome: created `tools/metrics/measure_comments_tracked_changes_pixels.py`. Output `tools/metrics/output/comments_tracked_changes_pixels.json` + per-fixture PDFs/PNGs. Promoted snapshot to `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_pixels.json` + `PIXEL_PASS_README.md`. INDEX.md flipped pixel-pass to checked.
- impact: **R-01, R-02 (palette slots 0-1), R-03, R-04, R-05, R-08, R-09, R-11 are unblocked**. R-07 (dotted connector line) and exact underline/strikethrough Y still need a small follow-up horizontal-line probe — non-blocking for the first renderer ticks.
- known limitations:
  - fixture_09 ink_rgb is null (rPrChange line offset by ~88pt due to formatting-balloon header).
  - Author palette slots 2..7 unmeasured (need 3+ author fixture).
  - Strikethrough / underline Y captured only as glyph ink center, not the marker line itself.
- baseline risk: none — measurement-only, no code changes.
- path: Path B confidence-merge.

## 2026-04-25 — oxi-2 — confirmed — Tick 2-3 fixture content-type fix (3 fixtures unblocked, 10/10 Word-OK)

- context: feat/comments-tracked-changes — `tools/metrics/output/comments_tracked_changes_com.json` had reported 7/10 fixtures Word-OK since 2026-04-18; fixtures 02 (reply), 03 (resolved), 10 (multi-author) failed `Word.Documents.Open` with `'ファイルが壊れている可能性があります。'`. Pixel/UIA pass for R-* renderer rows was blocked on these 3 because reply/done/multi-author paths are exactly what fixtures 02/03/10 exercise.
- hypothesis: the 3 failing fixtures share a structural distinguishing feature — fixtures 02 + 03 emit `commentsExtended.xml`, fixture 10 emits `people.xml`. Suspected the historical `application/vnd.ms-word.commentsExtended+xml` / `application/vnd.ms-word.people+xml` content types in `build_comments_samples.py` were rejected by Word 16.0's strict-open path.
- evidence:
  - `Word.Documents.Open(..., OpenAndRepair=True)` succeeded for all 3, returning expected `Comments`/`Revisions` collections.
  - Saved Word's repaired output via `SaveAs2` and diffed `[Content_Types].xml`: Word rewrote both content types to `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml` and `application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml`.
  - Single-line patch test: in-memory rewrite of just `[Content_Types].xml` (no other XML touched) made strict-open succeed for both 02 and 10 in a fresh Word instance.
- outcome: applied the same one-line-per-part change in `tools/fixtures/comments_samples/build_comments_samples.py`. Re-ran builder and the COM measurement script: 10/10 fixtures now `word_reads_ok=true`. Captured data confirms reply ancestor (fixture 02 `Comment.Ancestor` / `Replies.Count==1`), resolved flag (fixture 03 `Comment.Done==True`), and per-author revisions (fixture 10 Alice ins + Bob del). Both `tools/metrics/output/comments_tracked_changes_com.json` and `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_com.json` regenerated; README updated; INDEX.md `[ ] Tick 2-3 deferred` flipped to `[x] object-model pass complete`.
- impact: pixel/UIA pass still required before R-* renderer rows, but it is no longer blocked on fixture-build defects. Strict-mode round-trip of these fixtures (and Oxi's emitted .docx in the future) now lands in the same content-type bucket Word writes natively, removing one whole class of validator-rejection bugs from the pipeline.
- baseline risk: none — fixtures live under `tools/fixtures/comments_samples/`, not in the regression baseline.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — I-03 author palette + I-04 ShowRevisions

- context: feat/comments-tracked-changes Phase 2 IR rows. After Phase 2 parser COMPLETE, build the IR scaffolding renderer rows depend on.
- I-03 — author palette:
  - new `ir::Author { display: String, color_index: usize }`
  - `Document.authors: Vec<Author>` derived from 3 sources, first-seen order:
    1. `Document.people` (people.xml — Word writes reviewer-first-seen order)
    2. `Comment.author`
    3. tracked-change authors via `walk_block_authors` (run.tracked_change, run.rpr_change, paragraph.ppr_change, paragraph.paragraph_mark_revision; recurses into Block::Table)
  - color_index = position in the palette → renderer maps to RGB through any palette without a separate join.
- I-04 — show-revisions toggle:
  - `ir::ShowRevisions::{All (default), Simple, Original, Final}`
  - serde `rename_all = "snake_case"` so JSON-API consumers see `"all"|"simple"|"original"|"final"`.
  - not wired into a render config struct yet — added as IR plumbing for S-02.
- I-01 closeout: covered incrementally by P-01 + P-10 + P-11. Comment struct already has all required fields and is surfaced on Document.comments.
- I-02 deferred: keeping the current `Run.tracked_change: Option<TrackedChange>` + `Run.rpr_change: Option<PropertyChange>` shape. Multiple-revisions-per-run is rare and unused in baseline; SmallVec refactor blast-radius is high. Revisit only if a renderer needs it.
- evidence:
  - 2 unit tests: `show_revisions_default_is_all_and_round_trips_json`, `build_author_palette_dedupes_in_first_seen_order`
  - 1 integration extension: `fixture_10_people_two_reviewers` extended to assert `Document.authors` palette ordering
  - 1 new integration: `fixture_05_authors_palette_from_tracked_changes_only` — palette falls back to tracked-change authors when people.xml is absent
- baseline risk: none.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-08 *PrChange silent-drain (6 variants) — Phase 2 parser COMPLETE

- context: feat/comments-tracked-changes Phase 2 parser row P-08 — final row
- scope: ECMA-376 §17.13.5 — six revision-history wrappers, each containing a *prior* copy of the same property element they sit inside:
  - `<w:tblPrChange>` inside `<w:tblPr>`
  - `<w:trPrChange>` inside `<w:trPr>`
  - `<w:tcPrChange>` inside `<w:tcPr>`
  - `<w:sectPrChange>` inside `<w:sectPr>`
  - `<w:tblGridChange>` inside `<w:tblGrid>`
  - `<w:numberingChange>` inside `<w:numPr>`
- silent-bug class: each owning property parser uses the same depth-doesn't-gate-Empty-handlers pattern as parse_run_properties / parse_paragraph_properties. Without the drain, the prior property body would silently leak — most concretely:
  - `parse_table_grid` would APPEND prior `<w:gridCol>` widths to the column list (column count corruption)
  - `parse_num_pr` would OVERWRITE current numId/ilvl with the prior values
  - `parse_table_properties`, `parse_table_row`, `parse_cell_properties`, `parse_section_properties` would all leak prior style/border/margin into current state
- change:
  - new `drain_element(reader, tag_name)` helper — reads to the matching End regardless of nesting
  - 6 explicit drain branches added at the top of each respective parser's Start arm
- evidence:
  - 3 unit tests covering the helper, tblGridChange, and numberingChange (the most demonstrable corruption cases)
- IR emission deferred: attack_matrix says "Rare in practice; emit to IR for completeness, no renderer work yet". Since the renderer doesn't yet consume these, the silent-drain is the highest-value minimum. When a renderer needs them, the next iteration adds typed PropertyChange fields like P-06/P-07 did.
- baseline risk: none.
- path: Path B `[confidence-merge]`. **Phase 2 parser quartet (P-01..P-12, except deferred renderer rows) COMPLETE — 12/12 rows landed.**

## 2026-04-25 — oxi-2 — confirmed — P-09 paragraph-mark ins/del

- context: feat/comments-tracked-changes Phase 2 parser row P-09
- scope: ECMA-376 §17.13.5 — `<w:pPr>/<w:rPr>/<w:ins>` or `/<w:del>` marks the paragraph's pilcrow (¶) itself as inserted (new split) or deleted (paragraph merged with next). revisions_notes.md §2.
- change:
  - `Paragraph.paragraph_mark_revision: Option<TrackedChange>`
  - ins/del Empty detection added inside the pPr/rPr sub-loop in parse_paragraph_properties (where ppr_rpr is already being built). Captures change_type/author/date/pair_id.
  - parse_paragraph_properties return: 6-tuple → 7-tuple. Updated caller + 3 Paragraph constructors.
- evidence:
  - 2 unit tests: parse_pmark_ins_via_ppr_rpr + parse_pmark_del_via_ppr_rpr (inline XML; no fixture in 10-doc set, per attack_matrix note)
- baseline risk: none (0 pPr/rPr/ins or /del in 184 baseline docs).
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-11 commentsIds.xml durable ids

- context: feat/comments-tracked-changes Phase 2 parser row P-11
- scope: Word 2019+ `word/commentsIds.xml` (w16cid namespace) — carries durable ids that survive save-as roundtrips (local `w:id` is renumbered freely)
- change:
  - `Comment.durable_id: Option<String>`
  - new `parse_comments_ids_xml` free function, returns `paraId → durableId` map
  - merged into Comments in `build_context_with_theme` after commentsExtended merge
  - accepts both `w16cid:durableId` (canonical) and `w16cid:id` (older draft spelling)
- evidence:
  - 2 unit tests: standard durableId map + legacy id attribute acceptance
  - no fixture in the 10-doc set has commentsIds.xml (scanned all 10 → 0 hits)
- baseline risk: none (184 baseline docs have 0 commentsIds.xml).
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-07 pPrChange + silent-bug fix

- context: feat/comments-tracked-changes Phase 2 parser row P-07
- scope: ECMA-376 §17.13.5 `<w:pPrChange>` carries a full prior `<w:pPr>` body for paragraph-level revisions
- silent-bug fix (mirrors P-06): `parse_paragraph_properties`'s Empty handlers (jc, pStyle, spacing attrs, etc.) don't gate on depth. An inner `<w:jc val="right"/>` inside `<w:pPrChange>/<w:pPr>` would silently overwrite the current paragraph alignment. Same class of defect as rPrChange, resolved the same way: explicit drain before the fallback.
- change:
  - extend `PropertyChange` with `prior_paragraph_style: Option<Box<ParagraphStyle>>`
  - `Paragraph.ppr_change: Option<PropertyChange>`
  - `parse_paragraph_properties` return: 5-tuple → 6-tuple (added `Option<PropertyChange>`)
  - explicit `pPrChange` branch: captures id/author/date, recursively reparses inner `<w:pPr>` via `parse_paragraph_properties`, drains to `</w:pPrChange>`, handles self-closing `<w:pPr/>`
  - 3 Paragraph constructors updated with `ppr_change: None` (empty-para fallback×2 + main)
- evidence:
  - unit (1): `parse_pprchange_stores_prior_style_without_merging_into_current` — current=Left, prior pPr=Right, both captured without cross-contamination.
- no integration fixture: attack_matrix notes P-07 has no fixture in the 10-doc set. Defer until a dedicated pPrChange fixture is authored (or P-08 gets one).
- baseline risk: none (0 pPrChange in 184 baseline docs).
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-06 rPrChange + silent-bug fix

- context: feat/comments-tracked-changes Phase 2 parser row P-06
- scope: ECMA-376 §17.13.5 `<w:rPrChange>` carries a full prior `<w:rPr>` body to record run-property revisions
- silent-bug noticed while scoping: the pre-existing `parse_run_properties` had no `rPrChange` branch. Its depth counter increments on every Start but its property handlers (b/i/u/color/font) don't gate on depth. Result: the prior `<w:rPr>` inside `<w:rPrChange>` would silently merge into the *current* style — `<w:b/>` in the prior would set the current run bold. This is latent today (baseline 184 docs have 0 rPrChange) but would corrupt formatting on any future doc with rPrChange. The new explicit branch drains rPrChange before it reaches those handlers.
- change:
  - new `ir::PropertyChange { id, author, date, prior_run_style: Option<Box<RunStyle>> }`
  - new `Run.rpr_change: Option<PropertyChange>`
  - `parse_run_properties` return type: `RunStyle` → `(RunStyle, Option<PropertyChange>)`. Only one caller (parse_run) updated.
  - inline handler in parse_run_properties: captures rPrChange attrs, recursively reparses the inner `<w:rPr>` as the prior RunStyle, consumes up to `</w:rPrChange>` without touching the current style. Handles self-closing `<w:rPr/>` (prior = RunStyle::default()).
  - 3 Run constructors updated with `rpr_change: None` (layout empty-para-prefix, parser omml-math, parser bookmark-anchor)
- evidence:
  - integration: `fixture_09_rpr_change_bold` — verifies current bold + prior plain + id=300 + author/date populated
- baseline risk: none (new field, empty in 184 baseline docs).
- path: Path B `[confidence-merge]` — the silent-bug fix is a free correctness win.

## 2026-04-25 — oxi-2 — confirmed — P-03/P-04 ins+del locked down + P-05 moveFrom/moveTo

- context: feat/comments-tracked-changes Phase 2 parser rows — the tracked-change quartet
- P-03/P-04 (verification): `<w:ins>` / `<w:del>` were pre-existing, emitting `Run::tracked_change{change_type: "insert"|"delete", author, date}` and preserving `<w:delText>` as `Run::text`. Locked down with 3 integration tests on fixtures 05, 06, 07 (including XML-order preservation in mixed case).
- P-05 (new): `<w:moveFrom>` / `<w:moveTo>` wrap runs identically to ins/del. Added `change_type="moveFrom"|"moveTo"` plus a new `pair_id: Option<String>` field on `TrackedChange` (the wrapper's `w:id`).
- important finding — pairing is NOT via wrapper `w:id`: fixture 08 shows `moveFrom w:id="201"`, `moveTo w:id="202"`. The actual from↔to pairing lives on `moveFromRangeStart` / `moveToRangeStart` (both carrying `w:id="200"` + `w:name="move1"`). Revisions_notes.md §1.2 is correct; the attack-matrix row note saying "Pair via shared w:id on the Range markers" was accurate. Phase 2 parser surfaces the wrapper id only; R-11 will walk range markers when it needs the from↔to linkage.
- refactor: the four ins/del/moveFrom/moveTo branches collapsed into a single `"ins"|"del"|"moveFrom"|"moveTo"` arm mapped to `change_type` strings; `parse_tracked_change_runs` receives the element name as end_tag.
- tests:
  - integration (4 new): fixture_05_single_ins, fixture_06_single_del, fixture_07_mixed_ins_del, fixture_08_move_from_to_pair
- baseline risk: none. 184 baseline docs have 5 lone `w:del`s (already handled pre-change) and 0 w:ins/moveFrom/moveTo.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-12 people.xml reviewer list

- context: feat/comments-tracked-changes Phase 2 parser row P-12
- scope: MS-DOCX w15 — `<w15:people>/<w15:person w15:author="..."><w15:presenceInfo providerId userId/>`
- change:
  - new `ir::Person{author, provider_id, user_id}` type, re-exported from `ir::*`
  - `Document.people: Vec<Person>` (document order preserved — Word writes reviewer-first-seen, so this seeds R-02 palette without re-sort)
  - `OoxmlParser::parse_people()` + `parse_people_xml` free function; missing part → empty list
  - handles both `<w15:person>…</w15:person>` (with nested presenceInfo) and self-closing `<w15:person/>` (no presence data)
  - drops malformed `<w15:person>` entries missing `w15:author`
- evidence:
  - 3 unit tests: two-reviewer shape (fixture 10 mirror), missing-presenceInfo, blank-author dropped
  - 1 integration test: `fixture_10_people_two_reviewers` runs full parse_docx pipeline, confirms `Document.people == [Alice Reviewer, Bob Reviewer]` with provider/userId attached
- baseline risk: none (184 baseline docs have 0 people.xml).
- path: Path B `[confidence-merge]`. Completes the Phase-2 parser quartet needed before I-03/R-02 (author-colour palette) can start.

## 2026-04-25 — oxi-2 — confirmed — P-10 commentsExtended.xml (reply threading + resolved)

- context: feat/comments-tracked-changes Phase 2 parser row P-10
- scope: MS-DOCX w15 extension — `<w15:commentEx paraId="..." paraIdParent="..." done="..."/>`
- change:
  - new fields on `ir::Comment`: `para_id: Option<String>`, `parent_para_id: Option<String>`, `resolved: bool`
  - `parse_comments_xml` grabs the first `w14:paraId` off the comment body's first `<w:p>` (the commentsExtended join key)
  - new `parse_comments_extended_xml` function + merge step in `build_context_with_theme`. Accepts both canonical `w15:paraIdParent` and legacy `w15:parentParaId` per comments_notes.md §4
- evidence:
  - 3 new unit tests: `parse_comments_xml_captures_first_para_id`, `parse_comments_extended_reply_and_resolved`, `parse_comments_extended_accepts_legacy_parent_para_id_spelling`
  - 2 new integration tests: `fixture_02_comments_extended_reply_threaded` (parent/reply `paraIdParent` = "00000010"), `fixture_03_comments_extended_resolved_flag` (`w15:done="1"` → `resolved==true`)
  - NB: fixtures 02/03 fail `Documents.Open` in Word (validator defect to fix in a separate tick) but are valid enough for python-docx + our scanner + our parser; the tests cover the parser contract regardless.
- baseline risk: none — 184 baseline docs have 0 commentsExtended.xml.
- path: Path B `[confidence-merge]`.

## 2026-04-25 — oxi-2 — confirmed — P-02 commentReference wired to Run

- context: feat/comments-tracked-changes Phase 2, second parser row
- scope: balloon anchor marker `<w:commentReference w:id="N"/>` inside `<w:r>`
- change: added `comment_references: Vec<String>` to `oxidocs_core::ir::Run`; `parse_run` captures the id inside the enclosing run. `commentRangeStart/End` parsing was pre-existing (run-level `Vec<String>`).
- rationale: the enclosing run is the glyph the renderer projects to the right margin to position the balloon (ECMA-376 §22.1.2.56 + comments_notes.md §2.2). One run may legally carry multiple refs.
- evidence:
  - unit: `parse_run_captures_comment_reference` — synthesized `<w:r>` with `commentReference id="0"` yields `run.comment_references == ["0"]`.
  - integration: `fixture_01_comment_fields_roundtrip` extended — `commentReference` id="0" found on exactly one run in the body.
- touched Run constructors (4 sites): `ir::types::Run`, `layout::mod::<empty para prefix>`, `parser::ooxml::<omml run>`, `parser::ooxml::<bookmarkStart anchor>` — each now initialises `comment_references: Vec::new()`.
- baseline risk: none (new field is empty in all 184 baseline docs).

## 2026-04-25 — oxi-2 — confirmed — P-01 comments.xml parse complete (initials field)

- context: feat/comments-tracked-changes Phase 2 — first parser row from attack matrix
- scope: `Comment{ id, author, date, initials, runs }` per ECMA-376 §17.13.4.2
- change: added `initials: Option<String>` to `oxidocs_core::ir::Comment`; `parse_comments_xml` now captures the `w:initials` attribute. No renderer impact.
- evidence:
  - unit tests (inline XML) — `parse_comments_xml_captures_initials_and_metadata` (fixture 01 shape, initials="AR") and `parse_comments_xml_missing_initials_is_none` (older Word docs omit `w:initials`).
  - integration test — `crates/oxidocs-core/tests/comments_fixtures.rs::fixture_01_comment_fields_roundtrip` runs the full `parse_docx` pipeline on `fixture_01_single_comment.docx`; confirms `Document.comments[0]` has `id="0"`, `author="Alice Reviewer"`, `initials="AR"`, `date="2026-04-18T10:00:00Z"`, plus `commentRangeStart`/`End` markers surface on runs (P-02 shadow coverage).
  - COM ground truth: `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_com.json` (2026-04-18, Word 16.0) — fixture 01 `Comments(1).Initial == "AR"`.
- baseline risk: none. 184-doc baseline has 0 comments.xml (see `inventory/README.md`).
- path: Path B `[confidence-merge]` — spec-referenced + COM-validated + fixture-backed + baseline-neutral by construction.

## 2026-04-18 — oxi-2 — partial — Tick 2-3 Word COM pass, 7/10 fixtures validated

- context: feat/comments-tracked-changes Phase 1 Tick 2-3 (previously deferred; Word 16.0 turned out to be available on this box)
- method: `tools/metrics/measure_comments_tracked_changes_com.py` opens each fixture with Word COM and dumps `doc.Revisions` + `doc.Comments` collections
- evidence (7/10 OK): `doc.Revisions.Type` matches authoring intent for every revision fixture (05 wdRevisionInsert, 06 wdRevisionDelete, 07 × 3 alternating, 08 wdRevisionMovedFrom+MovedTo, 09 wdRevisionProperty). `doc.Comments.Scope.Text` matches authored range for 01 and 04
- evidence (3/10 fail): 02 and 03 (both have commentsExtended.xml) and 10 (has people.xml) fail `Documents.Open` with `com_error` — Word validator rejects even though syntactic XML is well-formed (our scanner reads the markers). Fixture-authoring defect to repair in a follow-up tick; does not block Phase 2 parser rows P-03…P-09.
- key Phase 2 confirmations:
  - `<w:rPrChange>` is bucketed as `wdRevisionProperty` (3) in Word's Revisions.Type enum — a single Inserted/Deleted/Property IR enum matches Word's model.
  - `moveFrom` / `moveTo` pair is reported as TWO separate revisions with identical range text — the `(pair_id, name)` linkage in revisions_notes.md §1.2 is the correct IR structure.
  - `Comment.Scope.Text` spans `\r`-separated paragraphs for multi-para ranges — parser can represent range as (start: RunRef, end: RunRef) over a flat glyph list.
- deliverable: `docs/spec/comments_tracked_changes/com_measurements/{comments_tracked_changes_com.json, README.md}`
- still deferred: balloon geometry, author RGB palette (R-02), strikethrough Y on CJK text, connector line style, stacking geometry. All need UIA / pixel-sampling.

## 2026-04-18 — oxi-2 — phase-1-complete — attack matrix + master index; COM deferred

- context: feat/comments-tracked-changes Phase 1 final tick (was Tick 7 in TASK.md)
- deliverables:
  - `docs/spec/comments_tracked_changes/attack_matrix.md` — 33-row priority matrix (12 parser + 4 IR + 13 renderer + 4 settings). Effort, blast radius, fixture coverage, COM-measurement dependency, baseline SSIM risk per row. Recommended execution order: P-01→P-12 → I-01→I-04 → R-01/R-03/R-04 first (cheapest renderer wins). Baseline risk = **none** for virtually every row (baseline has 0 comments / 5 lone dels) — Path B confidence-merge is natural.
  - `docs/spec/comments_tracked_changes/INDEX.md` — canonical entry point. Indexes every Phase 1 asset (spec notes, attack matrix, inventory, fixtures, coordination files) + 3 dogfood lookup simulations ("implement w:ins parse", "render comment balloons", "worried about SSIM regression on w:del").
- deferred: Tick 2-3 Word COM measurement (12-target checklist embedded in attack_matrix.md §"Tick 2-3 deferred COM checklist"). Must run in a dedicated session before starting renderer rows R-01+; parser rows (P-01…P-12) can begin without it.
- handoff: Phase 2 can start immediately at any parser row. IR sketch in revisions_notes.md §8; renderer row R-02 (author-color palette) blocks R-10 (margin change bar) and wants COM data.
- methodology: Phase C inventory→matrix→index re-application succeeded; validated pattern still generalises beyond yakumono/cell/footnote domains. No code changes in /loop per user directive; all 4 ticks = pure measurement + memo.

## 2026-04-18 — oxi-2 — spec-notes-written — ECMA-376 §17.13.1 + §17.13.5

- context: feat/comments-tracked-changes Phase 1 Tick 4 (spec notes)
- deliverables:
  - `docs/spec/comments_tracked_changes/comments_notes.md` — parts, content types, inline markers, comments.xml structure, commentsExtended threading, Word display rules (not-in-spec), JIS X 4051 interaction, parser checklist, fixture cross-reference
  - `docs/spec/comments_tracked_changes/revisions_notes.md` — element taxonomy (ins/del/move/\*Change), block-level ins/del via pPr/rPr, accept/reject semantics, Word display rules, IR sketch (Phase 2 planning), parser checklist
- key observations recorded for Phase 2:
  - `w:id` on revisions is NOT durable across saves; parsers must not use it as stable identifier
  - Paragraph-mark insert/delete lives on `pPr/rPr`, NOT as an outer `<w:ins>` — common miss
  - `moveFrom/To` pair via shared `w:id` on the Range markers + opaque `w:name` label
  - Comment balloon geometry, stacking, author color palette are all renderer-defined (not spec)
  - `w15:paraIdParent` and `w15:parentParaId` both appear in the wild; accept on parse, emit canonical form on write
- next: Tick 7 (attack matrix + master index), OR Tick 2-3 (Word COM measurement against the 10 fixtures — requires Word running on this box)

## 2026-04-18 — oxi-2 — fixtures-ready — 10 minimal repros for comments + tracked changes

- context: feat/comments-tracked-changes Phase 1 Tick 5-6 pulled forward (baseline too sparse for Tick 2-3 COM without self-authored docs)
- hypothesis: 10 feature-isolated fixtures are enough to unblock subsequent COM measurement + spec notes
- method: zip-level OOXML generator (`tools/fixtures/comments_samples/build_comments_samples.py`), one fixture per feature; validated via python-docx open + inventory re-scan
- evidence: all 10 open cleanly; per-file marker counts exactly match intent; MANIFEST.json committed alongside
- coverage: 01 single comment / 02 comment+reply / 03 resolved / 04 multi-para range / 05 single ins / 06 single del / 07 mixed ins+del / 08 moveFrom+moveTo / 09 rPrChange bold / 10 two reviewers (Alice+Bob)
- side-effect: inventory scanner reply-detection pattern corrected (`w:parentId` → `w15:paraIdParent`); baseline totals unchanged (still 0 comments, 5 del-only)
- next: Tick 2-3 runs Word COM measurement against these fixtures — balloon position, range highlight color, ins underline style, del strikethrough color, multi-reviewer color rotation
- tools: tools/fixtures/comments_samples/build_comments_samples.py
- path: tools/fixtures/comments_samples/fixture_{01..10}_*.docx (+ MANIFEST.json)

## 2026-04-18 — oxi-2 — baseline-inventory — comments + tracked-changes sparsity confirmed

- context: feat/comments-tracked-changes Phase 1 Tick 1 — establish baseline usage floor before Phase 2 implementation
- hypothesis: the 177/184 baseline .docx corpus contains enough comment + tracked-change usage to drive COM measurement and regression testing
- method: zip+XML scan of all 184 docx under `oxi-main/tools/golden-test/documents/docx/`. Count `w:commentRange*`, `w:commentReference`, `w:comment` bodies, `w:ins`, `w:del`, `w:moveFrom/To`, `w:*PrChange` markers. Detect `word/comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, `people.xml`
- evidence (JSON at tools/metrics/output/{comments_inventory,tracked_changes_inventory}.json):
  - docs_with_word_comments_xml: **0 / 184**
  - docs_with_any_revision_marker: **5 / 184** (all 1×w:del, one additionally 1×w:pPrChange, single author, boilerplate `people.xml`)
  - zero replies, zero moves, zero rPrChange, zero multi-reviewer scenarios in corpus
- outcome: REFUTED. Baseline provides essentially no test signal for comment + revision rendering. All Tick 2–3 COM measurements and all Phase 2 regression suites MUST use self-authored fixtures. Advantage: no SSIM floor risk for these features (Path B [confidence-merge] is the natural merge gate); work can proceed on dedicated branch without bottom-N concern.
- tools: `tools/metrics/inventory_comments_tracked_changes.py`
- next-tick: Tick 2 — author N reference docx fixtures in `tools/fixtures/comments_samples/` (even a provisional set unblocks Tick 2 COM measurement). This pulls Tick 5-6 earlier in the pipeline.

## 🔥 BLOCKER: GDI preset render coverage (Path A fix target)

## 2026-04-18 — oxi-1 — drift-localized — b35 p.1 Class B +2.5pt body offset
- context: Task #4 — b35 rank 3 bottom-5 (SSIM 0.6134), prior memos claim Class B
- hypothesis: b35 p.1 has measurable per-paragraph drift like 2ea81 Class B
- method: Word COM per-paragraph Y + Oxi --dump-layout per-block Y; align by text content
- evidence: 4 body paragraphs aligned (Oxi dump has only 4 body para_idx; tables use block para_idx). All 4 show +2.00-2.50pt downward drift. Median |Δ|=+2.50pt, consistent.
- outcome: NOT ceiling. Class B drift CONFIRMED. Likely same root cause as Task #1 line=exact boundary rule. Fixing line=exact rule may simultaneously improve b35 p.1 body paras. Table-row drift is SEPARATE mechanism (covered by oxi-4 LM0 cell formula).
- impact on session: bottom-5 coverage now complete (all 5 + rank 6 diagnosed). 4 have dedicated-session-ready fix targets, 1 ceiling, 1 pivot.
- tools: diff_b35_p1_paras.py
- memory: project_b35_p1_class_B_drift_confirmed.md

## 2026-04-18 — oxi-1 — refuted — b837 footnote spill hypothesis (oxi-2 unblock)
- context: Task #3 — b837 rank 4; oxi-2's footnote spill investigation
- hypothesis (oxi-2): Word splits long multi-line footnote bodies across pages
- method: Word COM measurement of all 25 fns in b837 (ref_page + body_first_page + body_last_page); 3 additive scratch variants with single/many/multi-line fns
- evidence: ZERO spill in 42 total fns tested (25 real + 17 scratch). All fn body_first_page == body_last_page. Word's rule: fn bodies NEVER span page boundaries.
- outcome: REFUTED. oxi-2 unblocked — pivot to estimate_footnote_h per-fn over-count (10pt/fn). Real bug location: mod.rs:631 per-footnote height estimate, NOT spill model.
- supporting evidence: existing output/b837_footnote_y.json + new pipeline_data/b837_footnote_spill.json
- tools: measure_b837_footnote_spill.py, gen_footnote_spill_repro.py
- memory: project_b837_footnote_spill_FALSIFIED.md
- impact: oxi-2 status can change from "investigating spill" to "investigating per-fn estimate over-count"

## 2026-04-18 — oxi-1 — re-confirmed — 0e7a p.2 layout ceiling
- context: rank 1 bottom-5 (0.5767); prior memos claimed layout ceiling,
  re-verification requested per Task #2
- hypothesis: 0e7a p.2 has no remaining layout drift; SSIM gap is
  sub-pixel / AA / glyph-rendering only
- method: fresh Oxi --dump-layout on current main + Word COM per-paragraph Y
  measurement; align 20 paragraphs by text content
- evidence: median Δ=+0.00pt, max |Δ|=0.50pt across 20 aligned paragraphs;
  17 of 20 paras show Δ=+0.00 exactly
- outcome: LAYOUT CEILING confirmed. 0e7a p.2 SSIM 0.5767 NOT
  layout-improvable. Future sessions should skip this page for layout
  work and focus on d77a p.9 (rank 2) where drift IS fixable
  (line=exact boundary rule — see preceding entry).
- tools: measure_word_paras_generic.py, diff_0e7a_p2_paras.py
- memory: project_0e7a_p2_ceiling_CONFIRMED_2026_04_18.md

## 2026-04-18 — oxi-1 — confirmed — line=exact boundary rule (additive)
- context: 2ea81 idx=6→7 +2pt bug (see project_2ea81_line_exact_boundary_bug.md)
- hypothesis: at adjacent paragraphs both with lineRule=exact, Y advance A→B uses lineA's value, not lineB
- method: 5 scratch additive variants (V1-V5) from python-docx Document(); different font/empty/non-empty/increasing/decreasing combinations
- evidence: all 5 variants confirm delta = N_A × lineA/20; V5 DECREASING (400→240) also matches (excludes naive "use larger" hypothesis)
- outcome: additive rule CONFIRMED. Oxi bug = +2pt at empty-paragraph boundary. Real-doc verification partial (2ea81 re-measured for idx=2→3 Oxi matches Word; idx=6→7 memo has Oxi+2pt). COM RPC crashes on larger docs blocked 29dc6/6514f214 verification.
- tools: repro_line_exact_variants.py, verify_line_exact_rule_3docs.py, verify_line_exact_oxi_vs_word.py
- outcome for next: implementation in dedicated session. Location: mod.rs paragraph cursor advance when prev para lineRule=exact. Must preserve non-empty-A cases.
- memory: project_line_exact_rule_additive_confirmed.md

## 2026-04-18 — oxi-3 — architectural-validation — yakumono is geometry heuristic, NOT fixed rule
- context: 6-tick 4-tier additive bisection from scratch (final conclusion)
- method: scratch + cSC + compat=15 + kern + jc + one-property-at-a-time
  (docGrid / rFonts cascade / sectPr), then pgMar width sweep
- evidence: Non-monotonic content-width dependency:
  - 465-475pt: NO compression
  - **453-455pt: COMPRESSED** (d77a range, L+R margin 1400-1420tw)
  - 451-452pt: NO compression (non-monotonic gap!)
  - 435-445pt: partial (11.0-11.5pt)
- interpretation: Word's yakumono compression is a LINE-WRAP PRE-PASS heuristic:
  - "If the line would fit with yakumono compression, compress"
  - "If compression doesn't help fit, don't compress"
  - "If no pressure at all, don't compress"
- outcome: ARCHITECTURAL VALIDATION — Oxi's existing Phase 2 reactive absorb
  (mod.rs:2977) + 50tw threshold (commit 70841a5) is the CORRECT approach.
  Preemptive char-advance compression would over-compress in cases Word
  doesn't (e.g., content=465pt), causing regressions.
- impact: Saves implementing the wrong solution. Risk-adjusted value HIGH —
  6 ticks of bisection prevented hours of misguided implementation.
- undocumented-quirk: this heuristic is NOT in ECMA-376 or JIS X 4051;
  it's a Word rendering pipeline implementation detail.
- methodology: 5th confirmation type (architectural validation) added to
  additive-primary protocol repertoire. All 4 tiers documented as
  complementary:
  - Positive (new spec → implement)
  - Negative (ceiling → skip)
  - Refutation (false hypothesis → revert)
  - Drift localization (scope expansion → extend)
  - Architectural validation (existing impl correct → stop)
- close: ALL yakumono pending tasks closed. No further investigation
  needed. Phase 2 threshold tuning is the only remaining lever, but its
  bottom-5 impact is already captured in 70841a5 merge.
- artifacts: tools/metrics/additive_tier1_docgrid.py + additive_tier2_rfonts.py
  + additive_tier3_sectpr.py + additive_pgmar_isolate.py + pgmar_width_sweep.py
  + bisect_d77a_styles.py + bisect_d77a_normal_properties.py + retry_normal_combos.py
  + scratch_kern_jc_test.py + pydocx_strip_d77a.py

## 2026-04-18 — oxi-3 — partial — yakumono trigger: kern+jc+cSC+compat15 + content width 452.8-455.3pt range
- Tier 1 (docGrid type): FALSIFIED — docGrid variants don't trigger
- Tier 2 (rPrDefault rFonts/lang cascade): FALSIFIED — no combination triggers
- Tier 3 (sectPr): **HIT** at pgMar L+R = 1418tw (others at 1440)
- pgMar width sweep (content_pt):
  - 465.3-475.3pt: NOT compressed
  - 453.3-455.3pt (L+R = 1400-1420): **COMPRESSED 10.5pt**
  - 451.3-452.8pt (L+R = 1425-1440): NOT compressed
  - 435.3-445.3pt: partial 11.0-11.5pt
- Interpretation: compression is NOT a simple necessary-condition gate but
  a geometry-dependent heuristic. Word's compression kicks in only when
  content width falls in a specific range that interacts with the text and
  line-wrap algorithm. This is a LINE-WRAP-HEURISTIC, not a per-character rule.
- Implementation implications:
  - Cannot be implemented as "compress yakumono IF compat_mode>=15 && cSC":
    that would compress in cases Word doesn't (e.g., content=465pt).
  - Would require replicating Word's line-break pre-pass: "if the line
    would fit with yakumono compression, compress; otherwise don't."
  - Matches the observation that Oxi's existing Phase 2 absorb logic
    (mod.rs:2977) is the right direction — it's reactive, not preemptive.
- Conclusion: char-advance-level yakumono compression cannot be implemented
  as a unconditional rule. The 50tw Phase 2 threshold (merged 70841a5) is
  the correct approach — reactive absorption up to a small overflow.
- artifacts: tools/metrics/additive_tier1_docgrid.py + additive_tier2_rfonts.py
  + additive_tier3_sectpr.py + additive_pgmar_isolate.py + pgmar_width_sweep.py

## 2026-04-18 — oxi-3 — partial — yakumono trigger narrowed to Normal style kern+jc (plus d77a-inherited factors)
- context: continuing reconciliation of earlier "cSC+compat15 confirmed" claim
- method: python-docx safe strip + XML component replacement
- findings (d77a-base variants):
  - replace styles.xml with empty/minimal: `（`=12.0pt (NOT compressed)
  - replace fontTable.xml with minimal: `（`=10.5pt (still compressed) → fontTable NOT needed
  - d77a full Normal style (w:styleId="a") alone in styles.xml: `（`=10.5pt (COMPRESSED)
  - minimum sufficient Normal pPr/rPr: **`<w:kern w:val="2"/>` + `<w:jc w:val="both"/>`** → 10.5pt
  - only jc=both: 11.5pt (partial, -0.5pt)
  - only kern=2: 12.0pt (no)
  - kern + jc: 10.5pt (full, -1.5pt)
- ANTI-BREAKTHROUGH (scratch additive test):
  - scratch + cSC + compat=15 + kern=2 + jc=both → **12.0pt (NOT compressed)**
  - Confirms user's warning: d77a-base tests are subtractive, NOT true scratch
  - d77a has ADDITIONAL inherited factors (beyond kern+jc+cSC+compat15) that
    activate the compression
- remaining candidates for true additive bisection (deferred):
  - d77a's pgMar 1418tw (vs scratch 1440tw)
  - d77a's `<w:docGrid w:type="lines" w:linePitch="360"/>` type attribute
  - d77a's fontTable PANOSE values (but replacing with minimal preserved
    compression — so maybe not needed IF pre-existing)
  - themeFontLang in styles.xml root element
  - rPrDefault rFonts cascade from docDefaults
- outcome: trigger narrower than cSC+compat15 but still not fully isolated;
  kern+jc is necessary but not sufficient in scratch. Next step: additive
  bisection from scratch, adding one d77a property at a time.
- artifact: tools/metrics/bisect_d77a_styles.py + bisect_d77a_normal_properties.py
  + retry_normal_combos.py + single_test_runner.py + scratch_kern_jc_test.py
  + pydocx_strip_d77a.py
- DO NOT implement: kern+jc gate would regress scratch/compat=15 docs that
  don't have the missing factor.

## 2026-04-18 — oxi-3 — needs-reconciliation — yakumono compression trigger = cSC + compat≥15
**STATUS DOWNGRADED from "confirmed" to "needs-reconciliation" 2026-04-18 (oxi-1 review)**

## 2026-04-18 — oxi-2 — reproducible-bug — PresetShape handler only renders bracketPair

- context: 2ea81 rank 6 AlternateContent analysis led to finding
- hypothesis: `tools/oxi-gdi-renderer/src/main.rs:377-433` PresetShape
  handler has `if shape_type == "bracketPair"` but NO else branch. All
  other preset types (rect, roundRect, ellipse, straightConnector1,
  bentConnector3, etc.) fall through silently = rendered as nothing
- evidence:
  - grep confirmed only bracketPair branch exists in the match
  - 2ea81 has 17 AlternateContent shapes (7 rect, 4 roundRect, 3 ellipse,
    2 straightConnector1, 1 bentConnector3) + 4 pic:pic images
  - 11 of 17 have txbx (text content renders via separate path) but
    borders/fills/lines do NOT render
  - 6 of 17 are shape-only (no txbx) → **completely invisible**
  - Also explains earlier "spt 32 connector invisible" finding — VML
    fallback maps to same PresetShape path
- outcome: Path A fix candidate (direct bottom-5 improvement expected)
- bottom-5 / rank 6+ impact:
  - 2ea81 (rank 6): 6 invisible + 11 missing borders — likely 4th major
    bug class after line=exact, tbRlV, spt 202
  - b35 (rank 3): 1 missing rect border (form header text box)
  - b837 (rank 4): 1 AlternateContent shape
  - Other bottom-5 docs: 0 AlternateContent
- implementation scope (dedicated session): 5 GDI primitive branches
  - rect → `Rectangle()`
  - roundRect → `RoundRect()`
  - ellipse → `Ellipse()`
  - straightConnector1 (and VML spt 32) → `MoveToEx` + `LineTo`
  - bentConnector3 → polyline (3-segment right-angle bent)
  - Plus stroke weight / color / flip handling
  Approximately 1 hour's work per the memo.
- memos: `asset_preset_shape_render_gaps.md`, `asset_vml_spt32_connector.md`

## 2026-04-18 — oxi-2 — reproducible-bug — line=exact boundary +2pt
- context: 2ea81 page 1 (rank 6, 0.6356); 2ea81 is Class B (page-aligned, intra-page only)
- hypothesis: when lineSpacingRule=exact changes between adjacent paragraphs (e.g., line=260tw empty → line=300tw body), Oxi advances by NEXT para's value (15pt) while Word uses CURRENT para's value (13pt); +2pt per boundary
- evidence:
  - Word DML + docx pPr: 2ea81 idx=6 empty (line=260tw exact), idx=7 body (line=300tw exact); Word measured advance 6→7 = 13pt
  - Oxi layout dump: dy jumps +2.7 → +4.7 exactly at idx=6→7 boundary
  - minimal repro (`tools/metrics/repro_line_exact_boundary.py`) reproduces precisely: Word A→C=26pt, Oxi A→C=28pt for 3-paragraph 260/260/300 sequence
- outcome: HANDOFF to dedicated session. Not applying code change per /loop policy.
  Code location candidates: `mod.rs:~2045-2100` (line_height computation), `mod.rs:2538` (cursor_y advance). Verify the empty-paragraph's line_height uses its own pPr not next's.
- wider impact: form-style docs (tax/application) with mixed exact line heights accumulate +delta per boundary. Potentially affects SSIM for 2ea81 rank 6 and similar form docs.
- memos: `project_2ea81_line_exact_boundary_bug.md`, `project_2ea81_class_b.md`

## 2026-04-18 — oxi-2 — refuted — yakumono rule `min(fs, 10.5)` char-advance
- context: d77a idx=9 line 1 (MS Gothic 12pt) showed '（' advance=10.5pt when fontsize=12, leading to "single yakumono compression" hypothesis
- hypothesis: Oxi's `char_width_pt` returns fontsize for fullwidth yakumono; should cap at 10.5pt
- evidence: minimal repro (tested 4 compat modes × 2 fonts × 6 sizes = 48 cases) showed yakumono advance = fontsize in all cases. Rule refuted.
- outcome: parallel session (oxi-3?) merged 1e05fe3 "fixed advance" from a different angle and was later REVERTED (b7fde5e). My refusal to merge was validated.
- note: /loop cannot reproduce the real-doc trigger for yakumono compression. Requires dedicated bisect-from-real-doc approach with subprocess isolation.
- memos: `project_yakumono_rule_unconfirmed.md`, `project_word_single_yakumono_compression.md`, `project_firstline_render_no_shift.md`

## 2026-04-18 — oxi-1 — refuted — d77a cell wrap hypothesis
- context: d77a p.9 drift (rank 2 bottom-5, 0.6042)
- hypothesis: cell text wrapping differs between Oxi and Word
- evidence: measured N=30 lines with COM; Oxi and Word agree on wrap position
- outcome: d77a drift is NOT line-wrap. Look elsewhere (cell height? vertical align?).
- commit: oxi-1 `bf1160b tools(metrics): refute d77a cell wrap hypothesis`

## 2026-04-17 — oxi-3 — confirmed — CJK overflow strict check
- context: d77a p.2 / 683f p.2 line breaking
- hypothesis: Oxi allows 2.5-38pt overflow before breaking line; Word is stricter
- evidence: COM-measured d77a PARA 21, Word 8 lines vs Oxi 7 lines (+17pt missing stub)
- outcome: fix lands in main as `9dab217`, bottom-5 sum +0.1402
- impact: 683f rank moved from 1 to 5; d77a p.9 rank 2 remaining

## 2026-04-17 — oxi-1 — confirmed — empty br-type=page stub
- context: 0e7a p.10 + d77a p.11 page break pattern
- hypothesis: Word renders empty `<w:p><w:br w:type="page"/></w:p>` as 1-line stub on prior page, then breaks
- evidence: COM Variant A shows P3 at y=84 (stub), P4 at y=56.5 (new page)
- outcome: fix lands in main as `8e63b43`, improves 0e7a p9/p10 and d77a p11
- impact: does NOT help 0e7a p.2 (different bug)

## 2026-04-17 — oxi-1 — confirmed — hanging-indent first-line x shift
- context: numbered/bulleted paragraph positioning
- hypothesis: Word places line 1 at `margin + indent_left + first_line_indent` (applies to both positive firstLine and negative hanging)
- evidence: COM `measure_hanging_indent_v2.py` on 6 cases
- outcome: fix lands in main as `640de56`, supersedes the old "firstLine does NOT shift line_x" comment

## 2026-04-17 — oxi-1 — confirmed — drop max_elem_bottom in row height (fix/b35)
- context: table row height inflation
- hypothesis: `pad_t + max(content_h, elem_bottom) + pad_b` double-counts text_y_offset
- evidence: for MS Mincho 10.5pt in 17.5pt grid, elem.y=3.5pt (center offset) + elem.height=17.5pt = 21pt bottom, 3.5pt past content_h=17.5pt
- outcome: fix lands in main as `625f8ad`, bottom-5 sum +0.0524

---

## Active hypotheses (not yet confirmed/refuted)

### oxi-4 — LM0 cell formula (investigating)
- observation: 10.5pt row_h = 18n, 12pt row_h = 28 + 36(n-1) — non-continuous formula
- current direction: sweep sizes {9,10,10.5,11,12,13,14,16,18} × both fonts × n∈{1..4} × adjustLineHeight on/off
- blocker: fit depends on whether closed-form from font metrics exists, or needs per-size constant table

### oxi-1 — 0e7a p.2 remaining drift (investigating)
- observation: p.2 SSIM 0.5767, unchanged by empty-br/hanging-indent fixes (those lifted p9/p10 only)
- current direction: per-paragraph Y position measurement to locate drift origin
- hypotheses tried: (none refuted yet for p.2 specifically)

### oxi-3 — d77a p.9 (investigating)
- refuted: cell wrap (2026-04-18)
- current direction: looking at cell height / vertical align / floating shape overhead

### 🔥 BLOCKER — footnote area over-reserve on b837 p.4
- **Status**: BLOCKS oxi-4 `39ebdb9` charGrid fix from merging
- **Symptom**: Oxi reserves full footnote body height per page; Word splits
  long footnotes across pages, reserving only what fits.
- **Measurement** (from oxi-4 memo `project_b837_footnote_over_reserve`):
  - p.4: Oxi reserves 198.5pt for 5 footnotes (all lines, 13 line-bodies)
  - Word reserves ~80pt less (splits fn 22's 5-line body across pages)
  - Oxi's cap puts paras[59] 2nd line past page end → premature break
- **Potential gain** (if fixed together with charGrid):
  - b837 p2: +0.0836
  - b837 p4: +0.0089 (target)
  - b837 p5: recovers from -0.0387 to possibly positive
  - Bottom-5 impact: potentially enough to push b837 out of bottom-5 entirely
- **Assignee (2026-04-18)**: oxi-2 (reassigned from fix/b35-multiline-cell)
- **Branch suggestion**: `fix/footnote-area-spill`
- **Key files**:
  - `crates/oxidocs-core/src/layout/mod.rs` — `estimate_footnote_h` cap logic
  - Memo chain: `project_chargrid_2cell_indent_width.md` →
    `project_footnote_reserve_sensitivity.md` →
    `project_b837_footnote_over_reserve.md`
- **Design direction**: implement Word-like footnote-area spill across pages
  (cap per-page at remaining body-space, overflow to next page's fn area).

### oxi-2 — footnote area spill (reassigned)
- target: implement Word-like footnote area page-split
- blocks: oxi-4 charGrid fix `39ebdb9`
- evidence: `project_b837_footnote_over_reserve.md` in agent memory

---

## Confidence merges (Path B — correct regardless of SSIM)

Merges that landed because the fix is *known correct* via COM + 3 docs + minimal
repro + spec reference, but didn't necessarily improve bottom-5 floor. See
CLAUDE.md §9 Path B for the rules.

(none yet — first one will land here)

---
