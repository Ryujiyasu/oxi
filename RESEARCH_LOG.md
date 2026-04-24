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
