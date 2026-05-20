# Tokumei Family Row-Drift Fix — Multi-Session Design Doc

**Session 135, 2026-05-20** — Multi-session arc started.
**Status**: Phase 2 IoU baseline = 0.8939. Goal Phase 2 → Phase 3 (mean_iou ≥ 0.99).
**Owner sessions**: S135 (design + measure) → S136+ (implementation per phase).

## Motivation

Four tokumei-family docs share a near-identical per-row drift signature
that drags the Phase 2 IoU baseline down by ~0.04 cross-doc:

| doc | mean_iou | n_paras | drift signature at i=22 |
|---|---|---|---|
| d4d126dfe1d9 | 0.5401 | 140 | -18.05pt |
| de6e32b5960b | 0.6485 | 142 | -8.18pt |
| 6514f214e482 | 0.6609 | 149 | -17.85pt |
| a1d6e4efa2e7 | 0.6904 | 143 | -17.85pt |

These are the **bottom 4 of the bottom 10 IoU docs** (excluding the 3a4f /
b837 / d1e8 / 31420af / d77a / 29dc6e group which have non-tokumei
distinct root causes). All four use the `tokumei-08-01` template
variant — same OOXML pattern, same drift.

Fixing this family cleanly is expected to add **+0.04 to +0.10 mean
IoU**, the largest single-fix opportunity at the Phase 2 entry.

## Current state — measured 2026-05-20

`tools/metrics/measure_tokumei_slow_drift.py` re-ran the existing 13
minimal repros (`tools/golden-test/repros/tokumei_slow_drift/TS_V100-V112.docx`,
authored S56). Today's results:

| Variant | spec | Cum drift (Word - Oxi) | Per-row |
|---|---|---|---|
| V100 baseline | line=240 exact + before/after=87 + vAlign=center | **-79.6pt** | **-2.74pt/row** |
| V101 no before/after | pure exact lineRule | +29.0pt | +1.0pt/row |
| V102 no vAlign | (same as V100, vAlign removed) | -79.6pt | -2.74pt/row (vAlign has no effect) |
| V103 lineRule=auto | (different mechanism) | -700.4pt | -24.15pt/row |
| V104 no adjustLineHeightInTable | (no effect) | -79.6pt | -2.74pt/row |
| V105 line=200 (10pt exact) | (line value invariant) | -79.85pt | -2.75pt/row |
| V106 line=300 (15pt exact) | (line value invariant) | -79.6pt | -2.74pt/row |
| V107 1 table 30 rows | within-table advance | 0.25pt | ~0pt |
| V108 6 tables × 5 rows | inter-table boundaries | 2.75pt | 0.55pt/table |
| V109 no border | (no border = no top_bw) | 0.0pt | 0pt |
| V110 explicit pad_t | (pad_t set explicitly) | -0.25pt | ~0pt |
| V111 lineRule=default | (different mechanism) | -473.9pt | -16.34pt/row |
| V112 body paragraphs | (Oxi gives n=0, broken script) | n/a | n/a |

### Key observations vs S56

1. **V100 sign FLIPPED since S56**: was +1.33pt/row OVER-advance, now is
   -2.74pt/row UNDER-advance. Intervening commits (notably S94
   `14e750d apply afterLines in cell render path`) changed the
   algorithm. The d4d126 real-doc signature now matches V100 exactly.

2. **Within-table advance is CORRECT** (V107 = 0.25pt over 30 rows = ~0).
   The drift is per-table-boundary only.

3. **Inter-table leak per boundary** (V108): 0.55pt × 5 boundaries =
   2.75pt. Smaller than V100's 2.74pt-per-table → V100's drift is NOT
   just inter-table; it's something specific to V100's spacing-before /
   spacing-after configuration.

4. **V101 - V100 cell-extent difference**:
   - Word per-row: 12pt (V101) vs 20.1pt (V100) → +8.1pt extra
   - Oxi per-row: 13pt (V101) vs 17.35pt (V100) → +4.35pt extra
   - **Oxi missing ~4.35pt** = exactly ONE space-before or space-after
     value (87 twip = 4.35pt).

## Root-cause hypothesis

**Hypothesis H1**: Oxi's `estimate_para_height_emit` for the first
paragraph in a cell SUBTRACTS `space_before` (per Day 33 part 17 fix
at [mod.rs:6302-6317](crates/oxidocs-core/src/layout/mod.rs#L6302-L6317)).
This makes row_height smaller by space_before, but the text RENDER
position still applies space_before, so text lands at the correct y
within the row. However, when the row's tcMar=0 or has no trHeight,
the NEXT row starts space_before earlier than Word, accumulating drift.

**Evidence**:
- V100 has cell para with `before=87`. Each cell is its own 1-row table
  → each row has its first-para's space_before subtracted from
  row_height → 30 rows × 4.35 = 130.5pt of drift expected. Observed
  -79.6pt (less — implies partial compensation elsewhere, possibly via
  cell `text_y_off` calc or pad_t shenanigans).
- V101 has no `before/after`. No subtraction happens → no missing
  space drift (just Bug A residual +1pt/row from top_bw add).
- V107 (1 table 30 rows): only ONE first-cell-para, so subtraction
  happens ONCE → -4.35pt total expected, but observed +0.25pt because
  other compensations (Bug A +0.5pt, etc.) net it close to zero.

## Implementation roadmap (multi-session)

### S135 (done) — research baseline
- Design doc (this file)
- Re-measure V100-V112 baseline (above)
- Memory note `session135_phase2_tokumei_iou_confirmed.md`

### S136 (done) — H1 verified via TR_V200-V203 + R1A re-measurement + env-gated A/B

**TR_V200-V203 minimal repros built** (`tools/metrics/gen_tokumei_sb_isolation_repro.py`,
`tools/golden-test/repros/tokumei_slow_drift/TR_V20{0,1,2,3}.docx`).
**Measurement** (`tools/metrics/measure_tokumei_sb_isolation.py`,
`pipeline_data/tokumei_sb_isolation_results.json`):

| Variant | OFF drift (current) | ON drift (revert) | Recovery |
|---|---|---|---|
| V200 before-only | -3.55pt/row | +0.53pt/row | +4.08pt/row |
| V201 after-only | +0.80pt/row | +0.80pt/row | unchanged |
| V202 before + trH=300tw | -1.05pt/row | +0.53pt/row | +1.58pt/row |
| V203 both (sanity) | -3.26pt/row | +0.82pt/row | +4.08pt/row |

**R1A re-measurement** (`tools/metrics/measure_r1a_para_y.py`,
`pipeline_data/r1a_para_y_results.json`) — the 8 minimal repros that
ORIGINALLY justified Day 33 part 17:

| Variant | cell_para_y | body_para_y | gap | sb? | exact? |
|---|---|---|---|---|---|
| R1A_baseline | 43.5 | 57.0 | 13.5 | no | no |
| R1A_spacing_before | **47.25** | 60.75 | 13.5 | yes | no |
| R1A_spacing_lineRule | **47.25** | 59.25 | 12.0 | yes | exact |
| R1A_all4 | **47.25** | 59.25 | 12.0 | yes | exact |

**Word DOES apply sb** — cell_para_y shifts from 43.5 to 47.25 (+4.35pt)
when sb=87. Day 33 part 17's "12.5pt row" claim conflated **row.Height**
(line-only) with **visual extent** (line + sb above + post-row gap).
The premise was wrong from the start. **H1 quantitatively confirmed.**

**Env-gated A/B fix applied** at `crates/oxidocs-core/src/layout/mod.rs:6302-6322`
+ `:6720-6740` — `OXI_SB_NO_SUPPRESS=1` disables both suppressions.
Default OFF preserves Day 33 part 17. Rebuilt GDI only (DWrite skipped
since we're testing Phase 1+2 only, not Phase 3 SSIM).

**4-doc A/B result** (with OXI_SB_NO_SUPPRESS=1):

| doc | OFF IoU (baseline) | ON IoU | Δ | Phase 1 |
|---|---|---|---|---|
| d4d126dfe1d9 | 0.5401 | 0.8239 | **+0.2838** | FAIL→FAIL (score 0.80→0.99) |
| 6514f214e482 | 0.6609 | 0.8227 | **+0.1618** | PASS→PASS |
| de6e32b5960b | 0.6485 | 0.7423 | **+0.0938** | PASS→PASS |
| a1d6e4efa2e7 | 0.6904 | 0.5961 | **-0.0943** | **PASS→FAIL** (page boundary shift) |

**Mean IoU on baseline 51 docs: 0.8939 → 0.9026 (+0.0087)** (4 docs' contribution).

**a1d6 regression analysis**: dy reaches -301pt at i=365-368 (page 7),
meaning Oxi pushed paragraphs to a different page than Word. The pre-fix
under-advance was compensating for some OTHER drift on a1d6's specific
content, making it land on correct pages despite within-page drift.
Post-fix Oxi over-advances → wrong page.

**Phase 2 merge gate violation**: Phase 1 sentinel "pagination
pass_rate doesn't regress" is violated (a1d6 PASS→FAIL). **Cannot ship
unconditional revert yet.**

### S137 (next) — investigate a1d6 over-correction

a1d6 and 6514f have **identical pre-fix drift fingerprints** (matching
to ±0.2pt across all rows). Yet post-fix:
- 6514f: page 7 final dy=-0.10pt (clean), PASS
- a1d6: page 7 final dy=-301pt (page shift), FAIL

What's structurally different? Hypotheses:
- (H2a) a1d6 has one extra paragraph somewhere that pushes the
  cumulative threshold past a page boundary that 6514f doesn't cross.
- (H2b) a1d6 has a different table with different sb pattern that
  responds differently to the fix.
- (H2c) Some non-tokumei content (image, footer, header) compresses
  differently.

Investigation:
- Compare a1d6 vs 6514f OOXML structure side-by-side
- COM-measure a1d6 at the exact page boundary where the shift happens
  (page 6→7 transition)
- Find the specific paragraph in a1d6 that "tips over" with the fix

### S138 (after S137) — refine fix to preserve a1d6 PASS

Options based on S137 findings:
- (A) **Per-paragraph compensation**: identify specific structural
  trigger that causes a1d6 over-correction; add a guard.
- (B) **Half-revert**: keep render-side suppression, revert only
  estimate-side (or vice versa). Test which combination preserves a1d6
  and improves d4d126/de6e/6514f.
- (C) **Different fix surface entirely**: maybe the bug isn't in Day 33
  part 17 but in a downstream calc that the suppression masks.

### S139 (after S138) — single-doc test on each tokumei doc

- Verify d4d126/de6e/6514f gains preserved
- Verify a1d6 PASS preserved
- Cross-doc test: scan all baseline docs for "first cell para with sb"
  pattern; predict which docs change positions.

### S140 — full baseline verify
- Rebuild **BOTH** `tools/oxi-dwrite-renderer/` AND
  `tools/oxi-gdi-renderer/`. Delete `pipeline_data/oxi_png/<doc>/` AND
  `pipeline_data/pagination_oxi/<doc>.json` for affected docs.
- Phase 1 (pagination): must stay ≥ 50/51 (current entry sentinel; was
  53/55 in S134 memo — same ratio, different denominator).
- Phase 2 (IoU): mean_iou must strictly increase.
- Phase 3 SSIM sentinel: must NOT drop > 0.005 mean.

### S141+ — handle remaining compensation regressions

If S139 finds the fix still breaks other PASS docs (besides a1d6), per
CLAUDE.md no-EXCEPTION-stacking: the spec is wrong, re-derive from
richer input space.

## Risk register

1. **Day 33 part 17 was added for a reason** — removing it may regress
   the docs it was protecting. Memory archaeology: find the original
   defect and reproduce it minimal-repro-style before changing.
2. **S56 Fix A reverted at -1.59 SSIM / 115 pages**. The current fix
   surface is in similar territory. Phase 2 IoU gate is more lenient
   than SSIM but the > 0.005 SSIM-drop sentinel still kills broad
   damage.
3. **3-doc + minimal repro rule (CLAUDE.md)**: fix is "hypothesis"
   until 3 distinct real docs + minimal repro confirm. d4d126 + de6e +
   6514f + a1d6 + TR_V200 = 5 confirmations. Sufficient.
4. **No EXCEPTION stacking**: if fix needs per-doc carve-outs, the spec
   is wrong. Re-derive from richer input space.

## Success criteria

- **Necessary**: d4d126 mean_iou ≥ 0.85 (from 0.54), tokumei-08-01 family
  ≥ 0.85 each.
- **Sufficient**: + Phase 1 still ≥ 53/55, SSIM mean drop ≤ 0.005.
- **Ideal**: Phase 2 mean_iou ≥ 0.92 (from 0.89), one step closer to
  Phase 3 transition.

## Links / context

- [[session135-phase2-tokumei-iou-confirmed]] — session note (this work)
- [[session134-phase2-transition]] — Phase 2 entry conditions
- [[session56-tokumei-slow-drift-localized]] — S56 investigation
- [[session56-day32-part10-correlation]] — pattern-cell correlation
- [[methodology-phase-based]] — gate definitions
- [`tools/metrics/measure_tokumei_slow_drift.py`](tools/metrics/measure_tokumei_slow_drift.py) — V100-V112 measurement
- [`pipeline_data/tokumei_slow_drift_results.json`](pipeline_data/tokumei_slow_drift_results.json) — current results
- [`crates/oxidocs-core/src/layout/mod.rs:6302-6317`](crates/oxidocs-core/src/layout/mod.rs#L6302-L6317) — Day 33 part 17 code
- [`crates/oxidocs-core/src/layout/mod.rs:6221-6224`](crates/oxidocs-core/src/layout/mod.rs#L6221-L6224) — Bug A (S56)
