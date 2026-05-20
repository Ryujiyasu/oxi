# LO-Aligned Layout Refactor — Multi-Week Plan

**Session 146, 2026-05-21** — Plan kickoff after S135-S145 diagnostic arc.

## Background

S135-S145 traced 5 accumulated bugs in Oxi's layout that compensate
each other for SOME docs and create regressions when individually
fixed:
1. **SB suppression**: Day 33 part 17 subtracts spacing-before from
   first cell para. Word does NOT (per TR_V200-V203 + R1A).
2. **Bug A**: per-table `cursor_y += top_bw` adds 0.5pt cumulative.
   Word doesn't add it the same way.
3. **char_grid_extra**: Oxi widens fullwidth CJK chars to grid pitch.
   Word does NOT for horizontal text (per LibreOffice
   `MS_WORD_COMP_GRID_METRICS` + `bUseGridKernPors`).
4. **Cell row-height calc**: includes/excludes spacing differently
   from Word (multiple Day 33 fixes).
5. **Inter-table boundary spacing**: small leak (~0.55pt per inter-table
   per V108 measurement).

Each individual fix creates regressions in other docs because they
were relying on the buggy behavior as compensation.

## Goal

Refactor Oxi to follow the LibreOffice algorithm faithfully for the
following operations:
- Char advance width: natural (= font_size × char_advance_ratio from font metrics)
- Spacing-before of first cell para: applied (no suppression)
- Per-table top_bw: not added separately (computed from cell padding)
- Row height: sum of (max content height per cell + padding) per row
- char_grid_extra: only when MS_WORD_COMP_GRID_METRICS=false (rare)

## Phases

### Phase 0 (S146, this session): Document the design + baseline
- This document
- Baseline measurement with all 5 env gates OFF
- Baseline measurement with LO-faithful bundle (H8+SB+BugA) ON
- Per-doc impact analysis

### Phase 1: Re-architect char width
- Make grid_char_pitch char-width adjustment opt-out by default
- Keep grid_line_pitch (line height) as-is
- Per Word/LO: natural char width is correct
- Test on baseline. Expect regressions in compensating docs.

### Phase 2: Re-architect first-cell-para SB handling
- Remove Day 33 part 17 SB suppression
- Apply spacing-before to first cell para like body
- Test. Expect regressions.

### Phase 3: Re-architect per-table top_bw
- Remove `cursor_y += top_bw` per-table
- Compute table top from cell padding instead
- Test. Expect regressions.

### Phase 4: Re-architect cell row-height calc
- Remove all Day 33 patches that subtract sb/sa from row height
- Use LO's natural row-height = max(cell_content_h per cell) + borders
- Test. Expect regressions.

### Phase 5: Re-architect inter-table spacing
- Body→table and table→table transitions per LO logic
- Spacing-collapse rule between body and table boundary

### Phase 6+: Address each regression
- Identify the original compensating bug per regression
- Fix that bug (which is also wrong from LO perspective)
- Continue until baseline IoU matches or exceeds pre-refactor

## Acceptance criteria

- mean IoU ≥ baseline 0.8893 (no overall regression)
- Phase 1 pass_rate ≥ 53/55 (no doc PASS → FAIL)
- Individual doc IoU regression ≤ -0.02 (no destruction)
- SSIM mean drop ≤ 0.005 (Phase 3 sentinel)

## Risk

- Multi-week scope per S56 + S145 prior conclusions
- Big-bang refactor often introduces new bugs
- Some Word behaviors may not be LO-faithful (LO is approximation
  too)
- Some compensations may not be traceable to single bugs

## Mitigation

- Phase-by-phase commits (not single mega-PR)
- Env-gated rollback at each phase
- Per-doc regression tracking
- Reference LibreOffice source for actual algorithm in disputed areas

## Pragmatic check

Before starting full refactor, attempt:
- **Phase 0 result**: H8+SB+BugA bundle. If it nets ≥ +0 IoU with ≤ 2
  regressions, ship as default-on. Else proceed with multi-phase
  refactor.

### Phase 0 measurement CORRECTED (S147, 2026-05-21)

S146 reported gen2 -0.98 but that was a measurement artifact (incomplete
pagination data from killed background task). S147 re-ran cleanly:

| Metric | OFF | Bundle (H8+SB+BugA) | Δ |
|---|---|---|---|
| mean IoU | 0.8893 | **0.8994** | **+0.0101** ✓ |
| Phase 1 pass | 53/55 | 51/55 | -2 |

**Bundle is actually MILDLY POSITIVE** on Phase 2 net. But Phase 1
regresses (a1d6+d4d126 PASS→FAIL) so cannot ship default-on yet.

### Phase 0 original (S146, INCORRECT data — kept for history)

Tested H8+SB+BugA bundle on 52 measured docs (out of 55):

**9 gains** (significant):
- d4d126: +0.29 (Phase 1 FAIL→PASS!)
- 31420af: +0.18
- 1636d28e: +0.18
- 6514f: +0.17
- de6e: +0.15
- 29dc6e: +0.15 (was -0.39 with SB+H7; H8 recovers it!)
- 15076df: +0.06
- cb8be715, bd90b00a: minor

**8 regressions** (significant):
- **gen2: 0.9789 → 0.0000 (-0.98) CATASTROPHIC**
- **04b88e: 0.9509 → 0.6211 (-0.33) BugA-induced**
- a1d6: -0.10 (BugA exposes another drift)
- d77a: -0.09 (BugA)
- b35: -0.07
- 34140b9c: -0.02
- 683ffcab, b5f706e9f6ad: minor

Net mean IoU: 0.8893 → 0.8753 (-0.0140) — **regresses overall**
Phase 1: 53/55 → 47/52 (with skips) = **regression**

**Conclusion**: Even LO-faithful bundle is NOT shippable. gen2 (80
synthetic docs aggregated) catastrophic and 04b88e -0.33 unacceptable.

This confirms the multi-phase refactor is required — single bundle
of "LO-aligned" fixes uncovers MORE compensating bugs we haven't
identified yet.

### Next concrete step (S147)

Investigate gen2 specifically — what about the SB suppression revert
makes 80 synthetic docs render to 0 IoU? This is likely a single
implementation bug that affects synthetic test docs uniformly. Fixing
it could unlock the gen2 line and reveal the true LO-aligned bundle
impact.

## Links

- [[session145-libreoffice-grid-kern-finding]]
- [[session144-29dc6e-compensation-analysis]]
- [[session143-bug-a-isolation-partial-clean]]
- [[session142-h7-breakthrough-all-tokumei-pass]]
- [[session141-h6-implemented-a1d6-still-fails]]
- [[session140-cellwrap-docgrid-root-cause]]
- [[session56-tokumei-slow-drift-localized]]
- `docs/design/tokumei_row_drift_fix.md` — S135 initial design
