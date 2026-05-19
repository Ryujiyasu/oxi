# jc=both Per-Character Compression — Design Doc

**Status**: S117 design (2026-05-19). Implementation across S117-S122 (estimated).

## Background

5-session arc S112-S116 isolated and measured Word's jc=both compression
mechanism for CJK text. Three failed ship attempts (S114 uniform fs/2,
S115 wrap-lookahead, S116 corrected savings) revealed:

1. Compression mechanism is **PER-CHAR ACROSS LINE** (not yakumono-only)
2. Compression triggers based on (cell_width, text_len, fs, cs) jointly
3. Char advances snap to 15tw (0.75pt) grid
4. Priority order: yakumono > fullwidth digits > kanji
5. Compression amounts (observed at fs=10.5 cs=-9):
   - '．','，','。' → ~50-55% reduction (9.6 → 5.25-6.0pt)
   - '１' fullwidth digit → ~6% reduction (9.6 → 9.0pt)
   - kanji → ~0.6% reduction (9.6 → 9.54pt)

## Data sources

- `tools/metrics/yakumono_grid/` — 52 variants: font × fs × cs × punct
  (S113 commit fabc6a4)
- `tools/metrics/jcboth_decision_grid/` — 64 variants: cell_width × text_len
  (S116 commit d637eb8)
- `tools/metrics/15076df_buildup/` — minimal repros + trigger isolation
  (S112 commit fa083b8)
- `tools/metrics/predict_yakumono_jc_both.py` — Python prediction oracle
  (100% accurate on overflow cases per S114 validation)

## Algorithm

### Input
- Sequence of (char, font_size, run_cs, advance_natural) tuples for a
  candidate line
- Budget = effective_wrap

### Output
- Final advance per char (post-compression)
- Bool: line fits (= sum of final advances ≤ budget)

### Steps

```
1. Natural sum = Σ advance_natural[i]
2. If natural_sum ≤ budget: return advance_natural unchanged (fits)
3. Overflow = natural_sum - budget
4. Build priority queue of compressible chars:
     Priority 1: yakumono ('．','，','。','、'..)
       - max compression = natural - fs/2 (floor at half-width)
     Priority 2: fullwidth digits ('０'..'９', 'Ａ'..'Ｚ', etc.)
       - max compression = natural * 0.06 (~ 5.5% reduction)
     Priority 3: kanji (CJK ideographs)
       - max compression = natural * 0.006 (~ 0.5% reduction)
5. Allocate compression budget = overflow
   For each priority level (high to low):
     For each char at this priority:
       this_max = max_compression for this char
       amount = min(this_max, remaining_budget / chars_remaining)
       final_advance[i] = natural[i] - amount
       snap final_advance[i] to 15tw grid (round to nearest)
       remaining_budget -= amount
       if remaining_budget ≤ 0: break out
6. If remaining_budget > 0 after all priorities: line still overflows.
   Return fits=false to wrap algorithm.
7. Else: return final_advance values, fits=true.
```

### Snap to 15tw

```rust
fn snap_15tw(pt: f32) -> f32 {
    let tw = pt * 20.0;
    let snapped_tw = (tw / 15.0).round() * 15.0;
    snapped_tw / 20.0
}
```

### Compression budget allocation

Word seems to distribute compression GREEDILY at priority levels but
uniformly within a priority level. So if there's 1 '．' and the overflow
exceeds the yakumono compression budget, all of yakumono compresses to
max, then digits start. This matches the cw=1800 tl=9 observation
(both '．' and '１' compress).

## Module structure

New file: `crates/oxidocs-core/src/layout/jc_both_compress.rs`

```rust
//! Word jc=both / jc=distribute per-character compression algorithm.
//!
//! Based on COM measurement of 52+64 minimal repros (S112-S116).
//! See docs/design/jc_both_per_char_compression.md for derivation.

pub struct CharContext {
    pub ch: char,
    pub natural_advance: f32,
    pub font_size: f32,
}

#[derive(Debug, Clone)]
pub struct CompressionResult {
    pub final_advance: Vec<f32>,
    pub fits: bool,
}

pub fn compute_compression(
    chars: &[CharContext],
    budget: f32,
    gate_active: bool,
) -> CompressionResult { ... }

fn snap_15tw(pt: f32) -> f32 { ... }
fn yakumono_max_compress(ch: char, fs: f32, natural: f32) -> f32 { ... }
fn digit_max_compress(ch: char, fs: f32, natural: f32) -> f32 { ... }
fn kanji_max_compress(ch: char, fs: f32, natural: f32) -> f32 { ... }

#[cfg(test)]
mod tests { /* against S113 + S116 grids */ }
```

## Integration plan (S118+)

Cell renderer (mod.rs:6700+):
- Pre-pass: gather chars-per-line via greedy natural-width wrap (current behavior)
- Per line: call compute_compression with the line's chars
- If fits=true: emit fragments with final_advance values
- If fits=false: try shorter line (already what wrap does)
- Render-time slack<0 compression (mod.rs:6957) becomes REDUNDANT — replace with
  the new algorithm result

count_cell_lines (mod.rs:7884):
- Same pre-pass logic but track only line count
- Use compute_compression to determine if a candidate line fits

Body renderer break_into_lines (mod.rs:4751):
- Same algorithm, also called per-line

## Testing

S117 adds unit tests in `jc_both_compress.rs` against the 52+64 = 116 grid
variants. Each test asserts (final_advance per '．','，','。','１', kanji
mean) within 0.75pt (1 snap unit) of Word's measurement.

S118+ adds integration tests:
- v8/15076df fits 10 chars on L1
- d77a Phase 1 doesn't regress
- 3a4f Phase 1 doesn't regress
- d1e8ac8 Phase 1 doesn't regress

S119+ Phase 1 + Phase 3 baseline check.

## Risk register

1. **Per-char compression model is approximate**. The 0.06 digit / 0.006
   kanji ratios are observed at ONE configuration (cw=1800 tl=9). Other
   configs may have different ratios. Mitigation: extend grid if early
   tests reveal divergence.

2. **15tw snap may not be universal**. The S113 grid showed '．' values
   snap to 15tw, but kanji_mean had non-15tw values. Possibly Word
   doesn't snap kanji individually; the mean is averaged across
   non-uniform individual advances. Mitigation: investigate via per-char
   COM measurement in S117 tests.

3. **Phase 1 regression risk remains**. Even with full algorithm, edge
   cases may diverge from Word. Mitigation: extensive grid + 4-5 known
   regression docs as oracle.

4. **Body vs cell renderer divergence**. The two paths have independent
   implementations. Sharing the same `compute_compression` function will
   converge them, which could cause regression in body-rendered docs
   that previously diverged from cell-rendered docs. Mitigation: gate
   body integration behind a feature flag initially.

5. **Performance**. Pre-pass adds an extra natural-width scan per line.
   For large docs (1000+ lines), this could add measurable latency.
   Mitigation: profile and optimize if needed.

## Session plan

- **S117** (this session): design doc + module skeleton + unit tests against grids
- **S118**: implementation of compute_compression matching grid
- **S119**: cell renderer integration (behind env var OXI_JCBOTH_REFACTOR=1)
- **S120**: count_cell_lines integration
- **S121**: body renderer integration (if applicable)
- **S122**: Phase 1 + Phase 3 verify, iterate
- **S123**: ship if clean, otherwise iterate

Estimated: 6-7 sessions. User has committed to this scope.

## Success criteria

- All 116 grid variants match Word within 0.75pt (1 snap unit)
- Phase 1 baseline ≥ 53/55 (no regression vs current)
- Phase 3 SSIM net ≥ 0 (no regression vs current)
- 15076df LLA: L12/L13 match Word (L1=10 chars)
- d77a / 3a4f / d1e8ac8 / 6514f214: no Phase 1 or SSIM regression

## Fallback criteria

If S122 doesn't meet success criteria, the algorithm needs further
measurement. Revert all integration changes, keep module + tests for
next attempt. Document divergent cases in research log.

## Linked

- [[session112-15076df-jc-both-trigger]]
- [[session113-yakumono-grid-jc-both-justification]]
- [[session114-attempted-revert]]
- [[session115-wrap-lookahead-phase1-regression]]
- [[session116-decision-grid-per-char-compression]]
