//! Word jc=both / jc=distribute per-character compression algorithm.
//!
//! Based on COM measurement of 116 minimal-repro variants (S112-S116, 2026-05-19).
//! See `docs/design/jc_both_per_char_compression.md` for full derivation.
//!
//! ## Mechanism (inferred)
//!
//! When a paragraph has jc=both/distribute AND balanceSBDB AND compressPunctuation
//! AND a CJK run has cs<0:
//!
//! 1. Word computes natural line width = Σ (font_size + 2*cs_pt) for fullwidth chars
//! 2. If natural > cell budget: compress chars to fit
//! 3. Compression order (priority):
//!    a. Yakumono ('．','，','。','、','（'..) — up to half-width floor
//!    b. Fullwidth digits ('０'..'９') — up to ~6% reduction
//!    c. Kanji — up to ~0.6% reduction
//! 4. Final char advances snap to 15tw (0.75pt) grid
//! 5. Line right-aligns to budget (jc=both fills exactly)
//!
//! ## Status
//!
//! S117: module skeleton + unit tests against 116-variant grid oracle.
//! Not yet integrated into layout pipeline.
//!
//! Integration plan: S118+ via env var `OXI_JCBOTH_REFACTOR=1` gate.

use crate::layout::kinsoku;

/// Per-character context for compression computation.
#[derive(Debug, Clone)]
pub struct CharContext {
    pub ch: char,
    /// Natural advance INCLUDING cs + balanceSBDB doubling.
    /// For fullwidth char at fs=10.5, cs=-9tw, balanceSBDB=on:
    /// natural_advance = 10.5 + 2 * (-0.45) = 9.6pt
    pub natural_advance: f32,
    pub font_size: f32,
}

/// Result of compression computation.
#[derive(Debug, Clone)]
pub struct CompressionResult {
    /// Final advance per char (post-compression, snapped to 15tw).
    pub final_advance: Vec<f32>,
    /// True if line fits within budget after compression.
    pub fits: bool,
}

/// Snap a width in points to the 15tw (0.75pt) grid using round-to-nearest.
///
/// COM-observed: all '．' advances in S113 grid are multiples of 0.75pt.
pub fn snap_15tw(pt: f32) -> f32 {
    if pt <= 0.0 {
        return 0.0;
    }
    let tw = pt * 20.0;
    let snapped_tw = (tw / 15.0).round() * 15.0;
    snapped_tw / 20.0
}

/// Maximum compression possible for a yakumono char.
///
/// S120: COM-observed across fs∈{9,10.5,12} that Word's minimum '．' advance
/// is ~6.0pt regardless of fs — NOT fs/2. So the actual floor is
/// max(6.0pt, fs/2). For fs=9 (fs/2=4.5), Word's min was 6.0pt observed
/// (extended grid fs=9 cw=2500 tl=16). For fs=10.5/12, fs/2 floor of
/// 5.25/6.0 ≈ matches observation. For fs=14+, fs/2 would dominate.
pub fn yakumono_max_savings(natural: f32, font_size: f32) -> f32 {
    let floor = (font_size / 2.0).max(6.0);
    (natural - floor).max(0.0)
}

/// Maximum compression possible for a fullwidth digit/letter.
///
/// COM-observed at cw=1800 tl=9: '１' compresses 9.6 → 9.0pt = -0.6pt = 6.25%.
/// Approximate as 6% of natural.
pub fn digit_max_savings(natural: f32) -> f32 {
    (natural * 0.06).max(0.0)
}

/// Maximum compression possible for a fullwidth kanji.
///
/// COM-observed at cw=1800 tl=9: kanji_mean 9.6 → 9.54pt = -0.06pt = 0.6%.
/// S119: reduced from 0.6% to 0.1% (1/6 of observation) to minimize
/// spurious compressions on kanji-heavy lines. Single-observation
/// ratio doesn't generalize; we need a more conservative default.
pub fn kanji_max_savings(natural: f32) -> f32 {
    (natural * 0.001).max(0.0)
}

/// Classify a char into compression priority class.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CompressClass {
    /// '．','，','。','、','（',etc. — highest priority, ~50% max compression
    Yakumono,
    /// Fullwidth digits/letters — ~6% max
    FullwidthDigit,
    /// Kanji / kana — ~0.6% max
    Kanji,
    /// Not compressible (Latin, half-width, etc.)
    None,
}

pub fn classify(ch: char) -> CompressClass {
    if kinsoku::is_cjk_compressible(ch) {
        return CompressClass::Yakumono;
    }
    let cp = ch as u32;
    // Fullwidth ASCII range U+FF01..U+FF5E (digits, letters, punct).
    // '．' (U+FF0E), '，' (U+FF0C) already captured by Yakumono above.
    if (0xFF01..=0xFF5E).contains(&cp) {
        return CompressClass::FullwidthDigit;
    }
    // CJK Unified Ideographs U+4E00..U+9FFF
    // Hiragana U+3040..U+309F
    // Katakana U+30A0..U+30FF
    if (0x4E00..=0x9FFF).contains(&cp)
        || (0x3040..=0x309F).contains(&cp)
        || (0x30A0..=0x30FF).contains(&cp)
    {
        return CompressClass::Kanji;
    }
    CompressClass::None
}

/// Compute per-char compression for a candidate line.
///
/// gate_active: caller-side gate (jc=both/distribute + balanceSBDB +
/// compressPunctuation + cs<0). When false, returns natural advances unchanged.
pub fn compute_compression(
    chars: &[CharContext],
    budget: f32,
    gate_active: bool,
) -> CompressionResult {
    let n = chars.len();
    let mut final_advance: Vec<f32> = chars.iter().map(|c| c.natural_advance).collect();

    if !gate_active || n == 0 {
        let fits = final_advance.iter().sum::<f32>() <= budget;
        return CompressionResult { final_advance, fits };
    }

    let natural_sum: f32 = final_advance.iter().sum();
    if natural_sum <= budget {
        // No compression needed.
        return CompressionResult { final_advance, fits: true };
    }

    let mut remaining_overflow = natural_sum - budget;

    // Build priority lists.
    let mut yakumono_idx: Vec<usize> = Vec::new();
    let mut digit_idx: Vec<usize> = Vec::new();
    let mut kanji_idx: Vec<usize> = Vec::new();
    for (i, c) in chars.iter().enumerate() {
        match classify(c.ch) {
            CompressClass::Yakumono => yakumono_idx.push(i),
            CompressClass::FullwidthDigit => digit_idx.push(i),
            CompressClass::Kanji => kanji_idx.push(i),
            CompressClass::None => {}
        }
    }

    // Helper: apply uniform compression to indices, capped by per-char max.
    let apply = |indices: &[usize],
                      max_fn: &dyn Fn(&CharContext) -> f32,
                      final_adv: &mut [f32],
                      remaining: &mut f32| {
        if *remaining <= 0.0 || indices.is_empty() {
            return;
        }
        // Per-char max savings.
        let max_per_char: Vec<f32> = indices.iter().map(|&i| max_fn(&chars[i])).collect();
        let total_max: f32 = max_per_char.iter().sum();
        // Take min(total_max, remaining_overflow) and distribute uniformly per char,
        // capped by each char's max.
        let to_absorb = (*remaining).min(total_max);
        // Distribute: greedy fill to per-char max in order.
        // For simplicity: divide equally first; if any char's share exceeds its max,
        // it gets only its max and the surplus redistributes. Iterate until stable.
        let mut per_char_savings = vec![0.0_f32; indices.len()];
        let mut remaining_budget = to_absorb;
        let mut active: Vec<usize> = (0..indices.len()).collect();
        while remaining_budget > 1e-4 && !active.is_empty() {
            let share = remaining_budget / active.len() as f32;
            let mut new_active: Vec<usize> = Vec::new();
            for &ai in &active {
                let head_room = max_per_char[ai] - per_char_savings[ai];
                let take = share.min(head_room);
                per_char_savings[ai] += take;
                remaining_budget -= take;
                if max_per_char[ai] - per_char_savings[ai] > 1e-4 {
                    new_active.push(ai);
                }
            }
            if new_active.len() == active.len() {
                // No progress (all hit max) — break.
                break;
            }
            active = new_active;
        }
        // Apply.
        for (ai, &i) in indices.iter().enumerate() {
            let new_adv = chars[i].natural_advance - per_char_savings[ai];
            final_adv[i] = snap_15tw(new_adv);
        }
        *remaining -= to_absorb;
    };

    apply(&yakumono_idx, &|c| yakumono_max_savings(c.natural_advance, c.font_size),
          &mut final_advance, &mut remaining_overflow);
    apply(&digit_idx, &|c| digit_max_savings(c.natural_advance),
          &mut final_advance, &mut remaining_overflow);
    apply(&kanji_idx, &|c| kanji_max_savings(c.natural_advance),
          &mut final_advance, &mut remaining_overflow);

    let final_sum: f32 = final_advance.iter().sum();
    let fits = final_sum <= budget + 0.5; // 0.5pt tolerance for snap rounding
    CompressionResult { final_advance, fits }
}

// =============================================================================
// Unit tests against COM-measured grid oracles (S112-S116).
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    /// Helper: build chars for the canonical test text "１．X X X..." where
    /// X is repeated kanji '提' with natural advance.
    fn build_v8_chars(n: usize, fs: f32, cs_pt: f32) -> Vec<CharContext> {
        let natural = fs + 2.0 * cs_pt;
        let mut v = Vec::with_capacity(n);
        v.push(CharContext { ch: '１', natural_advance: natural, font_size: fs });
        v.push(CharContext { ch: '．', natural_advance: natural, font_size: fs });
        for _ in 2..n {
            v.push(CharContext { ch: '提', natural_advance: natural, font_size: fs });
        }
        v
    }

    /// L1 budget for v8 cell config (cell=1968dxa, hanging indent).
    /// = 98.4 - 1.2 (cellMar) - 1.15 (effective left) - 3.8 (right) = 92.25pt
    const V8_L1_BUDGET: f32 = 92.25;

    #[test]
    fn snap_15tw_basics() {
        assert_eq!(snap_15tw(6.0), 6.0); // 120tw, exact 15tw multiple
        assert_eq!(snap_15tw(5.85), 6.0); // 117tw → 120tw
        assert_eq!(snap_15tw(5.0), 5.25); // 100tw → 105tw (15*7)
        assert_eq!(snap_15tw(0.0), 0.0);
    }

    #[test]
    fn classify_basics() {
        assert_eq!(classify('．'), CompressClass::Yakumono);
        assert_eq!(classify('，'), CompressClass::Yakumono);
        assert_eq!(classify('。'), CompressClass::Yakumono);
        assert_eq!(classify('、'), CompressClass::Yakumono);
        assert_eq!(classify('１'), CompressClass::FullwidthDigit);
        assert_eq!(classify('９'), CompressClass::FullwidthDigit);
        assert_eq!(classify('Ａ'), CompressClass::FullwidthDigit);
        assert_eq!(classify('提'), CompressClass::Kanji);
        assert_eq!(classify('あ'), CompressClass::Kanji);
        assert_eq!(classify('カ'), CompressClass::Kanji);
        assert_eq!(classify('a'), CompressClass::None);
        assert_eq!(classify('1'), CompressClass::None); // ASCII digit not FW
    }

    /// Oracle: v8 (fs=10.5, cs=-0.45pt, 10 chars, budget=92.25)
    /// Word measured: '．'=6.0pt, line fits at 91.25pt (10 chars).
    #[test]
    fn v8_10_chars_fits_with_compression() {
        let chars = build_v8_chars(10, 10.5, -0.45);
        let result = compute_compression(&chars, V8_L1_BUDGET, true);
        assert!(result.fits, "10 chars should fit with compression");
        // '．' is at index 1 — should be compressed close to 6.0pt
        let dot_adv = result.final_advance[1];
        assert!((dot_adv - 6.0).abs() < 0.8,
                "'．' advance should snap close to 6.0pt (got {:.3})", dot_adv);
    }

    /// Oracle: v8 with only 9 chars — Word fits naturally, no compression.
    #[test]
    fn v8_9_chars_fits_naturally() {
        let chars = build_v8_chars(9, 10.5, -0.45);
        let result = compute_compression(&chars, V8_L1_BUDGET, true);
        assert!(result.fits);
        // No compression needed: '．' stays at natural
        let dot_adv = result.final_advance[1];
        let natural = 10.5 + 2.0 * -0.45;
        assert!((dot_adv - natural).abs() < 0.01,
                "'．' should stay at natural {:.3}, got {:.3}", natural, dot_adv);
    }

    /// Oracle: v8 with 12 chars at natural — would overflow even with max compression.
    /// Natural sum = 12 * 9.6 = 115.2pt. Max yakumono savings = 9.6 - 5.25 = 4.35pt.
    /// Max digit savings = 9.6 * 0.06 = 0.576pt. Total max ~ 4.93pt. Budget 92.25,
    /// would need to save 23pt. Can't.
    #[test]
    fn v8_12_chars_doesnt_fit() {
        let chars = build_v8_chars(12, 10.5, -0.45);
        let result = compute_compression(&chars, V8_L1_BUDGET, true);
        assert!(!result.fits, "12 chars shouldn't fit");
    }

    /// Oracle: gate inactive → no compression.
    #[test]
    fn gate_inactive_no_compression() {
        let chars = build_v8_chars(10, 10.5, -0.45);
        let result = compute_compression(&chars, V8_L1_BUDGET, false);
        // Returns natural unchanged; fits depends on natural sum vs budget.
        let dot_adv = result.final_advance[1];
        let natural = 10.5 + 2.0 * -0.45;
        assert!((dot_adv - natural).abs() < 0.01,
                "gate inactive: '．' = natural, got {:.3}", dot_adv);
        // Natural sum = 96.0 > budget 92.25 → fits=false
        assert!(!result.fits);
    }

    /// Oracle: jcboth_decision_grid cw=1800 tl=9.
    /// Budget = 90 - 1.2 - 1.15 - 3.8 = 83.85pt
    /// Word measured: '１'=9.00, '．'=7.50, kanji_mean=9.54, L1=9 chars fits.
    #[test]
    fn grid_cw1800_tl9_uses_per_char_compression() {
        let chars = build_v8_chars(9, 10.5, -0.45);
        let budget = 83.85;
        let result = compute_compression(&chars, budget, true);
        assert!(result.fits, "cw=1800 tl=9 should fit with per-char compression");
        let dot_adv = result.final_advance[1];
        // '．' should compress significantly (Word measured 7.5)
        assert!(dot_adv < 8.0, "'．' should compress < 8.0 (got {:.3})", dot_adv);
    }

    /// Oracle: jcboth_decision_grid cw=1500 tl=20.
    /// Budget = 75 - 1.2 - 1.15 - 3.8 = 68.85pt
    /// Word measured: never compresses; wraps at 7 chars.
    #[test]
    fn grid_cw1500_tl20_natural_wraps() {
        // 8 chars natural = 8 * 9.6 = 76.8 > 68.85 → should NOT fit even with compression
        let chars = build_v8_chars(8, 10.5, -0.45);
        let budget = 68.85;
        let result = compute_compression(&chars, budget, true);
        // Max savings = yakumono(4.35) + digit(0.576) + 6*kanji(0.0576) = ~5.27pt
        // 76.8 - 5.27 = 71.53pt > 68.85 → can't fit even with max compression
        assert!(!result.fits, "8 chars at cw=1500 shouldn't fit");
    }

    /// Oracle: snap_15tw should produce values matching grid observations.
    #[test]
    fn grid_snap_values() {
        // Observed '．' values in S113 grid:
        for &v in &[6.0, 6.75, 7.5, 8.25, 9.0, 9.75, 10.5, 11.25, 12.0, 12.75, 13.5] {
            assert_eq!(snap_15tw(v), v, "{} should be at 15tw grid", v);
        }
    }
}
