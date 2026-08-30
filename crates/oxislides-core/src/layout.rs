// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Slide layout, independent of any platform's font system.
//!
//! The pptx layout has lived inside the native renderer, where every
//! measurement goes through a GDI device context. That is why the browser can
//! only *view* a deck: the rules that decide where a line breaks cannot run
//! where there is no GDI. The rules themselves, though, barely touch it -- the
//! renderer's `fits_line`, `measure_wrap` and `compute_shape_anchor_off` carry
//! an `HDC` only to hand it down to the advance probes. So the port is a seam,
//! not a rewrite: name what the layout needs to ASK about a face, and the same
//! arithmetic serves a GDI device, a metrics table, or a font parsed in wasm.
//!
//! [`FaceMetrics`] is that question. This module is the first piece of the
//! answer -- the break test, which is the most load-bearing rule the pptx side
//! has (`pptx-master-unit-break-law`, 221/221 against PowerPoint's own COM
//! sweep).

/// What the layout needs to ask about a face, whoever is answering.
///
/// The renderer answers with GDI; a browser build answers from the deck's own
/// embedded parts and a metrics table. Both must return the same numbers for
/// the same face, which is what makes the port verifiable: run the renderer
/// through this trait and its PNGs have to come out byte-identical.
pub trait FaceMetrics {
    /// The advance of `ch` in EM units for the face serving this request, or
    /// None when the answer is unknown -- a glyph the face lacks, a family
    /// nothing can serve. None is not zero: the caller must fall back to a
    /// coarser measurement rather than treat the character as empty.
    fn advance_em(&self, family: &str, bold: bool, italic: bool, ch: char) -> Option<f32>;

    /// Whether the face serving this request can draw every character of
    /// `text`. A run whose face is missing one glyph is measured elsewhere,
    /// because a fallback glyph does not advance by the base face's metrics.
    fn has_all_glyphs(&self, family: &str, bold: bool, italic: bool, text: &str) -> bool;
}

/// The PowerPoint-97 master unit: 1/8 pt, 576 to the inch.
pub const MASTER_UNITS_PER_PT: f64 = 8.0;

/// The width PowerPoint's line-break test compares against the box, in master
/// units, or None when any character's advance is unknown.
///
/// Each glyph's advance is rounded to the master unit ON ITS OWN and the
/// rounded values are summed -- not the other way round. That quantum and that
/// order are what a 221-arm COM sweep singled out, and it is what makes a line
/// whose exact width fits still break: d09's "Happy Holi!" is 546.399pt exact
/// against a 546.4128pt box, but 546.5 in master units, and PowerPoint breaks
/// it.
///
/// `spc` is the run's `a:rPr/@spc` tracking in points, added to every glyph
/// before the rounding, because PowerPoint tracks the advance and then rounds.
///
/// Characters outside the BMP are refused rather than guessed at: the probes
/// this stands in front of are UTF-16 and cannot address a surrogate pair.
pub fn master_units(
    metrics: &dyn FaceMetrics,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    spc: f32,
) -> Option<i64> {
    if text.chars().any(|c| c as u32 > 0xFFFF) {
        return None;
    }
    if !metrics.has_all_glyphs(family, bold, italic, text) {
        return None;
    }
    let mut sum: i64 = 0;
    for ch in text.chars() {
        let em = metrics.advance_em(family, bold, italic, ch)?;
        sum += f64::from((em * fs + spc) * MASTER_UNITS_PER_PT as f32).round() as i64;
    }
    Some(sum)
}

/// Whether a line of `mu` master units fits a box `width_pt` points wide.
///
/// The comparison is inclusive -- a line whose master sum is exactly the box
/// width stays whole -- and the epsilon is there because the width arrives as a
/// binary fraction of an EMU and can land a hair under its own value.
pub fn fits(mu: i64, width_pt: f32) -> bool {
    mu as f64 / MASTER_UNITS_PER_PT <= f64::from(width_pt) + 1e-6
}

/// The master-unit width as points, for reporting and for the callers that
/// compare it against something other than a box.
pub fn master_units_pt(mu: i64) -> f64 {
    mu as f64 / MASTER_UNITS_PER_PT
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A face with one advance for every character, so the arithmetic can be
    /// checked without a font.
    struct Flat(f32);
    impl FaceMetrics for Flat {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            Some(self.0)
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            true
        }
    }

    /// A face that knows nothing, to check that None propagates.
    struct Blind;
    impl FaceMetrics for Blind {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            None
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            true
        }
    }

    #[test]
    fn each_glyph_is_rounded_before_it_is_summed() {
        // 0.5 em at 12pt = 6.0pt = 48 master units exactly.
        let m = Flat(0.5);
        assert_eq!(master_units(&m, "abcd", 12.0, "X", false, false, 0.0), Some(4 * 48));
        // 0.51 em at 12pt = 6.12pt = 48.96 units -> 49 each, so the sum is 196,
        // not the 195.84 that rounding the total would give.
        let m = Flat(0.51);
        assert_eq!(master_units(&m, "abcd", 12.0, "X", false, false, 0.0), Some(196));
    }

    #[test]
    fn tracking_is_added_before_the_rounding() {
        let m = Flat(0.5);
        // +0.1pt per glyph: 6.1pt = 48.8 -> 49 units each.
        assert_eq!(master_units(&m, "ab", 12.0, "X", false, false, 0.1), Some(98));
    }

    #[test]
    fn an_unknown_advance_refuses_the_whole_line() {
        assert_eq!(master_units(&Blind, "ab", 12.0, "X", false, false, 0.0), None);
    }

    #[test]
    fn astral_characters_are_refused_not_guessed() {
        let m = Flat(0.5);
        assert_eq!(master_units(&m, "a\u{1F600}", 12.0, "X", false, false, 0.0), None);
    }

    #[test]
    fn the_fit_is_inclusive_at_equality() {
        // 546.5pt in master units against a box of exactly that width.
        assert!(fits(4372, 546.5));
        assert!(!fits(4373, 546.5));
    }

    #[test]
    fn d09_happy_holi_breaks_although_its_exact_width_fits() {
        // The sweep's own specimen: exact 546.399pt fits a 546.4128pt box,
        // the master sum 546.5 does not.
        assert!(546.399_f64 <= 546.4128);
        assert!(!fits(4372, 546.4128));
    }
}

/// The styles a paragraph's runs impose on a candidate line.
///
/// A line is a slice of the paragraph, so a character's style is found by
/// walking the runs from the paragraph's start: `line_start` is how many
/// characters earlier lines already took. The runs are contiguous and in
/// order, so this is a running total rather than a search.
pub struct RunStyles<'a> {
    pub runs: &'a [crate::ir::SlideRun],
    /// Characters of the paragraph already committed to earlier lines.
    pub line_start: usize,
}

/// The size, weight, slant and tracking the character at `at` is set in.
///
/// Falls back to the paragraph's own values for a character past the last run,
/// which is what a line ending in generated text (a field's cached value) hits.
fn style_at(
    styles: &RunStyles<'_>,
    at: usize,
    fs: f32,
    bold: bool,
    italic: bool,
    letter_spacing: bool,
) -> (f32, bool, bool, f32) {
    let mut seen = 0usize;
    for run in styles.runs {
        let n = run.text.chars().count();
        if at < seen + n {
            return (
                run.font_size.unwrap_or(fs),
                run.bold,
                run.italic,
                if letter_spacing { run.spacing.unwrap_or(0.0) } else { 0.0 },
            );
        }
        seen += n;
    }
    (fs, bold, italic, 0.0)
}

/// The paragraph's tracking, taken from its first run.
///
/// A paragraph whose runs disagree takes [`master_units_runs`] instead, so one
/// value for the whole line is right exactly when this is used.
pub fn para_spc(runs: &[crate::ir::SlideRun], letter_spacing: bool) -> f32 {
    if !letter_spacing {
        return 0.0;
    }
    runs.first().and_then(|r| r.spacing).unwrap_or(0.0)
}

/// Master units for `text` with every character measured at its OWN run's
/// size, weight and slant.
///
/// The single-style [`master_units`] measures a mixed line in one face, which
/// makes one bold word widen every line of its paragraph. The rounding is the
/// same -- per glyph, then summed -- only the style varies along the line.
///
/// Coverage is asked ONCE for the whole line, not per character: the
/// per-character form issues a face lookup per character per candidate prefix,
/// which is quadratic in the paragraph.
pub fn master_units_runs(
    metrics: &dyn FaceMetrics,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    styles: &RunStyles<'_>,
    letter_spacing: bool,
) -> Option<i64> {
    if text.chars().any(|c| c as u32 > 0xFFFF) {
        return None;
    }
    if !metrics.has_all_glyphs(family, bold, italic, text) {
        return None;
    }
    let mut sum: i64 = 0;
    for (i, ch) in text.chars().enumerate() {
        let (run_fs, run_bold, run_italic, run_track) =
            style_at(styles, styles.line_start + i, fs, bold, italic, letter_spacing);
        let em = metrics.advance_em(family, run_bold, run_italic, ch)?;
        sum += f64::from((em * run_fs + run_track) * MASTER_UNITS_PER_PT as f32).round() as i64;
    }
    Some(sum)
}

/// The pieces a line may break between.
///
/// Words, keeping the space that follows them so a trailing space stays with
/// the line it ends -- and, when `hyphen_breaks` is set, after a hyphen too, so
/// `"e-mail"` can break as `"e-"` + `"mail"`. A trailing hyphen does not split,
/// which would leave an empty tail piece.
pub fn break_pieces(text: &str, hyphen_breaks: bool) -> Vec<&str> {
    if !hyphen_breaks {
        return text.split_inclusive(' ').collect();
    }
    let mut out = Vec::new();
    for chunk in text.split_inclusive(' ') {
        let mut start = 0usize;
        for (i, ch) in chunk.char_indices() {
            if ch == '-' && i + 1 < chunk.len() {
                out.push(&chunk[start..i + 1]);
                start = i + 1;
            }
        }
        out.push(&chunk[start..]);
    }
    out
}

#[cfg(test)]
mod wrap_tests {
    use super::*;

    fn run(text: &str, size: Option<f32>, bold: bool) -> crate::ir::SlideRun {
        crate::ir::SlideRun {
            text: text.to_string(),
            font_size: size,
            bold,
            italic: false,
            underline: false,
            color: None,
            color_alpha: None,
            highlight: None,
            font_family: None,
            spacing: None,
        }
    }

    struct ByWeight;
    impl FaceMetrics for ByWeight {
        fn advance_em(&self, _: &str, bold: bool, _: bool, _: char) -> Option<f32> {
            Some(if bold { 0.6 } else { 0.5 })
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            true
        }
    }

    #[test]
    fn each_character_is_measured_in_its_own_run() {
        let runs = [run("ab", None, false), run("cd", None, true)];
        let st = RunStyles { runs: &runs, line_start: 0 };
        // 2 chars at 0.5em and 2 at 0.6em, 12pt: 2*48 + 2*57.6->58 = 212.
        let got = master_units_runs(&ByWeight, "abcd", 12.0, "X", false, false, &st, true);
        assert_eq!(got, Some(2 * 48 + 2 * 58));
    }

    #[test]
    fn line_start_shifts_which_run_owns_a_character() {
        let runs = [run("ab", None, false), run("cd", None, true)];
        // The line is the paragraph's tail, so both its characters are bold.
        let st = RunStyles { runs: &runs, line_start: 2 };
        assert_eq!(
            master_units_runs(&ByWeight, "cd", 12.0, "X", false, false, &st, true),
            Some(2 * 58)
        );
    }

    #[test]
    fn a_words_trailing_space_stays_with_it() {
        assert_eq!(break_pieces("a bc d", false), vec!["a ", "bc ", "d"]);
    }

    #[test]
    fn a_hyphen_opens_a_break_but_not_at_the_end() {
        assert_eq!(break_pieces("e-mail x-", true), vec!["e-", "mail ", "x-"]);
        assert_eq!(break_pieces("e-mail", false), vec!["e-mail"]);
    }
}
