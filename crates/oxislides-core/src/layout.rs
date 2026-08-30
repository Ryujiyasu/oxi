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

    /// Where the face puts its baseline inside one em, read from its own
    /// tables. None means "ask the offline table instead" -- which is what a
    /// source with no font system says, and what the renderer says for a face
    /// GDI cannot hand tables back for.
    fn baseline_offset_em(&self, _family: &str) -> Option<f32> {
        None
    }
}

/// The measured ascent split of the faces that were measured, in EM.
///
/// Taken before the first-baseline rule was derived and kept as the offline
/// answer for a face whose own tables cannot be read. Each is within 0.0005 of
/// what the rule computes for it.
pub fn table_baseline_offset_em(family: &str) -> f32 {
    match family.to_ascii_lowercase().as_str() {
        "arial" => 0.97274,
        "times new roman" => 0.96587,
        "calibri" => 0.93648,
        "segoe ui" => 0.97399,
        "georgia" => 0.96899,
        "verdana" => 0.99275,
        _ => 0.9685,
    }
}

/// How far the first baseline sits below the top of a text box.
///
/// `n` is the line-spacing multiple. The rule, derived over 31 arms: the
/// DESCENT the box reserves under the last line is capped at a quarter of the
/// box at single spacing and floored at a quarter above it, and the first
/// baseline is what is left. A face already deeper than that quarter gives up
/// a quarter of whatever the box loses.
///
/// `joined_rule` off restores the older split, whose parenthesisation is
/// load-bearing: `0.75 * (fs * 1.2 * n)` associates differently from
/// `0.75 * fs * 1.2 * n`, and the 1-ULP difference flipped a page of d37.
pub fn first_baseline_off(
    metrics: &dyn FaceMetrics,
    family: &str,
    fs: f32,
    n: f32,
    joined_rule: bool,
) -> f32 {
    let offset_em = metrics
        .baseline_offset_em(family)
        .unwrap_or_else(|| table_baseline_offset_em(family));
    if !joined_rule {
        return if (n - 1.0).abs() > 1e-4 {
            0.75 * (fs * 1.2 * n)
        } else {
            offset_em * fs
        };
    }
    let pitch = fs * 1.2;
    let natural_descent = pitch - offset_em * fs;
    let quarter = 0.25 * pitch;
    let descent = if n <= 1.0 {
        (natural_descent + quarter * (n - 1.0)).max(natural_descent.min(quarter * n))
    } else {
        natural_descent.max(quarter * n)
    };
    pitch * n - descent
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

/// What a wrap is allowed to do, beyond breaking at spaces.
pub struct WrapOpts {
    /// A line's trailing space HANGS past the right edge and is not part of
    /// the width the break is judged against. Measured on d28 slide 13
    /// (2026-08-18): "National Cemetery in Gettysburg, Pennsylvania. In just"
    /// is 1034px against a 1036px box and PowerPoint keeps it whole, but with
    /// the trailing space it is 1047px -- so a per-word accumulation broke
    /// before "just" and the paragraph needed 11 lines where PowerPoint needs
    /// 10. Off restores the old per-word accumulation.
    pub trim_trailing_space: bool,
    /// A single "word" wider than the line breaks INSIDE itself. d11 and d24
    /// slide 38 are 53 emoji with no space between them in a 490pt box;
    /// PowerPoint lays them out in four rows. 45 paragraphs across nine decks
    /// carry a space-free run of 30 characters or more.
    pub char_wrap: bool,
    /// A hyphen opens a break site, so `"e-mail"` may break as `"e-" + "mail"`.
    pub hyphen_breaks: bool,
}

/// Break `text` into lines that fit `first_width_pt`, then `rest_width_pt`.
///
/// The two widths differ because a hanging indent or a bullet narrows every
/// line after the first.
///
/// The platform appears only as two questions:
///
///   * `fits(candidate, width_pt, width_px, line_start)` -- does this candidate
///     line fit? `line_start` is how many characters of the paragraph earlier
///     lines already took, which is what maps a character back to its run.
///   * `measure_px(text)` -- the running width, consulted only when
///     `trim_trailing_space` is off.
///
/// Measuring the candidate PREFIX rather than summing per-word widths also
/// drops the per-word integer-pixel rounding, which pushed breaks the same way
/// the trailing space did.
pub fn wrap_lines<F, M>(
    text: &str,
    first_width_pt: f32,
    rest_width_pt: f32,
    scale: f64,
    opts: &WrapOpts,
    fits: F,
    measure_px: M,
) -> Vec<String>
where
    F: Fn(&str, f32, i32, usize) -> bool,
    M: Fn(&str) -> i32,
{
    let first_px = (first_width_pt as f64 * scale).round().max(1.0) as i32;
    let rest_px = (rest_width_pt as f64 * scale).round().max(1.0) as i32;
    let mut width_px = first_px;
    let mut width_pt = first_width_pt;
    let mut lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut current_w = 0i32;
    // Characters already committed to finished lines.
    let mut emitted = 0usize;
    for word in break_pieces(text, opts.hyphen_breaks) {
        let ok = if opts.trim_trailing_space {
            let mut candidate = current.clone();
            candidate.push_str(word);
            fits(candidate.trim_end(), width_pt, width_px, emitted)
        } else {
            current_w + measure_px(word) <= width_px
        };
        if !current.is_empty() && !ok {
            emitted += current.chars().count();
            lines.push(std::mem::take(&mut current));
            current_w = 0;
            width_px = rest_px;
            width_pt = rest_width_pt;
        }
        if opts.char_wrap && current.is_empty() {
            let mut rest = word;
            loop {
                let trimmed = rest.trim_end();
                if trimmed.is_empty() || fits(trimmed, width_pt, width_px, emitted) {
                    break;
                }
                // Longest prefix that fits, never empty so the loop ends.
                let mut last_ok = 0usize;
                for (i, ch) in rest.char_indices() {
                    let end = i + ch.len_utf8();
                    if fits(rest[..end].trim_end(), width_pt, width_px, emitted) {
                        last_ok = end;
                    } else {
                        break;
                    }
                }
                let take = if last_ok > 0 {
                    last_ok
                } else {
                    rest.char_indices().nth(1).map(|(i, _)| i).unwrap_or(rest.len())
                };
                emitted += rest[..take].chars().count();
                lines.push(rest[..take].to_string());
                rest = &rest[take..];
                width_px = rest_px;
                width_pt = rest_width_pt;
            }
            current.push_str(rest);
            current_w += measure_px(rest);
            continue;
        }
        current.push_str(word);
        current_w += measure_px(word);
    }
    if !current.is_empty() {
        lines.push(current);
    }
    if lines.is_empty() {
        lines.push(String::new());
    }
    lines
}

#[cfg(test)]
mod wrap_loop_tests {
    use super::*;

    const OPTS: WrapOpts = WrapOpts {
        trim_trailing_space: true,
        char_wrap: true,
        hyphen_breaks: false,
    };

    /// Every character is one point wide, so a width in points is a character
    /// budget and the breaks can be read by eye.
    fn one_pt_per_char(t: &str, w: f32, _px: i32, _start: usize) -> bool {
        t.chars().count() as f32 <= w
    }

    #[test]
    fn a_line_breaks_at_the_last_word_that_fits() {
        let got = wrap_lines("aa bb cc dd", 6.0, 6.0, 1.0, &OPTS, one_pt_per_char, |t| {
            t.len() as i32
        });
        assert_eq!(got, vec!["aa bb ", "cc dd"]);
    }

    #[test]
    fn the_trailing_space_does_not_count_against_the_box() {
        // "aa bb" is 5 characters and fits 5; its trailing space would not.
        let got = wrap_lines("aa bb cc", 5.0, 5.0, 1.0, &OPTS, one_pt_per_char, |t| {
            t.len() as i32
        });
        assert_eq!(got, vec!["aa bb ", "cc"]);
    }

    #[test]
    fn later_lines_are_judged_against_the_continuation_width() {
        let got = wrap_lines("aa bb cc dd", 6.0, 3.0, 1.0, &OPTS, one_pt_per_char, |t| {
            t.len() as i32
        });
        assert_eq!(got, vec!["aa bb ", "cc ", "dd"]);
    }

    #[test]
    fn a_word_wider_than_the_line_breaks_inside_itself() {
        let got = wrap_lines("aaaaaaa", 3.0, 3.0, 1.0, &OPTS, one_pt_per_char, |t| {
            t.len() as i32
        });
        assert_eq!(got, vec!["aaa", "aaa", "a"]);
    }

    #[test]
    fn char_wrap_off_leaves_the_long_word_whole() {
        let opts = WrapOpts { char_wrap: false, ..OPTS };
        let got = wrap_lines("aaaaaaa", 3.0, 3.0, 1.0, &opts, one_pt_per_char, |t| {
            t.len() as i32
        });
        assert_eq!(got, vec!["aaaaaaa"]);
    }

    #[test]
    fn empty_text_still_yields_one_empty_line() {
        let got = wrap_lines("", 10.0, 10.0, 1.0, &OPTS, one_pt_per_char, |t| t.len() as i32);
        assert_eq!(got, vec![String::new()]);
    }
}

/// A [`FaceMetrics`] that answers from the measured design-advance tables.
///
/// This is the answer a build with no font system can give: the `hmtx` tables
/// in [`crate::font_adv`], measured from the real files, cover the families
/// they cover and refuse everything else. Refusing is the point -- a browser
/// that guessed a width would break lines where PowerPoint does not, and a
/// caller that gets None can say so instead of drawing a lie.
///
/// The renderer keeps its own richer answer (the deck's embedded parts, the
/// Office cloud cache, then a GDI probe); this is what remains when none of
/// those exist.
pub struct TableMetrics;

impl FaceMetrics for TableMetrics {
    fn advance_em(&self, family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
        // The wide table first, because it knows the style; the renderer's
        // shared one is style-blind and only covers three faces.
        crate::font_adv_local::local_advance_em(family, bold, italic, ch)
            .or_else(|| crate::font_adv::hmtx_advance_em(family, ch))
    }

    fn has_all_glyphs(&self, family: &str, bold: bool, italic: bool, text: &str) -> bool {
        text.chars()
            .all(|c| self.advance_em(family, bold, italic, c).is_some())
    }
}

impl TableMetrics {
    /// Whether any face of `family` was measured, so a caller can tell a person
    /// which text on the page the engine laid out and which it could not.
    pub fn covers(family: &str) -> bool {
        crate::font_adv_local::local_family_supported(family)
            || crate::font_adv::family_supported(family)
    }
}

/// Break one paragraph's text into the lines a box `width_pt` wide holds.
///
/// The whole point of the port in one call: given a metrics source, this is
/// the same break the renderer makes, and it runs anywhere. `runs` carries the
/// paragraph's own runs so a mixed-weight line is measured per run.
///
/// Returns None when the metrics source cannot measure the text -- the caller
/// then has to fall back to whatever its platform offers, and should say that
/// the answer is not the engine's.
pub fn break_paragraph(
    metrics: &dyn FaceMetrics,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    width_pt: f32,
    runs: &[crate::ir::SlideRun],
) -> Option<Vec<String>> {
    // One probe first: if the source cannot measure the paragraph at all,
    // say so rather than returning a wrap built out of fallbacks.
    master_units(metrics, text, fs, family, bold, italic, 0.0)?;
    let opts = WrapOpts {
        trim_trailing_space: true,
        char_wrap: true,
        hyphen_breaks: false,
    };
    Some(wrap_lines(
        text,
        width_pt,
        width_pt,
        1.0,
        &opts,
        |candidate, w_pt, _px, emitted| {
            let styles = RunStyles { runs, line_start: emitted };
            let mu = if runs.len() > 1 {
                master_units_runs(metrics, candidate, fs, family, bold, italic, &styles, true)
            } else {
                master_units(metrics, candidate, fs, family, bold, italic, para_spc(runs, true))
            };
            mu.map(|mu| fits(mu, w_pt)).unwrap_or(true)
        },
        |_| 0,
    ))
}

#[cfg(test)]
mod paragraph_tests {
    use super::*;

    fn run(text: &str) -> crate::ir::SlideRun {
        crate::ir::SlideRun {
            text: text.to_string(),
            font_size: None,
            bold: false,
            italic: false,
            underline: false,
            color: None,
            color_alpha: None,
            highlight: None,
            font_family: None,
            spacing: None,
        }
    }

    #[test]
    fn a_paragraph_breaks_on_the_tables_own_advances() {
        let text = "The quick brown fox jumps over the lazy dog";
        let runs = [run(text)];
        // Arial 12pt: the whole string is about 220pt, so a 120pt box breaks it.
        let got = break_paragraph(&TableMetrics, text, 12.0, "Arial", false, false, 120.0, &runs)
            .expect("Arial is in the table");
        assert!(got.len() > 1, "expected a wrap, got {got:?}");
        assert_eq!(got.concat(), text);
    }

    #[test]
    fn a_family_the_tables_do_not_cover_is_refused_not_guessed() {
        let text = "hello";
        let runs = [run(text)];
        assert!(break_paragraph(&TableMetrics, text, 12.0, "Bebas Neue", false, false, 50.0, &runs)
            .is_none());
    }

    #[test]
    fn a_box_wide_enough_keeps_the_paragraph_whole() {
        let text = "one two";
        let runs = [run(text)];
        let got = break_paragraph(&TableMetrics, text, 12.0, "Arial", false, false, 400.0, &runs)
            .unwrap();
        assert_eq!(got, vec![text]);
    }
}

#[cfg(test)]
mod baseline_tests {
    use super::*;

    struct NoTables;
    impl FaceMetrics for NoTables {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            None
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            false
        }
    }

    struct Says(f32);
    impl FaceMetrics for Says {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            None
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            false
        }
        fn baseline_offset_em(&self, _: &str) -> Option<f32> {
            Some(self.0)
        }
    }

    #[test]
    fn a_face_that_can_read_its_own_tables_outranks_the_offline_one() {
        let a = first_baseline_off(&NoTables, "Arial", 40.0, 1.0, true);
        let b = first_baseline_off(&Says(0.90), "Arial", 40.0, 1.0, true);
        assert!((a - b).abs() > 0.1, "{a} vs {b}");
    }

    #[test]
    fn an_unmeasured_family_falls_back_to_the_generic_split() {
        assert_eq!(table_baseline_offset_em("Bebas Neue"), 0.9685);
    }

    #[test]
    fn the_quarter_caps_the_descent_at_single_spacing() {
        // Arial at 40pt: pitch 48, natural descent 48 - 38.9096 = 9.0904,
        // quarter 12 -- the natural descent is shallower, so it stands.
        let got = first_baseline_off(&NoTables, "Arial", 40.0, 1.0, true);
        assert!((got - 38.9096).abs() < 0.01, "{got}");
    }

    #[test]
    fn above_single_spacing_the_quarter_becomes_a_floor() {
        // 120% of 40pt: pitch 48, n 1.2 -> quarter*n = 14.4 > natural 9.09.
        let got = first_baseline_off(&NoTables, "Arial", 40.0, 1.2, true);
        assert!((got - (48.0 * 1.2 - 14.4)).abs() < 0.01, "{got}");
    }

    #[test]
    fn the_old_rule_keeps_its_parenthesisation() {
        let fs = 40.0f32;
        let n = 1.2f32;
        let got = first_baseline_off(&NoTables, "Arial", fs, n, false);
        assert_eq!(got, 0.75 * (fs * 1.2 * n));
    }
}

#[cfg(test)]
mod coverage_tests {
    use super::*;

    #[test]
    fn the_wide_table_reaches_families_the_shared_one_does_not() {
        assert!(!crate::font_adv::family_supported("Montserrat"));
        assert!(TableMetrics::covers("Montserrat"));
        assert!(TableMetrics::covers("Barlow Light"));
    }

    #[test]
    fn a_style_the_machine_had_is_measured_apart_from_the_upright() {
        let m = TableMetrics;
        let up = m.advance_em("Montserrat", false, false, 'a').unwrap();
        let bd = m.advance_em("Montserrat", true, false, 'a').unwrap();
        assert!(bd > up, "bold {bd} should advance more than regular {up}");
    }

    #[test]
    fn a_family_nobody_measured_is_still_refused() {
        assert!(!TableMetrics::covers("Zzyzx Nonexistent"));
        assert!(TableMetrics.advance_em("Zzyzx Nonexistent", false, false, 'a').is_none());
    }

    #[test]
    fn a_missing_style_falls_back_to_the_upright_face() {
        // Fira Sans was installed regular-only; asking bold must still answer.
        let m = TableMetrics;
        assert!(m.advance_em("Fira Sans", true, false, 'a').is_some());
    }
}

/// Where a paragraph's lines start, and how wide each may run.
///
/// All offsets are relative to P0, the inner left edge of the text area.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct IndentGeometry {
    /// Left offset of the FIRST line.
    pub first_x: f32,
    /// Left offset of every line after the first.
    pub rest_x: f32,
    /// Where a bullet marker is set, which a hanging indent puts left of the
    /// text it belongs to.
    pub marker_x: f32,
    /// The width the first line may use.
    pub first_width: f32,
    /// The width every later line may use.
    pub rest_width: f32,
}

/// Resolve `marL` and `indent` into line offsets and the widths that go with
/// them (Spec #8, measured).
///
///     para_left = P0 + marL
///     indent > 0   text_1st = para_left + indent,  marker = para_left
///     indent <= 0  text_1st = max(para_left, P0 - indent),
///                  marker   = text_1st + indent
///     continuation lines = para_left
///
/// ★The WIDTH matters as much as the offset. PowerPoint wraps every line
/// against the same RIGHT EDGE, so a line's usable width is the inner width
/// less ITS OWN left offset -- and a hanging indent makes line 0's offset
/// differ from the rest. Probe `wrapwidth` (2026-08-19, 7 arms): with
/// marL 18 / indent -18 and no bullet, line 0 starts at the inner left and
/// runs 232.07pt while the continuations start 18pt in and run at most 221.26,
/// both stopping at the same 316.8. Wrapping everything at the full width and
/// shifting afterwards lets a continuation run past the inset.
///
/// `apply_widths` off restores the older behaviour of wrapping every line at
/// the full inner width.
pub fn indent_geometry(
    inner_width: f32,
    mar_l: f32,
    indent: f32,
    apply_widths: bool,
) -> IndentGeometry {
    let (first_x, marker_x) = if indent > 0.0 {
        (mar_l + indent, mar_l)
    } else {
        let t = mar_l.max(-indent);
        (t, t + indent)
    };
    IndentGeometry {
        first_x,
        rest_x: mar_l,
        marker_x,
        first_width: (inner_width - if apply_widths { first_x } else { 0.0 }).max(1.0),
        rest_width: (inner_width - if apply_widths { mar_l } else { 0.0 }).max(1.0),
    }
}

/// The baseline-to-baseline step between two lines of DIFFERENT size.
///
/// A paragraph whose lines are all one size steps by `fs * 1.2 * n` exactly,
/// and must keep doing so down to the float association. When the sizes differ
/// the step is not that: `cursor` between lines is the BOTTOM of the previous
/// line box, and the next baseline sits its own ascent below it.
///
/// Probe `mixedpitch` (4 faces x 8 size pairs, 2026-08-18) fits
///
///     step = d * prev_size + a * next_size,   a + d = 1.2004
///
/// with d = 0.2284 (Arial) / 0.2322 (Georgia) / 0.2636 (Calibri) / 0.2088
/// (Verdana) -- each within 0.0015 of that face's own
/// `1.2 * descent / (ascent + descent)`, i.e. the 1.2 line height split by the
/// FONT's ascent:descent ratio. d28's title is 55pt then 66pt: PowerPoint steps
/// 159px at 150dpi, a flat rule gives 137px, and this gives 159.7px.
///
/// Expressed with the ascent this module already computes, the step from a line
/// of `prev` to a line of `next` is
///
///     (1.2 * prev * n - ascent(prev)) + ascent(next)
///
/// which is the previous box's descent plus the next line's ascent.
pub fn mixed_pitch_step(
    metrics: &dyn FaceMetrics,
    family: &str,
    prev: f32,
    next: f32,
    n: f32,
    joined_rule: bool,
) -> f32 {
    let ascent = |size: f32| first_baseline_off(metrics, family, size, n, joined_rule);
    (1.2 * prev * n - ascent(prev)) + ascent(next)
}

#[cfg(test)]
mod geometry_tests {
    use super::*;

    struct Arial;
    impl FaceMetrics for Arial {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            None
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            false
        }
    }

    #[test]
    fn a_positive_indent_pushes_the_first_line_in_and_leaves_the_marker() {
        let g = indent_geometry(200.0, 18.0, 18.0, true);
        assert_eq!(g.first_x, 36.0);
        assert_eq!(g.rest_x, 18.0);
        assert_eq!(g.marker_x, 18.0);
    }

    #[test]
    fn a_hanging_indent_pulls_the_marker_left_of_the_text() {
        let g = indent_geometry(200.0, 18.0, -18.0, true);
        assert_eq!(g.first_x, 18.0);
        assert_eq!(g.marker_x, 0.0);
        assert_eq!(g.rest_x, 18.0);
    }

    #[test]
    fn every_line_stops_at_the_same_right_edge() {
        // marL 18 / indent -18: line 0 starts at 18 and the rest at 18 too,
        // so both stop at 200 -- the widths differ only when the offsets do.
        let g = indent_geometry(200.0, 18.0, 36.0, true);
        assert_eq!(g.first_x + g.first_width, 200.0);
        assert_eq!(g.rest_x + g.rest_width, 200.0);
    }

    #[test]
    fn widths_off_gives_every_line_the_whole_inner_width() {
        let g = indent_geometry(200.0, 18.0, 36.0, false);
        assert_eq!(g.first_width, 200.0);
        assert_eq!(g.rest_width, 200.0);
    }

    #[test]
    fn a_same_size_step_is_the_flat_advance() {
        let step = mixed_pitch_step(&Arial, "Arial", 40.0, 40.0, 1.0, true);
        assert!((step - 40.0 * 1.2).abs() < 1e-4, "{step}");
    }

    #[test]
    fn a_growing_step_is_the_old_descent_plus_the_new_ascent() {
        // d28's 55 -> 66pt title: a flat rule would step 66, PowerPoint steps
        // about 76.4pt (159px at 150dpi).
        let step = mixed_pitch_step(&Arial, "Arial", 55.0, 66.0, 1.0, true);
        assert!(step > 55.0 * 1.2, "{step} should exceed the previous flat step");
        assert!((step - 76.4).abs() < 1.5, "{step}");
    }
}

/// The exact line height a paragraph asks for in POINTS, if it asks in points.
///
/// S-LNSPCPTS (2026-08-27): `a:lnSpc/a:spcPts` states a height outright and
/// outranks any multiple. S-LNSPCROUND (2026-08-29): PowerPoint rounds that
/// height to a whole point before using it -- read off its own PDF, where a
/// 12.984pt request steps 13.
pub fn exact_line_pt(
    para: &crate::ir::SlideParagraph,
    honour_points: bool,
    round_points: bool,
) -> Option<f32> {
    let v = para.line_spacing_pts.filter(|v| *v > 0.0 && honour_points)?;
    Some(if round_points { v.round() } else { v })
}

/// The size a paragraph is set at.
///
/// A run's explicit `sz` wins -- the LARGEST of them, since one big word sets
/// the line -- then whatever the placeholder chain inherited, then 18pt.
///
/// An EMPTY paragraph has no run to ask, and is sized by its paragraph MARK
/// (`a:endParaRPr`) instead; failing that it keeps the previous paragraph's
/// size, because an empty line between two paragraphs is as tall as the text
/// around it, not as tall as the default.
pub fn paragraph_font_size(
    para: &crate::ir::SlideParagraph,
    inherited: Option<f32>,
    prev_fs: Option<f32>,
    empty_para_rule: bool,
) -> f32 {
    let explicit = para
        .runs
        .iter()
        .filter_map(|r| r.font_size)
        .fold(None, |acc: Option<f32>, x| Some(acc.map_or(x, |a: f32| a.max(x))));
    if empty_para_rule && para.runs.iter().all(|r| r.text.is_empty()) {
        if let Some(fs) = para.end_para_size.or(prev_fs) {
            return fs;
        }
    }
    explicit.or(inherited).unwrap_or(18.0)
}

/// Where a line starts inside the width it may use, for its alignment.
///
/// Centre and right are measured from the width the line actually occupies, so
/// a line measured in a face it is not drawn in starts in the wrong place --
/// which is why the width handed in must come from the same per-run
/// measurement the wrap uses (S-RUNALIGN).
pub fn align_offset(alignment: crate::ir::SlideAlignment, area_w: f32, line_w: f32) -> f32 {
    match alignment {
        crate::ir::SlideAlignment::Center => ((area_w - line_w) / 2.0).max(0.0),
        crate::ir::SlideAlignment::Right => (area_w - line_w).max(0.0),
        _ => 0.0,
    }
}

#[cfg(test)]
mod size_tests {
    use super::*;
    use crate::ir::{SlideAlignment, SlideParagraph, SlideRun};

    /// A paragraph with nothing set, for the tests that only vary one field.
    pub(super) fn para_for_spacing() -> SlideParagraph {
        para(vec![])
    }

    fn para(runs: Vec<SlideRun>) -> SlideParagraph {
        SlideParagraph {
            runs,
            alignment: None,
            line_spacing: None,
            line_spacing_pts: None,
            space_before: None,
            space_after: None,
            lvl: 0,
            end_para_size: None,
            mar_l: None,
            indent: None,
            bullet: crate::ir::SlideBullet::default(),
        }
    }

    fn run(text: &str, size: Option<f32>) -> SlideRun {
        SlideRun {
            text: text.to_string(),
            font_size: size,
            bold: false,
            italic: false,
            underline: false,
            color: None,
            color_alpha: None,
            highlight: None,
            font_family: None,
            spacing: None,
        }
    }

    #[test]
    fn the_largest_explicit_size_sets_the_paragraph() {
        let p = para(vec![run("a", Some(12.0)), run("b", Some(40.0))]);
        assert_eq!(paragraph_font_size(&p, Some(18.0), None, true), 40.0);
    }

    #[test]
    fn an_inherited_size_is_used_when_no_run_states_one() {
        let p = para(vec![run("a", None)]);
        assert_eq!(paragraph_font_size(&p, Some(32.0), None, true), 32.0);
        assert_eq!(paragraph_font_size(&p, None, None, true), 18.0);
    }

    #[test]
    fn an_empty_paragraph_keeps_the_size_around_it() {
        let p = para(vec![run("", None)]);
        assert_eq!(paragraph_font_size(&p, Some(18.0), Some(44.0), true), 44.0);
        // With the rule off it falls back to what it inherited.
        assert_eq!(paragraph_font_size(&p, Some(18.0), Some(44.0), false), 18.0);
    }

    #[test]
    fn a_stated_line_height_is_rounded_to_a_whole_point() {
        let mut p = para(vec![run("a", None)]);
        p.line_spacing_pts = Some(12.984);
        assert_eq!(exact_line_pt(&p, true, true), Some(13.0));
        assert_eq!(exact_line_pt(&p, true, false), Some(12.984));
        assert_eq!(exact_line_pt(&p, false, true), None);
    }

    #[test]
    fn centre_and_right_are_measured_from_the_line_width() {
        assert_eq!(align_offset(SlideAlignment::Center, 100.0, 40.0), 30.0);
        assert_eq!(align_offset(SlideAlignment::Right, 100.0, 40.0), 60.0);
        assert_eq!(align_offset(SlideAlignment::Left, 100.0, 40.0), 0.0);
    }

    #[test]
    fn a_line_wider_than_its_area_is_not_pushed_left_of_it() {
        assert_eq!(align_offset(SlideAlignment::Center, 100.0, 140.0), 0.0);
        assert_eq!(align_offset(SlideAlignment::Right, 100.0, 140.0), 0.0);
    }
}

/// The outline level a paragraph inherits, master first and the placeholder's
/// own `a:lstStyle` over it, field by field.
///
/// Both lists are indexed by the paragraph's `lvl`, clamped to what the list
/// actually holds -- a deck may declare fewer levels than a paragraph asks for.
///
/// ★The LAYOUT placeholder's list must be resolved HERE, not only where the
/// text is drawn. Resolving it late wrapped d24's title at the master's 18pt
/// and then drew it at the layout's 60pt, so the line ran off its box instead
/// of breaking into the three PowerPoint gives it.
pub fn resolve_level(
    master: &[crate::ir::MasterStyleLevel],
    ph_levels: &[crate::ir::MasterStyleLevel],
    lvl: u32,
) -> crate::ir::MasterStyleLevel {
    let mut m = if master.is_empty() {
        crate::ir::MasterStyleLevel::default()
    } else {
        master[(lvl as usize).min(master.len() - 1)].clone()
    };
    if !ph_levels.is_empty() {
        let l = &ph_levels[(lvl as usize).min(ph_levels.len() - 1)];
        if l.font_size.is_some() {
            m.font_size = l.font_size;
        }
        if l.color.is_some() {
            m.color = l.color.clone();
        }
        if l.algn.is_some() {
            m.algn = l.algn;
        }
        if l.line_spacing.is_some() {
            m.line_spacing = l.line_spacing;
        }
        if l.bold.is_some() {
            m.bold = l.bold;
        }
    }
    m
}

/// The line-spacing multiple a paragraph is set at.
///
/// A height stated in POINTS wins and is carried as the equivalent multiple of
/// the 1.2 default, so `fs * 1.2 * n` reproduces it and everything downstream --
/// the first-baseline rule, the space-before fraction, the mixed-pitch step --
/// keeps working unchanged. Otherwise the paragraph's own percentage, then the
/// placeholder chain's, then single.
///
/// d36 s1's title asks 107.25pt against 97.58pt of type, so n = 0.9159 and the
/// step becomes 107.25pt against PowerPoint's measured 107.06. Stepping the
/// flat 117.10 put its three baselines 9.91 / 19.97 / 29.66pt low.
pub fn line_spacing_multiple(
    para: &crate::ir::SlideParagraph,
    level: &crate::ir::MasterStyleLevel,
    fs: f32,
    exact_pt: Option<f32>,
) -> f32 {
    exact_pt
        .filter(|_| fs > 0.0)
        .map(|pts| pts / (fs * 1.2))
        .or(para.line_spacing)
        .or(level.line_spacing)
        .unwrap_or(1.0)
}

#[cfg(test)]
mod level_tests {
    use super::*;
    use crate::ir::MasterStyleLevel;

    fn lvl(size: Option<f32>, spacing: Option<f32>) -> MasterStyleLevel {
        MasterStyleLevel {
            font_size: size,
            line_spacing: spacing,
            ..MasterStyleLevel::default()
        }
    }

    #[test]
    fn the_placeholders_own_list_overrides_the_master_field_by_field() {
        let master = [lvl(Some(18.0), Some(1.0))];
        let ph = [lvl(Some(60.0), None)];
        let got = resolve_level(&master, &ph, 0);
        assert_eq!(got.font_size, Some(60.0), "the layout's size wins");
        assert_eq!(got.line_spacing, Some(1.0), "and the master keeps what it did not override");
    }

    #[test]
    fn a_level_past_the_end_of_the_list_takes_the_last_one() {
        let master = [lvl(Some(18.0), None), lvl(Some(14.0), None)];
        assert_eq!(resolve_level(&master, &[], 9).font_size, Some(14.0));
    }

    #[test]
    fn no_lists_at_all_is_the_default_level() {
        assert_eq!(resolve_level(&[], &[], 0).font_size, None);
    }

    #[test]
    fn a_height_in_points_becomes_the_multiple_that_reproduces_it() {
        let p = super::size_tests::para_for_spacing();
        // 107.25pt asked against 97.58pt of type at 1.2.
        let n = line_spacing_multiple(&p, &lvl(None, None), 97.58 / 1.2, Some(107.25));
        assert!((n * (97.58 / 1.2) * 1.2 - 107.25).abs() < 1e-3, "{n}");
    }

    #[test]
    fn without_a_stated_height_the_paragraph_then_the_level_decides() {
        let p = super::size_tests::para_for_spacing();
        assert_eq!(line_spacing_multiple(&p, &lvl(None, Some(0.9)), 60.0, None), 0.9);
        assert_eq!(line_spacing_multiple(&p, &lvl(None, None), 60.0, None), 1.0);
    }
}
