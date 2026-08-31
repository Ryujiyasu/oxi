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

    /// Whether this source can serve `family` under its own name.
    ///
    /// The default is `true` -- "I cannot tell" must mean "keep the name the
    /// deck asked for". A source that substitutes on a guess would rename text
    /// PowerPoint draws in its own face, which is worse than leaving a name
    /// alone that nothing can serve.
    fn resolves(&self, _family: &str) -> bool {
        true
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
                // The run's OWN declaration, not `is_bold(bold)`: the `bold`
                // here is the PARAGRAPH's resolved weight, so inheriting it
                // would make a silent run bold merely because a sibling run
                // is. What a silent run inherits is the LEVEL, and the level
                // does not reach this far down.
                run.bold.unwrap_or(false),
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
            bold: if bold { Some(true) } else { None },
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

    fn baseline_offset_em(&self, family: &str) -> Option<f32> {
        // ★The face's own figure, not one number for every face. The renderer
        // asks GDI for `1.2 * asc / (asc + desc)`; this is the same arithmetic
        // read from the file at build time. Falling back to the six-family
        // table put Nunito's first baseline 0.063 em low and Barlow's 0.030 em
        // high, on every shape they set.
        crate::font_adv_local::local_baseline_offset_em(family)
    }

    // ★`resolves` is deliberately left at its default of `true`.
    //
    // Answering it with `covers` looked right and is wrong: a family the
    // tables do not carry may still be one the DECK EMBEDS, and PowerPoint
    // then draws the deck's own part. Substituting Calibri for it would rename
    // text PowerPoint sets in its real face -- and worse, it would make the
    // shape look measurable, because Calibri is in the tables. The honest
    // answer is to keep the name and let `break_paragraph` return None, which
    // marks the shape incomplete instead of laying out a fiction.
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
    // `<a:br/>` arrives as a newline in the run stream and ends the line where
    // it stands. Each segment wraps on its own, and the newline is kept on the
    // END of the line it closed so the caller's character accounting -- which
    // maps lines back to runs -- still lines up.
    //
    // ★It has to be handled before anything measures the text: a newline has
    // no advance, so the probe below refuses the whole paragraph on account of
    // it. That is what made d06 decline shapes whose face it could measure
    // perfectly well.
    if text.contains('\n') {
        let mut out: Vec<String> = Vec::new();
        let mut base = 0usize;
        for (si, seg) in text.split('\n').enumerate() {
            if si > 0 {
                match out.last_mut() {
                    Some(last) => last.push('\n'),
                    None => out.push("\n".to_string()),
                }
            }
            let mut part =
                break_segment(metrics, seg, fs, family, bold, italic, width_pt, runs, base)?;
            base += seg.chars().count() + 1;
            if part.is_empty() {
                // An empty segment is a blank line, not nothing.
                out.push(String::new());
            } else {
                out.append(&mut part);
            }
        }
        return Some(out);
    }
    break_segment(metrics, text, fs, family, bold, italic, width_pt, runs, 0)
}

/// One run of text with no hard break in it (see [`break_paragraph`]).
///
/// `base` is how many characters of the paragraph came before this segment, so
/// the run styles a candidate line is measured with are the right ones.
#[allow(clippy::too_many_arguments)]
fn break_segment(
    metrics: &dyn FaceMetrics,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    width_pt: f32,
    runs: &[crate::ir::SlideRun],
    base: usize,
) -> Option<Vec<String>> {
    // One probe first: if the source cannot measure the paragraph at all,
    // say so rather than returning a wrap built out of fallbacks.
    master_units(metrics, text, fs, family, bold, italic, 0.0)?;
    let opts = WrapOpts {
        trim_trailing_space: true,
        char_wrap: true,
        // ★A hyphen IS a break site for PowerPoint, and this said otherwise.
        // The renderer has had it since the `charwrap` probe, whose
        // `alpha-beta-gamma-...` in a 165.6pt box came back broken AFTER the
        // hyphens; the port hardcoded false, so d19 slide 37's
        // `slidescarnival.com/extra-free-resources-icons-and-maps` could not
        // start on the line that says `Find more icons at `, and every line
        // after it fell one line low.
        hyphen_breaks: true,
    };
    Some(wrap_lines(
        text,
        width_pt,
        width_pt,
        1.0,
        &opts,
        |candidate, w_pt, _px, emitted| {
            let styles = RunStyles { runs, line_start: base + emitted };
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
            bold: None,
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

    #[test]
    fn a_hard_break_ends_the_line_it_stands_on() {
        // `<a:br/>` reaches the IR as a newline in the run stream.
        let got = break_paragraph(
            &TableMetrics, "Imani Jackson\nJOB TITLE", 12.0, "Arial", false, false, 400.0,
            &[run("Imani Jackson\nJOB TITLE")],
        )
        .expect("a newline must not make the paragraph unmeasurable");
        assert_eq!(got, vec!["Imani Jackson\n".to_string(), "JOB TITLE".to_string()]);
    }

    #[test]
    fn each_side_of_a_hard_break_still_wraps() {
        let text = "The quick brown fox jumps\nover the lazy dog again and again";
        let got = break_paragraph(
            &TableMetrics, text, 12.0, "Arial", false, false, 90.0, &[run(text)],
        )
        .expect("measurable");
        assert!(got.len() > 2, "{got:?}");
        // Exactly one line carries the break, and it is the one it closed.
        assert_eq!(got.iter().filter(|l| l.ends_with('\n')).count(), 1);
    }

    #[test]
    fn a_break_with_nothing_after_it_leaves_a_blank_line() {
        let got = break_paragraph(
            &TableMetrics, "Alone\n", 12.0, "Arial", false, false, 400.0, &[run("Alone\n")],
        )
        .expect("measurable");
        assert_eq!(got, vec!["Alone\n".to_string(), String::new()]);
    }

    #[test]
    fn the_tables_reach_past_ascii_now() {
        // The one character a Western deck cannot avoid: U+2019, in 21 of the
        // corpora's decks. A face that lacked it made the whole shape decline.
        for ch in ['\u{2019}', '\u{00AE}', '\u{2014}', '\u{00E1}'] {
            assert!(
                TableMetrics.advance_em("Arial", false, false, ch).is_some(),
                "Arial should carry U+{:04X}",
                ch as u32
            );
        }
    }

    #[test]
    fn a_character_no_face_has_is_still_refused() {
        // An emoji is a different face's job; answering here would advance a
        // glyph that is not in the font.
        assert_eq!(
            TableMetrics.advance_em("Arial", false, false, '\u{1F600}'),
            None
        );
    }

    #[test]
    fn a_paragraph_with_a_curly_quote_can_be_broken() {
        let text = "That\u{2019}s a lot of money";
        assert!(
            break_paragraph(
                &TableMetrics, text, 12.0, "Arial", false, false, 400.0, &[run(text)]
            )
            .is_some(),
            "a curly apostrophe must not sink the paragraph"
        );
    }

    #[test]
    fn a_hyphen_opens_a_break_site() {
        // PowerPoint puts `Find more icons at slidescarnival.com/extra-` on
        // one line and the rest on the next; breaking only at spaces sends the
        // whole URL down a line and takes everything after it with it.
        let text = "Find more icons at slidescarnival.com/extra-free-resources";
        let got = break_paragraph(
            &TableMetrics, text, 9.0, "Arial", false, false, 120.0, &[run(text)],
        )
        .expect("measurable");
        assert!(got.iter().any(|l| l.ends_with('-')), "{got:?}");
        // and the break lands AFTER the hyphen, not at the last character
        // that would have fitted.
        for line in &got {
            if let Some(rest) = line.strip_suffix('-') {
                assert!(!rest.is_empty());
            }
        }
    }

    #[test]
    fn a_word_with_no_hyphen_still_breaks_only_at_spaces() {
        let text = "alpha beta gamma delta";
        let got = break_paragraph(
            &TableMetrics, text, 9.0, "Arial", false, false, 40.0, &[run(text)],
        )
        .expect("measurable");
        assert!(got.iter().all(|l| !l.trim_end().ends_with('-')), "{got:?}");
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
    fn a_missing_style_is_refused_rather_than_served_from_the_upright() {
        // Fira Sans was installed regular-only. Answering its bold request
        // from the regular face is the kind of quiet wrongness this whole
        // module exists to avoid: Merriweather Bold is ~1% wider than its
        // Regular, and serving one for the other put d08's 38pt titles up to
        // 9pt out while the layout still reported itself complete.
        let m = TableMetrics;
        assert!(m.advance_em("Fira Sans", false, false, 'a').is_some());
        assert!(m.advance_em("Fira Sans", true, false, 'a').is_none());
    }

    #[test]
    fn a_measured_family_supplies_its_own_baseline() {
        let m = TableMetrics;
        // Nunito sits far from the fallback and Barlow far the other way; both
        // must come from the face rather than from the six-family table.
        let nunito = m.baseline_offset_em("Nunito").expect("measured");
        let barlow = m.baseline_offset_em("Barlow").expect("measured");
        assert!((nunito - 0.88944).abs() < 1e-4, "{nunito}");
        assert!((barlow - 1.0).abs() < 1e-4, "{barlow}");
        assert!(m.baseline_offset_em("Zzyzx Nonexistent").is_none());
    }

    #[test]
    fn the_first_baseline_follows_the_face() {
        let m = TableMetrics;
        let nunito = first_baseline_off(&m, "Nunito", 36.0, 1.0, true);
        let barlow = first_baseline_off(&m, "Barlow", 36.0, 1.0, true);
        assert!(barlow > nunito, "{barlow} vs {nunito}");
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
/// ```text
/// para_left = P0 + marL
/// indent > 0   text_1st = para_left + indent,  marker = para_left
/// indent <= 0  text_1st = max(para_left, P0 - indent),
///              marker   = text_1st + indent
/// continuation lines = para_left
/// ```
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
/// ```text
/// step = d * prev_size + a * next_size,   a + d = 1.2004
/// ```
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
/// ```text
/// (1.2 * prev * n - ascent(prev)) + ascent(next)
/// ```
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
    // ★Not clamped at zero. A line WIDER than its area hangs out of it on
    // both sides when centred, and off the left when right-aligned -- the
    // offset simply goes negative. `overwide` probe, all five arms:
    //
    //   algn  box   truth      unclamped   clamped
    //   ctr    40   272.930    272.808     300.000
    //   r      40   245.690    245.615     300.000
    //   l      40   300.050    300.000     300.000   (agree)
    //   ctr   300   402.910    402.808     402.808   (agree)
    //   r     300   505.750    505.615     505.615   (agree)
    //
    // and seven slides of d32, which put a 167.65pt bullet in a 33.24pt
    // centred box: PowerPoint draws it 12.7pt left of the box, which
    // `(area - line) / 2` predicts to within 0.085pt on every one of them.
    //
    // `OXI_OVERWIDE_DISABLE` restores the clamp.
    static CLAMP: std::sync::OnceLock<bool> = std::sync::OnceLock::new();
    let clamp = *CLAMP.get_or_init(|| std::env::var("OXI_OVERWIDE_DISABLE").is_ok());
    let off = match alignment {
        crate::ir::SlideAlignment::Center => (area_w - line_w) / 2.0,
        crate::ir::SlideAlignment::Right => area_w - line_w,
        _ => 0.0,
    };
    if clamp {
        off.max(0.0)
    } else {
        off
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
            bold: None,
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
    fn a_line_wider_than_its_area_hangs_out_of_it() {
        // This asserted 0.0 for both until 2026-08-31, when the `overwide`
        // probe and seven slides of d32 said PowerPoint lets the line hang
        // out rather than pinning it to the left edge.
        assert_eq!(align_offset(SlideAlignment::Center, 100.0, 140.0), -20.0);
        assert_eq!(align_offset(SlideAlignment::Right, 100.0, 140.0), -40.0);
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
/// The face a paragraph's level chain supplies, if any.
///
/// The placeholder's own list first, then the master's -- the same order the
/// renderer's draw path walks, so the wrap measures the face that will be
/// drawn.
///
/// ★This is deliberately NOT part of `resolve_level`. That function overlays a
/// placeholder's level onto the master's and never carried the face, which is
/// how d15 slide 1's title came to be measured in the theme's Arial when the
/// level says Barlow and the truth PDF sets it in Barlow Bold.
pub fn level_family(
    master: &[crate::ir::MasterStyleLevel],
    ph_levels: &[crate::ir::MasterStyleLevel],
    lvl: u32,
) -> Option<String> {
    let lvl = lvl as usize;
    [ph_levels, master].into_iter().find_map(|levels| {
        levels
            .get(lvl.min(levels.len().saturating_sub(1)))
            .and_then(|l| l.font_family.clone())
    })
}

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

/// The family a run is actually set in.
///
/// PowerPoint falls back to **Calibri** for a typeface nothing can serve --
/// measured over Mali / Jua / a name invented for the probe, and it is not the
/// theme font (d19 asks for Mali over a theme that says something else and
/// still gets Calibri).
///
/// `substitute` off keeps every requested name, which is the arm that shows
/// what the substitution is worth.
pub fn effective_family(metrics: &dyn FaceMetrics, requested: &str, substitute: bool) -> String {
    if !substitute || requested.is_empty() {
        return requested.to_string();
    }
    if metrics.resolves(requested) {
        requested.to_string()
    } else {
        "Calibri".to_string()
    }
}

#[cfg(test)]
mod family_tests {
    use super::*;

    struct Has(&'static [&'static str]);
    impl FaceMetrics for Has {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            None
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            false
        }
        fn resolves(&self, family: &str) -> bool {
            self.0.iter().any(|f| f.eq_ignore_ascii_case(family))
        }
    }

    struct CannotTell;
    impl FaceMetrics for CannotTell {
        fn advance_em(&self, _: &str, _: bool, _: bool, _: char) -> Option<f32> {
            None
        }
        fn has_all_glyphs(&self, _: &str, _: bool, _: bool, _: &str) -> bool {
            false
        }
    }

    #[test]
    fn a_family_the_source_can_serve_keeps_its_name() {
        let m = Has(&["Arial"]);
        assert_eq!(effective_family(&m, "Arial", true), "Arial");
    }

    #[test]
    fn a_family_nothing_serves_becomes_calibri() {
        let m = Has(&["Arial"]);
        assert_eq!(effective_family(&m, "Mali", true), "Calibri");
    }

    #[test]
    fn a_source_that_cannot_tell_leaves_the_name_alone() {
        assert_eq!(effective_family(&CannotTell, "Mali", true), "Mali");
    }

    #[test]
    fn substitution_off_keeps_every_name() {
        let m = Has(&["Arial"]);
        assert_eq!(effective_family(&m, "Mali", false), "Mali");
    }

    #[test]
    fn an_empty_request_stays_empty() {
        let m = Has(&[]);
        assert_eq!(effective_family(&m, "", true), "");
    }
}

/// Decimal to uppercase Roman, greedily (1 = I .. 3999 = MMMCMXCIX).
pub fn to_roman(mut n: u32) -> String {
    const ROMAN: [(u32, &str); 13] = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
        (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ];
    let mut out = String::new();
    for (v, s) in ROMAN {
        while n >= v {
            out.push_str(s);
            n -= v;
        }
    }
    out
}

/// Decimal to a spreadsheet-style letter label (1 = A, 26 = Z, 27 = AA).
pub fn to_alpha(mut n: u32) -> String {
    let mut s = String::new();
    while n > 0 {
        let d = ((n - 1) % 26) as u8;
        s.insert(0, (b'A' + d) as char);
        n = (n - 1) / 26;
    }
    s
}

/// The text of an auto-numbered bullet, for `a:buAutoNum/@type` and a count.
pub fn autonum_text(kind: &str, n: u32) -> String {
    let body = if kind.starts_with("romanUc") {
        to_roman(n)
    } else if kind.starts_with("romanLc") {
        to_roman(n).to_lowercase()
    } else if kind.starts_with("alphaUc") {
        to_alpha(n)
    } else if kind.starts_with("alphaLc") {
        to_alpha(n).to_lowercase()
    } else {
        n.to_string()
    };
    if kind.ends_with("ParenBoth") {
        format!("({body})")
    } else if kind.ends_with("ParenR") {
        format!("{body})")
    } else if kind.ends_with("Period") {
        format!("{body}.")
    } else {
        body
    }
}

/// The next number in an auto-numbered list, and the counter state to keep.
///
/// Spec #11: the counter is per (level, kind). The sequence CONTINUES while
/// `start_at` stays the same -- absent staying absent, or the same value -- and
/// starts a NEW list whenever it changes, present to absent or to a different
/// number, resetting to `start_at` or 1.
///
/// Word truth: `autonum4` G with [None][5][None] renders 1, 5, 1 -- the second
/// None restarts, because its startAt differs from the [5] list's. And
/// `autonum` p1 level 0 runs 1,2,3..4 across interleaved levels, because the
/// (lvl, kind) key never changed its startAt.
///
/// `state` is the counter's `(last_start_at, count)`; the returned state
/// replaces it.
pub fn next_autonum(
    state: (Option<u32>, u32),
    start_at: Option<u32>,
) -> (u32, (Option<u32>, u32)) {
    let (last_start, count) = state;
    let n = if count == 0 || last_start != start_at {
        start_at.unwrap_or(1)
    } else {
        count
    };
    (n, (start_at, n + 1))
}

/// Where a marker sits and how far it pushes the first line.
///
/// The marker is set at the hanging-indent position, and the first line starts
/// after it when the marker is wider than the indent leaves room for --
/// otherwise a long number would overlap its own text.
pub fn marker_push(first_x: f32, marker_x: f32, marker_width: f32) -> f32 {
    first_x.max(marker_x + marker_width)
}

#[cfg(test)]
mod marker_tests {
    use super::*;

    #[test]
    fn roman_and_alpha_count_the_usual_way() {
        assert_eq!(to_roman(1), "I");
        assert_eq!(to_roman(4), "IV");
        assert_eq!(to_roman(1994), "MCMXCIV");
        assert_eq!(to_alpha(1), "A");
        assert_eq!(to_alpha(26), "Z");
        assert_eq!(to_alpha(27), "AA");
    }

    #[test]
    fn the_kind_decides_the_letters_and_the_punctuation() {
        assert_eq!(autonum_text("arabicPeriod", 3), "3.");
        assert_eq!(autonum_text("romanLcParenBoth", 4), "(iv)");
        assert_eq!(autonum_text("alphaUcParenR", 2), "B)");
        assert_eq!(autonum_text("arabicPlain", 7), "7");
        assert_eq!(autonum_text("somethingUnknown", 7), "7");
    }

    #[test]
    fn a_list_continues_while_its_start_stays_put() {
        let (n1, st) = next_autonum((None, 0), None);
        assert_eq!(n1, 1);
        let (n2, st) = next_autonum(st, None);
        assert_eq!(n2, 2);
        let (n3, _) = next_autonum(st, None);
        assert_eq!(n3, 3);
    }

    #[test]
    fn changing_the_start_begins_a_new_list() {
        // autonum4's G: [None] [5] [None] renders 1, 5, 1.
        let (a, st) = next_autonum((None, 0), None);
        let (b, st) = next_autonum(st, Some(5));
        let (c, _) = next_autonum(st, None);
        assert_eq!((a, b, c), (1, 5, 1));
    }

    #[test]
    fn a_marker_wider_than_its_indent_pushes_the_first_line() {
        assert_eq!(marker_push(18.0, 0.0, 30.0), 30.0);
        assert_eq!(marker_push(18.0, 0.0, 10.0), 18.0);
    }
}

/// One stretch of a line set in a single face.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct LineSegment {
    pub text: String,
    pub family: String,
    pub font_size: f32,
    pub bold: bool,
    pub italic: bool,
    /// Characters of the LINE before this segment.
    pub start: usize,
}

/// Split a line into the stretches its runs set it in.
///
/// ★A paragraph is not one face. d02 slide 2 opens every bullet with a bold
/// `Alternative:` and continues in the regular weight; the layout collapsed
/// that to `any(bold)` and measured all 80 characters bold, which put the last
/// one 29pt past where PowerPoint drew it. The BREAK was already per-run
/// (`master_units_runs`); it was the placing that was not.
///
/// `line_start` is how many characters of the paragraph precede this line.
/// `level_bold` and `level_italic` are what the paragraph's level says, which a
/// run that does not mention weight or slant still inherits. The paragraph's
/// own values fill in past the last run, which is what a line ending in
/// generated text hits.
pub fn line_segments(
    runs: &[crate::ir::SlideRun],
    line_start: usize,
    text: &str,
    fs: f32,
    family: &str,
    level_bold: bool,
    level_italic: bool,
) -> Vec<LineSegment> {
    let mut out: Vec<LineSegment> = Vec::new();
    for (i, ch) in text.chars().enumerate() {
        let at = line_start + i;
        let mut seen = 0usize;
        let mut style = (fs, family.to_string(), level_bold, level_italic);
        for run in runs {
            let n = run.text.chars().count();
            if at < seen + n {
                style = (
                    run.font_size.unwrap_or(fs),
                    run.font_family.clone().unwrap_or_else(|| family.to_string()),
                    run.is_bold(level_bold),
                    run.italic || level_italic,
                );
                break;
            }
            seen += n;
        }
        match out.last_mut() {
            Some(last)
                if last.font_size == style.0
                    && last.family == style.1
                    && last.bold == style.2
                    && last.italic == style.3 =>
            {
                last.text.push(ch);
            }
            _ => out.push(LineSegment {
                text: ch.to_string(),
                family: style.1,
                font_size: style.0,
                bold: style.2,
                italic: style.3,
                start: i,
            }),
        }
    }
    out
}

/// Whether the level chain asks for italic.
///
/// ★A LEVEL can ask for slant just as it asks for weight, and the layout was
/// only consulting the runs. d15 slide 5's quotation carries `i="1"` on its
/// layout's `lvl1pPr/defRPr` and nothing on any run, and PowerPoint sets it in
/// Barlow BOLD ITALIC -- whose advances are about 3% narrower than Barlow
/// Bold's, so the engine's lines ran up to 9.8pt long and broke a word early.
/// `italic` is a plain bool with no "unset", so the chain reads as "the first
/// level that says yes".
pub fn level_italic(
    master: &[crate::ir::MasterStyleLevel],
    ph_levels: &[crate::ir::MasterStyleLevel],
    lvl: u32,
) -> bool {
    let lvl = lvl as usize;
    [ph_levels, master].into_iter().any(|levels| {
        levels
            .get(lvl.min(levels.len().saturating_sub(1)))
            .is_some_and(|l| l.italic)
    })
}

/// One line of a shape's text, placed.
///
/// `x` and `baseline` are relative to the shape's own box, in points, so the
/// caller adds the shape's position and draws.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct PlacedLine {
    pub text: String,
    /// The stretches this line is set in, one per run it crosses.
    pub segments: Vec<LineSegment>,
    /// Left edge of the line, from the shape's left edge.
    pub x: f32,
    /// Baseline, from the shape's top edge.
    pub baseline: f32,
    pub font_size: f32,
    pub family: String,
    pub bold: bool,
    pub italic: bool,
    /// Which paragraph of the shape this line came from.
    pub para_index: usize,
    /// How many characters of that paragraph precede this line, so an editor
    /// can map a click back to a run.
    pub char_start: usize,
}

/// A shape's text, laid out.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct ShapeLayout {
    pub lines: Vec<PlacedLine>,
    /// The block's height, before any vertical anchoring.
    pub height: f32,
    /// Whether EVERY paragraph was measured by the engine.
    ///
    /// False means at least one fell back to a browser wrap, and a caller that
    /// draws this without saying so is showing a layout PowerPoint would not
    /// produce. It is not a small distinction: the tables cover 17 families of
    /// the corpus's 142.
    pub complete: bool,
}

/// Lay out one text shape's paragraphs into placed lines.
///
/// This is the loop the renderer runs, with every rule it uses now living in
/// this module: the level a paragraph inherits, the size it is set at, the
/// spacing multiple, where its first baseline sits, how its lines break, how
/// wide each may run, and where each starts for its alignment.
///
/// What it does NOT do yet is the part that needs a device: which face GDI
/// hands back for a name, and the ink. Both are the caller's, through
/// `metrics`.
pub fn layout_text_shape(
    metrics: &dyn FaceMetrics,
    shape: &crate::ir::Shape,
    paragraphs: &[crate::ir::SlideParagraph],
    master: &[crate::ir::MasterStyleLevel],
    ph_levels: &[crate::ir::MasterStyleLevel],
    default_family: &str,
) -> ShapeLayout {
    let inner_w = (shape.width - shape.l_ins - shape.r_ins).max(0.0);
    let mut lines: Vec<PlacedLine> = Vec::new();
    let mut cursor = shape.t_ins;
    let mut prev_fs: Option<f32> = None;
    let mut complete = true;

    for (pi, para) in paragraphs.iter().enumerate() {
        let level = resolve_level(master, ph_levels, para.lvl);
        let fs = paragraph_font_size(para, level.font_size, prev_fs, true);
        prev_fs = Some(fs);
        let n = line_spacing_multiple(para, &level, fs, exact_line_pt(para, true, true));
        let adv = fs * 1.2 * n;

        // Space before: the paragraph's own, else the level's fraction of the
        // advance.
        //
        // The first paragraph normally gets none -- its top IS the inset --
        // unless the shape sets `a:bodyPr/@spcFirstLastPara`, which says to
        // honour it. Probe `spcfirst` (8 arms): with the flag off, 0 / 6 / 10 /
        // 18pt all leave the first baseline where 0pt does; with it on they
        // move it 0 / 6.000 / 9.960 / 18.000pt down, and the gap to the second
        // paragraph is unchanged in every arm. d06 and d35 declare 10pt and
        // PowerPoint draws 9.815 / 9.834pt lower; d16 declares 6pt and draws
        // 6.065 lower -- it tracks the declared amount rather than being a
        // constant, which is what made it findable.
        if pi > 0 || shape.spc_first_last_para {
            cursor += para
                .space_before
                .or_else(|| level.spc_bef_pct.map(|p| p * adv))
                .unwrap_or(0.0);
        }

        // A run's own face, else the LEVEL's, else the theme's.
        //
        // ★The level was being skipped, which is where a title's face lives:
        // d15 slide 1 carries no face on any run and `Barlow` on the level,
        // and the truth PDF sets that title in Barlow Bold -- the engine was
        // measuring it in Arial (the theme's minor face) and placing every
        // glyph of the line accordingly. `bold` and the size already consult
        // the level; the face did not.
        let family = effective_family(
            metrics,
            para.runs
                .iter()
                .find_map(|r| r.font_family.clone())
                .or_else(|| level_family(master, ph_levels, para.lvl))
                .unwrap_or_else(|| default_family.to_string())
                .as_str(),
            true,
        );
        let lvl_bold = level.bold.unwrap_or(false);
        // An empty paragraph has only the level to go on; one with runs is
        // bold if any of its runs RESOLVES bold, which a `b="0"` run does not
        // even inside a bold level.
        let bold = if para.runs.is_empty() {
            lvl_bold
        } else {
            para.runs.iter().any(|r| r.is_bold(lvl_bold))
        };
        let italic =
            para.runs.iter().any(|r| r.italic) || level_italic(master, ph_levels, para.lvl);
        let geom = indent_geometry(
            inner_w,
            para.mar_l.unwrap_or(level.mar_l),
            para.indent.unwrap_or(level.indent),
            true,
        );

        let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        let broken = break_paragraph(
            metrics, &text, fs, &family, bold, italic, geom.first_width, &para.runs,
        );
        let broken = match broken {
            Some(b) => b,
            None => {
                complete = false;
                vec![text.clone()]
            }
        };

        let first_off = first_baseline_off(metrics, &family, fs, n, true);
        let align = para.alignment.or(level.algn).unwrap_or_default();
        let mut char_at = 0usize;
        for (li, line) in broken.iter().enumerate() {
            let width = if li == 0 { geom.first_width } else { geom.rest_width };
            let line_w = master_units(metrics, line.trim_end(), fs, &family, bold, italic, 0.0)
                .map(|mu| master_units_pt(mu) as f32)
                .unwrap_or(0.0);
            let x = if li == 0 { geom.first_x } else { geom.rest_x };
            lines.push(PlacedLine {
                text: line.clone(),
                segments: line_segments(
                    &para.runs,
                    char_at,
                    line,
                    fs,
                    &family,
                    level.bold.unwrap_or(false),
                    // The LEVEL's slant, not the paragraph's resolved one:
                    // passing the OR would make one italic run turn the whole
                    // line italic, which is the collapse the segments exist to
                    // undo.
                    level_italic(master, ph_levels, para.lvl),
                ),
                x: shape.l_ins + x + align_offset(align, width, line_w),
                baseline: cursor + first_off + li as f32 * adv,
                font_size: fs,
                family: family.clone(),
                bold,
                italic,
                para_index: pi,
                char_start: char_at,
            });
            char_at += line.chars().count();
        }
        cursor += broken.len() as f32 * adv;
        cursor += para.space_after.unwrap_or(0.0);
    }

    let height = (cursor - shape.t_ins).max(0.0);
    // Vertical anchoring shifts the whole block inside the inner box.
    let inner_h = (shape.height - shape.t_ins - shape.b_ins).max(0.0);
    //
    // ★Neither anchor clamps at zero. A block TALLER than its box still
    // centres on it: d24 slide 1's 60pt title needs 178pt in a 91pt box and
    // PowerPoint puts the block's centre on the box's, overflowing equally
    // above and below (measured 2026-08-18 from PowerPoint's own render); the
    // bottom anchor likewise holds the LAST baseline and lets the block run off
    // the top (probe `anchorb`, 12 arms). Clamping pins an overflowing block to
    // the box top -- d15 slide 5's five 30pt lines in a 64.6pt box came out
    // 75.15pt low, every line by the same amount, which is what a clamped
    // anchor looks like from the outside.
    let shift = match shape.anchor.as_deref() {
        Some("ctr") => (inner_h - height) / 2.0,
        Some("b") => inner_h - height,
        _ => 0.0,
    };
    if shift != 0.0 {
        for l in &mut lines {
            l.baseline += shift;
        }
    }
    ShapeLayout { lines, height, complete }
}

#[cfg(test)]
mod shape_tests {
    use super::*;
    use crate::ir::{Shape, ShapeContent, SlideParagraph, SlideRun};

    fn run(text: &str, size: Option<f32>) -> SlideRun {
        SlideRun {
            text: text.to_string(),
            font_size: size,
            bold: None,
            italic: false,
            underline: false,
            color: None,
            color_alpha: None,
            highlight: None,
            font_family: Some("Arial".to_string()),
            spacing: None,
        }
    }

    fn para(text: &str, size: f32) -> SlideParagraph {
        SlideParagraph {
            runs: vec![run(text, Some(size))],
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

    fn shape(w: f32, h: f32, anchor: Option<&str>) -> Shape {
        let mut s = Shape::default();
        s.width = w;
        s.height = h;
        s.l_ins = 0.0;
        s.r_ins = 0.0;
        s.t_ins = 0.0;
        s.b_ins = 0.0;
        s.anchor = anchor.map(|a| a.to_string());
        s.content = ShapeContent::TextBox { paragraphs: vec![] };
        s
    }

    #[test]
    fn a_paragraph_that_fits_is_one_line_at_its_own_ascent() {
        let p = [para("Hello", 12.0)];
        let got = layout_text_shape(&TableMetrics, &shape(400.0, 100.0, None), &p, &[], &[], "Arial");
        assert_eq!(got.lines.len(), 1);
        assert!(got.complete);
        // The baseline is the first-baseline offset, not the top.
        assert!(got.lines[0].baseline > 0.0);
        assert!((got.height - 12.0 * 1.2).abs() < 1e-3);
    }

    #[test]
    fn a_long_paragraph_wraps_and_steps_by_the_advance() {
        let p = [para("The quick brown fox jumps over the lazy dog", 12.0)];
        let got = layout_text_shape(&TableMetrics, &shape(120.0, 200.0, None), &p, &[], &[], "Arial");
        assert!(got.lines.len() > 1, "{:?}", got.lines);
        let step = got.lines[1].baseline - got.lines[0].baseline;
        assert!((step - 12.0 * 1.2).abs() < 1e-3, "{step}");
    }

    #[test]
    fn a_family_nothing_measured_marks_the_shape_incomplete() {
        let mut p = para("Hello", 12.0);
        p.runs[0].font_family = Some("Zzyzx Nonexistent".to_string());
        let got = layout_text_shape(&TableMetrics, &shape(400.0, 100.0, None), &[p], &[], &[], "Arial");
        assert!(!got.complete, "an unmeasurable family must be flagged");
    }

    #[test]
    fn centring_pushes_the_block_down_by_half_the_slack() {
        let p = [para("Hello", 12.0)];
        let top = layout_text_shape(&TableMetrics, &shape(400.0, 100.0, None), &p, &[], &[], "Arial");
        let ctr = layout_text_shape(&TableMetrics, &shape(400.0, 100.0, Some("ctr")), &p, &[], &[], "Arial");
        let slack = (100.0 - top.height) / 2.0;
        assert!((ctr.lines[0].baseline - top.lines[0].baseline - slack).abs() < 1e-3);
    }

    #[test]
    fn the_bottom_anchor_puts_the_last_line_on_the_floor() {
        let p = [para("Hello", 12.0)];
        let got = layout_text_shape(&TableMetrics, &shape(400.0, 100.0, Some("b")), &p, &[], &[], "Arial");
        assert!(got.lines[0].baseline > 80.0, "{:?}", got.lines[0]);
    }

    #[test]
    fn a_level_supplies_the_face_when_no_run_names_one() {
        // d15 slide 1's title: every run inherits, and the face lives on the
        // placeholder's level. Measuring it in the theme's face instead puts
        // every glyph of the line in the wrong place.
        let mut p = para("Hello", 12.0);
        p.runs[0].font_family = None;
        let mut level = crate::ir::MasterStyleLevel::default();
        level.font_family = Some("Georgia".to_string());
        let got = layout_text_shape(
            &TableMetrics, &shape(400.0, 100.0, None), &[p], &[], &[level], "Arial");
        assert_eq!(got.lines[0].family, "Georgia");
    }

    #[test]
    fn a_run_that_names_a_face_still_beats_the_level() {
        let p = para("Hello", 12.0);   // its run names Arial
        let mut level = crate::ir::MasterStyleLevel::default();
        level.font_family = Some("Georgia".to_string());
        let got = layout_text_shape(
            &TableMetrics, &shape(400.0, 100.0, None), &[p], &[], &[level], "Times New Roman");
        assert_eq!(got.lines[0].family, "Arial");
    }

    #[test]
    fn a_second_paragraph_starts_below_the_first() {
        let p = [para("One", 12.0), para("Two", 12.0)];
        let got = layout_text_shape(&TableMetrics, &shape(400.0, 200.0, None), &p, &[], &[], "Arial");
        assert_eq!(got.lines.len(), 2);
        assert!(got.lines[1].baseline > got.lines[0].baseline);
        assert_eq!(got.lines[1].para_index, 1);
    }

    #[test]
    fn a_line_that_crosses_runs_is_split_where_they_do() {
        let runs = vec![
            crate::ir::SlideRun { bold: Some(true), ..run("Alternative: ", None) },
            run("click the button", None),
        ];
        let segs = line_segments(&runs, 0, "Alternative: click the button",
                                 18.0, "Nunito", false, false);
        assert_eq!(segs.len(), 2, "{segs:?}");
        assert_eq!(segs[0].text, "Alternative: ");
        assert!(segs[0].bold);
        assert_eq!(segs[1].text, "click the button");
        assert!(!segs[1].bold);
        assert_eq!(segs[1].start, 13);
    }

    #[test]
    fn a_run_that_says_b_zero_is_not_bold_inside_a_bold_level() {
        // PowerPoint's own answer (`bzero` probe, 4/4 arms; d15 s11, d11 s11
        // and corpus 04 s11 in one shape each): `b="0"` turns the level's bold
        // OFF, and a run that says nothing takes it.
        let runs = vec![
            crate::ir::SlideRun { bold: Some(false), ..run("ZERO ", None) },
            run("SILENT", None),
        ];
        let segs = line_segments(&runs, 0, "ZERO SILENT", 18.0, "Arial", true, false);
        assert_eq!(segs.len(), 2, "{segs:?}");
        assert!(!segs[0].bold, "b=0 must beat a bold level");
        assert!(segs[1].bold, "a run that says nothing takes the level's bold");
    }

    #[test]
    fn a_later_line_takes_the_styles_it_actually_covers() {
        let runs = vec![
            crate::ir::SlideRun { bold: Some(true), ..run("HEAD", None) },
            run("tail text", None),
        ];
        // The second line starts four characters in, past the bold run.
        let segs = line_segments(&runs, 4, "tail text", 18.0, "Nunito", false, false);
        assert_eq!(segs.len(), 1);
        assert!(!segs[0].bold);
    }

    #[test]
    fn a_run_that_names_its_own_face_or_size_starts_a_segment() {
        let runs = vec![
            run("name", None),
            crate::ir::SlideRun {
                font_family: Some("Arial".to_string()),
                font_size: Some(9.0),
                ..run("JOB", None)
            },
        ];
        let segs = line_segments(&runs, 0, "nameJOB", 12.0, "IBM Plex Sans", false, false);
        assert_eq!(segs.len(), 2, "{segs:?}");
        assert_eq!(segs[1].family, "Arial");
        assert_eq!(segs[1].font_size, 9.0);
    }

    #[test]
    fn a_level_that_says_bold_reaches_a_run_that_does_not_mention_it() {
        let segs = line_segments(&[run("Title", None)], 0, "Title", 44.0, "Barlow", true, false);
        assert_eq!(segs.len(), 1);
        assert!(segs[0].bold, "the level's weight must reach the run");
    }

    #[test]
    fn a_block_taller_than_its_box_still_centres_on_it() {
        // Five lines of 30pt type in a 64.6pt box: PowerPoint overflows it
        // equally above and below rather than pinning it to the top.
        let p: Vec<_> = (0..5).map(|_| para("Hello", 30.0)).collect();
        let got = layout_text_shape(
            &TableMetrics, &shape(400.0, 64.6, Some("ctr")), &p, &[], &[], "Arial");
        assert!(got.height > 64.6, "the block must overflow: {}", got.height);
        let shift = (64.6 - got.height) / 2.0;
        let flat = layout_text_shape(
            &TableMetrics, &shape(400.0, 64.6, None), &p, &[], &[], "Arial");
        assert!(
            (got.lines[0].baseline - flat.lines[0].baseline - shift).abs() < 1e-3,
            "{} vs {} + {shift}",
            got.lines[0].baseline,
            flat.lines[0].baseline
        );
        assert!(got.lines[0].baseline < 0.0, "it starts above the box");
    }

    #[test]
    fn the_bottom_anchor_does_not_clamp_either() {
        let p: Vec<_> = (0..5).map(|_| para("Hello", 30.0)).collect();
        let got = layout_text_shape(
            &TableMetrics, &shape(400.0, 64.6, Some("b")), &p, &[], &[], "Arial");
        let flat = layout_text_shape(
            &TableMetrics, &shape(400.0, 64.6, None), &p, &[], &[], "Arial");
        let shift = 64.6 - got.height;
        assert!((got.lines[0].baseline - flat.lines[0].baseline - shift).abs() < 1e-3);
    }

    #[test]
    fn a_level_that_says_italic_reaches_a_run_that_does_not_mention_it() {
        let segs = line_segments(&[run("Quote", None)], 0, "Quote", 30.0, "Barlow",
                                 true, true);
        assert_eq!(segs.len(), 1);
        assert!(segs[0].bold && segs[0].italic);
    }

    #[test]
    fn the_level_supplies_the_slant_for_the_whole_shape() {
        // d15 slide 5: `i="1"` on the layout's level, nothing on any run, and
        // PowerPoint sets it in Barlow Bold Italic.
        let p = para("Quotations are commonly printed", 30.0);
        let mut level = crate::ir::MasterStyleLevel::default();
        level.italic = true;
        let got = layout_text_shape(
            &TableMetrics, &shape(400.0, 200.0, None), &[p], &[], &[level], "Arial");
        assert!(got.lines[0].italic, "the level's slant must reach the line");
        assert!(got.lines[0].segments.iter().all(|s| s.italic));
    }

    #[test]
    fn one_italic_run_does_not_turn_the_whole_line_italic() {
        let runs = vec![
            crate::ir::SlideRun { italic: true, ..run("Slanted ", None) },
            run("upright", None),
        ];
        let segs = line_segments(&runs, 0, "Slanted upright", 18.0, "Arial", false, false);
        assert_eq!(segs.len(), 2, "{segs:?}");
        assert!(segs[0].italic);
        assert!(!segs[1].italic);
    }

    #[test]
    fn the_first_paragraph_keeps_its_space_when_the_shape_asks() {
        let mut p = para("Alpha", 24.0);
        p.space_before = Some(10.0);
        let mut sh = shape(400.0, 300.0, None);
        let plain = layout_text_shape(&TableMetrics, &sh, &[p.clone()], &[], &[], "Arial");
        sh.spc_first_last_para = true;
        let kept = layout_text_shape(&TableMetrics, &sh, &[p], &[], &[], "Arial");
        assert!(
            (kept.lines[0].baseline - plain.lines[0].baseline - 10.0).abs() < 1e-3,
            "{} vs {}",
            kept.lines[0].baseline,
            plain.lines[0].baseline
        );
    }

    #[test]
    fn the_flag_does_not_disturb_the_gap_between_paragraphs() {
        let mut a = para("Alpha", 24.0);
        let mut b = para("Beta", 24.0);
        a.space_before = Some(10.0);
        b.space_before = Some(10.0);
        let mut sh = shape(400.0, 300.0, None);
        let plain = layout_text_shape(&TableMetrics, &sh, &[a.clone(), b.clone()],
                                      &[], &[], "Arial");
        sh.spc_first_last_para = true;
        let kept = layout_text_shape(&TableMetrics, &sh, &[a, b], &[], &[], "Arial");
        let gap = |l: &ShapeLayout| l.lines[1].baseline - l.lines[0].baseline;
        assert!((gap(&plain) - gap(&kept)).abs() < 1e-3);
    }

    #[test]
    fn a_shape_that_does_not_ask_still_drops_it() {
        let mut p = para("Alpha", 24.0);
        p.space_before = Some(18.0);
        let bare = para("Alpha", 24.0);
        let sh = shape(400.0, 300.0, None);
        let with_space = layout_text_shape(&TableMetrics, &sh, &[p], &[], &[], "Arial");
        let without = layout_text_shape(&TableMetrics, &sh, &[bare], &[], &[], "Arial");
        assert!((with_space.lines[0].baseline - without.lines[0].baseline).abs() < 1e-3);
    }
}

/// A [`FaceMetrics`] whose answers were supplied from outside.
///
/// The compiled tables carry the faces this machine could measure; a browser
/// can measure any face it is able to DRAW, which is a far larger set and the
/// only one that helps somebody else's deck. So the page measures the
/// characters a shape needs and hands them over, and the engine's rules run on
/// those -- the break law, the per-run measure, the indent geometry, all of it
/// unchanged.
///
/// Advances are in EM units, keyed by `(family, bold, italic)` and character.
/// A character the supplier did not measure is refused rather than guessed, so
/// a shape it cannot cover is reported incomplete exactly as before.
pub struct SuppliedMetrics {
    faces: std::collections::HashMap<(String, bool, bool), std::collections::HashMap<char, f32>>,
}

impl SuppliedMetrics {
    pub fn new() -> Self {
        Self { faces: std::collections::HashMap::new() }
    }

    /// Record what one face advances for one character.
    pub fn insert(&mut self, family: &str, bold: bool, italic: bool, ch: char, em: f32) {
        self.faces
            .entry((family.to_ascii_lowercase(), bold, italic))
            .or_default()
            .insert(ch, em);
    }

    pub fn is_empty(&self) -> bool {
        self.faces.is_empty()
    }
}

impl Default for SuppliedMetrics {
    fn default() -> Self {
        Self::new()
    }
}

impl FaceMetrics for SuppliedMetrics {
    fn advance_em(&self, family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
        let key = (family.to_ascii_lowercase(), bold, italic);
        self.faces
            .get(&key)
            .and_then(|m| m.get(&ch).copied())
            // The compiled tables remain the last word for the faces they do
            // carry, so a supplier that misses one is no worse off.
            //
            // ★Nothing else is invented. A style the supplier did not measure
            // is NOT served from the upright face: a bullet or a marker can
            // ask for a combination no run named, and answering it with the
            // wrong face's advances would lay the line out confidently wrong.
            // Unanswered means the shape is declined, which is the honest end.
            .or_else(|| TableMetrics.advance_em(family, bold, italic, ch))
    }

    fn has_all_glyphs(&self, family: &str, bold: bool, italic: bool, text: &str) -> bool {
        text.chars().all(|c| self.advance_em(family, bold, italic, c).is_some())
    }

    fn baseline_offset_em(&self, family: &str) -> Option<f32> {
        // ★A supplier answers about ADVANCES; it says nothing about where the
        // face puts its baseline, and leaving this at the trait default sent
        // every shape back to the six-family fallback -- so a per-face table
        // added to `TableMetrics` changed nothing at all until this was here.
        TableMetrics.baseline_offset_em(family)
    }

    // `resolves` keeps its default of `true`, for the reason spelled out on
    // `TableMetrics`: a name this source cannot measure may still be one the
    // deck embeds, and renaming it would be worse than declining the shape.
}

#[cfg(test)]
mod supplied_tests {
    use super::*;

    #[test]
    fn a_supplied_face_is_measured_by_what_was_given() {
        let mut m = SuppliedMetrics::new();
        for c in "ab".chars() {
            m.insert("Zzyzx", false, false, c, 0.5);
        }
        assert_eq!(m.advance_em("Zzyzx", false, false, 'a'), Some(0.5));
        assert_eq!(master_units(&m, "ab", 12.0, "Zzyzx", false, false, 0.0), Some(96));
    }

    #[test]
    fn a_character_nobody_measured_is_still_refused() {
        let mut m = SuppliedMetrics::new();
        m.insert("Zzyzx", false, false, 'a', 0.5);
        assert_eq!(m.advance_em("Zzyzx", false, false, 'q'), None);
    }

    #[test]
    fn a_style_that_was_not_measured_is_refused_not_borrowed() {
        let mut m = SuppliedMetrics::new();
        m.insert("Zzyzx", false, false, 'a', 0.5);
        assert_eq!(m.advance_em("Zzyzx", true, false, 'a'), None);
    }

    #[test]
    fn the_compiled_tables_still_answer_for_what_they_carry() {
        let m = SuppliedMetrics::new();
        assert!(m.advance_em("Arial", false, false, 'a').is_some());
    }

    #[test]
    fn a_supplier_still_answers_with_the_face_s_own_baseline() {
        let m = SuppliedMetrics::new();
        assert!(m.baseline_offset_em("Nunito").is_some());
        assert_eq!(
            m.baseline_offset_em("Nunito"),
            TableMetrics.baseline_offset_em("Nunito")
        );
    }
}

/// How far inside its bounding box a preset shape holds its text, as
/// (left, right, top, bottom) in points.
///
/// A `prstGeom` names a shape, and the shape names its own text rectangle --
/// which for most presets is smaller than the bounding box. Laying text out in
/// the box instead puts every centred or right-aligned line in the wrong
/// place: d35 s17's `homePlate` centres 'first' 7.86pt left of where the box
/// would, and d15 s17's (adj 50000) 6.59pt left. Both are what PowerPoint drew,
/// to within 0.13pt.
///
/// Measured by the `textrect` probe -- three alignments per preset, so the
/// left and right edges are each read directly rather than inferred from a
/// centre. Box 300x200pt, `Wq` at 18pt, no insets:
///
/// ```text
/// preset                 left    right    of a 300pt box
/// rect                  0.040  300.090   the box itself
/// ellipse              43.980  256.170   0.1466 .. 0.8539
/// teardrop             43.980  256.170   same as ellipse
/// pie                  43.980  256.170   same as ellipse
/// homePlate             0.040  250.020   right only
/// homePlate adj=30129   0.040  269.820
/// homePlate adj=50000   0.040  250.020
/// roundRect             9.810  290.150   all four sides
/// chevron             100.050  200.030   left and right
/// wedgeRectCallout      0.040  300.090   the box itself
/// ```
///
/// `ss` is the shorter side, which is what the DrawingML preset formulas
/// measure their adjustments against; the probe's box is 300x200 so the two
/// are distinguishable.
///
/// ★`homePlate` with no `a:gd` behaves as **adj = 50000**, not the 16667 the
/// preset definition documents -- the empty-`avLst` arm measures identical to
/// the explicit 50000 arm and 19.8pt away from the 30129 one. Recorded as
/// measured; the disagreement with the published default is not resolved.
pub fn geom_text_insets(
    shape_type: Option<&str>,
    adjustments: &std::collections::HashMap<String, f32>,
    width: f32,
    height: f32,
) -> (f32, f32, f32, f32) {
    let ss = width.min(height);
    let adj = |name: &str, default: f32| -> f32 {
        adjustments.get(name).copied().unwrap_or(default)
    };
    match shape_type {
        // The inscribed rectangle at 45 degrees: (1 - cos45) / 2 of each side.
        Some("ellipse") | Some("teardrop") | Some("pie") => {
            let k = (1.0 - std::f32::consts::FRAC_1_SQRT_2) / 2.0;
            (width * k, width * k, height * k, height * k)
        }
        // Only the point is taken out, and only from the right: the rectangle
        // ends halfway between the notch and the box edge.
        Some("homePlate") => (0.0, ss * adj("adj", 50_000.0) / 200_000.0, 0.0, 0.0),
        // Both points, and the full depth of each.
        Some("chevron") => {
            let dx = ss * adj("adj", 50_000.0) / 100_000.0;
            (dx, dx, 0.0, 0.0)
        }
        // The corner radius, less what the arc gives back.
        Some("roundRect") => {
            let idx = ss * adj("adj", 16_667.0) / 100_000.0 * 0.29289;
            (idx, idx, idx, idx)
        }
        _ => (0.0, 0.0, 0.0, 0.0),
    }
}

#[cfg(test)]
mod geom_tests {
    use super::*;
    use std::collections::HashMap;

    /// The probe's own box, so the numbers below are the measured ones.
    fn probe(preset: &str, adj: Option<f32>) -> (f32, f32, f32, f32) {
        let mut a = HashMap::new();
        if let Some(v) = adj {
            a.insert("adj".to_string(), v);
        }
        geom_text_insets(Some(preset), &a, 300.0, 200.0)
    }

    #[test]
    fn a_plain_rectangle_holds_text_in_its_whole_box() {
        assert_eq!(probe("rect", None), (0.0, 0.0, 0.0, 0.0));
    }

    #[test]
    fn an_ellipse_holds_text_in_the_box_inscribed_at_45_degrees() {
        // Probe: left 43.980 of 300, right 256.170.
        let (l, r, t, b) = probe("ellipse", None);
        assert!((l - 43.934).abs() < 0.1, "{l}");
        assert!((300.0 - r - 256.066).abs() < 0.1, "{r}");
        assert!((t - 29.289).abs() < 0.1, "{t}");
        assert!((b - 29.289).abs() < 0.1, "{b}");
        // A teardrop and a pie are the same ellipse.
        assert_eq!(probe("teardrop", None), (l, r, t, b));
        assert_eq!(probe("pie", None), (l, r, t, b));
    }

    #[test]
    fn a_home_plate_gives_up_only_its_point() {
        // Probe: right edge 269.820 at adj 30129, 250.020 at adj 50000, and
        // the same 250.020 with no adj at all.
        let (l, r, t, b) = probe("homePlate", Some(30_129.0));
        assert_eq!((l, t, b), (0.0, 0.0, 0.0));
        assert!((300.0 - r - 269.871).abs() < 0.1, "{r}");
        assert!((300.0 - probe("homePlate", Some(50_000.0)).1 - 250.0).abs() < 0.1);
        assert_eq!(probe("homePlate", None), probe("homePlate", Some(50_000.0)));
    }

    #[test]
    fn a_chevron_gives_up_a_point_at_each_end() {
        // Probe: left 100.050, right 200.030 of a 300pt box.
        let (l, r, ..) = probe("chevron", None);
        assert!((l - 100.0).abs() < 0.1, "{l}");
        assert!((r - 100.0).abs() < 0.1, "{r}");
    }

    #[test]
    fn a_rounded_rectangle_gives_up_its_corner_radius() {
        // Probe: left 9.810 of a 300pt box.
        let (l, r, t, b) = probe("roundRect", None);
        for v in [l, r, t, b] {
            assert!((v - 9.763).abs() < 0.1, "{v}");
        }
    }
}
