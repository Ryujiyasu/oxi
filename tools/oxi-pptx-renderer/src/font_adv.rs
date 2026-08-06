/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/. */

//! hmtx (design advance) width tables in EM units, measured from the TTF
//! files via `tools/metrics/gen_pptx_font_adv.py`.
//!
//! PowerPoint's PDF export places glyphs at their DESIGN advance (the TrueType
//! `hmtx` value), NOT at GDI's hinted / integer-pixel snapped advance. A
//! line's logical width in PowerPoint/Word is the sum of the hmtx advances of
//! its visible characters (trailing spaces excluded). GDI's
//! `GetCharABCWidthsFloatW` / `GetTextExtentPoint32W` return hinted,
//! pixel-rounded values (multiples of 1px @96dpi = 0.75pt), so a line that
//! PowerPoint measures at 254.04pt comes out 255.75pt in GDI (+1.71).
//!
//! Each table holds ASCII 32..=126 (95 entries) in code-point order.

/// Arial (Regular) hmtx advance, in EM units.
static ARIAL: [f32; 95] = [
    0.27783, 0.27783, 0.35498, 0.55615, 0.55615, 0.88916, 0.66699, 0.19092, 0.33301, 0.33301,
    0.38916, 0.58398, 0.27783, 0.33301, 0.27783, 0.27783, 0.55615, 0.55615, 0.55615, 0.55615,
    0.55615, 0.55615, 0.55615, 0.55615, 0.55615, 0.55615, 0.27783, 0.27783, 0.58398, 0.58398,
    0.58398, 0.55615, 1.01514, 0.66699, 0.66699, 0.72217, 0.72217, 0.66699, 0.61084, 0.77783,
    0.72217, 0.27783, 0.50000, 0.66699, 0.55615, 0.83301, 0.72217, 0.77783, 0.66699, 0.77783,
    0.72217, 0.66699, 0.61084, 0.72217, 0.66699, 0.94385, 0.66699, 0.66699, 0.61084, 0.27783,
    0.27783, 0.27783, 0.46924, 0.55615, 0.33301, 0.55615, 0.55615, 0.50000, 0.55615, 0.55615,
    0.27783, 0.55615, 0.55615, 0.22217, 0.22217, 0.50000, 0.22217, 0.83301, 0.55615, 0.55615,
    0.55615, 0.55615, 0.33301, 0.50000, 0.27783, 0.55615, 0.50000, 0.72217, 0.50000, 0.50000,
    0.50000, 0.33398, 0.25977, 0.33398, 0.58398,
];

/// Calibri (Regular) hmtx advance, in EM units. Measured via
/// `tools/metrics/extract_calibri_adv.py` (calibri.ttf, upm 2048).
static CALIBRI: [f32; 95] = [
    0.22607, 0.32568, 0.40088, 0.49805, 0.50684, 0.71484, 0.68213, 0.22070, 0.30322, 0.30322,
    0.49805, 0.49805, 0.24951, 0.30615, 0.25244, 0.38623, 0.50684, 0.50684, 0.50684, 0.50684,
    0.50684, 0.50684, 0.50684, 0.50684, 0.50684, 0.50684, 0.26758, 0.26758, 0.49805, 0.49805,
    0.49805, 0.46338, 0.89404, 0.57861, 0.54395, 0.53320, 0.61523, 0.48828, 0.45947, 0.63086,
    0.62305, 0.25195, 0.31885, 0.51953, 0.42041, 0.85498, 0.64551, 0.66211, 0.51660, 0.67285,
    0.54297, 0.45947, 0.48730, 0.64160, 0.56738, 0.88965, 0.51904, 0.48730, 0.46826, 0.30664,
    0.38623, 0.30664, 0.49805, 0.49805, 0.29102, 0.47900, 0.52539, 0.42285, 0.52539, 0.49756,
    0.30518, 0.47070, 0.52539, 0.22949, 0.23926, 0.45459, 0.22949, 0.79883, 0.52539, 0.52734,
    0.52539, 0.52539, 0.34863, 0.39111, 0.33496, 0.52539, 0.45166, 0.71484, 0.43311, 0.45264,
    0.39502, 0.31445, 0.46045, 0.31445, 0.49805,
];

/// Arial Bold hmtx advance, in EM units.
static ARIALBD: [f32; 95] = [
    0.27783, 0.33301, 0.47412, 0.55615, 0.55615, 0.88916, 0.72217, 0.23779, 0.33301, 0.33301,
    0.38916, 0.58398, 0.27783, 0.33301, 0.27783, 0.27783, 0.55615, 0.55615, 0.55615, 0.55615,
    0.55615, 0.55615, 0.55615, 0.55615, 0.55615, 0.55615, 0.33301, 0.33301, 0.58398, 0.58398,
    0.58398, 0.61084, 0.97510, 0.72217, 0.72217, 0.72217, 0.72217, 0.66699, 0.61084, 0.77783,
    0.72217, 0.27783, 0.55615, 0.72217, 0.61084, 0.83301, 0.72217, 0.77783, 0.66699, 0.77783,
    0.72217, 0.66699, 0.61084, 0.72217, 0.66699, 0.94385, 0.66699, 0.66699, 0.61084, 0.33301,
    0.27783, 0.33301, 0.58398, 0.55615, 0.33301, 0.55615, 0.61084, 0.55615, 0.61084, 0.55615,
    0.33301, 0.61084, 0.61084, 0.27783, 0.27783, 0.55615, 0.27783, 0.88916, 0.61084, 0.61084,
    0.61084, 0.61084, 0.38916, 0.55615, 0.33301, 0.61084, 0.55615, 0.77783, 0.55615, 0.55615,
    0.50000, 0.38916, 0.27979, 0.38916, 0.58398,
];

/// Whether `family` has an hmtx width table (so line widths / glyph
/// positions can be computed from the design advance instead of GDI).
pub fn family_supported(family: &str) -> bool {
    family.eq_ignore_ascii_case("arial")
        || family.eq_ignore_ascii_case("arialbd")
        || family.eq_ignore_ascii_case("calibri")
}

/// The hmtx design advance of `ch` for `family`, in EM units (None if the
/// family is unsupported or the character is outside ASCII 32..=126).
pub fn hmtx_advance_em(family: &str, ch: char) -> Option<f32> {
    let table = if family.eq_ignore_ascii_case("arial") {
        &ARIAL
    } else if family.eq_ignore_ascii_case("arialbd") {
        &ARIALBD
    } else if family.eq_ignore_ascii_case("calibri") {
        &CALIBRI
    } else {
        return None;
    };
    let idx = ch as u32 as usize;
    if (32..127).contains(&idx) {
        Some(table[idx - 32])
    } else {
        None
    }
}

/// The hmtx design advance of a BULLET MARKER character, in EM units.
/// These live OUTSIDE ASCII 32..=126 (so the table above never sees them),
/// but PowerPoint still places the bullet glyph at its design advance.
/// Measured via `tools/metrics/gen_pptx_font_adv.py`; identical for Arial
/// and Arial Bold.
fn bullet_hmtx_em(ch: char) -> Option<f32> {
    match ch {
        '•' => Some(0.35010), // U+2022 BULLET
        '–' => Some(0.55615), // U+2013 EN DASH (level-2 marker)
        '»' => Some(0.55615), // U+00BB RIGHT-POINTING DOUBLE ANGLE QUOTATION MARK
        '●' => Some(0.60400), // U+25CF BLACK CIRCLE
        _ => None,
    }
}

/// The hmtx design advance of a bullet marker character for `family`, in EM
/// units (None if the family is unsupported or the character is not a known
/// bullet marker).
pub fn bullet_advance_em(family: &str, ch: char) -> Option<f32> {
    if family_supported(family) {
        bullet_hmtx_em(ch)
    } else {
        None
    }
}

/// Logical width of `line` in points: the hmtx advance sum of its VISIBLE
/// characters (trailing spaces are excluded — PowerPoint does not include
/// the trailing space of a wrapped line in its logical width). The final
/// visible character's advance IS included.
///
/// Returns None when the family is unsupported, in which case the caller
/// should fall back to the GDI measurement.
pub fn line_hmtx_width_pt(line: &str, fs: f32, family: &str) -> Option<f32> {
    if !family_supported(family) {
        return None;
    }
    let visible = line.trim_end_matches(' ');
    let mut sum = 0.0f32;
    for c in visible.chars() {
        match hmtx_advance_em(family, c) {
            Some(em) => sum += em * fs,
            None => return None, // non-ASCII char -> not measurable with hmtx
        }
    }
    Some(sum)
}

/// Per-character pixel widths for `text` (every char incl. trailing spaces),
/// used as the `Dx` array of `ExtTextOutW` so glyphs land at the design
/// advance like PowerPoint. Pixels are rounded from the em value at `scale`
/// (px per pt). Returns None if the family is unsupported or any character
/// is outside the table (caller then falls back to plain `TextOutW`).
pub fn line_hmtx_dx_px(text: &str, fs: f32, family: &str, scale: f64) -> Option<Vec<i32>> {
    if !family_supported(family) {
        return None;
    }
    let mut dx = Vec::with_capacity(text.len());
    for c in text.chars() {
        match hmtx_advance_em(family, c) {
            Some(em) => dx.push((em * fs * scale as f32).round() as i32),
            None => return None,
        }
    }
    Some(dx)
}

/// Pixel width of `text` as the sum of its rounded per-char design advances
/// (every char incl. trailing spaces). None if unsupported / non-ASCII.
pub fn text_hmtx_px(text: &str, fs: f32, family: &str, scale: f64) -> Option<i32> {
    let dx = line_hmtx_dx_px(text, fs, family, scale)?;
    Some(dx.iter().sum())
}

/// Pixel width of the space character at the design advance. None if the
/// family is unsupported.
pub fn space_hmtx_px(fs: f32, family: &str, scale: f64) -> Option<i32> {
    let em = hmtx_advance_em(family, ' ')?;
    Some((em * fs * scale as f32).round() as i32)
}
