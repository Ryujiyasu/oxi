// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Row height conversions between points, stored twips and display pixels.

fn quarter_pixels_from_twips(twips: u64) -> u64 {
    twips.saturating_mul(4).saturating_add(7) / 15
}

/// Normalize an OOXML row height to the stored whole-twip precision.
pub fn file_points(points: f64) -> f32 {
    if !points.is_finite() || points < 0.0 {
        return points as f32;
    }
    ((points * 20.0).floor() / 20.0) as f32
}

/// Quantize a nonnegative VBA height, validating its rounded twip value.
/// Zero denotes a hidden row whose preceding stored height is retained.
pub fn input_to_stored_points(points: f64) -> Option<f32> {
    if !points.is_finite() || points < 0.0 {
        return None;
    }
    let twips = (points * 20.0 + 0.5).floor();
    if twips > 8190.0 {
        return None;
    }
    let quarters = quarter_pixels_from_twips(twips as u64);
    let stored_twips = (quarters * 15).div_ceil(4);
    Some(stored_twips as f32 / 20.0)
}

fn stored_quarter_pixels(points: f32) -> u64 {
    if !points.is_finite() || points <= 0.0 {
        return 0;
    }
    let twips = (f64::from(points) * 20.0 + 0.5).floor() as u64;
    quarter_pixels_from_twips(twips)
}

/// The logical point height exposed by Range.RowHeight.
pub fn logical_points(stored: f32) -> f64 {
    let quarters = stored_quarter_pixels(stored);
    (quarters.saturating_mul(3).saturating_add(2) / 4) as f64 / 4.0
}

/// Whole display pixels occupied by an explicitly visible row.
pub fn pixels(stored: f32) -> f32 {
    (stored_quarter_pixels(stored) / 4) as f32
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn heights_match_native_assignments_and_saved_points() {
        let mut checked = 0;
        for line in include_str!("../../../tests/fixtures/vba/row_height_units.tsv").lines() {
            if line.is_empty() || line.starts_with('#') { continue; }
            let fields: Vec<_> = line.split('\t').collect();
            let requested: f64 = fields[0].parse().unwrap();
            let number: u32 = fields[1].parse().unwrap();
            let stored: f32 = fields[2].parse().unwrap();
            let expected_logical: f64 = fields[3].parse().unwrap();
            let expected_physical: f64 = fields[4].parse().unwrap();
            let hidden = fields[5] == "True";
            let explicit_hidden = fields[6] == "True";
            let assigned = input_to_stored_points(requested);
            if number != 0 {
                assert_eq!(assigned, None, "{line}");
            } else if explicit_hidden {
                assert_eq!(assigned, Some(0.0), "{line}");
            } else {
                assert_eq!(assigned, Some(stored), "{line}");
            }
            let logical = if explicit_hidden { 0.0 } else { logical_points(stored) };
            let physical = if explicit_hidden { 0.0 } else { f64::from(pixels(stored)) * 0.75 };
            assert_eq!(logical, expected_logical, "{line}");
            assert_eq!(physical, expected_physical, "{line}");
            assert_eq!(explicit_hidden || pixels(stored) == 0.0, hidden, "{line}");
            checked += 1;
        }
        assert_eq!(checked, 136);
    }

    #[test]
    fn file_precision_matches_native_readonly_getters() {
        let mut checked = 0;
        for line in include_str!("../../../tests/fixtures/vba/row_height_file_units.tsv").lines() {
            if line.is_empty() || line.starts_with('#') { continue; }
            let fields: Vec<_> = line.split('\t').collect();
            let raw: f64 = fields[0].parse().unwrap();
            let expected_logical: f64 = fields[1].parse().unwrap();
            let expected_physical: f64 = fields[2].parse().unwrap();
            let expected_hidden = fields[3] == "True";
            let stored = file_points(raw);
            assert_eq!(logical_points(stored), expected_logical, "{line}");
            assert_eq!(f64::from(pixels(stored)) * 0.75, expected_physical, "{line}");
            assert_eq!(pixels(stored) == 0.0, expected_hidden, "{line}");
            checked += 1;
        }
        assert_eq!(checked, 95);
    }
}
