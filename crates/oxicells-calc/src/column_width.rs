// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Column width conversions, using the Normal font's digit width in pixels.

/// Turn character units into whole pixels, including the cell gutter.
pub fn characters_to_pixels(characters: f64, digit: f64) -> f64 {
    if characters < 1.0 {
        (characters * (digit + 5.0)).round()
    } else {
        (characters * digit + 5.0).round()
    }
}

/// Decode the gutter-inclusive width stored in a worksheet's column record.
pub fn stored_width_to_pixels(stored: f64, digit: f64) -> f64 {
    (((256.0 * stored + (128.0 / digit).trunc()) / 256.0) * digit).trunc()
}

/// Below one character the gutter shrinks together with the text area.
pub fn pixels_to_characters(pixels: f64, digit: f64) -> f64 {
    if pixels < digit + 5.0 {
        pixels / (digit + 5.0)
    } else {
        (pixels - 5.0) / digit
    }
}
