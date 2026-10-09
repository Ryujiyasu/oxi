// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Numeric OpenType MATH corner kerning, including script-style glyphs.
use serde::Deserialize;
use std::collections::HashMap;
use std::sync::OnceLock;

#[derive(Clone, Copy)]
pub enum Corner { TopRight, TopLeft, BottomRight, BottomLeft }

#[derive(Deserialize)]
struct Kern { heights: Vec<i16>, values: Vec<i16> }

#[derive(Deserialize)]
struct Corners {
    top_right: Option<Kern>, top_left: Option<Kern>,
    bottom_right: Option<Kern>, bottom_left: Option<Kern>,
}

#[derive(Deserialize)]
pub struct MathKerns {
    upm: u16,
    cmap: HashMap<u32, u16>,
    glyphs: HashMap<u16, Corners>,
}

impl MathKerns {
    pub fn cambria_math() -> &'static Self {
        static TABLE: OnceLock<MathKerns> = OnceLock::new();
        TABLE.get_or_init(|| serde_json::from_str(
            include_str!("data/cambria_math_kerns.json"))
            .expect("embedded math kerning must be valid"))
    }

    pub fn ordinary_gid(&self, c: char) -> Option<u16> {
        if super::runtime::resolve_registered("Cambria Math", false, false).is_some() {
            return super::runtime::registered_glyph("Cambria Math", false, false, c, 0).map(|g| g.index);
        }
        self.cmap.get(&(c as u32)).copied()
    }

    /// Correction heights are relative to this glyph's own baseline and size.
    /// At a boundary, select the next interval, as specified by OpenType MATH.
    pub fn value(&self, gid: Option<u16>, corner: Corner, height: f32, size: f32) -> f32 {
        if let Some(value) = gid.and_then(|gid| super::runtime::registered_corner_kern(
            "Cambria Math", false, false, gid, corner, height, size)) { return value; }
        if size <= 0.0 || self.upm == 0 { return 0.0; }
        let Some(corners) = gid.and_then(|g|self.glyphs.get(&g)) else { return 0.0; };
        let table = match corner {
            Corner::TopRight => &corners.top_right, Corner::TopLeft => &corners.top_left,
            Corner::BottomRight => &corners.bottom_right, Corner::BottomLeft => &corners.bottom_left,
        };
        let Some(table) = table else { return 0.0; };
        let du = height * self.upm as f32 / size;
        let index = table.heights.partition_point(|h|*h as f32 <= du);
        table.values.get(index).map_or(0.0,|v|*v as f32 * size / self.upm as f32)
    }
}
