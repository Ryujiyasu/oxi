// SPDX-License-Identifier: MIT OR Apache-2.0

//! Memory-only font programs and selected-glyph painting commands.
//! No system paths, font names, text substitutions or outline data tables
//! are part of the implementation. The caller supplies font bytes.

use serde::Serialize;
use std::collections::HashMap;
use ttf_parser::{Face, GlyphId, OutlineBuilder};

#[derive(Debug, PartialEq)]
pub enum FontError {
    InvalidFont,
    UnknownFamily,
    StyleMismatch,
    MissingProgram,
    InvalidGlyph,
    GeometryMismatch,
    MissingOutline,
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(tag = "op", content = "points")]
pub enum PathCommand {
    Move([f32; 2]),
    Line([f32; 2]),
    Quad([f32; 4]),
    Curve([f32; 6]),
    Close,
}

#[derive(Debug, Serialize)]
pub struct GlyphPainting {
    pub index: u16,
    pub units_per_em: u16,
    pub bounds_em: [f32; 4],
    /// Commands use normalized font coordinates, with positive y upwards.
    /// Canvas applies translate(origin, baseline) then scale(size, -size).
    pub commands: Vec<PathCommand>,
}

struct PathBuilder {
    scale: f32,
    commands: Vec<PathCommand>,
}

impl OutlineBuilder for PathBuilder {
    fn move_to(&mut self, x: f32, y: f32) {
        self.commands.push(PathCommand::Move([x * self.scale, y * self.scale]));
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.commands.push(PathCommand::Line([x * self.scale, y * self.scale]));
    }
    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.commands.push(PathCommand::Quad([
            x1 * self.scale, y1 * self.scale, x * self.scale, y * self.scale,
        ]));
    }
    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.commands.push(PathCommand::Curve([
            x1 * self.scale, y1 * self.scale, x2 * self.scale,
            y2 * self.scale, x * self.scale, y * self.scale,
        ]));
    }
    fn close(&mut self) {
        self.commands.push(PathCommand::Close);
    }
}

struct Program {
    bytes: Vec<u8>,
    face_index: u32,
}

#[derive(Default)]
pub struct FontPrograms {
    programs: Vec<Program>,
    faces: HashMap<(String, bool, bool), usize>,
}

fn family_key(name: &str) -> String {
    name.split_whitespace().collect::<Vec<_>>().join(" ").to_lowercase()
}

impl FontPrograms {
    pub(crate) fn program(&self, family: &str, bold: bool, italic: bool) -> Option<(&[u8], u32)> {
        let index = *self.faces.get(&(family_key(family), bold, italic))?;
        let program = &self.programs[index];
        Some((&program.bytes, program.face_index))
    }

    pub fn register(
        &mut self, family: &str, bold: bool, italic: bool,
        bytes: Vec<u8>, face_index: u32,
    ) -> Result<(), FontError> {
        let face = Face::parse(&bytes, face_index).map_err(|_| FontError::InvalidFont)?;
        if face.is_bold() != bold || face.is_italic() != italic {
            return Err(FontError::StyleMismatch);
        }
        let aliases: Vec<String> = face.names().into_iter()
            .filter(|n| n.name_id == 1 || n.name_id == 16)
            .filter_map(|n| n.to_string()).map(|n| family_key(&n)).collect();
        if !aliases.contains(&family_key(family)) {
            return Err(FontError::UnknownFamily);
        }
        let index = self.programs.len();
        for alias in aliases {
            self.faces.insert((alias, bold, italic), index);
        }
        self.programs.push(Program { bytes, face_index });
        Ok(())
    }

    pub fn selected_glyph(
        &self, family: &str, bold: bool, italic: bool,
        index: u16, expected_bounds_em: [f32; 4],
    ) -> Result<GlyphPainting, FontError> {
        let key = (family_key(family), bold, italic);
        let program = self.faces.get(&key).map(|&i| &self.programs[i])
            .ok_or(FontError::MissingProgram)?;
        let face = Face::parse(&program.bytes, program.face_index)
            .map_err(|_| FontError::InvalidFont)?;
        if index >= face.number_of_glyphs() {
            return Err(FontError::InvalidGlyph);
        }
        let upm = face.units_per_em();
        let scale = 1.0 / f32::from(upm);
        let mut builder = PathBuilder { scale, commands: Vec::new() };
        let bounds = face.outline_glyph(GlyphId(index), &mut builder);
        let declared = oxidocs_core::font::static_glyph_bounds(&program.bytes, program.face_index, index);
        let bounds_em = declared.or_else(|| bounds.map(|r| [
            i32::from(r.x_min), i32::from(r.y_min), i32::from(r.x_max), i32::from(r.y_max),
        ])).map(|bounds| bounds.map(|value| value as f32 * scale)).unwrap_or([0.0; 4]);
        if expected_bounds_em.iter().zip(bounds_em).any(|(&expected, actual)|
            !expected.is_finite() || (expected - actual).abs() > 1e-6)
        {
            return Err(FontError::GeometryMismatch);
        }
        if bounds.is_none() && (declared.is_some_and(|values| values != [0; 4])
            || face.glyph_bounding_box(GlyphId(index)).is_some()) {
            return Err(FontError::MissingOutline);
        }
        Ok(GlyphPainting { index, units_per_em: upm, bounds_em, commands: builder.commands })
    }
}
