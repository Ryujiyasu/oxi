// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! OpenType MATH variant selection and connected glyph assembly.
//! Coordinates remain in font design units. A caller positions the resulting
//! glyphs relative to the mathematical baseline and paints the glyph indices
//! with the SAME font face; scaling a Unicode character is not equivalent.

use serde::Deserialize;

#[derive(Debug, Clone, Copy, Deserialize, PartialEq, Eq)]
pub enum Direction { Vert, Horiz }

#[derive(Debug, Clone, Deserialize)]
pub struct Glyph {
    pub gid: u16,
    pub advance_width: u32,
    pub bounds: [i32; 4],
    /// A ready-made variant's correction is separate from an assembly's.
    #[serde(default)]
    pub italic_correction: Option<i32>,
    /// OpenType ExtendedShapeCoverage for this actual selected glyph.
    #[serde(default)]
    pub extended_shape: bool,
}

#[derive(Debug, Clone, Deserialize)]
pub struct Variant {
    #[serde(flatten)]
    pub glyph: Glyph,
    pub advance_measurement: u32,
}

#[derive(Debug, Clone, Deserialize)]
pub struct Part {
    #[serde(flatten)]
    pub glyph: Glyph,
    pub start: u32,
    pub end: u32,
    pub full_advance: u32,
    pub flags: u16,
}

#[derive(Debug, Deserialize)]
pub struct Construction {
    pub direction: Direction,
    pub base: Glyph,
    pub codepoint: Option<u32>,
    pub variants: Vec<Variant>,
    pub assembly: Vec<Part>,
    #[serde(default)]
    pub italic_correction: i32,
}

#[derive(Debug, Deserialize)]
pub struct StretchTable {
    pub upm: u32,
    pub min_connector_overlap: u32,
    pub constructions: Vec<Construction>,
}

#[derive(Debug, Clone)]
pub struct Placement {
    pub glyph: Glyph,
    /// Offset along growth axis, from the left or bottom of the assembly.
    pub advance_offset: f64,
}

#[derive(Debug, Clone)]
pub struct StretchPlan {
    pub direction: Direction,
    pub advance_measurement: f64,
    pub italic_correction: i32,
    pub placements: Vec<Placement>,
    pub assembled: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StretchError {
    InvalidTarget,
    InvalidConstruction,
    MissingAssembly,
    PartLimitExceeded,
}

impl StretchTable {
    pub fn cambria_math() -> std::sync::Arc<Self> {
        if let Some(program) = super::runtime::registered_math("Cambria Math") {
            return program.stretch.clone();
        }
        if let Some(metrics) = super::runtime::resolve_registered("Cambria Math", false, false) {
            return std::sync::Arc::new(Self { upm: u32::from(metrics.units_per_em),
                min_connector_overlap: 0, constructions: Vec::new() });
        }
        static TABLE: std::sync::OnceLock<std::sync::Arc<StretchTable>> = std::sync::OnceLock::new();
        TABLE.get_or_init(|| std::sync::Arc::new(serde_json::from_str(include_str!("data/cambria_math_stretch_metrics.json"))
            .expect("embedded MATH stretch metrics"))).clone()
    }

    pub fn construction(&self, direction: Direction, codepoint: char) -> Option<&Construction> {
        self.constructions.iter().find(|c| c.direction == direction
            && c.codepoint == Some(codepoint as u32))
    }

    pub fn plan(&self, c: &Construction, target: f64) -> Result<StretchPlan, StretchError> {
        if !target.is_finite() || target < 0.0 { return Err(StretchError::InvalidTarget); }
        if self.upm == 0 { return Err(StretchError::InvalidConstruction); }
        if let Some(v) = c.variants.iter().filter(|v| v.advance_measurement as f64 >= target)
            .min_by_key(|v| v.advance_measurement) {
            return Ok(StretchPlan { direction: c.direction,
                advance_measurement: v.advance_measurement as f64,
                italic_correction: c.italic_correction,
                placements: vec![Placement { glyph: v.glyph.clone(), advance_offset: 0.0 }],
                assembled: false });
        }
        if c.assembly.is_empty() { return Err(StretchError::MissingAssembly); }
        if c.assembly.iter().any(|p| p.full_advance == 0 || p.start > p.full_advance
            || p.end > p.full_advance) { return Err(StretchError::InvalidConstruction); }
        let extenders = c.assembly.iter().filter(|p| p.flags & 1 != 0).count();
        // A resource bound reports failure rather than silently clipping a
        // requested assembly to the largest ready-made variant.
        const MAX_PARTS: usize = 4096;
        for copies in 0..=MAX_PARTS {
            let count = c.assembly.len() - extenders + extenders * copies;
            if count > MAX_PARTS { return Err(StretchError::PartLimitExceeded); }
            let parts: Vec<&Part> = c.assembly.iter().flat_map(|p|
                std::iter::repeat_n(p, if p.flags & 1 != 0 { copies } else { 1 })).collect();
            if parts.is_empty() { continue; }
            let max_overlap: Vec<f64> = parts.windows(2)
                .map(|p| p[0].end.min(p[1].start) as f64).collect();
            let minimum = self.min_connector_overlap as f64;
            if max_overlap.iter().any(|o| *o < minimum) {
                if extenders == 0 || copies >= 1 { return Err(StretchError::InvalidConstruction); }
                continue;
            }
            let total: f64 = parts.iter().map(|p| p.full_advance as f64).sum();
            let shortest = total - max_overlap.iter().sum::<f64>();
            let longest = total - minimum * max_overlap.len() as f64;
            if shortest <= 0.0 { return Err(StretchError::InvalidConstruction); }
            if target > longest {
                if extenders == 0 { return Err(StretchError::MissingAssembly); }
                continue;
            }
            let length = target.max(shortest);
            let mut overlap = max_overlap.clone();
            let mut extra = length - shortest;
            // Equal extension, with redistribution after a short connector
            // reaches its limit. Every join stays within both font bounds.
            while extra > 1e-7 {
                let active: Vec<usize> = overlap.iter().enumerate()
                    .filter_map(|(i,o)| (*o-minimum > 1e-7).then_some(i)).collect();
                if active.is_empty() { return Err(StretchError::InvalidConstruction); }
                let share = extra / active.len() as f64;
                let mut used = 0.0;
                for i in active {
                    let take = share.min(overlap[i]-minimum);
                    overlap[i] -= take;
                    used += take;
                }
                if used <= 0.0 { return Err(StretchError::InvalidConstruction); }
                extra -= used;
            }
            let mut offset = 0.0;
            let placements = parts.iter().enumerate().map(|(i,p)| {
                let result = Placement { glyph: p.glyph.clone(), advance_offset: offset };
                if let Some(o) = overlap.get(i) { offset += p.full_advance as f64 - o; }
                result
            }).collect();
            return Ok(StretchPlan { direction: c.direction, advance_measurement: length,
                italic_correction: c.italic_correction, placements, assembled: true });
        }
        Err(StretchError::PartLimitExceeded)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table() -> StretchTable {
        serde_json::from_str(include_str!("data/cambria_math_stretch_metrics.json")).unwrap()
    }

    #[test]
    fn every_ready_made_variant_boundary_is_preserved() {
        let t=table();
        assert_eq!(t.constructions.len(),93);
        for c in &t.constructions {
            for (i,v) in c.variants.iter().enumerate() {
                let p=t.plan(c,v.advance_measurement as f64).unwrap();
                assert!(!p.assembled);
                assert_eq!(p.placements[0].glyph.gid,v.glyph.gid);
                if i+1<c.variants.len() {
                    let p=t.plan(c,v.advance_measurement as f64+0.01).unwrap();
                    assert_eq!(p.placements[0].glyph.gid,c.variants[i+1].glyph.gid);
                }
            }
        }
    }

    #[test]
    fn word_radical_variants_retain_nominal_size_and_real_width() {
        // Saved Word PDF glyph identity: these are two DIFFERENT 12pt glyphs.
        // This checks availability and scale; deriving the Word radicand target
        // and baseline remains the layout caller's separate responsibility.
        let t=table();let c=t.construction(Direction::Vert,'\u{221a}').unwrap();
        for (target,gid,width,ink_height) in [(4569.0,3495,9.005859375,26.765625),
                                             (6829.0,3496,9.17578125,40.0078125)] {
            let p=t.plan(c,target).unwrap();let g=&p.placements[0].glyph;
            assert_eq!(g.gid,gid);
            assert!((g.advance_width as f64*12.0/t.upm as f64-width).abs()<1e-8);
            assert!(((g.bounds[3]-g.bounds[1]) as f64*12.0/t.upm as f64-ink_height).abs()<1e-8);
        }
    }

    #[test]
    fn large_shapes_use_connected_parts_without_truncation() {
        let t=table();
        let mut tested=0;
        let mut invalid=0;
        for c in &t.constructions {
            if c.assembly.is_empty() { continue; }
            let longest=c.variants.iter().map(|v|v.advance_measurement).max().unwrap_or(0);
            for multiplier in [2.0,4.0] {
                let target=(longest as f64+1.0)*multiplier;
                if c.assembly.windows(2).any(|p| p[0].end.min(p[1].start) < t.min_connector_overlap) {
                    assert_eq!(t.plan(c,target).unwrap_err(),StretchError::InvalidConstruction);
                    invalid+=1;
                    continue;
                }
                let p=t.plan(c,target).unwrap();
                assert!(p.assembled);
                assert!(p.advance_measurement+1e-7>=target);
                let last=p.placements.last().unwrap();
                let last_part=c.assembly.iter().find(|part|part.glyph.gid==last.glyph.gid).unwrap();
                assert!((last.advance_offset+last_part.full_advance as f64-p.advance_measurement).abs()<1e-6);
                for pair in p.placements.windows(2) {
                    let left=c.assembly.iter().find(|part|part.glyph.gid==pair[0].glyph.gid).unwrap();
                    let right=c.assembly.iter().find(|part|part.glyph.gid==pair[1].glyph.gid).unwrap();
                    let overlap=pair[0].advance_offset+left.full_advance as f64-pair[1].advance_offset;
                    assert!(overlap+1e-6>=t.min_connector_overlap as f64);
                    assert!(overlap<=left.end.min(right.start) as f64+1e-6);
                }
                tested+=1;
            }
        }
        assert!(tested>50);
        assert_eq!(invalid,4,"Two real-font constructors have 133du connectors below the required 200du overlap; both requests must reject them");
    }

    #[test]
    fn unequal_connectors_redistribute_extra_length() {
        let t:StretchTable=serde_json::from_str(r#"{"upm":1000,"min_connector_overlap":10,
            "constructions":[{"direction":"Horiz","base":{"gid":1,"advance_width":100,"bounds":[0,0,100,100]},
            "codepoint":null,"variants":[],"italic_correction":0,"assembly":[
              {"gid":1,"advance_width":100,"bounds":[0,0,100,100],"start":0,"end":20,"full_advance":100,"flags":0},
              {"gid":2,"advance_width":100,"bounds":[0,0,100,100],"start":20,"end":80,"full_advance":100,"flags":0},
              {"gid":3,"advance_width":100,"bounds":[0,0,100,100],"start":80,"end":0,"full_advance":100,"flags":0}]}]}"#).unwrap();
        let p=t.plan(&t.constructions[0],270.0).unwrap();
        assert_eq!(p.advance_measurement,270.0);
        assert_eq!(p.placements[1].advance_offset,90.0);
        assert_eq!(p.placements[2].advance_offset,170.0);
    }

    #[test]
    fn absent_assembly_and_nonfinite_requests_report_failure() {
        let t=table();let c=t.constructions.iter().find(|c|c.assembly.is_empty()).unwrap();
        assert_eq!(t.plan(c,f64::NAN).unwrap_err(),StretchError::InvalidTarget);
        assert_eq!(t.plan(c,-1.0).unwrap_err(),StretchError::InvalidTarget);
        assert_eq!(t.plan(c,1e6).unwrap_err(),StretchError::MissingAssembly);
    }
}
