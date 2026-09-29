// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use super::{normalize_family_name, FontMetricsRegistry};
use serde::Deserialize;
use std::{collections::HashMap, sync::OnceLock};

#[derive(Deserialize)]
struct Coverage {
    faces: HashMap<String, Vec<[u32; 2]>>,
    traditional_initial: Vec<[u32; 2]>,
    priority: Vec<String>,
}

fn data() -> &'static Coverage {
    static DATA: OnceLock<Coverage> = OnceLock::new();
    DATA.get_or_init(|| serde_json::from_str(include_str!("data/cjk_fallback_coverage.json"))
        .expect("valid CJK coverage metadata"))
}

fn contains(ranges: &[[u32; 2]], cp: u32) -> bool {
    let index = ranges.partition_point(|range| range[0] <= cp);
    index > 0 && cp <= ranges[index - 1][1]
}

/// Resolve a missing Han glyph, retaining a contiguous fallback where possible.
/// Missing coverage is unknown, so it must never manufacture a missing glyph.
pub(crate) fn select(
    registry: &FontMetricsRegistry, base: &str, ch: char, previous: Option<&str>,
) -> Option<&'static str> {
    let cp = ch as u32;
    if !matches!(cp, 0x3400..=0x9fff | 0xf900..=0xfaff | 0x20000..=0x323af) {
        return None;
    }
    let data = data();
    let normalized = normalize_family_name(base);
    let base_coverage = data.faces.get(super::render_family_name(base))
        .or_else(|| data.faces.get(&normalized))?;
    if contains(base_coverage, cp) { return None; }
    if let Some(previous) = previous {
        if let Some(name) = data.priority.iter().find(|name| name.as_str() == previous) {
            if registry.supports_family(name)
                && data.faces.get(name).is_some_and(|ranges| contains(ranges, cp)) {
                return Some(name.as_str());
            }
        }
    }
    data.priority.iter().find(|name| {
        registry.supports_family(name)
            && data.faces.get(*name).is_some_and(|ranges| contains(ranges, cp))
            && (name.as_str() != "PMingLiU" || contains(&data.traditional_initial, cp))
    }).map(String::as_str)
}
