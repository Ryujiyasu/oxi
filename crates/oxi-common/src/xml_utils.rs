/// Extract local name from a potentially namespaced XML tag (e.g., "w:body" -> "body")
pub fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

/// Get attribute value by local name
pub fn get_attr(e: &quick_xml::events::BytesStart, attr_name: &str) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        if key == attr_name {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

/// Get attribute value by raw key (including namespace prefix)
pub fn get_raw_attr(e: &quick_xml::events::BytesStart, key_name: &str) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
        if key == key_name || key.ends_with(&format!(":{}", key_name)) {
            return Some(String::from_utf8_lossy(&attr.value).to_string());
        }
    }
    None
}

/// Parse a twips value (1/20 pt) to points
pub fn twips_to_pt(twips: f32) -> f32 {
    twips / 20.0
}

/// Parse a half-point value to points
pub fn half_pt_to_pt(half_pt: f32) -> f32 {
    half_pt / 2.0
}

/// Parse EMU (English Metric Units) to points (1 pt = 12700 EMU)
pub fn emu_to_pt(emu: f32) -> f32 {
    emu / 12700.0
}
