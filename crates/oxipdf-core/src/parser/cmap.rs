//! CMap (Character Map) parser for PDF fonts.
//!
//! CMaps map character codes from a PDF content stream to Unicode code points.
//! This is essential for extracting text from PDFs that use CIDFont-based
//! fonts (common for CJK text like Japanese).

use std::collections::HashMap;

/// A parsed CMap that maps character codes to Unicode strings.
#[derive(Debug, Clone, Default)]
pub struct CMap {
    /// Maps a single character code to a Unicode string.
    pub char_map: HashMap<u32, String>,
    /// Ranges that map a contiguous block of codes to Unicode.
    pub ranges: Vec<CMapRange>,
}

/// A range entry from `beginbfrange`/`endbfrange`.
#[derive(Debug, Clone)]
pub struct CMapRange {
    pub start: u32,
    pub end: u32,
    pub mapping: RangeMapping,
}

#[derive(Debug, Clone)]
pub enum RangeMapping {
    /// Sequential offset from a base Unicode code point.
    Offset(u32),
    /// Explicit array of Unicode strings for each code in the range.
    Array(Vec<String>),
}

impl CMap {
    /// Look up a character code and return the Unicode string.
    pub fn decode(&self, code: u32) -> Option<String> {
        // Check direct mappings first.
        if let Some(s) = self.char_map.get(&code) {
            return Some(s.clone());
        }

        // Check ranges.
        for range in &self.ranges {
            if code >= range.start && code <= range.end {
                let offset = code - range.start;
                return match &range.mapping {
                    RangeMapping::Offset(base) => {
                        char::from_u32(base + offset).map(|c| c.to_string())
                    }
                    RangeMapping::Array(arr) => arr.get(offset as usize).cloned(),
                };
            }
        }

        None
    }

    /// Decode a sequence of bytes using this CMap.
    /// Tries 2-byte codes first (common for CJK), falls back to 1-byte.
    pub fn decode_bytes(&self, bytes: &[u8]) -> String {
        let mut result = String::new();
        let mut i = 0;

        while i < bytes.len() {
            // Try 2-byte code first (CJK fonts typically use 2-byte encoding).
            if i + 1 < bytes.len() {
                let code = ((bytes[i] as u32) << 8) | (bytes[i + 1] as u32);
                if let Some(s) = self.decode(code) {
                    result.push_str(&s);
                    i += 2;
                    continue;
                }
            }

            // Fall back to 1-byte code.
            let code = bytes[i] as u32;
            if let Some(s) = self.decode(code) {
                result.push_str(&s);
            } else {
                result.push(bytes[i] as char);
            }
            i += 1;
        }

        result
    }
}

/// Parse a CMap from its text representation.
pub fn parse_cmap(data: &[u8]) -> CMap {
    let text = String::from_utf8_lossy(data);
    let mut cmap = CMap::default();

    let mut lines = text.lines().peekable();
    while let Some(line) = lines.next() {
        let trimmed = line.trim();

        if trimmed.ends_with("beginbfchar") {
            // Parse individual character mappings.
            let count = parse_leading_int(trimmed);
            for _ in 0..count {
                if let Some(mapping_line) = lines.next() {
                    if let Some((code, unicode)) = parse_bfchar_line(mapping_line.trim()) {
                        cmap.char_map.insert(code, unicode);
                    }
                }
            }
        } else if trimmed.ends_with("beginbfrange") {
            // Parse range mappings.
            let count = parse_leading_int(trimmed);
            for _ in 0..count {
                if let Some(range_line) = lines.next() {
                    if let Some(range) = parse_bfrange_line(range_line.trim()) {
                        cmap.ranges.push(range);
                    }
                }
            }
        }
    }

    cmap
}

fn parse_leading_int(s: &str) -> usize {
    s.split_whitespace()
        .next()
        .and_then(|n| n.parse().ok())
        .unwrap_or(0)
}

/// Parse a bfchar line: `<XXXX> <YYYY>` → (code, unicode_string)
fn parse_bfchar_line(line: &str) -> Option<(u32, String)> {
    let parts: Vec<&str> = line.split('>').collect();
    if parts.len() < 2 {
        return None;
    }
    let code = parse_hex_token(parts[0])?;
    let unicode_hex = parts[1].trim().trim_start_matches('<');
    let unicode_str = hex_to_unicode_string(unicode_hex)?;
    Some((code, unicode_str))
}

/// Parse a bfrange line: `<XXXX> <YYYY> <ZZZZ>` or `<XXXX> <YYYY> [<A> <B> ...]`
fn parse_bfrange_line(line: &str) -> Option<CMapRange> {
    let parts: Vec<&str> = line.split('>').collect();
    if parts.len() < 3 {
        return None;
    }

    let start = parse_hex_token(parts[0])?;
    let end = parse_hex_token(parts[1])?;

    let rest = parts[2..].join(">");
    let rest = rest.trim();

    if rest.starts_with('[') {
        // Array mapping.
        let inner = rest.trim_start_matches('[').trim_end_matches(']');
        let array: Vec<String> = inner
            .split('>')
            .filter_map(|part| {
                let hex = part.trim().trim_start_matches('<');
                if hex.is_empty() {
                    None
                } else {
                    hex_to_unicode_string(hex)
                }
            })
            .collect();
        Some(CMapRange {
            start,
            end,
            mapping: RangeMapping::Array(array),
        })
    } else {
        // Offset mapping.
        let base_hex = rest.trim_start_matches('<').trim_end_matches('>');
        let base = u32::from_str_radix(base_hex.trim(), 16).ok()?;
        Some(CMapRange {
            start,
            end,
            mapping: RangeMapping::Offset(base),
        })
    }
}

fn parse_hex_token(s: &str) -> Option<u32> {
    let hex = s.trim().trim_start_matches('<');
    if hex.is_empty() {
        return None;
    }
    u32::from_str_radix(hex.trim(), 16).ok()
}

fn hex_to_unicode_string(hex: &str) -> Option<String> {
    let hex = hex.trim();
    if hex.is_empty() {
        return None;
    }
    // Each 4 hex digits = one UTF-16 code unit.
    let mut result = String::new();
    let mut i = 0;
    while i + 3 < hex.len() {
        if let Ok(code) = u16::from_str_radix(&hex[i..i + 4], 16) {
            if let Some(c) = char::from_u32(code as u32) {
                result.push(c);
            }
        }
        i += 4;
    }
    // Handle remaining 2 hex digits (single byte).
    if i + 1 < hex.len() && i + 2 <= hex.len() {
        if let Ok(code) = u8::from_str_radix(&hex[i..i + 2], 16) {
            result.push(code as char);
        }
    }
    if result.is_empty() {
        None
    } else {
        Some(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_bfchar() {
        let cmap_data = b"\
/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
2 beginbfchar
<0048> <0048>
<3042> <3042>
endbfchar
endcmap
";
        let cmap = parse_cmap(cmap_data);
        assert_eq!(cmap.decode(0x0048), Some("H".to_string()));
        assert_eq!(cmap.decode(0x3042), Some("\u{3042}".to_string())); // あ
    }

    #[test]
    fn test_parse_bfrange() {
        let cmap_data = b"\
1 beginbfrange
<0041> <005A> <0041>
endbfrange
";
        let cmap = parse_cmap(cmap_data);
        assert_eq!(cmap.decode(0x0041), Some("A".to_string()));
        assert_eq!(cmap.decode(0x0042), Some("B".to_string()));
        assert_eq!(cmap.decode(0x005A), Some("Z".to_string()));
        assert_eq!(cmap.decode(0x005B), None); // out of range
    }

    #[test]
    fn test_decode_bytes_cjk() {
        let cmap_data = b"\
3 beginbfchar
<3042> <3042>
<3044> <3044>
<3046> <3046>
endbfchar
";
        let cmap = parse_cmap(cmap_data);
        // "あいう" as 2-byte codes
        let bytes = [0x30, 0x42, 0x30, 0x44, 0x30, 0x46];
        let text = cmap.decode_bytes(&bytes);
        assert_eq!(text, "あいう");
    }

    #[test]
    fn test_bfrange_array() {
        let cmap_data = b"\
1 beginbfrange
<0001> <0003> [<3042> <3044> <3046>]
endbfrange
";
        let cmap = parse_cmap(cmap_data);
        assert_eq!(cmap.decode(0x0001), Some("あ".to_string()));
        assert_eq!(cmap.decode(0x0002), Some("い".to_string()));
        assert_eq!(cmap.decode(0x0003), Some("う".to_string()));
    }

    #[test]
    fn test_empty_cmap() {
        let cmap = parse_cmap(b"some random data");
        assert_eq!(cmap.decode(0x0041), None);
    }
}
