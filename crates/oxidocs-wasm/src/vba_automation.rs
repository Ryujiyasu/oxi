// SPDX-License-Identifier: MIT OR Apache-2.0

//! Portable ANSI name conversion for a target Windows Automation host.

#[derive(Clone, Copy)]
pub(crate) enum AnsiCodePage {
    Japanese,
    Western,
}

impl AnsiCodePage {
    pub(crate) fn from_id(id: u16) -> Result<Self, String> {
        match id {
            932 => Ok(Self::Japanese),
            1252 => Ok(Self::Western),
            _ => Err(format!("unsupported VBA host ANSI code page: {id}")),
        }
    }

    pub(crate) fn names_equal(self, declared: &str, requested: &str) -> bool {
        if declared.eq_ignore_ascii_case(requested) {
            return true;
        }
        match (self.encode(declared), self.encode(requested)) {
            (Some(left), Some(right)) => sort_key(&left) == sort_key(&right),
            _ => false,
        }
    }

    fn encode(self, text: &str) -> Option<Vec<u8>> {
        let table: &[u8] = match self {
            Self::Japanese => include_bytes!("data/vba_automation_cp932.bin"),
            Self::Western => include_bytes!("data/vba_automation_cp1252.bin"),
        };
        let mut result = Vec::with_capacity(text.len());
        for unit in text.encode_utf16() {
            if unit < 128 {
                result.push(unit as u8);
                continue;
            }
            let mut low = 0;
            let mut high = table.len() / 4;
            while low < high {
                let middle = (low + high) / 2;
                let code = u16::from_le_bytes([table[middle * 4], table[middle * 4 + 1]]);
                if code < unit { low = middle + 1; } else { high = middle; }
            }
            let entry = table.get(low * 4..low * 4 + 4)?;
            if u16::from_le_bytes([entry[0], entry[1]]) != unit {
                // A replacement '?' cannot name a declared VBA member. Avoid
                // giving distinct unmappable source identifiers the same key.
                return None;
            }
            let encoded = u16::from_le_bytes([entry[2], entry[3]]);
            if encoded > 255 { result.push((encoded >> 8) as u8); }
            result.push(encoded as u8);
        }
        Some(result)
    }
}

fn sort_key(source: &[u8]) -> Vec<u8> {
    let weights = include_bytes!("data/vba_automation_weights.bin");
    let mut primary = Vec::new();
    let mut secondary = Vec::new();
    let mut tertiary = Vec::new();
    let mut special = Vec::new();
    let mut position = 0usize;
    for &byte in source {
        let row = &weights[byte as usize * 10..][..10];
        let count = row[0] as usize;
        let units = count / 2;
        primary.extend_from_slice(&row[1..1 + count]);
        for (level, value) in [(&mut secondary, row[5]), (&mut tertiary, row[6])] {
            if value != 0 { level.push(value); }
            level.extend(std::iter::repeat_n(2, units.saturating_sub(usize::from(value != 0))));
        }
        if row[7] != 0 {
            let offset = 0xffffusize.wrapping_sub(position);
            special.extend_from_slice(&[(offset >> 8) as u8, offset as u8, row[8], row[9]]);
        }
        position += units;
    }
    while secondary.last() == Some(&2) { secondary.pop(); }
    while tertiary.last() == Some(&2) { tertiary.pop(); }
    primary.push(1);
    primary.extend(secondary);
    primary.push(1);
    primary.extend(tertiary);
    primary.extend_from_slice(&[1, 1]);
    primary.extend(special);
    primary.push(0);
    primary
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn composed_automation_keys_match_native_windows_vectors() {
        let cases: Vec<serde_json::Value> = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/automation_sort_keys.json"
        )).unwrap();
        for case in cases {
            let bytes = |name: &str| -> Vec<u8> {
                case[name].as_array().unwrap().iter().map(|b| b.as_u64().unwrap() as u8).collect()
            };
            let input = bytes("input");
            assert_eq!(sort_key(&input), bytes("key"), "{input:?}");
        }
    }

    #[test]
    fn host_code_pages_keep_encoding_and_dispatch_comparison_separate() {
        let japanese = AnsiCodePage::Japanese;
        let western = AnsiCodePage::Western;
        assert!(japanese.names_equal("陰", "餌"));
        assert!(japanese.names_equal("佳", "霞"));
        assert!(japanese.names_equal("過", "鋭s"));
        assert!(japanese.names_equal("画", "餌e"));
        assert!(japanese.names_equal("億A", "陰"));
        assert!(japanese.names_equal("Cafe", "Café"));
        assert!(!western.names_equal("Cafe", "Café"));
        assert!(western.names_equal("Straße", "STRASSE"));
        assert!(!japanese.names_equal("Ａ", "ａ"));
        assert!(!japanese.names_equal("カ", "ｶ"));
        assert!(!japanese.names_equal("Ω", "ω"));
        assert!(!japanese.names_equal("陰", "売"));
        assert!(!western.names_equal("陰", "餌"));
        assert!(!japanese.names_equal("I", "ı"));
        assert!(!japanese.names_equal("K", "K"));
        assert!(!japanese.names_equal("ReadText", "ReadText😀"));
        assert!(AnsiCodePage::from_id(0).is_err());
    }
}
