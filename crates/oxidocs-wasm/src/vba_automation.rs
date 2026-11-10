// SPDX-License-Identifier: MIT OR Apache-2.0

//! Portable Windows host name conversion and comparison.

#[derive(Clone, Copy)]
pub(crate) enum AnsiCodePage {
    Japanese,
    Western,
}

impl AnsiCodePage {
    pub(crate) fn character_codec(self, locale: Option<i64>) -> Option<oxivba_core::AnsiCodec> {
        let system_page = match self { Self::Japanese => 932, Self::Western => 1252 };
        oxivba_core::AnsiCodec::for_locale(locale, system_page, 1252)
    }

    fn encoded_unit(self, unit: u16) -> Option<u16> {
        if unit<128 {return Some(unit);}
        let table: &[u8]=match self {
            Self::Japanese=>include_bytes!("data/vba_automation_cp932.bin"),
            Self::Western=>include_bytes!("data/vba_automation_cp1252.bin"),
        };
        let mut low=0;let mut high=table.len()/4;
        while low<high {
            let middle=(low+high)/2;
            let code=u16::from_le_bytes([table[middle*4],table[middle*4+1]]);
            if code<unit {low=middle+1;} else {high=middle;}
        }
        let entry=table.get(low*4..low*4+4)?;
        (u16::from_le_bytes([entry[0],entry[1]])==unit).then(||u16::from_le_bytes([entry[2],entry[3]]))
    }

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
        let mut result = Vec::with_capacity(text.len());
        for unit in text.encode_utf16() {
            // Names cannot use the replacement fallback: distinct unmappable
            // source identifiers must not acquire the same lookup key.
            let encoded = self.encoded_unit(unit)?;
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


fn workbook_ascii_name_match(own: &[u16], request: &[u16]) -> Option<bool> {
    let own = own.split(|unit| *unit == 0).next().unwrap_or_default();
    let request = request.split(|unit| *unit == 0).next().unwrap_or_default();
    if request.is_empty() { return Some(false); }
    if !own.iter().chain(request).all(|unit| (0x20..=0x7e).contains(unit)) { return None; }
    let equal = |target: &[u16]| target.len() == request.len()
        && target.iter().zip(request).all(|(a,b)| (*a as u8).eq_ignore_ascii_case(&(*b as u8)));
    Some(equal(own) || own.iter().rposition(|unit| *unit == u16::from(b'.')).is_some_and(|at| equal(&own[..at])))
}

#[derive(serde::Deserialize)]
struct WorkbookNameMetadata {
    decompositions: Vec<(u32, Vec<u32>)>,
    specials: Vec<(u32, u8, u8, u8)>,
    repeat_skip: Vec<u32>,
    repeat_reset: Vec<u32>,
    previous_primary: Vec<(u32, Vec<u8>)>,
}
struct WorkbookNameWeights {
    bytes: Vec<u8>,
    offsets: Vec<u32>,
    metadata: WorkbookNameMetadata,
}
struct WorkbookNameRow<'a> {
    primary: &'a [u8],
    secondary: &'a [u8],
    case: &'a [u8],
    extra: &'a [u8],
    special: &'a [u8],
    slots: usize,
}
impl WorkbookNameWeights {
    fn from_archive(archive: &[u8]) -> Self {
        use std::io::Read;
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(archive)).expect("internal name weights archive");
        let mut bytes = Vec::new();
        archive.by_name("rows").expect("internal name rows").take(16 * 1024 * 1024).read_to_end(&mut bytes).expect("internal name row data");
        assert!(bytes.len() >= 8 && &bytes[..4] == b"WBW1");
        let count = u32::from_le_bytes(bytes[4..8].try_into().unwrap()) as usize;
        assert_eq!(count, 0x110000);
        let mut offsets = Vec::with_capacity(count);
        let mut at = 8usize;
        for _ in 0..count {
            offsets.push(u32::try_from(at).expect("name weight offset"));
            let header = bytes.get(at..at + 6).expect("internal name row header");
            at += 6 + header[..5].iter().map(|n| usize::from(*n)).sum::<usize>();
            assert!(at <= bytes.len());
        }
        assert_eq!(at, bytes.len());
        let metadata = serde_json::from_reader(archive.by_name("metadata").expect("internal name metadata")).expect("internal name metadata data");
        Self { bytes, offsets, metadata }
    }
    fn row(&self, scalar: u32) -> WorkbookNameRow<'_> {
        let mut at = self.offsets[scalar as usize] as usize;
        let header = &self.bytes[at..at + 6];
        at += 6;
        let mut part = |length: u8| {
            let start = at;
            at += usize::from(length);
            &self.bytes[start..at]
        };
        WorkbookNameRow { primary: part(header[0]), secondary: part(header[1]), case: part(header[2]),
            extra: part(header[3]), special: part(header[4]), slots: usize::from(header[5]) }
    }
    fn special(&self, scalar: u32) -> Option<(u8, u8, u8)> {
        self.metadata.specials.binary_search_by_key(&scalar, |row| row.0).ok().map(|at| {
            let row = self.metadata.specials[at];
            (row.1, row.2, row.3)
        })
    }
    fn decomposition(&self, scalar: u32) -> Option<&[u32]> {
        self.metadata.decompositions.binary_search_by_key(&scalar, |row| row.0).ok()
            .map(|at| self.metadata.decompositions[at].1.as_slice())
    }
    fn repeat_skips(&self, scalar: u32) -> bool {
        self.metadata.repeat_skip.binary_search(&scalar).is_ok()
    }
    fn previous_primary(&self, scalar: u32) -> Option<&[u8]> {
        self.metadata.previous_primary.binary_search_by_key(&scalar, |row| row.0).ok()
            .map(|at| self.metadata.previous_primary[at].1.as_slice())
    }
    fn key(&self, text: &[u16]) -> WorkbookNameKey {
        let mut builder = WorkbookNameBuilder::new(self);
        for scalar in char::decode_utf16(text.iter().copied()) {
            let scalar = scalar.map(u32::from).unwrap_or_else(|error| u32::from(error.unpaired_surrogate()));
            if scalar == 0 { break; }
            builder.push(scalar);
        }
        builder.finish()
    }
    fn matches(&self, own_name: &[u16], request: &[u16]) -> bool {
        let request = request.split(|unit| *unit == 0).next().unwrap_or_default();
        if request.is_empty() { return false; }
        if let Some(found) = workbook_ascii_name_match(own_name, request) { return found; }
        let own_name = own_name.split(|unit| *unit == 0).next().unwrap_or_default();
        let stem = own_name.iter().rposition(|unit| *unit == u16::from(b'.')).map(|at| &own_name[..at]);
        let wanted = [Some(self.key(own_name)), stem.map(|stem| self.key(stem))];
        let mut first_match = [None, None];
        let mut builder = WorkbookNameBuilder::new(self);
        let mut consumed = 0usize;
        for scalar in char::decode_utf16(request.iter().copied()) {
            let (scalar, units) = match scalar {
                Ok(scalar) => (u32::from(scalar), scalar.len_utf16()),
                Err(error) => (u32::from(error.unpaired_surrogate()), 1),
            };
            consumed += units;
            builder.push(scalar);
            for index in 0..2 {
                if first_match[index].is_none() && wanted[index].as_ref().is_some_and(|key| builder.equals(key)) {
                    first_match[index] = Some(consumed);
                }
            }
        }
        first_match.contains(&Some(request.len()))
    }
}
#[derive(Debug, PartialEq, Eq)]
struct WorkbookNameKey {
    primary: Vec<u8>,
    secondary: Vec<u8>,
    case: Vec<u8>,
    extra: [Vec<u8>; 4],
    special: Vec<(usize, Vec<u8>)>,
}
struct WorkbookNameBuilder<'a> {
    weights: &'a WorkbookNameWeights,
    key: WorkbookNameKey,
    position: usize,
    previous: Option<Vec<u8>>,
    previous_kana: Option<[u8; 4]>,
}
const WORKBOOK_NAME_EXTRA_DEFAULTS: [u8; 4] = [0xc6, 3, 0xe4, 0xc5];
fn workbook_name_extra(value: &[u8]) -> [Vec<u8>; 4] {
    if value.is_empty() { return std::array::from_fn(|_| Vec::new()); }
    let first = value.iter().position(|b| *b == 0xff).expect("name extra separator");
    let next = first + 1 + value[first + 1..].iter().position(|b| *b == 2).expect("name repetition separator");
    let third = next + 1 + value[next + 1..].iter().position(|b| *b == 0xff).expect("name kana separator");
    let fourth = third + 1 + value[third + 1..].iter().position(|b| *b == 0xff).expect("name width separator");
    assert_eq!(fourth + 1, value.len());
    [value[..first].to_vec(), value[first + 1..next].to_vec(), value[next + 1..third].to_vec(), value[third + 1..fourth].to_vec()]
}
fn workbook_name_trimmed(value: &[u8], default: u8) -> &[u8] {
    let end = value.iter().rposition(|b| *b != default).map_or(0, |at| at + 1);
    &value[..end]
}
impl<'a> WorkbookNameBuilder<'a> {
    fn new(weights: &'a WorkbookNameWeights) -> Self {
        Self { weights, key: WorkbookNameKey { primary: Vec::new(), secondary: Vec::new(), case: Vec::new(),
            extra: std::array::from_fn(|_| Vec::new()), special: Vec::new() },
            position: 0, previous: None, previous_kana: None }
    }
    fn push(&mut self, scalar: u32) {
        let saved = self.weights.repeat_skips(scalar).then(|| (self.previous.clone(), self.previous_kana));
        self.push_expanded(scalar);
        if self.weights.metadata.repeat_reset.binary_search(&scalar).is_ok() {
            self.previous = None;
            self.previous_kana = None;
        } else if let Some((previous, kana)) = saved {
            self.previous = previous;
            self.previous_kana = kana;
        } else if let Some(primary) = self.weights.previous_primary(scalar) {
            self.previous = Some(primary.to_vec());
            self.previous_kana = None;
        }
    }
    fn push_expanded(&mut self, scalar: u32) {
        let special = self.weights.special(scalar);
        if special.is_none() {
            if let Some(parts) = self.weights.decomposition(scalar) {
                for child in parts { self.push(*child); }
                return;
            }
        }
        let row = self.weights.row(scalar);
        let mut primary = row.primary.to_vec();
        let mut secondary = row.secondary.to_vec();
        let mut case = row.case.to_vec();
        let mut slots = row.slots;
        let mut extra = workbook_name_extra(row.extra);
        let mut generated = false;
        if let (Some((kind, width, diacritic)), Some(previous)) = (special.filter(|(kind, _, _)| *kind != 0), self.previous.as_ref()) {
            generated = true;
            primary = previous.clone();
            secondary = vec![diacritic];
            case.clear();
            slots = 1;
            extra = std::array::from_fn(|_| Vec::new());
            if let Some([small, _, kana, _]) = self.previous_kana {
                if kind == 5 { *primary.last_mut().expect("preceding kana") &= 0x87; }
                extra = [vec![small], vec![kind], vec![kana], vec![width]];
            }
        }
        if primary.is_empty() && !secondary.is_empty() && case.is_empty() && row.extra.is_empty() && row.special.is_empty() {
            if self.key.secondary.is_empty() { self.key.secondary.push(0); }
            let previous = self.key.secondary.last_mut().unwrap();
            for mark in secondary { *previous = previous.wrapping_add(mark); }
            return;
        }
        let has_extra = extra.iter().any(|part| !part.is_empty()) || !row.extra.is_empty();
        let mut last_extra = None;
        if has_extra {
            let count = slots.max(1).max(extra.iter().map(Vec::len).max().unwrap_or(0));
            let mut last = [0; 4];
            for index in 0..4 {
                extra[index].resize(count, WORKBOOK_NAME_EXTRA_DEFAULTS[index]);
                last[index] = *extra[index].last().unwrap();
                self.key.extra[index].extend_from_slice(&extra[index]);
            }
            last_extra = Some(last);
        }
        if !row.special.is_empty() && !generated { self.key.special.push((self.position, row.special.to_vec())); }
        self.key.primary.extend_from_slice(&primary);
        self.key.secondary.extend_from_slice(&secondary);
        self.key.secondary.extend(std::iter::repeat_n(2, slots.saturating_sub(secondary.len())));
        self.key.case.extend_from_slice(&case);
        self.key.case.extend(std::iter::repeat_n(2, slots.saturating_sub(case.len())));
        self.position += slots;
        if special.is_none() || generated {
            if !primary.is_empty() {
                let width = if slots != 0 && primary.len().is_multiple_of(slots) { primary.len() / slots } else { primary.len() };
                let previous = primary[primary.len() - width..].to_vec();
                self.previous_kana = if previous[0] == 0x22 { last_extra } else { None };
                self.previous = Some(previous);
            } else if !row.special.is_empty() {
                self.previous = Some(vec![6, row.special[row.special.len() - 2]]);
                self.previous_kana = None;
            }
        }
    }
    fn equals(&self, other: &WorkbookNameKey) -> bool {
        self.key.primary == other.primary && self.key.special == other.special
            && workbook_name_trimmed(&self.key.secondary, 2) == other.secondary
            && workbook_name_trimmed(&self.key.case, 2) == other.case
            && (0..4).all(|index| workbook_name_trimmed(&self.key.extra[index], WORKBOOK_NAME_EXTRA_DEFAULTS[index]) == other.extra[index])
    }
    fn finish(mut self) -> WorkbookNameKey {
        self.key.secondary.truncate(workbook_name_trimmed(&self.key.secondary, 2).len());
        self.key.case.truncate(workbook_name_trimmed(&self.key.case, 2).len());
        for (extra, default) in self.key.extra.iter_mut().zip(WORKBOOK_NAME_EXTRA_DEFAULTS) {
            let length = workbook_name_trimmed(extra, default).len();
            extra.truncate(length);
        }
        self.key
    }
}

pub(crate) fn workbook_name_matches(own_name: &str, request: &[u16]) -> bool {
    let own_name: Vec<u16> = own_name.encode_utf16().collect();
    if let Some(found) = workbook_ascii_name_match(&own_name, request) { return found; }
    static WEIGHTS: std::sync::OnceLock<WorkbookNameWeights> = std::sync::OnceLock::new();
    let weights = WEIGHTS.get_or_init(|| {
        let data = include_bytes!("data/vba_automation_weights.bin");
        WorkbookNameWeights::from_archive(&data[256 * 10..])
    });
    weights.matches(&own_name, request)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_workbook_name_lookup_matches_excel() {
        let corpus: serde_json::Value = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/automation_sort_keys.json"
        )).unwrap();
        for case in corpus["workbooks"].as_array().unwrap() {
            let own_name = case["file_name"].as_str().unwrap();
            let request: Vec<u16> = case["request"].as_str().unwrap().encode_utf16().collect();
            assert_eq!(workbook_name_matches(own_name, &request), case["expected"].as_bool().unwrap(),
                "workbook {:?}, request {:?}", own_name, request);
        }
    }


    #[test]
    fn composed_automation_keys_match_native_windows_vectors() {
        let corpus: serde_json::Value = serde_json::from_str(include_str!(
            "../../../tests/fixtures/vba/automation_sort_keys.json"
        )).unwrap();
        for case in corpus["ansi"].as_array().unwrap() {
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
