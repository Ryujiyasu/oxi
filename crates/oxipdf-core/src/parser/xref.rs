use std::collections::HashMap;

use crate::error::PdfError;
use super::object::ObjRef;

/// Cross-reference table: maps object references to byte offsets.
#[derive(Debug, Clone)]
pub struct XrefTable {
    pub entries: HashMap<u32, XrefEntry>,
}

#[derive(Debug, Clone, Copy)]
#[allow(dead_code)]
pub enum XrefEntry {
    /// In-use object at the given byte offset.
    InUse { offset: u64, gen: u16 },
    /// Free object.
    Free { next_free: u32, gen: u16 },
    /// Object stored in an object stream (PDF 1.5+).
    Compressed { stream_obj: u32, index: u16 },
}

impl XrefTable {
    pub fn new() -> Self {
        Self {
            entries: HashMap::new(),
        }
    }

    /// Get the byte offset of an in-use object.
    pub fn get_offset(&self, obj_num: u32) -> Option<u64> {
        match self.entries.get(&obj_num) {
            Some(XrefEntry::InUse { offset, .. }) => Some(*offset),
            _ => None,
        }
    }
}

/// The PDF trailer dictionary contents.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct Trailer {
    /// Total number of entries in the xref table.
    pub size: u32,
    /// Reference to the document catalog.
    pub root: ObjRef,
    /// Reference to the document info dictionary (optional).
    pub info: Option<ObjRef>,
    /// File identifiers (optional).
    pub id: Option<(Vec<u8>, Vec<u8>)>,
    /// Byte offset of the previous xref section (for incremental updates).
    pub prev: Option<u64>,
}

/// Find the `startxref` offset by scanning backwards from end of file.
pub fn find_startxref(data: &[u8]) -> Result<u64, PdfError> {
    // PDF spec: startxref is near the end, within the last 1024 bytes.
    let search_start = data.len().saturating_sub(1024);
    let tail = &data[search_start..];

    let marker = b"startxref";
    let pos = tail
        .windows(marker.len())
        .rposition(|w| w == marker)
        .ok_or(PdfError::Parse("startxref not found".into()))?;

    // Skip "startxref" + whitespace, read the offset number.
    let after = &tail[pos + marker.len()..];
    let offset_str: String = after
        .iter()
        .skip_while(|b| b.is_ascii_whitespace())
        .take_while(|b| b.is_ascii_digit())
        .map(|&b| b as char)
        .collect();

    offset_str
        .parse::<u64>()
        .map_err(|_| PdfError::Parse(format!("invalid startxref offset: {offset_str}")))
}

/// Parse a traditional (non-stream) xref table starting at `offset`.
pub fn parse_xref_table(data: &[u8], offset: u64) -> Result<(XrefTable, usize), PdfError> {
    let mut pos = offset as usize;
    let mut table = XrefTable::new();

    // Skip "xref" keyword + whitespace.
    if data.len() < pos + 4 || &data[pos..pos + 4] != b"xref" {
        return Err(PdfError::Parse("expected 'xref' keyword".into()));
    }
    pos += 4;
    pos = skip_whitespace(data, pos);

    // Parse subsections.
    loop {
        if pos >= data.len() || data[pos] == b't' {
            // 't' for "trailer"
            break;
        }

        // Read "start_obj count"
        let (start_obj, new_pos) = read_u32(data, pos)?;
        pos = skip_whitespace(data, new_pos);
        let (count, new_pos) = read_u32(data, pos)?;
        pos = skip_whitespace(data, new_pos);

        for i in 0..count {
            if pos + 20 > data.len() {
                return Err(PdfError::Parse("xref entry truncated".into()));
            }
            let line = &data[pos..pos + 20];
            let offset_val: u64 = std::str::from_utf8(&line[0..10])
                .map_err(|_| PdfError::Parse("invalid xref offset".into()))?
                .trim()
                .parse()
                .map_err(|_| PdfError::Parse("invalid xref offset number".into()))?;
            let gen: u16 = std::str::from_utf8(&line[11..16])
                .map_err(|_| PdfError::Parse("invalid xref gen".into()))?
                .trim()
                .parse()
                .map_err(|_| PdfError::Parse("invalid xref gen number".into()))?;
            let flag = line[17];

            let obj_num = start_obj + i;
            let entry = if flag == b'n' {
                XrefEntry::InUse {
                    offset: offset_val,
                    gen,
                }
            } else {
                XrefEntry::Free {
                    next_free: offset_val as u32,
                    gen,
                }
            };
            table.entries.insert(obj_num, entry);

            pos += 20;
        }
        pos = skip_whitespace(data, pos);
    }

    Ok((table, pos))
}

fn skip_whitespace(data: &[u8], mut pos: usize) -> usize {
    while pos < data.len() && data[pos].is_ascii_whitespace() {
        pos += 1;
    }
    pos
}

fn read_u32(data: &[u8], pos: usize) -> Result<(u32, usize), PdfError> {
    let s: String = data[pos..]
        .iter()
        .take_while(|b| b.is_ascii_digit())
        .map(|&b| b as char)
        .collect();
    if s.is_empty() {
        return Err(PdfError::Parse(format!(
            "expected integer at offset {pos}"
        )));
    }
    let val = s
        .parse::<u32>()
        .map_err(|_| PdfError::Parse(format!("invalid u32: {s}")))?;
    Ok((val, pos + s.len()))
}
