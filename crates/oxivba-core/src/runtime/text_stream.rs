// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A TextStream's encoding and character cursor, separate from file bytes.

use crate::{files::{FileError, Files, OpenMode}, AnsiCodec};
use std::collections::VecDeque;

#[derive(Debug, Clone, Copy)]
enum Encoding {
    Ansi(AnsiCodec),
    Utf16 { big_endian: bool },
}

#[derive(Debug, Clone)]
pub(super) struct Stream {
    pub number: i64,
    pub mode: OpenMode,
    pub line: i64,
    pub column: i64,
    encoding: Encoding,
    pending: VecDeque<u16>,
}

impl Stream {
    pub fn open(files: &mut Files, number: i64, format: i64, codec: AnsiCodec) -> Result<Self, FileError> {
        if !matches!(format, 0 | -1 | -2) {
            files.close(&[number]);
            return Err(FileError::InvalidArgument);
        }
        let handle = files.handle(number)?;
        let mode = handle.mode;
        let bytes = files.bytes(&handle.path)?;
        let unicode = format == -1 || format == -2 && bytes.starts_with(&[0xff, 0xfe]) && bytes.len() % 2 == 0;
        let encoding = if unicode {
            Encoding::Utf16 { big_endian: mode == OpenMode::Input && bytes.starts_with(&[0xfe, 0xff]) }
        } else { Encoding::Ansi(codec) };
        let has_bom = matches!(encoding, Encoding::Utf16 { .. })
            && (bytes.starts_with(&[0xff, 0xfe]) || bytes.starts_with(&[0xfe, 0xff]));
        let empty = bytes.is_empty();
        let mut stream = Self { number, mode, line: 1, column: 1, encoding, pending: VecDeque::new() };
        if mode == OpenMode::Input {
            if has_bom { files.advance_stream(number, 2)?; }
        } else if unicode && empty {
            files.write_stream(number, &[0xff, 0xfe])?;
        } else if mode == OpenMode::Append && !unicode {
            // ANSI append resumes its character position. Unicode append starts
            // a fresh character cursor while retaining the existing file bytes.
            for unit in codec.decode(bytes) { stream.advance_character(unit); }
        }
        Ok(stream)
    }

    fn decoded(&self, files: &Files, position: usize) -> Result<(Vec<u16>, usize), FileError> {
        let handle = files.handle(self.number)?;
        let bytes = files.bytes(&handle.path)?;
        let bytes = bytes.get(position..).unwrap_or_default();
        Ok(match self.encoding {
            Encoding::Ansi(codec) => codec.decode_prefix(bytes),
            Encoding::Utf16 { big_endian } => {
                if bytes.len() < 2 { (Vec::new(), 0) }
                else {
                    let pair = [bytes[0], bytes[1]];
                    (vec![if big_endian { u16::from_be_bytes(pair) } else { u16::from_le_bytes(pair) }], 2)
                }
            }
        })
    }

    fn peek(&self, files: &Files, count: usize) -> Result<Vec<u16>, FileError> {
        let mut units: Vec<_> = self.pending.iter().copied().take(count).collect();
        let mut position = files.handle(self.number)?.position;
        while units.len() < count {
            let (next, consumed) = self.decoded(files, position)?;
            if consumed == 0 { break; }
            position += consumed;
            units.extend(next.into_iter().take(count - units.len()));
        }
        Ok(units)
    }

    pub fn at_end(&self, files: &Files) -> Result<bool, FileError> {
        Ok(self.peek(files, 1)?.is_empty())
    }

    pub fn at_end_of_line(&self, files: &Files) -> Result<bool, FileError> {
        let units = self.peek(files, 2)?;
        Ok(units.is_empty() || units[0] == 10 || units.as_slice() == [13, 10])
    }

    fn advance_character(&mut self, unit: u16) {
        if unit == 10 { self.line += 1; self.column = 1; }
        else { self.column += 1; }
    }

    fn next(&mut self, files: &mut Files) -> Result<Option<u16>, FileError> {
        if self.pending.is_empty() {
            let position = files.handle(self.number)?.position;
            let (units, consumed) = self.decoded(files, position)?;
            if consumed == 0 { return Ok(None); }
            files.advance_stream(self.number, consumed)?;
            self.pending.extend(units);
        }
        let next = self.pending.pop_front();
        if let Some(unit) = next { self.advance_character(unit); }
        Ok(next)
    }

    pub fn read(&mut self, files: &mut Files, count: Option<usize>, line_only: bool) -> Result<Vec<u16>, FileError> {
        if self.at_end(files)? { return Err(FileError::PastEnd); }
        let mut units = Vec::new();
        while count.is_none_or(|count| units.len() < count) {
            let Some(unit) = self.next(files)? else { break; };
            if line_only && unit == 10 {
                if units.last() == Some(&13) { units.pop(); }
                break;
            }
            units.push(unit);
        }
        Ok(units)
    }

    pub fn write(&mut self, files: &mut Files, units: &[u16]) -> Result<(), FileError> {
        let bytes = match self.encoding {
            Encoding::Ansi(codec) => codec.encode_checked(units).ok_or(FileError::InvalidArgument)?,
            Encoding::Utf16 { .. } => units.iter().flat_map(|unit| unit.to_le_bytes()).collect(),
        };
        files.write_stream(self.number, &bytes)?;
        for &unit in units { self.advance_character(unit); }
        Ok(())
    }
}
