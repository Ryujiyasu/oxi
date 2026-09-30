// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The files a macro opens with `Open`, kept in memory for the run.
//!
//! A browser has no disk to give a macro, so the files it writes live here
//! and are there to read back until the run ends -- which is what a macro
//! writing a report and reading it back, or keeping a scratch file, needs.
//! Paths are compared the way Windows compares them: case aside, `/` and
//! `\` alike. Text is kept as UTF-8, so a byte count (`LOF`, `FileLen`)
//! matches Excel's for ASCII and not for text its ANSI code page writes in
//! two bytes.

use std::collections::BTreeMap;

/// How a file was opened.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OpenMode {
    Input,
    Output,
    Append,
}

/// One open file.
#[derive(Debug, Clone)]
pub struct Handle {
    pub path: String,
    pub mode: OpenMode,
    /// Where the next read starts, in bytes.
    pub position: usize,
    /// The column the next `Print` field starts at, counted from 0.
    pub column: usize,
}

/// Why a file operation failed, as the VBA error number it raises.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileError {
    /// 52, Bad file name or number.
    BadNumber = 52,
    /// 53, File not found.
    NotFound = 53,
    /// 55, File already open.
    AlreadyOpen = 55,
    /// 58, File already exists.
    AlreadyExists = 58,
    /// 62, Input past end of file.
    PastEnd = 62,
}

impl FileError {
    pub fn number(self) -> i64 {
        self as i64
    }

    pub fn description(self) -> &'static str {
        match self {
            FileError::BadNumber => "Bad file name or number",
            FileError::NotFound => "File not found",
            FileError::AlreadyOpen => "File already open",
            FileError::AlreadyExists => "File already exists",
            FileError::PastEnd => "Input past end of file",
        }
    }
}

/// The files of a run and the numbers they are open under.
#[derive(Debug, Default, Clone)]
pub struct Files {
    contents: BTreeMap<String, (String, Vec<u8>)>,
    open: BTreeMap<i64, Handle>,
    /// What `Dir` with no arguments goes on to answer.
    listing: Vec<String>,
}

/// A path as Windows compares it.
fn key(path: &str) -> String {
    path.trim().replace('/', "\\").to_lowercase()
}

/// The last part of a path.
fn file_name(path: &str) -> &str {
    path.rsplit(['\\', '/']).next().unwrap_or(path)
}

/// Whether a file name answers a `Dir` or `Kill` pattern of `*` and `?`.
fn matches_pattern(name: &str, pattern: &str) -> bool {
    let name: Vec<char> = name.to_lowercase().chars().collect();
    let pattern: Vec<char> = pattern.to_lowercase().chars().collect();
    fn walk(name: &[char], pattern: &[char]) -> bool {
        match pattern.split_first() {
            None => name.is_empty(),
            Some(('*', rest)) => (0..=name.len()).any(|skip| walk(&name[skip..], rest)),
            Some(('?', rest)) => !name.is_empty() && walk(&name[1..], rest),
            Some((one, rest)) => name.first() == Some(one) && walk(&name[1..], rest),
        }
    }
    walk(&name, &pattern)
}

impl Files {
    /// `FreeFile`: the lowest number not in use, from 1, or from 256 for
    /// `FreeFile(1)`.
    pub fn free_number(&self, upper_range: bool) -> i64 {
        let start = if upper_range { 256 } else { 1 };
        (start..start + 255).find(|number| !self.open.contains_key(number)).unwrap_or(0)
    }

    pub fn open(&mut self, path: &str, mode: OpenMode, number: i64) -> Result<(), FileError> {
        if !(1..=511).contains(&number) {
            return Err(FileError::BadNumber);
        }
        if self.open.contains_key(&number) {
            return Err(FileError::AlreadyOpen);
        }
        let name = key(path);
        if name.is_empty() {
            return Err(FileError::BadNumber);
        }
        match mode {
            OpenMode::Input if !self.contents.contains_key(&name) => return Err(FileError::NotFound),
            OpenMode::Output => {
                self.contents.insert(name.clone(), (path.trim().to_string(), Vec::new()));
            }
            OpenMode::Append => {
                self.contents.entry(name.clone()).or_insert_with(|| (path.trim().to_string(), Vec::new()));
            }
            OpenMode::Input => {}
        }
        let column = match mode {
            OpenMode::Append => {
                let bytes = &self.contents[&name].1;
                let last_line = bytes.iter().rposition(|byte| *byte == b'\n').map_or(0, |at| at + 1);
                bytes.len() - last_line
            }
            _ => 0,
        };
        self.open.insert(number, Handle { path: name, mode, position: 0, column });
        Ok(())
    }

    /// `Close`: the numbers given, or every file when none is.
    pub fn close(&mut self, numbers: &[i64]) {
        if numbers.is_empty() {
            self.open.clear();
        } else {
            for number in numbers {
                self.open.remove(number);
            }
        }
    }

    pub fn handle(&self, number: i64) -> Result<&Handle, FileError> {
        self.open.get(&number).ok_or(FileError::BadNumber)
    }

    /// Write text to a file open for output, keeping track of its column.
    pub fn write(&mut self, number: i64, text: &str) -> Result<(), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if handle.mode == OpenMode::Input {
            return Err(FileError::BadNumber);
        }
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        file.1.extend_from_slice(text.as_bytes());
        match text.rfind('\n') {
            Some(at) => handle.column = text[at + 1..].chars().count(),
            None => handle.column += text.chars().count(),
        }
        Ok(())
    }

    pub fn column(&self, number: i64) -> Result<usize, FileError> {
        Ok(self.handle(number)?.column)
    }

    /// `LOF`: the file's length in bytes.
    pub fn length_of_open(&self, number: i64) -> Result<usize, FileError> {
        let handle = self.handle(number)?;
        Ok(self.contents.get(&handle.path).map_or(0, |file| file.1.len()))
    }

    /// `EOF`: whether reading has reached the end.
    pub fn at_end(&self, number: i64) -> Result<bool, FileError> {
        let handle = self.handle(number)?;
        let length = self.contents.get(&handle.path).map_or(0, |file| file.1.len());
        Ok(handle.position >= length)
    }

    /// `Loc` of a sequential file: the bytes read so far over 128.
    pub fn location(&self, number: i64) -> Result<i64, FileError> {
        let handle = self.handle(number)?;
        Ok((handle.position as i64 + 127) / 128)
    }

    fn reading(&mut self, number: i64) -> Result<(&[u8], &mut usize), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if handle.mode != OpenMode::Input {
            return Err(FileError::BadNumber);
        }
        let bytes = self.contents.get(&handle.path).map(|file| file.1.as_slice()).unwrap_or(&[]);
        Ok((bytes, &mut handle.position))
    }

    /// `Line Input #`: the next line, without its line break.
    pub fn read_line(&mut self, number: i64) -> Result<String, FileError> {
        let (bytes, position) = self.reading(number)?;
        if *position >= bytes.len() {
            return Err(FileError::PastEnd);
        }
        let start = *position;
        let mut end = start;
        while end < bytes.len() && bytes[end] != b'\r' && bytes[end] != b'\n' {
            end += 1;
        }
        let line = String::from_utf8_lossy(&bytes[start..end]).into_owned();
        let mut next = end;
        if next < bytes.len() && bytes[next] == b'\r' {
            next += 1;
        }
        if next < bytes.len() && bytes[next] == b'\n' {
            next += 1;
        }
        *position = next;
        Ok(line)
    }

    /// `Input #`: the next field as written, and whether it was quoted.
    pub fn read_field(&mut self, number: i64) -> Result<(String, bool), FileError> {
        let (bytes, position) = self.reading(number)?;
        let mut at = *position;
        while at < bytes.len() && matches!(bytes[at], b' ' | b'\t') {
            at += 1;
        }
        if at >= bytes.len() {
            return Err(FileError::PastEnd);
        }
        let (field, quoted) = if bytes[at] == b'"' {
            let start = at + 1;
            let mut end = start;
            while end < bytes.len() && bytes[end] != b'"' {
                end += 1;
            }
            at = (end + 1).min(bytes.len());
            // Past the closing quote, up to the next delimiter.
            while at < bytes.len() && !matches!(bytes[at], b',' | b'\r' | b'\n') {
                at += 1;
            }
            (String::from_utf8_lossy(&bytes[start..end]).into_owned(), true)
        } else {
            let start = at;
            while at < bytes.len() && !matches!(bytes[at], b',' | b'\r' | b'\n') {
                at += 1;
            }
            (String::from_utf8_lossy(&bytes[start..at]).trim().to_string(), false)
        };
        // One delimiter goes with the field: a comma, or a line break.
        if at < bytes.len() && bytes[at] == b',' {
            at += 1;
        } else {
            if at < bytes.len() && bytes[at] == b'\r' {
                at += 1;
            }
            if at < bytes.len() && bytes[at] == b'\n' {
                at += 1;
            }
        }
        *position = at;
        Ok((field, quoted))
    }

    /// `FileLen`: a closed or open file's length.
    pub fn length(&self, path: &str) -> Result<usize, FileError> {
        self.contents.get(&key(path)).map(|file| file.1.len()).ok_or(FileError::NotFound)
    }

    fn names_matching(&self, pattern: &str) -> Vec<String> {
        let wanted = key(pattern);
        let (folder, leaf) = match wanted.rfind('\\') {
            Some(at) => (&wanted[..=at], &wanted[at + 1..]),
            None => ("", wanted.as_str()),
        };
        self.contents
            .iter()
            .filter(|(held, _)| {
                let (held_folder, held_leaf) = match held.rfind('\\') {
                    Some(at) => (&held[..=at], &held[at + 1..]),
                    None => ("", held.as_str()),
                };
                held_folder == folder && matches_pattern(held_leaf, leaf)
            })
            .map(|(_, (written, _))| file_name(written).to_string())
            .collect()
    }

    /// `Dir(pattern)`: the first file answering it, or "" -- and `Dir()`
    /// the next.
    pub fn dir(&mut self, pattern: Option<&str>) -> String {
        if let Some(pattern) = pattern {
            self.listing = self.names_matching(pattern);
            self.listing.reverse();
        }
        self.listing.pop().unwrap_or_default()
    }

    /// `Kill`: every file answering the pattern.
    pub fn kill(&mut self, pattern: &str) -> Result<(), FileError> {
        let wanted = key(pattern);
        let folder = wanted.rfind('\\').map_or(String::new(), |at| wanted[..=at].to_string());
        let names = self.names_matching(pattern);
        if names.is_empty() {
            return Err(FileError::NotFound);
        }
        for name in names {
            let full = key(&format!("{folder}{name}"));
            if self.open.values().any(|handle| handle.path == full) {
                return Err(FileError::AlreadyOpen);
            }
            self.contents.remove(&full);
        }
        Ok(())
    }

    /// `FileCopy`.
    pub fn copy(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        let bytes = self.contents.get(&key(source)).ok_or(FileError::NotFound)?.1.clone();
        self.contents.insert(key(destination), (destination.trim().to_string(), bytes));
        Ok(())
    }

    /// `Name ... As ...`.
    pub fn rename(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if self.contents.contains_key(&key(destination)) {
            return Err(FileError::AlreadyExists);
        }
        let (_, bytes) = self.contents.remove(&key(source)).ok_or(FileError::NotFound)?;
        self.contents.insert(key(destination), (destination.trim().to_string(), bytes));
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_file_written_reads_back_line_by_line_and_field_by_field() {
        let mut files = Files::default();
        files.open("C:\\tmp\\a.txt", OpenMode::Output, 1).unwrap();
        files.write(1, "one\r\n\"q, x\",2.5,#TRUE#\r\n").unwrap();
        files.close(&[1]);
        assert_eq!(files.length("c:/TMP/A.TXT"), Ok(24));
        files.open("C:\\tmp\\a.txt", OpenMode::Input, 1).unwrap();
        assert_eq!(files.read_line(1).unwrap(), "one");
        assert_eq!(files.read_field(1).unwrap(), ("q, x".to_string(), true));
        assert_eq!(files.read_field(1).unwrap(), ("2.5".to_string(), false));
        assert_eq!(files.read_field(1).unwrap(), ("#TRUE#".to_string(), false));
        assert!(files.at_end(1).unwrap());
        assert_eq!(files.dir(Some("C:\\tmp\\*.txt")), "a.txt");
        assert_eq!(files.dir(None), "");
    }
}
