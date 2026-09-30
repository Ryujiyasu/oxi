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
    /// `For Binary`: bytes read and written at a position.
    Binary,
    /// `For Random Len = n`: records of n bytes.
    Random(usize),
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
    /// 76, Path not found.
    PathNotFound = 76,
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
            FileError::PathNotFound => "Path not found",
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
    /// Folders made with MkDir or CreateFolder, by key.
    folders: std::collections::BTreeSet<String>,
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
            OpenMode::Append | OpenMode::Binary | OpenMode::Random(_) => {
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

    /// A number for a TextStream, from 1000 on, never one FreeFile hands out.
    pub fn open_stream(&mut self, path: &str, mode: OpenMode) -> Result<i64, FileError> {
        let number = (1000..).find(|number| !self.open.contains_key(number)).unwrap_or(1000);
        let name = key(path);
        match mode {
            OpenMode::Input if !self.contents.contains_key(&name) => return Err(FileError::NotFound),
            OpenMode::Output => {
                self.contents.insert(name.clone(), (path.trim().to_string(), Vec::new()));
            }
            _ => {
                self.contents.entry(name.clone()).or_insert_with(|| (path.trim().to_string(), Vec::new()));
            }
        }
        self.open.insert(number, Handle { path: name, mode, position: 0, column: 0 });
        Ok(number)
    }

    pub fn exists(&self, path: &str) -> bool {
        self.contents.contains_key(&key(path))
    }

    /// Whether the folder was made, or any kept file lies in it.
    pub fn folder_exists(&self, path: &str) -> bool {
        let folder = key(path);
        let folder = folder.trim_end_matches('\\');
        self.folders.contains(folder)
            || self.folders.iter().any(|held| held.starts_with(&format!("{folder}\\")))
            || self.contents.keys().any(|held| held.rsplit_once('\\').is_some_and(|(parent, _)| parent == folder || parent.starts_with(&format!("{folder}\\"))))
    }

    /// `MkDir` / `CreateFolder`: 58 when it is there already.
    pub fn make_folder(&mut self, path: &str) -> Result<(), FileError> {
        let folder = key(path).trim_end_matches('\\').to_string();
        if self.folders.contains(&folder) {
            return Err(FileError::AlreadyExists);
        }
        self.folders.insert(folder);
        Ok(())
    }

    /// `RmDir` / `DeleteFolder`: the folder and, for DeleteFolder, what is
    /// in it; 76 when there is no such folder.
    pub fn remove_folder(&mut self, path: &str, with_contents: bool) -> Result<(), FileError> {
        let folder = key(path).trim_end_matches('\\').to_string();
        if !self.folder_exists(&folder) {
            return Err(FileError::PathNotFound);
        }
        let inside = format!("{folder}\\");
        if with_contents {
            self.contents.retain(|held, _| !held.starts_with(&inside));
            self.folders.retain(|held| !held.starts_with(&inside));
        }
        self.folders.remove(&folder);
        Ok(())
    }

    /// The written names of the files directly in a folder, in order.
    pub fn files_in(&self, path: &str) -> Vec<String> {
        let folder = key(path).trim_end_matches('\\').to_string();
        self.contents
            .iter()
            .filter(|(held, _)| held.rsplit_once('\\').is_some_and(|(parent, _)| parent == folder))
            .map(|(_, (written, _))| written.clone())
            .collect()
    }

    /// What is left to read, for ReadAll.
    pub fn read_rest(&mut self, number: i64) -> Result<String, FileError> {
        let (bytes, position) = self.reading(number)?;
        if *position >= bytes.len() {
            return Err(FileError::PastEnd);
        }
        let rest = String::from_utf8_lossy(&bytes[*position..]).into_owned();
        *position = bytes.len();
        Ok(rest)
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

    /// `Loc`: for a binary file the last byte read or written, for a
    /// random one the last record, and for a sequential one the bytes read
    /// so far over 128.
    pub fn location(&self, number: i64) -> Result<i64, FileError> {
        let handle = self.handle(number)?;
        Ok(match handle.mode {
            OpenMode::Binary => handle.position as i64,
            OpenMode::Random(length) => (handle.position / length.max(1)) as i64,
            _ => (handle.position as i64 + 127) / 128,
        })
    }

    /// `Seek(n)`: where the next read or write goes, from 1 -- a byte, or a
    /// record in a random file.
    pub fn next_position(&self, number: i64) -> Result<i64, FileError> {
        let handle = self.handle(number)?;
        Ok(match handle.mode {
            OpenMode::Random(length) => (handle.position / length.max(1)) as i64 + 1,
            _ => handle.position as i64 + 1,
        })
    }

    /// `Seek #n, position`, and the position a `Get` or `Put` names.
    pub fn move_to(&mut self, number: i64, position: i64) -> Result<(), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if position < 1 {
            return Err(FileError::BadNumber);
        }
        handle.position = match handle.mode {
            OpenMode::Random(length) => (position as usize - 1) * length,
            _ => position as usize - 1,
        };
        Ok(())
    }

    /// The record length of a random file, or None.
    pub fn record_length(&self, number: i64) -> Result<Option<usize>, FileError> {
        Ok(match self.handle(number)?.mode {
            OpenMode::Random(length) => Some(length),
            _ => None,
        })
    }

    /// `Put`: bytes at the current position, the file growing (with
    /// noughts across any gap) to take them.
    pub fn put(&mut self, number: i64, bytes: &[u8]) -> Result<(), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if !matches!(handle.mode, OpenMode::Binary | OpenMode::Random(_)) {
            return Err(FileError::BadNumber);
        }
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        let end = handle.position + bytes.len();
        if file.1.len() < end {
            file.1.resize(end, 0);
        }
        file.1[handle.position..end].copy_from_slice(bytes);
        handle.position = end;
        Ok(())
    }

    /// After a whole `Get` or `Put` in a random file, the next record: one
    /// read or written short still takes its whole length.
    pub fn finish_record(&mut self, number: i64) {
        if let Some(handle) = self.open.get_mut(&number) {
            if let OpenMode::Random(length) = handle.mode {
                let length = length.max(1);
                handle.position = handle.position.div_ceil(length) * length;
            }
        }
    }

    /// `Get`: so many bytes from the current position, noughts past the end.
    pub fn get(&mut self, number: i64, count: usize) -> Result<Vec<u8>, FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if !matches!(handle.mode, OpenMode::Binary | OpenMode::Random(_)) {
            return Err(FileError::BadNumber);
        }
        let file = self.contents.get(&handle.path).map(|file| file.1.as_slice()).unwrap_or(&[]);
        let mut read = vec![0u8; count];
        for (index, slot) in read.iter_mut().enumerate() {
            if let Some(byte) = file.get(handle.position + index) {
                *slot = *byte;
            }
        }
        handle.position += count;
        Ok(read)
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

    /// `Input(n, #f)`: the next n characters, line breaks and all.
    pub fn read_characters(&mut self, number: i64, count: usize) -> Result<String, FileError> {
        let (bytes, position) = self.reading(number)?;
        if *position + count > bytes.len() {
            return Err(FileError::PastEnd);
        }
        let read = String::from_utf8_lossy(&bytes[*position..*position + count]).into_owned();
        *position += count;
        Ok(read)
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
