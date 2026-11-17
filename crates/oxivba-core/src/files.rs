// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The files a macro opens with `Open`, kept in memory for the run.
//!
//! A browser has no disk to give a macro, so the files it writes live here
//! and are there to read back until the run ends -- which is what a macro
//! writing a report and reading it back, or keeping a scratch file, needs.
//! Paths are compared the way Windows compares them: case aside, `/` and
//! `\` alike. Sequential VBA output is encoded by the host ANSI profile;
//! binary reads and writes retain the original bytes.

use std::collections::BTreeMap;
mod wildcards;

/// Process directory facts, independent of project variables and file handles.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirectoryState {
    pub current_drive: char,
    /// One absolute remembered directory per available drive.
    pub directories: Vec<String>,
}

/// Closed filesystem contents, independent of a project's variables, handles
/// and Dir iterator. The host can carry this across project replacement.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileSystemSnapshot {
    pub folders: Vec<String>,
    pub files: Vec<(String, Vec<u8>)>,
    /// Attribute overrides; omitted entries use archive32/file or directory16.
    pub attributes: Vec<(String, i64)>,
}

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
    /// The access granted when the file opened, independent of later flags.
    pub can_read: bool,
    pub can_write: bool,
    pub share_read: bool,
    pub share_write: bool,
}

#[derive(Debug, Clone)]
struct RecordRead {
    start: usize,
    logical_start: usize,
    offset: usize,
    length: usize,
    span: usize,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct FileCursor {
    physical: usize,
    input_origin: usize,
    read_eof: bool,
    record: Option<RecordRead>,
}

impl FileCursor {
    fn text_end_position(&self, at: usize, length: usize) -> usize {
        let start = self.input_origin.saturating_add(at.saturating_sub(self.input_origin) / 512 * 512);
        start.saturating_add(512).min(length).saturating_add(at.saturating_sub(start)).saturating_add(1)
    }
}

#[derive(Debug, Clone)]
struct RecordLock {
    owner: (u64, i64),
    path: String,
    start: u64,
    end: u64,
}

/// Why a file operation failed, as the VBA error number it raises.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileError {
    /// 5, Invalid procedure call or argument.
    InvalidArgument = 5,
    /// 7, Out of memory.
    OutOfMemory = 7,
    /// 52, Bad file name or number.
    BadNumber = 52,
    /// 53, File not found.
    NotFound = 53,
    /// 54, Bad file mode.
    BadMode = 54,
    /// 55, File already open.
    AlreadyOpen = 55,
    /// 58, File already exists.
    AlreadyExists = 58,
    /// 59, Bad record length.
    BadRecordLength = 59,
    /// 62, Input past end of file.
    PastEnd = 62,
    /// 63, Bad record number.
    BadRecordNumber = 63,
    /// 68, Device unavailable.
    DeviceUnavailable = 68,
    /// 70, Permission denied.
    PermissionDenied = 70,
    /// 75, Path/File access error.
    PathFileAccess = 75,
    /// 76, Path not found.
    PathNotFound = 76,
    /// 458, Variable uses an Automation type not supported in Visual Basic.
    UnsupportedType = 458,
}

impl FileError {
    pub fn number(self) -> i64 {
        self as i64
    }

    pub fn description(self) -> &'static str {
        match self {
            FileError::InvalidArgument => "Invalid procedure call or argument",
            FileError::OutOfMemory => "Out of memory",
            FileError::BadNumber => "Bad file name or number",
            FileError::NotFound => "File not found",
            FileError::BadMode => "Bad file mode",
            FileError::BadRecordLength => "Bad record length",
            FileError::UnsupportedType => "Variable uses an Automation type not supported in Visual Basic",
            FileError::AlreadyOpen => "File already open",
            FileError::AlreadyExists => "File already exists",
            FileError::PastEnd => "Input past end of file",
            FileError::BadRecordNumber => "Bad record number",
            FileError::DeviceUnavailable => "Device unavailable",
            FileError::PermissionDenied => "Permission denied",
            FileError::PathFileAccess => "Path/File access error",
            FileError::PathNotFound => "Path not found",
        }
    }
}

#[derive(Debug, Clone)]
struct PendingWrite {
    written: f64,
    offset: usize,
    buffered: Vec<u8>,
}

#[derive(Debug, Clone)]
struct FileEntry {
    length: usize,
    modified: f64,
}

/// The files of a run and the numbers they are open under.
#[derive(Debug, Clone)]
pub struct Files {
    contents: BTreeMap<String, (String, Vec<u8>)>,
    open: BTreeMap<(u64, i64), Handle>,
    project_scope: u64,
    cursors: BTreeMap<(u64, i64), FileCursor>,
    /// Stream IDs are not reused while old TextStream references can exist.
    next_stream_number: i64,
    /// What `Dir` with no arguments goes on to answer.
    listings: BTreeMap<u64, Vec<String>>,
    /// Folders made with MkDir or CreateFolder, by key.
    folders: std::collections::BTreeSet<String>,
    current_drive: char,
    drive_directories: BTreeMap<char, String>,
    attributes: BTreeMap<String, i64>,
    volume_labels: BTreeMap<String,String>,
    record_locks: Vec<RecordLock>,
    modified: BTreeMap<String, f64>,
    directory_entries: BTreeMap<String, f64>,
    file_entries: BTreeMap<String, FileEntry>,
    current_time: f64,
    clock: Option<crate::clock::Clock>,
    pending_writes: BTreeMap<(u64, i64), PendingWrite>,
}

impl Default for Files {
    fn default() -> Self {
        Self {
            contents: BTreeMap::new(), open: BTreeMap::new(), project_scope: 0, next_stream_number: 1000, listings: BTreeMap::new(),
            folders: Default::default(), current_drive: 'C',
            drive_directories: BTreeMap::from([('C', "C:\\".to_string())]),
            attributes: BTreeMap::new(),
            volume_labels: BTreeMap::new(),
            record_locks: Vec::new(),
            modified: BTreeMap::new(), directory_entries: BTreeMap::new(), current_time: 0.0, clock: None, pending_writes: Default::default(),
            cursors: BTreeMap::new(), file_entries: BTreeMap::new(),
        }
    }
}


pub(crate) fn whole_second_date(serial: f64) -> f64 {
    let seconds = serial.abs().fract() * 86_400.0;
    let whole = |seconds| serial.trunc() + serial.signum() * seconds / 86_400.0;
    // Whole-second DATEs can land below their second after the fractional
    // part is separated from a distant day. Preserve an exact round trip
    // instead of adding a fixed epsilon that fails near the Date limits.
    if whole(seconds.round()) == serial { serial } else { whole(seconds.floor()) }
}

/// The last part of a path.
fn file_name(path: &str) -> &str {
    path.rsplit(['\\', '/']).next().unwrap_or(path)
}

/// Whether a file name answers a `Dir` or `Kill` pattern of `*` and `?`.
fn matches_pattern(name: &str, pattern: &str) -> bool {
    wildcards::matches(name,pattern)
}

fn channel_key(project: u64, number: i64) -> (u64, i64) {
    (if number >= 1000 { 0 } else { project }, number)
}

impl Files {
    /// Select the code project's numbered channels and Dir cursor.
    /// The owned namespace and native stream table stay in this same Files.
    pub fn select_project(&mut self, project: u64) {
        self.project_scope = project;
    }

    pub fn project_scope(&self) -> u64 { self.project_scope }

    /// A project joins the Application before acquiring its own file state.
    pub(crate) fn pristine_application_context(&self) -> bool {
        self.contents.is_empty() && self.open.is_empty() && self.listings.is_empty()
            && self.folders.is_empty() && self.attributes.is_empty()
            && self.volume_labels.is_empty() && self.record_locks.is_empty()
            && self.modified.is_empty() && self.directory_entries.is_empty() && self.pending_writes.is_empty()
            && self.file_entries.is_empty() && self.next_stream_number==1000 && self.current_drive=='C'
            && self.drive_directories.len()==1
            && self.drive_directories.get(&'C').is_some_and(|path|path=="C:\\")
    }




    /// Validate and normalize an absolute path supplied by a host snapshot.
    /// Relative mounts use the ordinary file-context API instead.
    pub fn normalize_snapshot_path(&self, path: &str) -> Result<String, FileError> {
        let bytes=path.as_bytes();
        let drive=bytes.len()>=3 && bytes[0].is_ascii_alphabetic() && bytes[1]==b':' && matches!(bytes[2],b'\\'|b'/');
        let unc=path.replace('/',"\\");
        let unc=unc.strip_prefix("\\\\").is_some_and(|tail| {
            let mut parts=tail.split('\\');
            matches!((parts.next(),parts.next()),(Some(server),Some(share)) if !server.is_empty() && !share.is_empty() && server!="." && server!=".." && share!="." && share!="..")
        });
        if !(drive||unc) || path.chars().enumerate().any(|(at,ch)| ch.is_control() || matches!(ch,'?'|'*'|'"'|'<'|'>'|'|') || ch==':' && !(drive && at==1)) {
            return Err(FileError::InvalidArgument);
        }
        Ok(self.absolute_path(path))
    }

    /// Files whose buffered sequential output is published when this environment closes.
    pub fn buffered_output_files(&self) -> Vec<String> {
        self.pending_writes.iter().filter(|(_, pending)| !pending.buffered.is_empty())
            .filter_map(|(key, _)| self.open.get(key))
            .filter_map(|handle| self.contents.get(&handle.path).map(|(path, _)| path.clone()))
            .collect::<std::collections::BTreeSet<_>>().into_iter().collect()
    }

    pub fn closed_snapshot(&self) -> FileSystemSnapshot {
        FileSystemSnapshot {
            folders:self.folders.iter().map(|path| if path.len()==2 && path.ends_with(':') {format!("{path}\\")} else {path.clone()}).collect(),
            files:self.contents.values().cloned().collect(),
            attributes:self.attributes.iter().map(|(path,value)| {
                let path=self.contents.get(path).map(|(path,_)|path.clone()).unwrap_or_else(||if path.len()==2 && path.ends_with(':') {format!("{path}\\")} else {path.clone()});
                (path,*value)
            }).collect(),
        }
    }

    /// Replace the closed namespace atomically. No VBA handles or enumeration
    /// cursors are restored; the process directory is a separate host fact.
    pub fn restore_closed_snapshot(&mut self, snapshot: &FileSystemSnapshot) -> Result<(), FileError> {
        let mut folders=std::collections::BTreeSet::new();
        let mut contents=BTreeMap::new();
        for path in &snapshot.folders {
            let path=self.normalize_snapshot_path(path)?;
            if !folders.insert(path.to_lowercase().trim_end_matches('\\').to_string()) {return Err(FileError::InvalidArgument);}
        }
        for (path,bytes) in &snapshot.files {
            let path=self.normalize_snapshot_path(path)?;
            let key=path.to_lowercase();
            if path.ends_with('\\') || folders.contains(&key) || contents.insert(key,(path,bytes.clone())).is_some() {return Err(FileError::InvalidArgument);}
        }
        let mut attributes=BTreeMap::new();
        for (path,value) in &snapshot.attributes {
            let path=self.normalize_snapshot_path(path)?;
            let key=path.to_lowercase().trim_end_matches('\\').to_string();
            let folder=folders.contains(&key);
            if !(folder || contents.contains_key(&key)) || *value<0 || value & !55!=0
                || (value & 16!=0)!=folder || attributes.insert(key,*value).is_some() {
                return Err(FileError::InvalidArgument);
            }
        }
        self.folders=folders;
        self.contents=contents;
        self.attributes=attributes;
        let current_time=self.current_time();
        self.modified=self.contents.keys().chain(self.folders.iter()).map(|key|(key.clone(),current_time)).collect();
        self.pending_writes.clear();
        self.file_entries.clear();
        self.directory_entries.clear();
        self.open.clear();
        self.cursors.clear();
        self.record_locks.clear();
        self.listings.clear();
        Ok(())
    }

    pub fn directory_state(&self) -> DirectoryState {
        DirectoryState {
            current_drive: self.current_drive,
            directories: ('A'..='Z').filter(|drive| self.drive_exists(*drive))
                .map(|drive| self.current_directory(Some(&drive.to_string())).expect("available drive")).collect(),
        }
    }

    /// Restore an explicitly supplied process environment atomically. Folder
    /// and file mounts remain separate; this does not create files or handles.
    pub fn restore_directory_state(&mut self, state: &DirectoryState) -> Result<(), FileError> {
        if !state.current_drive.is_ascii_alphabetic() { return Err(FileError::InvalidArgument); }
        let current=state.current_drive.to_ascii_uppercase();
        let mut directories=BTreeMap::new();
        for path in &state.directories {
            let bytes=path.as_bytes();
            if bytes.len()<3 || !bytes[0].is_ascii_alphabetic() || bytes[1]!=b':' || !matches!(bytes[2],b'\\'|b'/')
                || path.chars().enumerate().any(|(at,ch)| ch.is_control() || matches!(ch,'?'|'*'|'"'|'<'|'>'|'|') || ch==':' && at!=1) {
                return Err(FileError::InvalidArgument);
            }
            let drive=char::from(bytes[0]).to_ascii_uppercase();
            if directories.insert(drive,self.absolute_path(path)).is_some() { return Err(FileError::InvalidArgument); }
        }
        if !directories.contains_key(&current) { return Err(FileError::InvalidArgument); }
        self.current_drive=current;
        self.drive_directories=directories;
        Ok(())
    }

    fn drive_exists(&self, drive: char) -> bool {
        let prefix = format!("{}:", drive.to_ascii_lowercase());
        self.drive_directories.contains_key(&drive)
            || self.volume_labels.contains_key(&prefix)
            || self.folders.iter().any(|path| path.starts_with(&prefix))
            || self.contents.keys().any(|path| path.starts_with(&prefix))
    }

    /// Windows keeps one current directory for each drive. ChDir updates the
    /// addressed drive's directory; ChDrive chooses which drive is active.
    pub fn current_directory(&self, drive: Option<&str>) -> Result<String, FileError> {
        let drive = match drive.and_then(|value| value.chars().next()) {
            Some(value) if value.is_ascii_alphabetic() => value.to_ascii_uppercase(),
            Some(_) => return Err(FileError::InvalidArgument),
            None => self.current_drive,
        };
        if !self.drive_exists(drive) { return Err(FileError::DeviceUnavailable); }
        Ok(self.drive_directories.get(&drive).cloned().unwrap_or_else(|| format!("{drive}:\\")))
    }

    pub fn change_drive(&mut self, drive: &str) -> Result<(), FileError> {
        let Some(first) = drive.chars().next() else { return Ok(()); };
        if !first.is_ascii_alphabetic() { return Err(FileError::InvalidArgument); }
        let first = first.to_ascii_uppercase();
        if !self.drive_exists(first) { return Err(FileError::DeviceUnavailable); }
        self.current_drive = first;
        Ok(())
    }

    pub fn change_directory(&mut self, path: &str) -> Result<(), FileError> {
        if path.is_empty() { return Err(FileError::PathNotFound); }
        let path = self.absolute_path(path);
        if !self.folder_exists(&path) { return Err(FileError::PathNotFound); }
        self.directory_entries.remove(path.to_lowercase().trim_end_matches('\\'));
        if path.as_bytes().get(1) == Some(&b':') {
            let drive = char::from(path.as_bytes()[0]).to_ascii_uppercase();
            self.drive_directories.insert(drive, path);
        }
        Ok(())
    }

    /// Resolve against the active drive and its remembered directory, then
    /// collapse dot segments. Filename components follow Win32's handling of
    /// trailing spaces and periods; written download names remain separate.
    pub fn absolute_path(&self, path: &str) -> String {
        let path = path.replace('/', "\\");
        let absolute = if path.starts_with("\\\\") {
            path
        } else if path.as_bytes().get(1) == Some(&b':') && path.as_bytes()[0].is_ascii_alphabetic() {
            let drive = char::from(path.as_bytes()[0]).to_ascii_uppercase();
            let tail = &path[2..];
            if tail.starts_with('\\') { format!("{drive}:{tail}") }
            else {
                let base = self.drive_directories.get(&drive).cloned().unwrap_or_else(|| format!("{drive}:\\"));
                format!("{}\\{tail}", base.trim_end_matches('\\'))
            }
        } else if path.starts_with('\\') {
            format!("{}:{path}", self.current_drive)
        } else {
            let base = self.drive_directories.get(&self.current_drive).cloned().unwrap_or_else(|| format!("{}:\\", self.current_drive));
            format!("{}\\{path}", base.trim_end_matches('\\'))
        };
        let (prefix, tail, floor) = if let Some(tail) = absolute.strip_prefix("\\\\") {
            ("\\\\".to_string(), tail, 2)
        } else {
            (format!("{}\\", &absolute[..2]), &absolute[3..], 0)
        };
        let mut parts = Vec::new();
        for part in tail.split('\\') {
            match part {
                "" | "." => {},
                ".." => { if parts.len() > floor { parts.pop(); } },
                _ => { let part=part.trim_end_matches([' ', '.']); if !part.is_empty() { parts.push(part); } },
            }
        }
        format!("{prefix}{}", parts.join("\\"))
    }

    fn key(&self, path: &str) -> String {
        self.absolute_path(path.trim()).to_lowercase()
    }

    fn pattern_key(&self, path: &str) -> String {
        let path=path.replace('/',"\\");
        let start=path.rfind('\\').map_or(0,|at|at+1).max(if path.as_bytes().get(1)==Some(&b':') {2} else {0});
        let leaf=&path[start..];
        if leaf.contains(['*','?']) && leaf.ends_with('.') {
            let temporary=self.absolute_path(&format!("{}x",&path[..start]));
            let parent=temporary.rsplit_once('\\').map_or("",|(parent,_)|parent);
            format!("{parent}\\{leaf}").to_lowercase()
        } else {self.key(&path)}
    }

    fn volume_root(&self, path: &str) -> String {
        let path=self.absolute_path(path).to_lowercase();
        if let Some(tail)=path.strip_prefix("\\\\") {
            format!("\\\\{}",tail.split('\\').take(2).collect::<Vec<_>>().join("\\"))
        } else {path[..2].to_string()}
    }

    /// Host volume facts are separate from a closed namespace or cwd snapshot.
    pub fn volume_facts(&self) -> Vec<(String,String)> {
        self.volume_labels.iter().map(|(root,label)| {
            let root=if root.len()==2 && root.ends_with(':') {format!("{}\\",root.to_ascii_uppercase())} else {root.clone()};
            (root,label.clone())
        }).collect()
    }

    pub fn restore_volume_facts(&mut self, facts: &[(String,String)]) -> Result<(),FileError> {
        let mut volumes=BTreeMap::new();
        for (root,label) in facts {
            let path=self.normalize_snapshot_path(root)?;
            let root=self.volume_root(&path);
            if path.trim_end_matches('\\').to_lowercase()!=root || volumes.insert(root,label.clone()).is_some() {
                return Err(FileError::InvalidArgument);
            }
        }
        self.volume_labels=volumes;Ok(())
    }


    /// Supply the host's local OLE clock for newly created or closed files.
    pub fn set_current_time(&mut self, serial: f64) {
        if serial.is_finite() { self.current_time = whole_second_date(serial); self.clock = None; }
    }

    /// Select the same advancing host clock used by the interpreter.
    pub fn set_clock(&mut self, clock: Option<crate::clock::Clock>) {
        self.clock = clock;
    }

    fn current_time(&self) -> f64 {
        whole_second_date(self.clock.as_ref().map_or(self.current_time, crate::clock::Clock::now))
    }

    /// Local modification times are separate host facts, like volume labels.
    /// Keeping them separate preserves existing closed-snapshot callers.
    pub fn modified_facts(&self) -> Vec<(String, f64)> {
        self.modified.iter().map(|(key, serial)| {
            let path = self.contents.get(key).map(|(path, _)| path.clone())
                .unwrap_or_else(|| if key.len() == 2 && key.ends_with(':') { format!("{key}\\") } else { key.clone() });
            (path, *serial)
        }).collect()
    }

    /// Seed a mounted entry's real modification time without writing its bytes.
    pub fn set_modified_time(&mut self, path: &str, serial: f64) -> Result<(), FileError> {
        if !serial.is_finite() || !(serial > -657435.0 && serial < 2958466.0) { return Err(FileError::InvalidArgument); }
        let path = self.normalize_snapshot_path(path)?;
        let key = path.to_lowercase().trim_end_matches('\\').to_string();
        if !self.contents.contains_key(&key) && !self.folder_exists(&path) { return Err(FileError::InvalidArgument); }
        self.directory_entries.remove(&key);
        self.modified.insert(key, whole_second_date(serial));
        Ok(())
    }

    /// Validate the entire time snapshot before replacing any existing facts.
    pub fn restore_modified_facts(&mut self, facts: &[(String, f64)]) -> Result<(), FileError> {
        let mut next = self.clone();
        next.modified.clear();
        let paths: Vec<_> = next.contents.keys().chain(next.folders.iter()).cloned().collect();
        let current_time=next.current_time();
        for key in paths { next.modified.insert(key, current_time); }
        let mut seen = std::collections::BTreeSet::new();
        for (path, serial) in facts {
            let key = next.normalize_snapshot_path(path)?.to_lowercase().trim_end_matches('\\').to_string();
            if !seen.insert(key) { return Err(FileError::InvalidArgument); }
            next.set_modified_time(path, *serial)?;
        }
        self.modified = next.modified;
        self.directory_entries.clear();
        Ok(())
    }

    /// FileDateTime uses the local last-write time and native pathname errors.
    pub fn date_modified(&mut self, path: &str) -> Result<f64, FileError> {
        let path = path.split('\0').next().unwrap_or_default();
        if path.trim_end_matches(' ').is_empty() { return Err(FileError::NotFound); }
        if path.contains(['*', '?']) { return Err(FileError::BadNumber); }
        if path.chars().enumerate().any(|(at, ch)|
            ch.is_control() || matches!(ch, '"' | '<' | '>' | '|') || ch == ':' && at != 1) {
            return Err(FileError::InvalidArgument);
        }
        let key = self.absolute_path(path).to_lowercase().trim_end_matches('\\').to_string();
        let leaf=key.rsplit('\\').next().unwrap_or_default().split('.').next().unwrap_or_default();
        if matches!(leaf,"con"|"prn"|"aux"|"nul") || (leaf.starts_with("com") || leaf.starts_with("lpt"))
            && leaf.len()==4 && matches!(leaf.as_bytes()[3],b'1'..=b'9') { return Err(FileError::InvalidArgument); }
        if path.ends_with(['\\', '/']) {
            return Err(if self.contents.contains_key(&key) { FileError::InvalidArgument } else { FileError::NotFound });
        }
        self.publish_parent_entry(&key);
        if !self.contents.contains_key(&key) && !self.folder_exists(&key) {
            let parent = key.rsplit_once('\\').map_or("", |(parent, _)| parent);
            return Err(if self.folder_exists(parent) { FileError::NotFound } else { FileError::PathNotFound });
        }
        Ok(self.file_entries.get(&key).map(|entry| entry.modified)
            .or_else(|| self.directory_entries.get(&key).or_else(|| self.modified.get(&key)).copied())
            .unwrap_or_else(|| self.current_time()))
    }

    /// Physical directory dates and the entry returned by FileDateTime can
    /// differ until a metadata query refreshes that entry.
    pub fn directory_entry_facts(&self) -> Vec<(String, f64)> {
        self.directory_entries.iter().map(|(path, value)| {
            let path = if path.len() == 2 && path.ends_with(':') { format!("{path}\\") } else { path.clone() };
            (path, *value)
        }).collect()
    }

    pub fn restore_directory_entry_facts(&mut self, facts: &[(String, f64)]) -> Result<(), FileError> {
        let mut entries = BTreeMap::new();
        for (path, value) in facts {
            let path = self.normalize_snapshot_path(path)?.to_lowercase().trim_end_matches('\\').to_string();
            if !value.is_finite() || !(*value > -657435.0 && *value < 2958466.0)
                || !self.folder_exists(&path) || self.contents.contains_key(&path)
                || entries.insert(path, whole_second_date(*value)).is_some() {
                return Err(FileError::InvalidArgument);
            }
        }
        self.directory_entries = entries;
        Ok(())
    }

    fn publish_parent_entry(&mut self, path: &str) {
        if let Some((parent, _)) = path.rsplit_once('\\') { self.directory_entries.remove(parent); }
    }

    fn parent_changed(&mut self, path: &str, publish: bool) {
        if let Some((parent, _)) = path.rsplit_once('\\') {
            if self.folder_exists(parent) {
                let current_time=self.current_time();
                let previous = self.modified.insert(parent.into(), current_time).unwrap_or(current_time);
                if publish { self.directory_entries.remove(parent); }
                else if previous != current_time { self.directory_entries.entry(parent.into()).or_insert(previous); }
            }
        }
    }

    /// GetAttr refreshes a directory's entry without changing its physical time.
    pub fn query_attributes(&mut self, path: &str) -> Result<i64, FileError> {
        let value = self.attributes(path)?;
        if value & 16 != 0 { self.directory_entries.remove(self.key(path).trim_end_matches('\\')); }
        else { self.refresh_file_entry(path)?; }
        Ok(value)
    }

    pub fn query_length(&mut self, path: &str) -> Result<usize, FileError> {
        let key = self.key(path);
        self.publish_parent_entry(&key);
        let length = self.length(path)?;
        Ok(self.file_entries.get(&key).map_or(length, |entry| entry.length))
    }

    /// A handle-based metadata query refreshes the entry used by FileLen.
    pub fn fresh_length(&mut self, path: &str) -> Result<usize, FileError> {
        self.publish_parent_entry(&self.key(path));
        self.refresh_file_entry(path)
    }

    fn refresh_file_entry(&mut self, path: &str) -> Result<usize, FileError> {
        let key = self.key(path);
        let length = self.length(path)?;
        if self.open.values().any(|handle| handle.path == key) {
            let modified = self.modified.get(&key).copied().unwrap_or_else(|| self.current_time());
            let entry = self.file_entries.entry(key).or_insert(FileEntry { length, modified });
            if entry.length != length { *entry = FileEntry { length, modified }; }
        }
        Ok(length)
    }

    fn cache_file_entry(&mut self, path: &str, truncate: bool) {
        if truncate { self.file_entries.remove(path); }
        let length = self.contents.get(path).map_or(0, |file| file.1.len());
        let modified = self.modified.get(path).copied().unwrap_or_else(|| self.current_time());
        self.file_entries.entry(path.into()).or_insert(FileEntry { length, modified });
    }

    pub fn fresh_modified_time(&mut self, path: &str) -> Result<f64, FileError> {
        self.query_attributes(path)?;
        self.publish_parent_entry(&self.key(path));
        let key = self.key(path);
        Ok(self.file_entries.get(key.trim_end_matches('\\')).map(|entry| entry.modified)
            .or_else(|| self.modified.get(key.trim_end_matches('\\')).copied())
            .unwrap_or_else(|| self.current_time()))
    }

    fn publish_write_time(&mut self, path: &str, written: f64) {
        for (key, pending) in &mut self.pending_writes {
            if self.open.get(key).is_some_and(|handle| handle.path == path) {
                pending.written = written;
            }
        }
    }

    fn mark_handle_written(&mut self, number: i64, path: &str, count: usize) {
        self.mark_written(path);
        if count != 0 {
            let written = self.current_time();
            self.pending_writes.insert(channel_key(self.project_scope, number),
                PendingWrite { written, offset: 0, buffered: Vec::new() });
            self.modified.insert(path.into(), written);
            self.publish_write_time(path, written);
        }
    }

    fn write_physical(&mut self, path: &str, offset: usize, bytes: &[u8]) {
        let file = self.contents.get_mut(path).expect("open file contents");
        let end = offset + bytes.len();
        if file.1.len() < end { file.1.resize(end, 0); }
        file.1[offset..end].copy_from_slice(bytes);
    }

    fn buffer_sequential(&mut self, key: (u64, i64), path: &str, offset: usize, bytes: &[u8]) {
        self.mark_written(path);
        if bytes.is_empty() { return; }
        let written = self.current_time();
        let mut pending = self.pending_writes.remove(&key)
            .unwrap_or(PendingWrite { written, offset, buffered: Vec::new() });
        if pending.buffered.is_empty() { pending.offset = offset; }
        debug_assert_eq!(pending.offset + pending.buffered.len(), offset);
        pending.buffered.extend_from_slice(bytes);
        let total = pending.buffered.len();
        let keep = if bytes.len() >= 1024 { 0 } else if total > 512 { total % 512 } else { total };
        let count = total - keep;
        if count != 0 {
            self.write_physical(path, pending.offset, &pending.buffered[..count]);
            pending.buffered.drain(..count);
            pending.offset += count;
            pending.written = written;
            self.modified.insert(path.into(), written);
            self.publish_write_time(path, written);
        }
        self.pending_writes.insert(key, pending);
    }

    fn flush_pending_key(&mut self, key: (u64, i64)) {
        if let Some(mut pending) = self.pending_writes.remove(&key) {
            if !pending.buffered.is_empty() {
                let path = self.open.get(&key).expect("open buffered output").path.clone();
                let written = self.current_time();
                self.write_physical(&path, pending.offset, &pending.buffered);
                pending.offset += pending.buffered.len();
                pending.buffered.clear();
                pending.written = written;
                self.modified.insert(path.clone(), written);
                self.publish_write_time(&path, written);
            }
            self.pending_writes.insert(key, pending);
        }
    }

    fn close_key(&mut self, key: (u64, i64)) {
        self.flush_pending_key(key);
        if let Some(handle) = self.open.remove(&key) {
            let written = if let Some(pending) = self.pending_writes.remove(&key) {
                self.publish_write_time(&handle.path, pending.written);
                self.modified.insert(handle.path.clone(), pending.written);
                true
            } else { false };
            let resized = self.file_entries.get(&handle.path).is_some_and(|entry|
                self.contents.get(&handle.path).is_some_and(|file| file.1.len() != entry.length));
            if written { self.file_entries.remove(&handle.path); }
            if written || resized { let _ = self.refresh_file_entry(&handle.path); }
            if !self.open.values().any(|other| other.path == handle.path) {
                self.file_entries.remove(&handle.path);
            }
        }
        self.cursors.remove(&key);
    }

    pub fn attributes(&self, path: &str) -> Result<i64, FileError> {
        let key=self.key(path).trim_end_matches('\\').to_string();
        let default=if self.contents.contains_key(&key) {32}
            else if self.folder_exists(&format!("{key}\\")) {16}
            else {return Err(FileError::NotFound);};
        Ok(self.attributes.get(&key).copied().unwrap_or(default))
    }

    /// SetAttr replaces the writable bits; the directory bit is structural.
    pub fn set_attributes(&mut self, path: &str, value: i64) -> Result<(), FileError> {
        if value<0 || value & !39!=0 {return Err(FileError::InvalidArgument);}
        let previous=self.attributes(path)?;
        let key=self.key(path).trim_end_matches('\\').to_string();
        if self.open.values().any(|handle|handle.path==key && handle.mode!=OpenMode::Input) {return Err(FileError::AlreadyOpen);}
        self.attributes.insert(key,value | (previous & 16));
        Ok(())
    }

    /// FSO masks unsupported bits and permits updates through open handles.
    pub fn set_fso_attributes(&mut self, path: &str, value: i64) -> Result<(), FileError> {
        let previous=self.attributes(path)?;
        let key=self.key(path).trim_end_matches('\\').to_string();
        self.attributes.insert(key,(value & 39) | (previous & 16));
        Ok(())
    }

    fn is_read_only_key(&self, key: &str) -> bool {
        self.attributes.get(key).copied().unwrap_or(0) & 1!=0
    }

    fn mark_written(&mut self, key: &str) {
        let attributes=self.attributes.get(key).copied().unwrap_or(32) | 32;
        self.attributes.insert(key.to_string(),attributes);
    }

    /// Publish complete binary contents to a closed file in an existing folder.
    /// Validation happens before replacement, so a refused write is atomic.
    pub fn replace_bytes(&mut self, path: &str, bytes: Vec<u8>) -> Result<(), FileError> {
        let name = self.key(path);
        if name.is_empty() || name.ends_with('\\') {
            return Err(FileError::BadNumber);
        }
        if self.open.values().any(|handle| handle.path == name) {
            return Err(FileError::AlreadyOpen);
        }
        if self.is_read_only_key(&name) {return Err(FileError::PermissionDenied);}
        if let Some((parent, _)) = name.rsplit_once('\\') {
            if !parent.is_empty() && !self.folder_exists(parent) {
                return Err(FileError::PathNotFound);
            }
        }
        if !self.contents.contains_key(&name) { self.parent_changed(&name, false); }
        self.mark_written(&name);
        let current_time=self.current_time(); self.modified.insert(name.clone(),current_time);
        self.contents.insert(name, (self.absolute_path(path.trim()), bytes));
        Ok(())
    }

    /// Read the bytes actually held by the VBA filesystem, including edits
    /// made through an open binary handle after a host produced the file.
    pub fn bytes(&self, path: &str) -> Result<&[u8], FileError> {
        self.contents.get(&self.key(path)).map(|(_, bytes)| bytes.as_slice()).ok_or(FileError::NotFound)
    }

    /// Absolute paths and complete contents of files created or modified since a snapshot.
    /// Renaming a newly produced file returns its final name; deleted files
    /// have no downloadable contents.
    pub fn changed_files(&self, before: &Self) -> Vec<(String, Vec<u8>)> {
        self.contents.iter().filter_map(|(key, (path, bytes))| {
            (before.contents.get(key).is_none_or(|(_, old)| old != bytes))
                .then(|| (path.clone(), bytes.clone()))
        }).collect()
    }

    pub fn deleted_files(&self, before: &Self) -> Vec<String> {
        before.contents.iter().filter(|(key, _)| !self.contents.contains_key(*key))
            .map(|(_, (path, _))| path.clone()).collect()
    }

    /// `FreeFile`: the lowest number not in use, from 1, or from 256 for
    /// `FreeFile(1)`.
    pub fn free_number(&self, upper_range: bool) -> i64 {
        let start = if upper_range { 256 } else { 1 };
        (start..start + 255).find(|number| !self.open.contains_key(&channel_key(self.project_scope, *number))).unwrap_or(0)
    }

    pub fn open(&mut self, path: &str, mode: OpenMode, number: i64) -> Result<(), FileError> {
        self.open_with_access(path,mode,number,None)
    }

    pub fn open_with_access(&mut self, path: &str, mode: OpenMode, number: i64, access: Option<crate::ast::FileAccess>) -> Result<(), FileError> {
        self.open_with_options(path,mode,number,access,None)
    }

    pub fn open_with_options(&mut self, path: &str, mode: OpenMode, number: i64, access: Option<crate::ast::FileAccess>, lock: Option<crate::ast::FileLock>) -> Result<(), FileError> {
        if !(1..=511).contains(&number) {
            return Err(FileError::BadNumber);
        }
        if self.open.contains_key(&channel_key(self.project_scope, number)) {
            return Err(FileError::AlreadyOpen);
        }
        let name = self.key(path);
        if path.trim().is_empty() {
            return Err(FileError::BadNumber);
        }
        if self.open.iter().any(|(held_number,held)|held_number.1<=511 && held.path==name &&
            (matches!(mode,OpenMode::Output|OpenMode::Append) || matches!(held.mode,OpenMode::Output|OpenMode::Append))) {
            return Err(FileError::AlreadyOpen);
        }
        let newly_created=!self.contents.contains_key(&name);
        let read_only=self.is_read_only_key(&name);
        let (can_read,can_write)=match access {
            Some(crate::ast::FileAccess::Read)=>(true,false),
            Some(crate::ast::FileAccess::Write)=>(false,true),
            Some(crate::ast::FileAccess::ReadWrite)=>(true,true),
            None=>match mode {
                OpenMode::Input=>(true,false),
                OpenMode::Output|OpenMode::Append=>(false,true),
                // VBA retries default binary/random access as read-only.
                OpenMode::Binary|OpenMode::Random(_)=>(true,!read_only),
            },
        };
        if read_only && can_write {return Err(FileError::PathFileAccess);}
        let (share_read,share_write)=match lock {
            None|Some(crate::ast::FileLock::Shared)=>(true,true),
            Some(crate::ast::FileLock::Read)=>(false,true),
            Some(crate::ast::FileLock::Write)=>(true,false),
            Some(crate::ast::FileLock::ReadWrite)=>(false,false),
        };
        if self.sharing_conflict(&name,can_read,can_write,share_read,share_write) {
            return Err(FileError::PermissionDenied);
        }
        match mode {
            OpenMode::Input if !self.contents.contains_key(&name) => return Err(FileError::NotFound),
            OpenMode::Output => {
                self.mark_written(&name);
                self.contents.insert(name.clone(), (self.absolute_path(path.trim()), Vec::new()));
            }
            OpenMode::Append | OpenMode::Binary | OpenMode::Random(_) => {
                let path=self.absolute_path(path.trim());
                self.contents.entry(name.clone()).or_insert_with(|| (path, Vec::new()));
            }
            OpenMode::Input => {}
        }
        if newly_created { self.parent_changed(&name, false); }
        if newly_created || mode==OpenMode::Output { let current_time=self.current_time(); self.modified.insert(name.clone(),current_time); }
        let column = match mode {
            OpenMode::Append => {
                let bytes = &self.contents[&name].1;
                let last_line = bytes.iter().rposition(|byte| *byte == b'\n').map_or(0, |at| at + 1);
                bytes.len() - last_line
            }
            _ => 0,
        };
        self.cache_file_entry(&name, mode == OpenMode::Output);
        self.open.insert(channel_key(self.project_scope, number), Handle { path: name, mode, position: 0, column,can_read,can_write,share_read,share_write });
        self.cursors.insert(channel_key(self.project_scope, number), FileCursor::default());
        Ok(())
    }

    /// A number for a TextStream, from 1000 on, never one FreeFile hands out.
    pub fn open_stream(&mut self, path: &str, mode: OpenMode) -> Result<i64, FileError> {
        let number = self.next_stream_number;
        self.next_stream_number = number.checked_add(1).ok_or(FileError::BadNumber)?;
        let name = self.key(path);
        if mode!=OpenMode::Input && self.is_read_only_key(&name) {return Err(FileError::PermissionDenied);}
        // Native FSO streams permit concurrent readers. ForWriting and
        // ForAppending deny writers, including writers that opened first.
        let newly_created=!self.contents.contains_key(&name);
        let can_read=mode==OpenMode::Input;let can_write=!can_read;
        let share_read=true;let share_write=can_read;
        if self.sharing_conflict(&name,can_read,can_write,share_read,share_write) {return Err(FileError::PermissionDenied);}
        match mode {
            OpenMode::Input if !self.contents.contains_key(&name) => return Err(FileError::NotFound),
            OpenMode::Output => {
                self.mark_written(&name);
                self.contents.insert(name.clone(), (self.absolute_path(path.trim()), Vec::new()));
            }
            _ => {
                let path=self.absolute_path(path.trim());
                self.contents.entry(name.clone()).or_insert_with(|| (path, Vec::new()));
            }
        }
        if newly_created { self.parent_changed(&name, false); }
        if newly_created || mode==OpenMode::Output { let current_time=self.current_time(); self.modified.insert(name.clone(),current_time); }
        let column = if mode == OpenMode::Append {
            let bytes = &self.contents[&name].1;
            let start = bytes.iter().rposition(|byte| matches!(*byte, b'\r' | b'\n')).map_or(0, |at| at + 1);
            String::from_utf8_lossy(&bytes[start..]).encode_utf16().count()
        } else { 0 };
        self.cache_file_entry(&name, mode == OpenMode::Output);
        self.open.insert(channel_key(self.project_scope, number), Handle { path: name, mode, position: 0, column,can_read,can_write,share_read,share_write });
        Ok(number)
    }

    fn sharing_conflict(&self, path: &str, can_read: bool, can_write: bool, share_read: bool, share_write: bool) -> bool {
        self.open.values().any(|held|held.path==path &&
            (can_read && !held.share_read || can_write && !held.share_write || held.can_read && !share_read || held.can_write && !share_write))
    }

    pub fn exists(&self, path: &str) -> bool {
        self.contents.contains_key(&self.key(path))
    }

    /// Whether the folder was made, or any kept file lies in it.
    pub fn folder_exists(&self, path: &str) -> bool {
        let folder = self.key(path);
        let folder = folder.trim_end_matches('\\');
        if folder.len()==2 && folder.as_bytes()[1]==b':' {
            return self.drive_exists(char::from(folder.as_bytes()[0]).to_ascii_uppercase());
        }
        self.folders.contains(folder)
            || self.folders.iter().any(|held| held.starts_with(&format!("{folder}\\")))
            || self.contents.keys().any(|held| held.rsplit_once('\\').is_some_and(|(parent, _)| parent == folder || parent.starts_with(&format!("{folder}\\"))))
    }

    /// Import a host directory fact together with its existing ancestors. This
    /// seeds the environment rather than performing VBA's one-level MkDir.
    pub fn mount_folder(&mut self, path: &str) -> Result<(), FileError> {
        if path.trim().is_empty() {return Err(FileError::InvalidArgument);}
        let path=self.normalize_snapshot_path(&self.absolute_path(path.trim()))?;
        let mut key=path.trim_end_matches('\\').to_lowercase();
        let floor=if let Some(tail)=key.strip_prefix("\\\\") {
            2+tail.split('\\').take(2).map(str::len).sum::<usize>()+1
        } else {2};
        let mut folders=Vec::new();
        loop {
            if self.contents.contains_key(&key) {return Err(FileError::AlreadyExists);}
            folders.push(key.clone());
            if key.len()<=floor {break;}
            let parent=key.rfind('\\').ok_or(FileError::InvalidArgument)?;
            key.truncate(parent);
        }
        let current_time=self.current_time();
        for folder in &folders { self.modified.entry(folder.clone()).or_insert(current_time); }
        self.folders.extend(folders);
        Ok(())
    }

    /// Create one directory, with an existing parent. The caller distinguishes
    /// VBA MkDir's error75 from FSO CreateFolder's error58 for an existing name.
    pub fn make_folder(&mut self, path: &str) -> Result<(), FileError> {
        if path.is_empty() {return Err(FileError::PathNotFound);}
        let folder = self.key(path).trim_end_matches('\\').to_string();
        if self.folder_exists(&format!("{folder}\\")) || self.contents.contains_key(&folder) {
            return Err(FileError::AlreadyExists);
        }
        let parent=folder.rsplit_once('\\').map(|(parent,_)|parent).ok_or(FileError::PathNotFound)?;
        if !self.folder_exists(&format!("{parent}\\")) {return Err(FileError::PathNotFound);}
        self.parent_changed(&folder, false);
        let current_time=self.current_time(); self.modified.insert(folder.clone(),current_time);
        self.folders.insert(folder);
        Ok(())
    }

    /// RmDir refuses a nonempty or current directory without changing it. FSO
    /// deletion visits entries in name order and retains completed work when a
    /// later entry is locked; a caught error must expose those partial effects.
    pub fn remove_folder(&mut self, path: &str, with_contents: bool) -> Result<(), FileError> {
        if with_contents {return self.delete_folder_with_force(path,false);}
        if path.is_empty() {return Err(FileError::PathNotFound);}
        let folder = self.key(path).trim_end_matches('\\').to_string();
        if !self.folder_exists(&format!("{folder}\\")) {
            return Err(FileError::PathNotFound);
        }
        if self.is_read_only_key(&folder) {return Err(FileError::PathFileAccess);}
        let inside = format!("{folder}\\");
        let cwd=self.key(&self.current_directory(None)?).trim_end_matches('\\').to_string();
        if folder==cwd || cwd.starts_with(&inside) || folder.len()==2 && folder.ends_with(':')
            || self.contents.keys().any(|held|held.starts_with(&inside))
            || self.folders.iter().any(|held|held.starts_with(&inside)) {
            return Err(FileError::PathFileAccess);
        }
        self.folders.remove(&folder);
        self.attributes.remove(&folder);
        self.modified.remove(&folder);
        self.directory_entries.remove(&folder);
        self.parent_changed(&folder, false);
        Ok(())
    }

    fn child_folders(&self, parent: &str) -> std::collections::BTreeSet<String> {
        let prefix=format!("{parent}\\");
        self.folders.iter().map(String::as_str)
            .chain(self.contents.keys().filter_map(|path|path.rsplit_once('\\').map(|(parent,_)|parent)))
            .filter_map(|path|path.strip_prefix(&prefix).and_then(|tail|tail.split('\\').next())
                .filter(|name|!name.is_empty()).map(|name|format!("{prefix}{name}"))).collect()
    }

    fn delete_directory_tree(&mut self, folder: &str, force: bool) -> Result<(), FileError> {
        let mut entries:BTreeMap<String,bool>=self.child_folders(folder).into_iter().map(|path|(path,true)).collect();
        for path in self.contents.keys().filter(|path|path.rsplit_once('\\').is_some_and(|(parent,_)|parent==folder)) {
            entries.insert(path.clone(),false);
        }
        for (path,directory) in entries {
            if directory {self.delete_directory_tree(&path,force)?;}
            else {
                if self.open.values().any(|handle|handle.path==path) || !force && self.is_read_only_key(&path) {return Err(FileError::PermissionDenied);}
                self.contents.remove(&path);
                self.attributes.remove(&path);
                self.modified.remove(&path);
                self.parent_changed(&path, true);
            }
        }
        let cwd=self.key(&self.current_directory(None)?).trim_end_matches('\\').to_string();
        if folder==cwd || cwd.starts_with(&format!("{folder}\\")) {return Err(FileError::PermissionDenied);}
        if !force && self.is_read_only_key(folder) {return Err(FileError::PermissionDenied);}
        self.folders.remove(folder);
        self.attributes.remove(folder);
        self.modified.remove(folder);
        self.directory_entries.remove(folder);
        self.parent_changed(folder, true);
        Ok(())
    }

    pub fn delete_folder_with_force(&mut self, path: &str, force: bool) -> Result<(), FileError> {
        if path.is_empty() {return Err(FileError::InvalidArgument);}
        let supplied=path.replace('/',"\\");
        let leaf=supplied.trim_end_matches('\\').rsplit('\\').next().unwrap_or_default();
        if matches!(leaf,"."|"..") {return Err(FileError::PathNotFound);}
        if supplied.ends_with('\\') {
            let absolute=supplied.starts_with("\\\\") || supplied.as_bytes().get(1)==Some(&b':') && supplied.as_bytes().get(2)==Some(&b'\\');
            return Err(if absolute {FileError::PathNotFound} else {FileError::InvalidArgument});
        }
        let pattern=self.key(path);
        let (parent,leaf)=pattern.rsplit_once('\\').ok_or(FileError::PathNotFound)?;
        let targets:Vec<_>=self.child_folders(parent).into_iter().filter(|folder|matches_pattern(file_name(folder),leaf)).collect();
        if targets.is_empty() {
            // FSO's folder matcher accepts a matching ordinary file but does
            // not remove it. A nonexistent name/pattern still raises76.
            return if self.contents.keys().any(|file|file.rsplit_once('\\').is_some_and(|(held,name)|held==parent && matches_pattern(name,leaf))) {
                Ok(())
            } else {Err(FileError::PathNotFound)};
        }
        for folder in targets {self.delete_directory_tree(&folder,force)?;}
        Ok(())
    }

    /// The resolved absolute paths of the files directly in a folder, in order.
    pub fn files_in(&self, path: &str) -> Vec<String> {
        let folder = self.key(path).trim_end_matches('\\').to_string();
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

    /// Close internal handles, including streams, or every handle for an empty list.
    pub fn close(&mut self, numbers: &[i64]) {
        if numbers.is_empty() {
            for key in self.open.keys().copied().collect::<Vec<_>>() { self.close_key(key); }
            self.cursors.clear();
            self.record_locks.clear();
        } else {
            for number in numbers {
                self.close_key(channel_key(self.project_scope,*number));
            }
            self.record_locks.retain(|held|!numbers.iter().any(|number| held.owner == channel_key(self.project_scope, *number)));
        }
    }

    /// End abandons this project's Dir enumeration and numbered channels.
    /// Native TextStream handles and directory facts survive independently.
    pub(crate) fn reset_project(&mut self) {
        self.close_vba(&[]);
        self.listings.remove(&self.project_scope);
    }

    /// VBA Close and Reset affect numbered Open channels, not FSO streams.
    pub(crate) fn close_vba(&mut self, numbers: &[i64]) {
        if numbers.is_empty() {
            let project = self.project_scope;
            let keys:Vec<_>=self.open.keys().filter(|(owner,number)|*owner==project && (1..=511).contains(number)).copied().collect();
            for key in keys { self.close_key(key); }
            self.record_locks.retain(|held| held.owner.0 != project || !(1..=511).contains(&held.owner.1));
        } else {
            self.close(numbers);
        }
    }

    pub fn handle(&self, number: i64) -> Result<&Handle, FileError> {
        self.open.get(&channel_key(self.project_scope, number)).ok_or(FileError::BadNumber)
    }

    pub(crate) fn cursor_snapshot(&self, number: i64) -> Option<FileCursor> {
        self.cursors.get(&channel_key(self.project_scope, number)).cloned()
    }

    pub(crate) fn restore_byte_position(&mut self, number: i64, path: &str, position: usize, cursor: Option<FileCursor>) {
        if let Some(handle) = self.open.get_mut(&channel_key(self.project_scope, number)).filter(|handle| handle.path == path) {
            handle.position = position;
            if let Some(cursor) = cursor { self.cursors.insert(channel_key(self.project_scope, number), cursor); }
        }
    }

    pub fn lock_records(&mut self, number: i64, start: Option<i64>, end: Option<i64>, unlock: bool) -> Result<(),FileError> {
        let handle=self.handle(number)?.clone();
        let (start,end)=if start.is_none() && end.is_none() {(0,u64::MAX)} else {
            let start=start.unwrap_or(1);let end=end.unwrap_or(start);
            if start<1 || end<start {return Err(FileError::BadRecordNumber);}
            let unit=match handle.mode {OpenMode::Random(length)=>length.max(1) as u64,_=>1};
            ((start as u64-1).checked_mul(unit).ok_or(FileError::BadRecordNumber)?,(end as u64).checked_mul(unit).ok_or(FileError::BadRecordNumber)?)
        };
        if unlock {
            let index=self.record_locks.iter().position(|held|held.owner==channel_key(self.project_scope, number) && held.path==handle.path && held.start==start && held.end==end).ok_or(FileError::PermissionDenied)?;
            self.record_locks.remove(index);
        } else {
            if self.record_locks.iter().any(|held|held.path==handle.path && held.start<end && start<held.end) {return Err(FileError::PermissionDenied);}
            self.record_locks.push(RecordLock {owner:channel_key(self.project_scope, number),path:handle.path.clone(),start,end});
        }
        Ok(())
    }

    fn check_record_access(&self, number: i64, handle: &Handle, count: usize) -> Result<(),FileError> {
        let count=match handle.mode {OpenMode::Random(length)=>count.max(length),_=>count};
        let start=self.cursors.get(&channel_key(self.project_scope, number)).map_or(handle.position, |cursor| cursor.physical) as u64;
        let end=start.saturating_add(count as u64);
        if count!=0 && self.record_locks.iter().any(|held|held.path==handle.path && held.owner!=channel_key(self.project_scope, number) && held.start<end && start<held.end) {
            Err(FileError::PermissionDenied)
        } else {Ok(())}
    }

    /// Move the raw cursor of a native TextStream after decoding characters.
    pub(crate) fn advance_stream(&mut self, number: i64, count: usize) -> Result<(), FileError> {
        let handle = self.open.get_mut(&channel_key(self.project_scope, number)).ok_or(FileError::BadNumber)?;
        handle.position = handle.position.checked_add(count).ok_or(FileError::BadRecordNumber)?;
        Ok(())
    }

    /// Append already encoded TextStream data without a UTF-8 conversion.
    pub(crate) fn write_stream(&mut self, number: i64, bytes: &[u8]) -> Result<(), FileError> {
        let handle = self.open.get_mut(&channel_key(self.project_scope, number)).ok_or(FileError::BadNumber)?;
        if handle.mode == OpenMode::Input { return Err(FileError::BadMode); }
        if !handle.can_write { return Err(FileError::PathFileAccess); }
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        file.1.extend_from_slice(bytes);
        let path = handle.path.clone();
        self.mark_handle_written(number, &path, bytes.len());
        Ok(())
    }

    /// Write text to a file open for output, keeping track of its column.
    pub fn write(&mut self, number: i64, text: &str) -> Result<(), FileError> {
        if (1..=511).contains(&number) && matches!(self.handle(number)?.mode, OpenMode::Output | OpenMode::Append) {
            return self.write_sequential(number, text, text.as_bytes());
        }
        let handle = self.open.get_mut(&channel_key(self.project_scope, number)).ok_or(FileError::BadNumber)?;
        if handle.mode == OpenMode::Input {
            return Err(FileError::BadNumber);
        }
        if !handle.can_write {return Err(FileError::PathFileAccess);}
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        file.1.extend_from_slice(text.as_bytes());
        match text.rfind('\n') {
            Some(at) => handle.column = text[at + 1..].chars().count(),
            None => handle.column += text.chars().count(),
        }
        let path=handle.path.clone();
        self.mark_handle_written(number,&path,text.len());
        Ok(())
    }

    pub(crate) fn write_sequential(&mut self, number: i64, text: &str, bytes: &[u8]) -> Result<(), FileError> {
        let handle = self.handle(number)?;
        if !(1..=511).contains(&number) { return Err(FileError::BadNumber); }
        if !matches!(handle.mode, OpenMode::Output | OpenMode::Append) { return Err(FileError::BadMode); }
        if !handle.can_write { return Err(FileError::PathFileAccess); }
        let key = channel_key(self.project_scope, number);
        let length = self.length_of_open(number)?;
        let handle = self.open.get_mut(&key).expect("checked handle");
        let cursor = self.cursors.get_mut(&key).ok_or(FileError::BadNumber)?;
        let start = if handle.mode == OpenMode::Append { length } else { cursor.physical };
        let end = start.checked_add(bytes.len()).ok_or(FileError::BadRecordNumber)?;
        let logical = handle.position.checked_add(bytes.len()).ok_or(FileError::BadRecordNumber)?;
        handle.position = logical;
        cursor.physical = end;
        match text.rfind('\n') {
            Some(at) => handle.column = text[at + 1..].chars().count(),
            None => handle.column += text.chars().count(),
        }
        let path = handle.path.clone();
        self.buffer_sequential(key, &path, start, bytes);
        Ok(())
    }

    pub(crate) fn vba_eof(&mut self, number: i64) -> Result<bool, FileError> {
        if !(1..=511).contains(&number) { return Err(FileError::BadNumber); }
        let mode = self.handle(number)?.mode;
        let key = channel_key(self.project_scope, number);
        match mode {
            OpenMode::Output | OpenMode::Append => Ok(true),
            OpenMode::Binary | OpenMode::Random(_) => self.cursors.get(&key)
                .map(|cursor| cursor.read_eof).ok_or(FileError::BadNumber),
            OpenMode::Input => {
                let handle = self.open.get_mut(&key).expect("checked handle");
                let cursor = self.cursors.get_mut(&key).ok_or(FileError::BadNumber)?;
                if cursor.read_eof { return Ok(true); }
                let bytes = self.contents.get(&handle.path).map_or(&[][..], |file| file.1.as_slice());
                cursor.physical = handle.position;
                while bytes.get(cursor.physical) == Some(&0) { cursor.physical += 1; }
                handle.position = cursor.physical;
                if bytes.get(cursor.physical) == Some(&26) {
                    let at = cursor.physical;
                    cursor.physical += 1;
                    handle.position = cursor.text_end_position(at, bytes.len());
                    cursor.read_eof = true;
                    Ok(true)
                } else { Ok(cursor.physical >= bytes.len()) }
            }
        }
    }

    pub fn column(&self, number: i64) -> Result<usize, FileError> {
        Ok(self.handle(number)?.column)
    }

    fn flush_pending_write(&mut self, number: i64) {
        self.flush_pending_key(channel_key(self.project_scope, number));
    }

    /// VBA's LOF also flushes buffered sequential output.
    pub fn query_length_of_open(&mut self, number: i64) -> Result<usize, FileError> {
        let length = self.length_of_open(number)?;
        self.flush_pending_write(number);
        Ok(length)
    }

    /// Querying Seek flushes output without moving the next position.
    pub fn query_next_position(&mut self, number: i64) -> Result<i64, FileError> {
        let position = self.next_position(number)?;
        self.flush_pending_write(number);
        Ok(position)
    }

    /// `LOF`: the file's length in bytes.
    pub fn length_of_open(&self, number: i64) -> Result<usize, FileError> {
        let handle = self.handle(number)?;
        let length = self.contents.get(&handle.path).map_or(0, |file| file.1.len());
        Ok(self.pending_writes.get(&channel_key(self.project_scope, number))
            .filter(|pending| !pending.buffered.is_empty())
            .map_or(length, |pending| length.max(pending.offset + pending.buffered.len())))
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
            OpenMode::Input => handle.position.div_ceil(128).max(1) as i64,
            OpenMode::Output | OpenMode::Append => (handle.position / 128) as i64,
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
        let handle = self.open.get_mut(&channel_key(self.project_scope, number)).ok_or(FileError::BadNumber)?;
        if position < 1 {
            return Err(FileError::BadNumber);
        }
        handle.position = match handle.mode {
            OpenMode::Random(length) => (position as usize - 1) * length,
            _ => position as usize - 1,
        };
        if let Some(cursor) = self.cursors.get_mut(&channel_key(self.project_scope, number)) {
            cursor.physical = handle.position;
            cursor.input_origin = handle.position;
            cursor.read_eof = false;
            cursor.record = None;
        }
        self.flush_pending_write(number);
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
        let handle = self.handle(number)?;
        if !(1..=511).contains(&number) { return Err(FileError::BadNumber); }
        if !matches!(handle.mode, OpenMode::Binary | OpenMode::Random(_)) { return Err(FileError::BadMode); }
        if !handle.can_write { return Err(FileError::PathFileAccess); }
        if matches!(handle.mode, OpenMode::Random(length) if bytes.len() > length) { return Err(FileError::BadRecordLength); }
        self.check_record_access(number, handle, bytes.len())?;
        let key = channel_key(self.project_scope, number);
        let handle = self.open.get_mut(&key).expect("checked handle");
        let cursor = self.cursors.get_mut(&key).ok_or(FileError::BadNumber)?;
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        let start = cursor.physical;
        let end = start.checked_add(bytes.len()).ok_or(FileError::BadRecordNumber)?;
        let logical = handle.position.checked_add(bytes.len()).ok_or(FileError::BadRecordNumber)?;
        if file.1.len() < end { file.1.resize(end, 0); }
        file.1[start..end].copy_from_slice(bytes);
        cursor.physical = end;
        cursor.record = None;
        handle.position = logical;
        let path = handle.path.clone();
        self.mark_handle_written(number,&path,bytes.len());
        Ok(())
    }

    /// Complete one random record. Short reads advance its logical record
    /// counter while the underlying stream advances only by bytes available.
    pub fn finish_record(&mut self, number: i64) {
        let key = channel_key(self.project_scope, number);
        if let (Some(handle), Some(cursor)) = (self.open.get_mut(&key), self.cursors.get_mut(&key)) {
            if let OpenMode::Random(length) = handle.mode {
                if let Some(record) = cursor.record.take() {
                    let available = self.contents.get(&handle.path).map_or(0, |file| file.1.len()).saturating_sub(record.start);
                    cursor.physical = record.start.saturating_add(available.min(record.span));
                    handle.position = record.logical_start.saturating_add(record.length);
                } else {
                    let length = length.max(1);
                    handle.position = handle.position.div_ceil(length).saturating_mul(length);
                    cursor.physical = handle.position;
                }
            }
        }
    }

    pub(crate) fn abort_get_record(&mut self, number: i64) {
        let key = channel_key(self.project_scope, number);
        if let (Some(handle), Some(cursor)) = (self.open.get_mut(&key), self.cursors.get_mut(&key)) {
            if let Some(record) = cursor.record.take() { handle.position = record.logical_start; }
        }
    }

    pub fn get(&mut self, number: i64, count: usize) -> Result<Vec<u8>, FileError> {
        self.get_bytes(number, count)
    }

    fn get_bytes(&mut self, number: i64, count: usize) -> Result<Vec<u8>, FileError> {
        let handle = self.handle(number)?;
        if !(1..=511).contains(&number) { return Err(FileError::BadNumber); }
        if !matches!(handle.mode, OpenMode::Binary | OpenMode::Random(_)) { return Err(FileError::BadMode); }
        if !handle.can_read { return Err(FileError::PathFileAccess); }
        let random = match handle.mode { OpenMode::Random(length) => Some(length.max(1)), _ => None };
        self.check_record_access(number, handle, random.unwrap_or(count))?;
        let key = channel_key(self.project_scope, number);
        let handle = self.open.get_mut(&key).expect("checked handle");
        let cursor = self.cursors.get_mut(&key).ok_or(FileError::BadNumber)?;
        let file = self.contents.get(&handle.path).map(|file| file.1.as_slice()).unwrap_or(&[]);
        let mut read = vec![0u8; count];
        if let Some(length) = random {
            let record = cursor.record.get_or_insert_with(|| RecordRead {
                start: cursor.physical, logical_start: handle.position, offset: 0, length, span: length,
            });
            cursor.read_eof = file.len().saturating_sub(record.start) < record.span;
            let end = record.offset.checked_add(count).ok_or(FileError::BadRecordLength)?;
            if end > record.length { return Err(FileError::BadRecordLength); }
            record.span = record.span.max(end);
            cursor.read_eof = file.len().saturating_sub(record.start) < record.span;
            for (index, slot) in read.iter_mut().enumerate() {
                if let Some(byte) = record.start.checked_add(record.offset).and_then(|at| at.checked_add(index)).and_then(|at| file.get(at)) { *slot = *byte; }
            }
            record.offset = end;
            handle.position = record.logical_start.saturating_add(end);
        } else {
            let available = file.len().saturating_sub(cursor.physical).min(count);
            if available != 0 { read[..available].copy_from_slice(&file[cursor.physical..cursor.physical + available]); }
            cursor.physical = cursor.physical.saturating_add(available);
            handle.position = handle.position.checked_add(count).ok_or(FileError::BadRecordNumber)?;
            if count != 0 { cursor.read_eof = available < count; }
        }
        Ok(read)
    }

    fn reading(&mut self, number: i64) -> Result<(&[u8], &mut usize), FileError> {
        let handle = self.open.get_mut(&channel_key(self.project_scope, number)).ok_or(FileError::BadNumber)?;
        if handle.mode != OpenMode::Input {
            return Err(FileError::BadNumber);
        }
        if !handle.can_read {return Err(FileError::PathFileAccess);}
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

    pub(crate) fn input_bytes(&mut self, number: i64, count: usize, byte_result: bool) -> Result<Vec<u8>, FileError> {
        if !(1..=511).contains(&number) { return Err(FileError::BadNumber); }
        let handle = self.handle(number)?;
        if !matches!(handle.mode, OpenMode::Input | OpenMode::Binary) { return Err(FileError::BadMode); }
        if !handle.can_read { return Err(FileError::PathFileAccess); }
        self.check_record_access(number, handle, count)?;
        let key = channel_key(self.project_scope, number);
        let handle = self.open.get_mut(&key).expect("checked handle");
        let cursor = self.cursors.get_mut(&key).ok_or(FileError::BadNumber)?;
        let file = self.contents.get(&handle.path).ok_or(FileError::BadNumber)?.1.as_slice();
        let mut result = Vec::new();
        if handle.mode == OpenMode::Binary {
            let available = file.len().saturating_sub(cursor.physical).min(count);
            let length = if byte_result { count } else { available };
            result.try_reserve_exact(length).map_err(|_| FileError::OutOfMemory)?;
            result.resize(length, 0);
            if available != 0 { result[..available].copy_from_slice(&file[cursor.physical..cursor.physical + available]); }
            cursor.physical = cursor.physical.saturating_add(available);
            handle.position = handle.position.saturating_add(available);
            if count != 0 { cursor.read_eof = available < count; }
            return Ok(result);
        }
        result.try_reserve_exact(count.min(file.len())).map_err(|_| FileError::OutOfMemory)?;
        if !cursor.read_eof { cursor.physical = handle.position; }
        while result.len() < count {
            if cursor.read_eof { return Err(FileError::PastEnd); }
            let Some(&byte) = file.get(cursor.physical) else {
                handle.position = cursor.physical;
                return Err(FileError::PastEnd);
            };
            let at = cursor.physical;
            cursor.physical += 1;
            if byte == 26 {
                handle.position = cursor.text_end_position(at, file.len());
                cursor.read_eof = true;
                return Err(FileError::PastEnd);
            }
            handle.position = cursor.physical;
            if byte != 0 { result.push(byte); }
        }
        Ok(result)
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
        self.contents.get(&self.key(path)).map(|file| file.1.len()).ok_or(FileError::NotFound)
    }

    fn names_matching(&self, pattern: &str) -> Vec<String> {
        let wanted = self.pattern_key(pattern);
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
        self.dir_with_attributes(pattern,0)
    }

    pub fn dir_with_attributes(&mut self, pattern: Option<&str>, attributes: i64) -> String {
        self.dir_checked(pattern,attributes).unwrap_or_default()
    }

    /// Dir distinguishes its terminal empty answer from a later invalid call.
    pub fn dir_checked(&mut self, pattern: Option<&str>, attributes: i64) -> Result<String,FileError> {
        if attributes<0 || attributes & !63!=0 {return Err(FileError::InvalidArgument);}
        if let Some(pattern) = pattern {
            if attributes==8 {
                let root=self.volume_root(pattern);
                let label=self.volume_labels.get(&root).cloned().unwrap_or_default();
                self.listings.insert(self.project_scope, Vec::new());
                return Ok(label);
            }
            let pattern=if pattern.is_empty() || pattern.ends_with(['\\','/']) {format!("{pattern}*")} else {pattern.to_string()};
            let wanted=self.pattern_key(&pattern);
            let folder=wanted.rfind('\\').map_or("",|at|&wanted[..=at]);
            let mut entries:BTreeMap<String,(String,i64)>=self.names_matching(&pattern).into_iter().map(|name| {
                let key=format!("{folder}{}",name.to_lowercase());
                let held=self.attributes.get(&key).copied().unwrap_or(32);
                (name.to_lowercase(),(name,held))
            }).collect();
            if attributes & 16!=0 {
                let parent=folder.trim_end_matches('\\');
                if self.folder_exists(&format!("{parent}\\")) {
                    let leaf=wanted.rsplit('\\').next().unwrap_or_default();
                    for path in self.child_folders(parent) {
                        let name=file_name(&path);
                        if matches_pattern(name,leaf) {entries.insert(name.to_lowercase(),(name.to_string(),self.attributes.get(&path).copied().unwrap_or(16)));}
                    }
                    if parent!=self.volume_root(&format!("{parent}\\")) {
                        for name in [".",".."] {
                            if matches_pattern(name,leaf) {
                                let key=if name=="." {parent.to_string()} else {self.key(&format!("{parent}\\.."))};
                                entries.insert(name.to_string(),(name.to_string(),self.attributes.get(&key).copied().unwrap_or(16)));
                            }
                        }
                    }
                }
            }
            let mut listing:Vec<_>=entries.into_values().filter(|(_,held)| {
                (held & 2==0 || attributes & 2!=0) && (held & 4==0 || attributes & 4!=0)
            }).map(|(name,_)|name).collect();
            listing.reverse();self.listings.insert(self.project_scope, listing);
        }
        let Some(listing)=self.listings.get_mut(&self.project_scope) else {return Err(FileError::InvalidArgument);};
        if let Some(name)=listing.pop() {Ok(name)} else {self.listings.remove(&self.project_scope);Ok(String::new())}
    }

    /// `Kill`: every file answering the pattern.
    pub fn kill(&mut self, pattern: &str) -> Result<(), FileError> {
        self.delete_matching_files(pattern,false,false)
    }

    pub fn delete_files(&mut self, pattern: &str, force: bool) -> Result<(), FileError> {
        self.delete_matching_files(pattern,force,true)
    }

    fn delete_matching_files(&mut self, pattern: &str, force: bool, fso: bool) -> Result<(), FileError> {
        let wanted = self.key(pattern);
        let folder = wanted.rfind('\\').map_or(String::new(), |at| wanted[..=at].to_string());
        let names = self.names_matching(pattern);
        if names.is_empty() {
            return Err(FileError::NotFound);
        }
        for name in names {
            let full = self.key(&format!("{folder}{name}"));
            if self.open.values().any(|handle| handle.path == full) {
                return Err(if fso {FileError::PermissionDenied} else {FileError::AlreadyOpen});
            }
            if !force && self.is_read_only_key(&full) {return Err(if fso {FileError::PermissionDenied} else {FileError::PathFileAccess});}
            self.contents.remove(&full);
            self.attributes.remove(&full);
            self.modified.remove(&full);
            self.parent_changed(&full, true);
        }
        Ok(())
    }

    /// FSO folder copying retains completed entries when a later copy fails.
    pub fn copy_folder(&mut self, source: &str, destination: &str, overwrite: bool) -> Result<(), FileError> {
        if source.is_empty() || destination.is_empty() { return Err(FileError::InvalidArgument); }
        if source.ends_with(['\\', '/']) { return Err(FileError::PathNotFound); }
        if destination.contains(['*', '?']) { return Err(FileError::InvalidArgument); }
        let source = self.key(source);
        let wildcard = source.contains(['*', '?']);
        let sources = if wildcard {
            let (parent, pattern) = source.rsplit_once('\\').ok_or(FileError::PathNotFound)?;
            if parent.contains(['*', '?']) { return Err(FileError::PathNotFound); }
            self.child_folders(parent).into_iter().filter(|path| matches_pattern(file_name(path), pattern)).collect::<Vec<_>>()
        } else if self.folder_exists(&source) && !self.contents.contains_key(&source) { vec![source] }
        else { return Err(FileError::PathNotFound); };
        if sources.is_empty() { return Err(FileError::PathNotFound); }
        let container = wildcard || destination.ends_with(['\\', '/']);
        if container && !self.folder_exists(destination) { return Err(FileError::PathNotFound); }
        for source in sources {
            let target = if container { format!("{}\\{}", destination.trim_end_matches(['\\', '/']), file_name(&source)) }
                else { destination.into() };
            let target_key = self.key(&target);
            if self.contents.contains_key(&target_key) || !overwrite && self.folder_exists(&target_key) { return Err(FileError::AlreadyExists); }
            if source == target_key { continue; }
            if target_key.starts_with(&format!("{source}\\")) { return Err(FileError::PermissionDenied); }
            let result = self.copy_directory_tree(&source, &target, overwrite);
            self.publish_parent_entry(self.key(destination).trim_end_matches('\\'));
            result?;
        }
        Ok(())
    }

    fn copy_directory_tree(&mut self, source: &str, destination: &str, overwrite: bool) -> Result<(), FileError> {
        let target = self.key(destination);
        if self.contents.contains_key(&target) { return Err(FileError::AlreadyExists); }
        if !self.folder_exists(&target) {
            let attributes = self.attributes(source)?;
            self.make_folder(destination)?;
            self.attributes.insert(target.clone(), attributes);
            // CopyFolder leaves nested directory entries cached until queried.
        } else if !overwrite { return Err(FileError::AlreadyExists); }
        let mut entries: BTreeMap<String, (String, bool)> = self.child_folders(source).into_iter().map(|path| (path.clone(), (path, true))).collect();
        for (key, (path, _)) in &self.contents {
            if key.rsplit_once('\\').is_some_and(|(parent, _)| parent == source) { entries.insert(key.clone(), (path.clone(), false)); }
        }
        for (_, (path, folder)) in entries {
            let target = format!("{}\\{}", destination.trim_end_matches(['\\', '/']), file_name(&path));
            if folder { self.copy_directory_tree(&path, &target, overwrite)?; }
            else { self.copy_one(&path, &target, true, overwrite, false)?; }
        }
        Ok(())
    }

    /// `FileCopy`.
    pub fn copy(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if source.is_empty() {return Err(FileError::PathFileAccess);}
        self.copy_with_options(source,destination,false)
    }

    pub fn copy_fso(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        self.copy_fso_with_overwrite(source,destination,true)
    }

    pub fn copy_fso_with_overwrite(&mut self, source: &str, destination: &str, overwrite: bool) -> Result<(), FileError> {
        if source.is_empty() || destination.is_empty() {return Err(FileError::InvalidArgument);}
        let sources=self.fso_source_files(source)?;
        let folder_target=destination.ends_with(['\\','/']) || source.contains(['*','?']) && self.folder_exists(destination);
        if folder_target && !self.folder_exists(destination) {return Err(FileError::PathNotFound);}
        for source in sources {
            let destination=if folder_target {format!("{}\\{}",destination.trim_end_matches(['\\','/']),file_name(&source))} else {destination.to_string()};
            self.copy_one(&source,&destination,true,overwrite,true)?;
        }
        Ok(())
    }

    fn copy_with_options(&mut self, source: &str, destination: &str, fso: bool) -> Result<(), FileError> {
        self.copy_one(source,destination,fso,true,fso)
    }

    fn copy_one(&mut self, source: &str, destination: &str, fso: bool, overwrite: bool, publish_parent: bool) -> Result<(), FileError> {
        let bytes = self.contents.get(&self.key(source)).ok_or(FileError::NotFound)?.1.clone();
        let key=self.key(destination);
        let source_key=self.key(source);
        if fso {
            if self.sharing_conflict(&source_key,true,false,true,true) || self.sharing_conflict(&key,false,true,true,true) {return Err(FileError::PermissionDenied);}
        } else if self.open.values().any(|held|held.path==source_key || held.path==key) {
            return Err(FileError::AlreadyOpen);
        }
        if fso && !overwrite && self.contents.contains_key(&key) {return Err(FileError::AlreadyExists);}
        if !fso && destination.ends_with(['\\','/']) {return Err(FileError::PathNotFound);}
        if source_key==key {return if fso {Ok(())}else{Err(FileError::PermissionDenied)};}
        if let Some((parent,_))=key.rsplit_once('\\') {
            if !self.folder_exists(&format!("{parent}\\")) {return Err(FileError::PathNotFound);}
        }
        if self.folder_exists(&format!("{key}\\")) {return Err(if fso {FileError::PermissionDenied}else{FileError::PathFileAccess});}
        if self.is_read_only_key(&key) {return Err(if fso {FileError::PermissionDenied}else{FileError::PathFileAccess});}
        let attributes=if fso {self.attributes(source)? | 32}else{32};
        if !self.contents.contains_key(&key) { self.parent_changed(&key, publish_parent); }
        self.attributes.insert(key.clone(),attributes);
        let modified=self.modified.get(&source_key).copied().unwrap_or_else(|| self.current_time());
        self.modified.insert(key.clone(),modified);
        self.contents.insert(key, (self.absolute_path(destination.trim()), bytes));
        Ok(())
    }

    /// `Name ... As ...`.
    pub fn rename(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if source.is_empty() {return Err(FileError::PathFileAccess);}
        self.rename_with_options(source,destination,false)
    }

    pub fn rename_fso(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if source.is_empty() || destination.is_empty() {return Err(FileError::InvalidArgument);}
        let sources=self.fso_source_files(source)?;
        let folder_target=destination.ends_with(['\\','/']) || source.contains(['*','?']) && self.folder_exists(destination);
        if folder_target && !self.folder_exists(destination) {return Err(FileError::PathNotFound);}
        for source in sources {
            let destination=if folder_target {format!("{}\\{}",destination.trim_end_matches(['\\','/']),file_name(&source))}else{destination.to_string()};
            self.rename_with_options(&source,&destination,true)?;
        }
        Ok(())
    }

    fn fso_source_files(&self, source: &str) -> Result<Vec<String>,FileError> {
        if !source.contains(['*','?']) {
            let key=self.key(source);
            return self.contents.get(&key).map(|(path,_)|vec![path.clone()]).ok_or(FileError::NotFound);
        }
        let pattern=self.absolute_path(source);let parent=pattern.rsplit_once('\\').map_or("",|(parent,_)|parent);
        let names=self.names_matching(source);
        if names.is_empty() {return Err(FileError::NotFound);}
        Ok(names.into_iter().map(|name|format!("{parent}\\{name}")).collect())
    }

    fn rename_with_options(&mut self, source: &str, destination: &str, fso: bool) -> Result<(), FileError> {
        let key=self.key(destination);let old=self.key(source);
        if old!=key && self.contents.contains_key(&key) {
            return Err(FileError::AlreadyExists);
        }
        let attributes=self.attributes(source)? | 32;
        if self.open.values().any(|held|held.path==old) {return Err(if fso {FileError::PermissionDenied} else {FileError::AlreadyOpen});}
        if !self.contents.contains_key(&old) {return Err(FileError::NotFound);}
        if old==key {
            let destination=self.absolute_path(destination.trim());
            if let Some((path,_))=self.contents.get_mut(&old) {
                if *path!=destination {*path=destination;self.attributes.insert(old,attributes);}
            }
            return Ok(());
        }
        if let Some((parent,_))=key.rsplit_once('\\') {
            if !self.folder_exists(&format!("{parent}\\")) {return Err(if fso {FileError::PathNotFound}else{FileError::NotFound});}
        }
        if self.folder_exists(&format!("{key}\\")) {return Err(FileError::PermissionDenied);}
        let (_, bytes) = self.contents.remove(&old).ok_or(FileError::NotFound)?;
        let same_parent = old.rsplit_once('\\').map(|(parent, _)| parent) == key.rsplit_once('\\').map(|(parent, _)| parent);
        self.parent_changed(&old, fso || same_parent);
        self.parent_changed(&key, true);
        let modified=self.modified.remove(&old).unwrap_or_else(|| self.current_time());
        self.modified.insert(key.clone(),modified);
        self.attributes.remove(&old);
        self.attributes.insert(key.clone(),attributes);
        self.contents.insert(key, (self.absolute_path(destination.trim()), bytes));
        Ok(())
    }
}

#[cfg(test)]
mod tests {

    #[test]
    fn shared_writers_keep_the_last_write_time_in_either_close_order() {
        for second_writes in [false, true] {
            for first_closed in [1, 2] {
                let mut files = Files::default();
                files.mount_folder(r"C:\work").unwrap();
                files.replace_bytes(r"C:\work\data.bin", b"OLD".to_vec()).unwrap();
                files.set_modified_time(r"C:\work\data.bin", 43890.50283564815).unwrap();
                files.open(r"C:\work\data.bin", OpenMode::Binary, 1).unwrap();
                files.open(r"C:\work\data.bin", OpenMode::Binary, 2).unwrap();
                files.set_current_time(46343.5);
                files.put(1, b"AAA").unwrap();
                files.set_current_time(46343.625);
                if second_writes { files.put(2, b"BBB").unwrap(); }
                files.set_current_time(46343.75);
                files.close(&[first_closed]);
                let written = if second_writes { 46343.625 } else { 46343.5 };
                let published = if !second_writes && first_closed == 2 { 43890.50283564815 } else { written };
                assert_eq!(files.date_modified(r"C:\work\data.bin"), Ok(published));
                files.set_current_time(46344.5);
                files.close(&[3 - first_closed]);
                assert_eq!(files.date_modified(r"C:\work\data.bin"), Ok(written), "second_writes={second_writes}, first_closed={first_closed}");
            }
        }
    }

    #[test]
    fn sequential_buffer_boundaries_keep_the_last_flush_time() {
        let cases: &[(&[usize], bool)] = &[
            (&[3], true), (&[513], true), (&[1024], false), (&[4096], false),
            (&[256, 256], true), (&[256, 256, 256], true),
            (&[512, 512], false), (&[512, 512, 512], true),
            (&[513, 513], true), (&[1024, 1024], false),
        ];
        for mode in [OpenMode::Output, OpenMode::Append] {
            for &(chunks, buffered) in cases {
                let mut files = Files::default();
                files.mount_folder(r"C:\work").unwrap();
                files.replace_bytes(r"C:\work\data.bin", b"OLD".to_vec()).unwrap();
                files.set_current_time(46343.5);
                files.open(r"C:\work\data.bin", mode, 1).unwrap();
                let mut written = 46343.5;
                for (index, &count) in chunks.iter().enumerate() {
                    written = 46343.5 + index as f64 / 8.0;
                    files.set_current_time(written);
                    let bytes = vec![b'A'; count];
                    let text = String::from_utf8(bytes.clone()).unwrap();
                    files.write_sequential(1, &text, &bytes).unwrap();
                }
                files.set_current_time(46344.5);
                files.close(&[1]);
                let expected = if buffered { 46344.5 } else { written };
                assert_eq!(files.date_modified(r"C:\work\data.bin"), Ok(expected), "{mode:?} {chunks:?}");
            }
        }
    }

    #[test]
    fn sequential_queries_preserve_the_flush_time_at_later_close() {
        for query in 0..4 {
            let mut files = Files::default();
            files.mount_folder(r"C:\work").unwrap();
            files.set_current_time(46343.5);
            files.open(r"C:\work\data.bin", OpenMode::Output, 1).unwrap();
            files.write_sequential(1, "ABC", b"ABC").unwrap();
            files.set_current_time(46343.625);
            match query {
                1 => { assert_eq!(files.query_length_of_open(1), Ok(3)); }
                2 => { assert_eq!(files.query_next_position(1), Ok(4)); }
                3 => { files.move_to(1, 1).unwrap(); }
                _ => {}
            }
            files.set_current_time(46343.75);
            files.close(&[1]);
            let expected = if query == 0 { 46343.75 } else { 46343.625 };
            assert_eq!(files.date_modified(r"C:\work\data.bin"), Ok(expected), "query={query}");
        }
    }

    #[test]
    fn refused_output_queries_do_not_flush_pending_writes() {
        let mut files = Files::default();
        files.mount_folder(r"C:\work").unwrap();
        files.set_current_time(46343.5);
        files.open(r"C:\work\data.bin", OpenMode::Output, 1).unwrap();
        files.write_sequential(1, "ABC", b"ABC").unwrap();
        files.set_current_time(46343.625);
        assert_eq!(files.query_length_of_open(2), Err(FileError::BadNumber));
        assert_eq!(files.query_next_position(2), Err(FileError::BadNumber));
        assert_eq!(files.move_to(1, 0), Err(FileError::BadNumber));
        files.set_current_time(46343.75);
        files.close(&[1]);
        assert_eq!(files.date_modified(r"C:\work\data.bin"), Ok(46343.75));
    }

    #[test]
    fn flushed_writes_keep_the_write_time_when_closed_later() {
        use std::cell::Cell;
        use std::rc::Rc;
        for (mode, stream) in [(OpenMode::Binary, false), (OpenMode::Random(16), false),
            (OpenMode::Output, true), (OpenMode::Append, true)] {
            let clock = Rc::new(Cell::new(46343.5));
            let observed = Rc::clone(&clock);
            let mut files = Files::default();
            files.mount_folder(r"C:\work").unwrap();
            files.replace_bytes(r"C:\work\data.bin", b"OLD".to_vec()).unwrap();
            files.set_modified_time(r"C:\work\data.bin", 43890.50283564815).unwrap();
            files.set_clock(Some(crate::clock::Clock::new(move || observed.get())));
            let number = if stream { files.open_stream(r"C:\work\data.bin", mode).unwrap() }
                else { files.open(r"C:\work\data.bin", mode, 1).unwrap(); 1 };
            if stream { files.write(number, "ABC").unwrap(); }
            else { files.put(number, b"ABC").unwrap(); }
            clock.set(46343.75);
            files.close(&[number]);
            assert_eq!(files.date_modified(r"C:\work\data.bin"), Ok(46343.5), "{mode:?} stream={stream}");
        }
    }

    #[test]
    fn copy_folder_matches_native_target_and_overwrite_rules() {
        let mut files = Files::default();
        files.mount_folder(r"C:\work\source\branch").unwrap();
        files.mount_folder(r"C:\work\existing").unwrap();
        files.replace_bytes(r"C:\work\source\data.bin", b"ABC".to_vec()).unwrap();
        files.replace_bytes(r"C:\work\source\branch\child.bin", b"DEF".to_vec()).unwrap();
        files.set_modified_time(r"C:\work\source\data.bin", 43890.50283564815).unwrap();
        files.set_modified_time(r"C:\work", 43890.50283564815).unwrap();
        files.set_current_time(47485.89925925926);
        files.copy_folder(r"C:\work\source", r"C:\work\target", true).unwrap();
        assert_eq!(files.bytes(r"C:\work\target\branch\child.bin"), Ok(b"DEF".as_slice()));
        assert_eq!(files.date_modified(r"C:\work\target\data.bin"), Ok(43890.50283564815));
        assert_eq!(files.date_modified(r"C:\work"), Ok(47485.89925925926));
        assert_eq!(files.copy_folder(r"C:\work\source", r"C:\work\existing", false), Err(FileError::AlreadyExists));
        assert!(!files.exists(r"C:\work\existing\data.bin"));
        files.copy_folder(r"C:\work\source", r"C:\work\source", true).unwrap();
        assert_eq!(files.copy_folder(r"C:\work\source\", r"C:\work\another", true), Err(FileError::PathNotFound));
        files.copy_folder(r"C:\work\s*", r"C:\work\existing\", false).unwrap();
        assert_eq!(files.bytes(r"C:\work\existing\source\data.bin"), Ok(b"ABC".as_slice()));
    }

    #[test]
    fn copy_folder_partial_failure_follows_name_order_and_retains_completed_entries() {
        let mut files = Files::default();
        files.mount_folder(r"C:\work\source\middle").unwrap();
        files.mount_folder(r"C:\work\target\middle").unwrap();
        for (path, bytes) in [(r"C:\work\source\a.bin", b"A".as_slice()),
            (r"C:\work\source\z.bin", b"Z".as_slice()),
            (r"C:\work\source\middle\locked.bin", b"NEW".as_slice()),
            (r"C:\work\target\middle\locked.bin", b"OLD".as_slice())] {
            files.replace_bytes(path, bytes.to_vec()).unwrap();
        }
        files.set_attributes(r"C:\work\target\middle\locked.bin", 1).unwrap();
        files.set_modified_time(r"C:\work\target", 43890.50283564815).unwrap();
        files.set_current_time(47485.89925925926);
        assert_eq!(files.copy_folder(r"c:\WORK\source", r"C:\work\target", true), Err(FileError::PermissionDenied));
        assert_eq!(files.bytes(r"C:\work\target\a.bin"), Ok(b"A".as_slice()));
        assert!(!files.exists(r"C:\work\target\z.bin"));
        assert_eq!(files.bytes(r"C:\work\target\middle\locked.bin"), Ok(b"OLD".as_slice()));
        assert_eq!(files.date_modified(r"C:\work\target"), Ok(43890.50283564815));
    }

    #[test]
    fn copy_folder_failure_after_creation_publishes_the_destination_parent_date() {
        let mut files = Files::default();
        files.mount_folder(r"C:\work\source").unwrap();
        files.replace_bytes(r"C:\work\source\a.bin", b"A".to_vec()).unwrap();
        files.replace_bytes(r"C:\work\source\blocked.bin", b"B".to_vec()).unwrap();
        files.set_modified_time(r"C:\work", 43890.50283564815).unwrap();
        files.set_current_time(47485.89925925926);
        files.open_with_options(r"C:\work\source\blocked.bin", OpenMode::Binary, 1,
            Some(crate::ast::FileAccess::Read), Some(crate::ast::FileLock::ReadWrite)).unwrap();
        assert_eq!(files.copy_folder(r"C:\work\source", r"C:\work\new", true), Err(FileError::PermissionDenied));
        assert_eq!(files.date_modified(r"C:\work"), Ok(47485.89925925926));
        assert_eq!(files.bytes(r"C:\work\new\a.bin"), Ok(b"A".as_slice()));
        assert!(!files.exists(r"C:\work\new\blocked.bin"));
    }

    fn directory_date_fixture() -> Files {
        let mut files = Files::default();
        files.mount_folder(r"C:\work\branch").unwrap();
        files.replace_bytes(r"C:\work\data.bin", b"ABC".to_vec()).unwrap();
        files.replace_bytes(r"C:\work\branch\child.bin", b"DEF".to_vec()).unwrap();
        for path in [r"C:\work", r"C:\work\branch", r"C:\work\data.bin", r"C:\work\branch\child.bin"] {
            files.set_modified_time(path, 43890.50283564815).unwrap();
        }
        files.set_current_time(47485.89925925926);
        files.open(r"C:\work\new.bin", OpenMode::Output, 1).unwrap();
        files.close(&[1]);
        files
    }

    #[test]
    fn directory_dates_retain_native_entry_values_until_a_refreshing_query() {
        let old = 43890.50283564815;
        let current = 47485.89925925926;
        for query in 0..9 {
            let mut files = directory_date_fixture();
            for _ in 0..3 { assert_eq!(files.date_modified(r"C:\work"), Ok(old)); }
            assert_eq!(files.modified_facts().iter().find(|(path, _)| path == r"c:\work").unwrap().1, current);
            files.set_current_time(current + 10.0);
            match query {
                0 => { files.query_attributes(r"C:\work").unwrap(); }
                1 => { files.date_modified(r"C:\work\data.bin").unwrap(); }
                2 => { assert_eq!(files.date_modified(r"C:\work\missing.bin"), Err(FileError::NotFound)); }
                3 => { files.query_length(r"C:\work\data.bin").unwrap(); }
                4 => { files.change_directory(r"C:\work").unwrap(); }
                5 => { assert_eq!(files.fresh_modified_time(r"C:\work"), Ok(current)); }
                6 => { files.query_attributes(r"C:\work\data.bin").unwrap(); }
                7 => { files.dir(Some(r"C:\work\*")); }
                8 => { files.date_modified(r"C:\work\branch\child.bin").unwrap(); }
                _ => unreachable!(),
            }
            assert_eq!(files.date_modified(r"C:\work"), Ok(if query < 6 { current } else { old }), "query {query}");
        }
    }

    #[test]
    fn directory_entry_snapshots_keep_absolute_drive_roots() {
        let mut files = Files::default();
        files.mount_folder(r"C:\").unwrap();
        files.set_modified_time(r"C:\", 43890.50283564815).unwrap();
        files.set_current_time(47485.89925925926);
        files.open(r"C:\new.bin", OpenMode::Output, 1).unwrap();
        files.close(&[1]);
        let entries = files.directory_entry_facts();
        assert_eq!(entries, vec![(r"c:\".into(), 43890.50283564815)]);
        let mut next = Files::default();
        next.restore_closed_snapshot(&files.closed_snapshot()).unwrap();
        next.restore_modified_facts(&files.modified_facts()).unwrap();
        next.restore_directory_entry_facts(&entries).unwrap();
        assert_eq!(next.fresh_modified_time(r"C:\"), Ok(47485.89925925926));
    }

    #[test]
    fn directory_entry_snapshots_preserve_cached_and_physical_dates_atomically() {
        let files = directory_date_fixture();
        let closed = files.closed_snapshot();
        let modified = files.modified_facts();
        let entries = files.directory_entry_facts();
        assert!(!entries.is_empty());
        let mut restored = Files::default();
        restored.restore_closed_snapshot(&closed).unwrap();
        restored.restore_modified_facts(&modified).unwrap();
        restored.restore_directory_entry_facts(&entries).unwrap();
        assert_eq!(restored.date_modified(r"C:\work"), Ok(43890.50283564815));
        for facts in [vec![(r"C:\work\data.bin".into(), 1.0)], vec![(r"C:\absent".into(), 1.0)],
            vec![(r"C:\work".into(), f64::NAN)], vec![(r"C:\work".into(), -657435.0)],
            vec![(r"C:\work".into(), 2958466.0)],
            vec![(r"C:\work".into(), 1.0), ("c:/WORK".into(), 2.0)]] {
            assert_eq!(restored.restore_directory_entry_facts(&facts), Err(FileError::InvalidArgument));
            assert_eq!(restored.directory_entry_facts(), entries);
            assert_eq!(restored.modified_facts(), modified);
        }
        restored.query_attributes(r"C:\work").unwrap();
        assert_eq!(restored.date_modified(r"C:\work"), Ok(47485.89925925926));
    }


    #[test]
    fn modification_dates_preserve_whole_seconds_across_the_date_range() {
        let path = r"C:\work\data.bin";
        let mut files = Files::default();
        files.mount_folder(r"C:\work").unwrap();
        files.replace_bytes(path, b"ABC".to_vec()).unwrap();
        // Include both supported limits and the independently measured
        // 1601, 3000, 5000 and 9998 dates, with all seconds of each day.
        for days in [-657434_i64, -109204, -1, 0, 43890, 401770, 1132255, 2957737, 2958465] {
            for second in 0..86_400 {
                let direction = if days < 0 { -1.0 } else { 1.0 };
                let expected = days as f64 + direction * f64::from(second) / 86_400.0;
                files.set_modified_time(path, expected).unwrap();
                assert_eq!(files.date_modified(path), Ok(expected), "day {days}, second {second}");
            }
            for second in [0, 191, 43200, 86399] {
                let direction = if days < 0 { -1.0 } else { 1.0 };
                let expected = days as f64 + direction * f64::from(second) / 86_400.0;
                for subsecond in [0.001, 0.5, 0.999] {
                    files.set_modified_time(path, expected + direction * subsecond / 86_400.0).unwrap();
                    assert_eq!(files.date_modified(path), Ok(expected));
                }
            }
        }
    }

    #[test]
    fn modification_times_follow_native_open_write_close_copy_and_name() {
        let old=43890.50283564815;
        let current=47485.89925925926;
        for mode in [OpenMode::Binary,OpenMode::Output,OpenMode::Append,OpenMode::Random(1)] {
            let mut files=Files::default();
            files.set_current_time(current);
            files.mount_folder(r"C:\work").unwrap();
            files.replace_bytes(r"C:\work\data.bin",b"ABC".to_vec()).unwrap();
            files.set_modified_time(r"C:\work\data.bin",old).unwrap();
            files.open(r"C:\work\data.bin",mode,1).unwrap();
            let during=if mode==OpenMode::Output {current}else{old};
            assert_eq!(files.date_modified(r"C:\work\data.bin"),Ok(during));
            match mode {
                OpenMode::Output|OpenMode::Append=>files.write_sequential(1,"Z\r\n",b"Z\r\n").unwrap(),
                _=>files.put(1,b"Z").unwrap(),
            }
            assert_eq!(files.date_modified(r"C:\work\data.bin"),Ok(during));
            files.close(&[1]);
            assert_eq!(files.date_modified(r"C:\work\data.bin"),Ok(current));
            files.set_modified_time(r"C:\work\data.bin",old).unwrap();
            files.copy(r"C:\work\data.bin",r"C:\work\copy.bin").unwrap();
            assert_eq!(files.date_modified(r"C:\work\copy.bin"),Ok(old));
            files.rename(r"C:\work\copy.bin",r"C:\work\renamed.bin").unwrap();
            assert_eq!(files.date_modified(r"C:\work\renamed.bin"),Ok(old));
            assert_eq!(files.date_modified(r"C:\work\copy.bin"),Err(FileError::NotFound));
            files.make_folder(r"C:\work\created").unwrap();
            assert_eq!(files.date_modified(r"C:\work\created"),Ok(current));
        }
    }

    #[test]
    fn modification_snapshots_validate_atomically_and_survive_project_scoped_close() {
        let old=43890.50283564815;
        let current=47485.89925925926;
        let mut files=Files::default();
        files.set_current_time(current);
        files.mount_folder(r"C:\work").unwrap();
        files.replace_bytes(r"C:\work\data.bin",b"ABC".to_vec()).unwrap();
        files.set_modified_time(r"C:\work\data.bin",old).unwrap();
        let namespace=files.closed_snapshot();let times=files.modified_facts();
        let mut restored=Files::default();
        restored.restore_closed_snapshot(&namespace).unwrap();
        restored.restore_modified_facts(&times).unwrap();
        assert_eq!(restored.modified_facts(),times);
        for facts in [vec![(r"C:\missing".into(),old)],vec![(r"C:\work\data.bin".into(),f64::NAN)],
            vec![(r"C:\work\data.bin".into(),old),(r"c:/WORK/DATA.bin".into(),current)],
            vec![(r"C:\work\data.bin".into(),-657435.0)],vec![(r"C:\work\data.bin".into(),2958466.0)]] {
            assert_eq!(restored.restore_modified_facts(&facts),Err(FileError::InvalidArgument));
            assert_eq!(restored.modified_facts(),times);
        }
        files.select_project(1);files.open(r"C:\work\data.bin",OpenMode::Binary,1).unwrap();
        files.put(1,b"Z").unwrap();
        files.select_project(2);files.open(r"C:\work\data.bin",OpenMode::Input,1).unwrap();
        files.close_vba(&[]);
        assert_eq!(files.date_modified(r"C:\work\data.bin"),Ok(old));
        files.select_project(1);files.close_vba(&[]);
        assert_eq!(files.date_modified(r"C:\work\data.bin"),Ok(current));
    }

    use super::*;

    #[test]
    fn empty_file_paths_preserve_api_errors_and_existing_bytes() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        assert_eq!(files.copy("","C:\\work\\copy"),Err(FileError::PathFileAccess));
        assert_eq!(files.rename("","C:\\work\\moved"),Err(FileError::PathFileAccess));
        assert_eq!(files.copy_fso("","C:\\work\\copy"),Err(FileError::InvalidArgument));
        assert_eq!(files.copy_fso("C:\\work\\held",""),Err(FileError::InvalidArgument));
        assert_eq!(files.rename_fso("","C:\\work\\moved"),Err(FileError::InvalidArgument));
        assert_eq!(files.rename_fso("C:\\work\\held",""),Err(FileError::InvalidArgument));
        assert_eq!(files.bytes("C:\\work\\held").unwrap(),b"seed");assert!(!files.exists("C:\\work\\copy"));assert!(!files.exists("C:\\work\\moved"));
    }

    #[test]
    fn fso_copy_preserves_source_flags_expands_names_and_honors_overwrite() {
        let mut files=Files::default();files.mount_folder("C:\\work\\out").unwrap();
        files.replace_bytes("C:\\work\\UPPER.TXT",b"seed".to_vec()).unwrap();files.set_attributes("C:\\work\\UPPER.TXT",1).unwrap();
        files.copy_fso("C:\\work\\*.txt","C:\\work\\out\\").unwrap();
        assert_eq!(files.bytes("C:\\work\\out\\UPPER.TXT").unwrap(),b"seed");assert_eq!(files.attributes("C:\\work\\out\\UPPER.TXT"),Ok(33));
        files.set_attributes("C:\\work\\UPPER.TXT",32).unwrap();files.set_attributes("C:\\work\\out\\UPPER.TXT",32).unwrap();
        assert_eq!(files.copy_fso_with_overwrite("C:\\work\\UPPER.TXT","C:\\work\\out\\",false),Err(FileError::AlreadyExists));
        files.set_attributes("C:\\work\\out\\UPPER.TXT",1).unwrap();
        assert_eq!(files.copy_fso("C:\\work\\UPPER.TXT","C:\\work\\out\\"),Err(FileError::PermissionDenied));
        files.set_attributes("C:\\work\\UPPER.TXT",1).unwrap();
        files.copy("C:\\work\\UPPER.TXT","C:\\work\\basic").unwrap();assert_eq!(files.attributes("C:\\work\\basic"),Ok(32));
        assert_eq!(files.copy_fso("C:\\work\\UPPER.TXT","C:\\missing\\new"),Err(FileError::PathNotFound));
        assert!(!files.exists("C:\\missing\\new"));assert_eq!(files.bytes("C:\\work\\UPPER.TXT").unwrap(),b"seed");
    }

    #[test]
    fn move_self_is_noop_and_invalid_parents_do_not_consume_the_source() {
        let mut files=Files::default();files.mount_folder("C:\\work\\out").unwrap();files.replace_bytes("C:\\work\\held.txt",b"seed".to_vec()).unwrap();
        files.rename("C:\\work\\held.txt","C:\\work\\held.txt").unwrap();files.rename_fso("C:\\work\\held.txt","C:\\work\\held.txt").unwrap();
        assert_eq!(files.rename("C:\\work\\held.txt","C:\\work\\missing\\new"),Err(FileError::NotFound));
        assert_eq!(files.rename_fso("C:\\work\\held.txt","C:\\work\\missing\\new"),Err(FileError::PathNotFound));
        assert_eq!(files.bytes("C:\\work\\held.txt").unwrap(),b"seed");
        files.rename_fso("C:\\work\\*.txt","C:\\work\\out\\").unwrap();assert!(!files.exists("C:\\work\\held.txt"));assert_eq!(files.bytes("C:\\work\\out\\held.txt").unwrap(),b"seed");
    }

    #[test]
    fn open_file_copy_and_move_keep_basic_and_fso_errors_and_bytes() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();
        files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();files.open("C:\\work\\held",OpenMode::Binary,1).unwrap();
        assert_eq!(files.copy("C:\\work\\held","C:\\work\\copy"),Err(FileError::AlreadyOpen));
        assert!(!files.exists("C:\\work\\copy"));
        assert_eq!(files.rename("C:\\work\\held","C:\\work\\moved"),Err(FileError::AlreadyOpen));
        assert_eq!(files.rename_fso("C:\\work\\held","C:\\work\\moved"),Err(FileError::PermissionDenied));
        assert!(!files.exists("C:\\work\\moved"));assert_eq!(files.bytes("C:\\work\\held").unwrap(),b"seed");
        files.copy_fso("C:\\work\\held","C:\\work\\copy").unwrap();assert_eq!(files.bytes("C:\\work\\copy").unwrap(),b"seed");
        files.close(&[]);files.replace_bytes("C:\\work\\copy",b"old".to_vec()).unwrap();files.open("C:\\work\\copy",OpenMode::Binary,1).unwrap();
        assert_eq!(files.copy("C:\\work\\held","C:\\work\\copy"),Err(FileError::AlreadyOpen));
        assert_eq!(files.bytes("C:\\work\\copy").unwrap(),b"old");
        files.copy_fso("C:\\work\\held","C:\\work\\copy").unwrap();
        assert_eq!(files.bytes("C:\\work\\copy").unwrap(),b"seed");
    }

    #[test]
    fn fso_writers_share_readers_but_refuse_other_writers_before_truncation() {
        use crate::ast::{FileAccess,FileLock};
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        files.open_with_options("C:\\work\\held",OpenMode::Binary,1,Some(FileAccess::Write),Some(FileLock::Shared)).unwrap();
        assert_eq!(files.open_stream("C:\\work\\held",OpenMode::Output),Err(FileError::PermissionDenied));
        assert_eq!(files.bytes("C:\\work\\held").unwrap(),b"seed");
        let reader=files.open_stream("C:\\work\\held",OpenMode::Input).unwrap();files.close(&[1,reader]);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,1,Some(FileAccess::Read),Some(FileLock::Shared)).unwrap();
        let writer=files.open_stream("C:\\work\\held",OpenMode::Append).unwrap();
        assert_eq!(files.open("C:\\work\\held",OpenMode::Binary,2),Err(FileError::PermissionDenied));
        let reader=files.open_stream("C:\\work\\held",OpenMode::Input).unwrap();files.close(&[writer,reader]);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Write),Some(FileLock::Shared)).unwrap();
    }

    #[test]
    fn sequential_and_share_conflicts_are_refused_before_replacement() {
        use crate::ast::{FileAccess,FileLock};
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        files.open("C:\\work\\held",OpenMode::Input,1).unwrap();
        assert_eq!(files.open("C:\\work\\held",OpenMode::Output,2),Err(FileError::AlreadyOpen));
        assert_eq!(files.bytes("C:\\work\\held"),Ok(b"seed".as_slice()));files.close(&[]);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,1,Some(FileAccess::Read),Some(FileLock::Write)).unwrap();
        assert_eq!(files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Write),Some(FileLock::Shared)),Err(FileError::PermissionDenied));
        assert_eq!(files.free_number(false),2);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Read),Some(FileLock::Shared)).unwrap();files.close(&[2]);
        assert_eq!(files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Read),Some(FileLock::ReadWrite)),Err(FileError::PermissionDenied));
    }

    #[test]
    fn record_lock_ownership_exact_unlock_and_close_release_are_preserved() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        files.open("C:\\work\\held",OpenMode::Binary,1).unwrap();files.open("C:\\work\\held",OpenMode::Binary,2).unwrap();
        files.lock_records(1,Some(1),Some(2),false).unwrap();
        assert_eq!(files.lock_records(1,Some(1),Some(2),false),Err(FileError::PermissionDenied));
        assert_eq!(files.lock_records(1,Some(1),Some(1),true),Err(FileError::PermissionDenied));
        assert_eq!(files.get(2,1),Err(FileError::PermissionDenied));assert_eq!(files.next_position(2),Ok(1));
        assert_eq!(files.put(2,b"X"),Err(FileError::PermissionDenied));assert_eq!(files.bytes("C:\\work\\held"),Ok(b"seed".as_slice()));
        assert_eq!(files.get(1,1),Ok(b"s".to_vec()));
        files.close(&[1]);assert_eq!(files.get(2,1),Ok(b"s".to_vec()));
        assert_eq!(files.lock_records(2,Some(0),None,false),Err(FileError::BadRecordNumber));
        assert_eq!(files.lock_records(2,Some(4),Some(2),false),Err(FileError::BadRecordNumber));
        files.lock_records(2,None,None,false).unwrap();let state=files.closed_snapshot();files.restore_closed_snapshot(&state).unwrap();
        files.open("C:\\work\\held",OpenMode::Binary,1).unwrap();files.lock_records(1,None,None,false).unwrap();
    }

    #[test]
    fn dir_cursor_distinguishes_terminal_empty_and_invalid_continuation() {
        let mut files=Files::default();files.mount_folder("C:\\work\\empty").unwrap();files.change_directory("C:\\work").unwrap();
        files.replace_bytes("held.txt",vec![]).unwrap();
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some("held.txt"),0),Ok("held.txt".into()));
        assert_eq!(files.dir_checked(None,0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some("missing*"),0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some(""),0),Ok("held.txt".into()));
        assert_eq!(files.dir_checked(Some("empty"),16),Ok("empty".into()));
        assert_eq!(files.dir_checked(Some("*"),16),Ok(".".into()));
        assert_eq!(files.dir_checked(None,0),Ok("..".into()));
        assert_eq!(files.dir_checked(None,0),Ok("empty".into()));
    }

    #[test]
    fn volume_queries_use_host_facts_and_have_their_own_terminal_empty() {
        let mut files=Files::default();
        files.restore_volume_facts(&[("C:\\".into(),String::new()),("D:\\".into(),"Fixture volume".into())]).unwrap();
        assert_eq!(files.dir_checked(Some("D:\\held.txt"),8),Ok("Fixture volume".into()));
        assert_eq!(files.dir_checked(None,0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some("C:\\held.txt"),8),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        let state=files.volume_facts();
        for invalid in [vec![("relative".into(),"x".into())],vec![("C:\\folder".into(),"x".into())],vec![("C:\\".into(),"x".into()),("c:/".into(),"y".into())]] {
            assert_eq!(files.restore_volume_facts(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(files.volume_facts(),state);
        }
    }

    #[test]
    fn file_access_refusals_preserve_bytes_and_handle_positions() {
        use crate::ast::FileAccess;
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();
        files.replace_bytes("C:\\work\\held.txt",b"seed".to_vec()).unwrap();
        files.set_attributes("C:\\work\\held.txt",1).unwrap();
        for mode in [OpenMode::Binary,OpenMode::Random(4)] {
            files.open("C:\\work\\held.txt",mode,1).unwrap();
            assert_eq!(files.put(1,b"X"),Err(FileError::PathFileAccess));
            assert_eq!(files.next_position(1),Ok(1));
            assert_eq!(files.bytes("C:\\work\\held.txt"),Ok(b"seed".as_slice()));
            assert_eq!(files.get(1,2),Ok(b"se".to_vec()));files.close(&[1]);
        }
        for access in [FileAccess::Write,FileAccess::ReadWrite] {
            assert_eq!(files.open_with_access("C:\\work\\held.txt",OpenMode::Binary,1,Some(access)),Err(FileError::PathFileAccess));
            assert_eq!(files.free_number(false),1);
        }
        files.set_attributes("C:\\work\\held.txt",0).unwrap();
        files.open_with_access("C:\\work\\held.txt",OpenMode::Binary,1,Some(FileAccess::Read)).unwrap();
        assert_eq!(files.put(1,b"X"),Err(FileError::PathFileAccess));
        assert_eq!(files.bytes("C:\\work\\held.txt"),Ok(b"seed".as_slice()));
    }

    #[test]
    fn dir_flags_filter_hidden_and_system_without_losing_cursor_state() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();
        for (name,attributes) in [("a.txt",1),("b.txt",2),("c.txt",4),("d.txt",0)] {
            let path=format!("C:\\work\\{name}");files.replace_bytes(&path,vec![]).unwrap();files.set_attributes(&path,attributes).unwrap();
        }
        assert_eq!(files.dir(Some("C:\\work\\*.txt")),"a.txt");
        assert_eq!(files.dir(None),"d.txt");assert_eq!(files.dir(None),"");
        assert_eq!(files.dir_with_attributes(Some("C:\\work\\b.txt"),2),"b.txt");
        assert_eq!(files.dir_with_attributes(Some("C:\\work\\c.txt"),4),"c.txt");
        assert_eq!(files.dir_with_attributes(Some("C:\\work\\d.txt"),1),"d.txt");
    }

    #[test]
    fn attributes_survive_transport_and_refused_updates_are_atomic() {
        let mut files=Files::default();files.mount_folder("C:\\work\\empty").unwrap();
        files.replace_bytes("C:\\work\\held.txt",b"held".to_vec()).unwrap();
        files.set_attributes("C:\\work\\held.txt",1).unwrap();
        files.set_fso_attributes("C:\\work\\empty",-1).unwrap();
        let state=files.closed_snapshot();
        let mut restored=Files::default();restored.restore_closed_snapshot(&state).unwrap();
        assert_eq!(restored.attributes("c:/WORK/held.txt"),Ok(1));
        assert_eq!(restored.attributes("c:/WORK/empty"),Ok(55));
        assert_eq!(restored.replace_bytes("C:\\work\\held.txt",vec![]),Err(FileError::PermissionDenied));
        assert_eq!(restored.kill("C:\\work\\held.txt"),Err(FileError::PathFileAccess));
        assert_eq!(restored.closed_snapshot(),state);
        for attributes in [vec![("C:\\work\\missing".into(),0)],vec![("C:\\work\\held.txt".into(),16)],
            vec![("C:\\work\\empty".into(),0)],vec![("C:\\work\\held.txt".into(),64)],
            vec![("C:\\work\\held.txt".into(),-1)],vec![("C:\\work\\held.txt".into(),1),("c:/WORK/HELD.txt".into(),0)]] {
            let invalid=FileSystemSnapshot {attributes,..state.clone()};
            assert_eq!(restored.restore_closed_snapshot(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(restored.closed_snapshot(),state);
        }
        // Legacy snapshots without overrides replace metadata with defaults.
        restored.restore_closed_snapshot(&FileSystemSnapshot {attributes:vec![],..state.clone()}).unwrap();
        assert_eq!(restored.attributes("C:\\work\\held.txt"),Ok(32));
        assert_eq!(restored.attributes("C:\\work\\empty"),Ok(16));
    }

    #[test]
    fn readonly_directory_failure_retains_native_partial_deletion() {
        let mut files=Files::default();files.mount_folder("C:\\work\\tree\\child").unwrap();
        files.replace_bytes("C:\\work\\tree\\child\\nested.txt",vec![1]).unwrap();
        files.replace_bytes("C:\\work\\tree\\leaf.txt",vec![2]).unwrap();
        files.set_attributes("C:\\work\\tree\\child",1).unwrap();
        assert_eq!(files.delete_folder_with_force("C:\\work\\tree",false),Err(FileError::PermissionDenied));
        assert!(!files.exists("C:\\work\\tree\\child\\nested.txt"));
        assert!(files.exists("C:\\work\\tree\\leaf.txt"));
        assert_eq!(files.attributes("C:\\work\\tree\\child"),Ok(17));
        files.delete_folder_with_force("C:\\work\\tree",true).unwrap();
        assert!(!files.folder_exists("C:\\work\\tree"));
        assert!(!files.closed_snapshot().attributes.iter().any(|(path,_)|path.to_lowercase().contains("tree")));
    }

    #[test]
    fn directory_operations_preserve_refused_writes_and_native_partial_deletion_order() {
        let mut files=Files::default();
        for folder in ["C:\\work\\empty","C:\\work\\exists","C:\\work\\nonempty\\child","C:\\work\\nonempty\\empty-child"] {
            files.mount_folder(folder).unwrap();
        }
        for (path,bytes) in [("C:\\work\\held.txt",b"held".to_vec()),("C:\\work\\nonempty\\leaf.txt",b"leaf".to_vec()),("C:\\work\\nonempty\\child\\nested.txt",b"nested".to_vec())] {
            files.replace_bytes(path,bytes).unwrap();
        }
        let before=files.closed_snapshot();
        assert_eq!(files.make_folder("C:\\work\\missing\\child"),Err(FileError::PathNotFound));
        assert_eq!(files.make_folder("C:\\work\\held.txt"),Err(FileError::AlreadyExists));
        assert_eq!(files.remove_folder("C:\\work\\nonempty",false),Err(FileError::PathFileAccess));
        assert_eq!(files.closed_snapshot(),before);
        // Native case53: sorted enumeration reaches a locked ordinary file
        // after the earlier empty directories, before the later subtree.
        files.open("C:\\work\\held.txt",OpenMode::Input,1).unwrap();
        assert_eq!(files.remove_folder("C:\\work",true),Err(FileError::PermissionDenied));
        assert!(!files.folder_exists("C:\\work\\empty"));
        assert!(!files.folder_exists("C:\\work\\exists"));
        assert_eq!(files.bytes("C:\\work\\nonempty\\child\\nested.txt").unwrap(),b"nested");
        files.close(&[1]);
        // Native case44: earlier child directories are removed, while the
        // locked leaf and its containing directory remain available.
        files.open("C:\\work\\nonempty\\leaf.txt",OpenMode::Input,1).unwrap();
        assert_eq!(files.remove_folder("C:\\work\\nonempty",true),Err(FileError::PermissionDenied));
        assert!(!files.folder_exists("C:\\work\\nonempty\\child"));
        assert!(!files.folder_exists("C:\\work\\nonempty\\empty-child"));
        assert_eq!(files.bytes("C:\\work\\nonempty\\leaf.txt").unwrap(),b"leaf");
        assert!(files.folder_exists("C:\\work\\nonempty"));
    }

    #[test]
    fn closed_namespace_snapshots_preserve_empty_folders_and_bytes_atomically() {
        let mut files=Files::default();
        files.mount_folder("C:\\work\\empty").unwrap();
        files.change_directory("C:\\work").unwrap();
        files.replace_bytes("binary.dat",vec![0,255,128,1]).unwrap();
        files.open("binary.dat",OpenMode::Input,1).unwrap();
        let state=files.closed_snapshot();
        for invalid in [
            FileSystemSnapshot {folders:vec!["relative".into()],files:state.files.clone(),attributes:vec![]},
            FileSystemSnapshot {folders:vec!["C:\\work".into(),"c:/WORK".into()],files:vec![],attributes:vec![]},
            FileSystemSnapshot {folders:vec![],files:vec![("C:\\same".into(),vec![1]),("c:/SAME".into(),vec![2])],attributes:vec![]},
            FileSystemSnapshot {folders:vec![],files:vec![("C:\\bad|name".into(),vec![])],attributes:vec![]},
        ] {
            assert_eq!(files.restore_closed_snapshot(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(files.closed_snapshot(),state);
            assert!(files.handle(1).is_ok());
        }
        let mut restored=Files::default();restored.restore_closed_snapshot(&state).unwrap();
        assert!(restored.folder_exists("c:/WORK/EMPTY"));
        assert_eq!(restored.bytes("C:\\work\\binary.dat").unwrap(),&[0,255,128,1]);
        files.restore_closed_snapshot(&state).unwrap();
        assert_eq!(files.free_number(false),1);
        assert_eq!(files.current_directory(None).unwrap(),"C:\\work");
    }

    #[test]
    fn directory_snapshots_restore_drive_memory_and_reject_invalid_input_atomically() {
        let mut files=Files::default();
        files.mount_folder("C:\\first").unwrap();files.mount_folder("D:\\other").unwrap();
        files.change_directory("C:\\first").unwrap();files.change_directory("D:\\other").unwrap();files.change_drive("D").unwrap();
        let state=files.directory_state();
        let mut restored=Files::default();restored.restore_directory_state(&state).unwrap();
        assert_eq!(restored.current_directory(None).unwrap(),"D:\\other");
        assert_eq!(restored.absolute_path("C:held.txt"),"C:\\first\\held.txt");
        for invalid in [
            DirectoryState {current_drive:'Z',directories:state.directories.clone()},
            DirectoryState {current_drive:'?',directories:state.directories.clone()},
            DirectoryState {current_drive:'C',directories:vec!["relative".into()]},
            DirectoryState {current_drive:'C',directories:vec!["C:\\first".into(),"c:\\other".into()]},
            DirectoryState {current_drive:'C',directories:vec!["C:\\bad|path".into()]},
        ] {
            assert_eq!(restored.restore_directory_state(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(restored.directory_state(),state);
        }
    }

    #[test]
    fn file_exports_and_folder_listings_keep_their_resolved_paths() {
        let mut files=Files::default();
        files.mount_folder("C:\\work\\sub").unwrap();
        let before=files.clone();
        files.change_directory("C:\\work").unwrap();
        files.open("held.txt",OpenMode::Output,1).unwrap();
        files.write(1,"ROOT").unwrap();files.close(&[1]);
        files.change_directory("sub").unwrap();
        files.open("held.txt",OpenMode::Output,1).unwrap();
        files.write(1,"CHILD").unwrap();files.close(&[1]);
        assert_eq!(files.files_in("C:\\work"),vec!["C:\\work\\held.txt"]);
        assert_eq!(files.changed_files(&before),vec![("C:\\work\\held.txt".into(),b"ROOT".to_vec()),("C:\\work\\sub\\held.txt".into(),b"CHILD".to_vec())]);
        let previous=files.clone();
        files.rename("held.txt","renamed.txt").unwrap();
        assert_eq!(files.deleted_files(&previous),vec!["C:\\work\\sub\\held.txt"]);
        assert_eq!(files.changed_files(&previous),vec![("C:\\work\\sub\\renamed.txt".into(),b"CHILD".to_vec())]);
        files.copy("renamed.txt","..\\copied.txt").unwrap();
        assert_eq!(files.files_in("C:\\work"),vec!["C:\\work\\copied.txt","C:\\work\\held.txt"]);
    }

    #[test]
    fn current_directories_resolve_files_without_rebinding_open_handles() {
        let mut files=Files::default();
        for folder in ["C:\\first", "C:\\second", "D:\\other"] { files.mount_folder(folder).unwrap(); }
        files.change_directory("C:\\first").unwrap();
        files.open("value.txt",OpenMode::Output,1).unwrap();
        files.write(1,"first").unwrap();
        files.change_directory("C:\\second").unwrap();
        files.open(".\\value.txt",OpenMode::Output,2).unwrap();
        files.write(2,"second").unwrap();
        files.write(1,"!").unwrap();
        files.close(&[1,2]);
        assert_eq!(files.bytes("..\\first\\value.txt").unwrap(),b"first!");
        assert_eq!(files.bytes("C:value.txt").unwrap(),b"second");
        assert_eq!(files.bytes("\\second\\value.txt").unwrap(),b"second");
        files.change_directory("D:\\other").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"C:\\second");
        files.change_drive("D:").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"D:\\other");
        files.change_directory("C:\\first").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"D:\\other");
        files.change_drive("C:\\ignored").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"C:\\first");
        assert_eq!(files.clone().current_directory(Some("D")).unwrap(),"D:\\other");
        assert_eq!(files.change_directory("missing"),Err(FileError::PathNotFound));
        assert_eq!(files.current_directory(None).unwrap(),"C:\\first");
        assert_eq!(files.change_drive("?"),Err(FileError::InvalidArgument));
        assert_eq!(files.current_directory(Some("Z")),Err(FileError::DeviceUnavailable));
    }

    #[test]
    fn host_binary_output_is_immediately_readable_and_replacement_is_atomic() {
        let mut files = Files::default();
        assert_eq!(files.replace_bytes("C:\\missing\\copy.xlsm", vec![1]), Err(FileError::PathNotFound));
        assert!(!files.exists("C:\\missing\\copy.xlsm"));
        files.make_folder("C:\\copies").unwrap();
        files.replace_bytes("C:\\copies\\copy.xlsm", vec![b'P', b'K', 0, 255]).unwrap();
        assert_eq!(files.length("c:/COPIES/COPY.xlsm"), Ok(4));
        assert_eq!(files.dir(Some("C:\\copies\\*.xlsm")), "copy.xlsm");
        files.open("c:/copies/copy.xlsm", OpenMode::Binary, 1).unwrap();
        assert_eq!(files.get(1, 4).unwrap(), vec![b'P', b'K', 0, 255]);
        assert_eq!(files.replace_bytes("C:\\copies\\copy.xlsm", vec![2]), Err(FileError::AlreadyOpen));
        assert_eq!(files.bytes("C:\\copies\\copy.xlsm").unwrap(), &[b'P', b'K', 0, 255]);
        files.move_to(1, 3).unwrap();
        files.put(1, &[7]).unwrap();
        assert_eq!(files.bytes("C:\\copies\\copy.xlsm").unwrap(), &[b'P', b'K', 7, 255]);
        files.close(&[1]);
        files.replace_bytes("C:\\copies\\copy.xlsm", vec![2]).unwrap();
        assert_eq!(files.bytes("C:\\copies\\copy.xlsm").unwrap(), &[2]);
    }

    #[test]
    fn sequential_location_matches_native_block_boundaries() {
        let controls = [(0, 1), (0, 2), (127, 128), (127, 129),
            (128, 129), (128, 130), (513, 514), (513, 515)];
        for (mode, initial, expected) in [
            (OpenMode::Input, 1, [1, 1, 1, 1, 1, 2, 5, 5]),
            (OpenMode::Output, 0, [0, 0, 0, 1, 1, 1, 4, 4]),
            (OpenMode::Append, 0, [0, 0, 0, 1, 1, 1, 4, 4]),
        ] {
            for ((size, position), expected) in controls.into_iter().zip(expected) {
                let mut files = Files::default();
                files.mount_folder("C:\\work").unwrap();
                files.replace_bytes("C:\\work\\location.bin", vec![65; size]).unwrap();
                files.open("C:\\work\\location.bin", mode, 1).unwrap();
                assert_eq!(files.location(1), Ok(initial), "initial {mode:?}, size {size}");
                files.move_to(1, position).unwrap();
                assert_eq!(files.location(1), Ok(expected), "{mode:?}, size {size}, Seek {position}");
            }
        }
    }

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
